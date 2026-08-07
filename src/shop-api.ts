export type ShopProductVariant = {
  id: string;
  category: string;
  name: string;
  description: string;
};

export type ShopProduct = {
  id: string;
  title: string;
  description?: string;
  imageUrl?: string;
  variants: ShopProductVariant[];
};

const mockProductIds = [
  "11111111-1111-4111-8111-111111111111",
  "22222222-2222-4222-8222-222222222222",
  "33333333-3333-4333-8333-333333333333",
  "44444444-4444-4444-8444-444444444444",
  "55555555-5555-4555-8555-555555555555",
  "66666666-6666-4666-8666-666666666666",
  "77777777-7777-4777-8777-777777777777",
  "88888888-8888-4888-8888-888888888888",
];

const mockProducts: ShopProduct[] = Array.from({ length: 8 }, (_, index) => {
  const itemNumber = index + 1;
  const productId = mockProductIds[index];

  return {
    id: productId,
    title: `Produkt ${itemNumber}`,
    description: `Detailbeschreibung fuer Produkt ${itemNumber}.`,
    imageUrl: `https://adaptivecards.io/content/cats/${itemNumber}.png`,
    variants: [
      {
        id: `${productId}-size`,
        category: "Groesse",
        name: "M",
        description: "Standardgroesse M",
      },
      {
        id: `${productId}-color`,
        category: "Farbe",
        name: "Schwarz",
        description: "Schwarze Variante",
      },
    ],
  };
});

function findMockProduct(productId: string): ShopProduct | undefined {
  const product = mockProducts.find((item) => item.id === productId);

  if (product) {
    return product;
  }

  const legacyProductIdMatch = /^product-(\d+)$/.exec(productId);
  const legacyProductIndex = legacyProductIdMatch
    ? Number.parseInt(legacyProductIdMatch[1], 10) - 1
    : -1;

  return mockProducts[legacyProductIndex];
}

function requiredString(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim().length > 0
    ? value
    : fallback;
}

function optionalString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0
    ? value
    : undefined;
}

function toRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object"
    ? (value as Record<string, unknown>)
    : {};
}

function extractProductList(payload: unknown): unknown[] {
  if (Array.isArray(payload)) {
    return payload;
  }

  const record = toRecord(payload);
  const candidates = [
    record.products,
    record.items,
    record.data,
    record.results,
  ];
  const list = candidates.find(Array.isArray);

  return list ?? [];
}

function normalizeProductVariants(value: unknown): ShopProductVariant[] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.map((variant, index) => {
    const record = toRecord(variant);

    return {
      id: requiredString(record.id, `variant-${index + 1}`),
      category: requiredString(record.category, "Variante"),
      name: requiredString(record.name, `Variante ${index + 1}`),
      description: requiredString(record.description, ""),
    };
  });
}

function normalizeProduct(value: unknown, index = 0): ShopProduct {
  const record = toRecord(value);
  const id = requiredString(record.id, `product-${index + 1}`);

  return {
    id,
    title: requiredString(record.name, `Produkt ${index + 1}`),
    description: optionalString(record.description),
    imageUrl: optionalString(
      record.imageUrl ??
        record.image ??
        record.thumbnailUrl ??
        record.thumbnail,
    ),
    variants: normalizeProductVariants(record.productVariants),
  };
}

function getApiBaseUrl(): string | undefined {
  return optionalString(process.env.SHOP_API_BASE_URL);
}

function getHeaders(): HeadersInit {
  const headers: Record<string, string> = {
    Accept: "application/json",
  };

  if (process.env.SHOP_API_TOKEN) {
    headers.Authorization = `Bearer ${process.env.SHOP_API_TOKEN}`;
  }

  if (process.env.SHOP_API_KEY) {
    headers[process.env.SHOP_API_KEY_HEADER || "x-api-key"] =
      process.env.SHOP_API_KEY;
  }

  return headers;
}

async function fetchJson(path: string): Promise<unknown> {
  const baseUrl = getApiBaseUrl();

  if (!baseUrl) {
    throw new Error("SHOP_API_BASE_URL is not configured");
  }

  const response = await fetch(new URL(path, baseUrl), {
    headers: getHeaders(),
  });

  if (!response.ok) {
    throw new Error(`Shop API request failed: ${response.status}`);
  }

  return response.json() as Promise<unknown>;
}

export async function getShopProducts(): Promise<ShopProduct[]> {
  const baseUrl = getApiBaseUrl();

  if (!baseUrl) {
    return mockProducts;
  }

  const productsPath = process.env.SHOP_API_PRODUCTS_PATH || "/products";
  const payload = await fetchJson(productsPath);

  return extractProductList(payload).map(normalizeProduct);
}

export async function getShopProduct(productId: string): Promise<ShopProduct> {
  const baseUrl = getApiBaseUrl();

  if (!baseUrl) {
    const product = findMockProduct(productId);

    if (!product) {
      throw new Error(`Unknown product id: ${productId}`);
    }

    return product;
  }

  const detailPathTemplate =
    process.env.SHOP_API_PRODUCT_DETAIL_PATH || "/products/:id";
  const detailPath = detailPathTemplate.replace(
    ":id",
    encodeURIComponent(productId),
  );
  const payload = await fetchJson(detailPath);

  return normalizeProduct(payload);
}
