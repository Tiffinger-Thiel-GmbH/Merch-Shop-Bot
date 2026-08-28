// cardBuilder.ts
import type {
  CreateOrderItemDTO,
  ProductDTO,
  ProductListDTO,
  ProductVariantCategoriesDTO,
  ProductVariantListDTO,
} from "./api/merchApi";

type AdaptiveCardElement = Record<string, unknown>;

const COMPACT_THRESHOLD = 6;
const PAGE_SIZE = 10;

// ---------- Produktkarte ----------

export function buildProductsCard(data: ProductListDTO, page = 0): object {
  const { items, totalCount } = data;

  if (totalCount === 0) {
    return buildEmptyStateCard("Keine Produkte gefunden.");
  }

  const pageItems = items.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);
  const hasMore = totalCount > (page + 1) * PAGE_SIZE;
  const columnsPerRow = totalCount <= COMPACT_THRESHOLD ? 2 : 3;

  const body: AdaptiveCardElement[] = [
    {
      type: "TextBlock",
      text: `${totalCount} ${totalCount === 1 ? "Produkt" : "Produkte"} gefunden`,
      weight: "Bolder",
      size: "Medium",
      wrap: true,
    },
    buildProductGrid(pageItems, columnsPerRow),
  ];

  const actions: AdaptiveCardElement[] = [];
  if (hasMore) {
    actions.push({
      type: "Action.Submit",
      title: "Weitere anzeigen",
      data: { action: "nextPage", page: page + 1 },
    });
  }

  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body,
    actions,
  };
}

function buildProductGrid(
  items: ProductDTO[],
  columnsPerRow: number,
): AdaptiveCardElement {
  const rows: AdaptiveCardElement[] = [];
  for (let i = 0; i < items.length; i += columnsPerRow) {
    const rowItems = items.slice(i, i + columnsPerRow);
    rows.push({
      type: "ColumnSet",
      columns: rowItems.map((p) => buildProductColumn(p)),
    });
  }
  return { type: "Container", items: rows };
}

function buildProductColumn(p: ProductDTO): AdaptiveCardElement {
  return {
    type: "Column",
    width: "stretch",
    selectAction: {
      type: "Action.Submit",
      data: { action: "selectProduct", productId: p.id },
    },
    items: [
      {
        type: "Image",
        url: p.imageUrl,
        size: "Large",
        style: "default",
        height: "stretch",
        horizontalAlignment: "Center",
      },
      {
        type: "TextBlock",
        text: p.name,
        wrap: true,
        weight: "Bolder",
        size: "Small",
        horizontalAlignment: "Center",
      },
    ],
  };
}

// ---------- Variantenkarte ----------

export function buildVariantsCard(
  productId: string,
  categoriesData: ProductVariantCategoriesDTO,
  _selectedCategory: string,
  variantsData: ProductVariantListDTO,
  quantity: number,
): object {
  const { categories } = categoriesData;
  const { items } = variantsData;
  const categoryInputIds = categories.map((category, index) => ({
    category,
    inputId: `variant_${index}`,
  }));

  const body: AdaptiveCardElement[] = [
    {
      type: "TextBlock",
      text: "Varianten wählen",
      weight: "Bolder",
      size: "Medium",
      wrap: true,
    },
  ];

  if (items.length === 0) {
    body.push({
      type: "TextBlock",
      text: "Keine Varianten für dieses Produkt verfügbar.",
      isSubtle: true,
      wrap: true,
    });
  } else if (categories.length === 0) {
    body.push({
      type: "TextBlock",
      text: "Keine Varianten-Kategorien für dieses Produkt verfügbar.",
      isSubtle: true,
      wrap: true,
    });
  } else {
    for (const { category, inputId } of categoryInputIds) {
      const categoryVariants = items.filter(
        (item) => item.category === category,
      );
      if (categoryVariants.length === 0) {
        continue;
      }

      body.push({
        type: "Input.ChoiceSet",
        id: inputId,
        label: category,
        style: "compact",
        choices: categoryVariants.map((variant) => ({
          title: variant.name,
          value: variant.productVariantId,
        })),
      });
    }
    body.push({
      type: "Input.Number",
      id: "quantity",
      label: "Menge",
      min: 1,
      max: 10,
      value: quantity,
    });
  }

  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body,
    actions: [
      {
        type: "Action.Submit",
        title: "Auswahl senden",
        data: {
          action: "submitProductSelection",
          productId,
          variantInputIds: categoryInputIds,
        },
      },
      {
        type: "Action.Submit",
        title: "Zurück zu Produkten",
        data: { action: "backToProducts" },
      },
    ],
  };
}

// ---------- Gemeinsame Helper ----------

function buildEmptyStateCard(message: string): object {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [{ type: "TextBlock", text: message, isSubtle: true, wrap: true }],
  };
}
