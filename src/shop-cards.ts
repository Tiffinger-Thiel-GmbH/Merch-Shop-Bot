import { IAdaptiveCard } from "@microsoft/teams.cards";
import { ShopProduct } from "./shop-api";

const fallbackImageUrl = "https://adaptivecards.io/content/cats/1.png";

function createProductColumn(product: ShopProduct) {
  return {
    type: "Column",
    width: "stretch",
    items: [
      {
        type: "Image",
        id: `image-${product.id}`,
        url: product.imageUrl || fallbackImageUrl,
        altText: product.title,
        size: "Stretch",
        selectAction: {
          type: "Action.Submit",
          title: product.title,
          data: {
            action: "openMessage",
            id: product.id,
          },
        },
      },
      {
        type: "TextBlock",
        text: product.title,
        weight: "Bolder",
        wrap: true,
      },
      ...(product.variants.length
        ? [
            {
              type: "TextBlock",
              text: `${product.variants.length} Varianten`,
              spacing: "None",
              wrap: true,
            },
          ]
        : []),
    ],
  };
}

export function createShopCard(products: ShopProduct[]): IAdaptiveCard {
  const rows: unknown[] = [];

  for (let index = 0; index < products.length; index += 2) {
    rows.push({
      type: "ColumnSet",
      spacing: index === 0 ? "Small" : "Medium",
      columns: products.slice(index, index + 2).map(createProductColumn),
    });
  }

  const card = {
    type: "AdaptiveCard",
    $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
    body: [
      {
        type: "TextBlock",
        text: "Shop",
        weight: "Bolder",
        size: "Large",
        wrap: true,
      },
      ...rows,
    ],
  };

  return card as IAdaptiveCard;
}

export function createProductDetailCard(product: ShopProduct): IAdaptiveCard {
  const variantFacts = product.variants.map((variant) => ({
    title: `${variant.category}:`,
    value: `${variant.name}${
      variant.description ? ` - ${variant.description}` : ""
    }`,
  }));

  const card = {
    type: "AdaptiveCard",
    $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
    body: [
      {
        type: "TextBlock",
        text: product.title,
        weight: "Bolder",
        size: "Large",
        wrap: true,
      },
      {
        type: "Image",
        url: product.imageUrl || fallbackImageUrl,
        altText: product.title,
        size: "Stretch",
      },
      {
        type: "TextBlock",
        text: product.description || "Keine Beschreibung verfuegbar.",
        wrap: true,
      },
      ...(variantFacts.length
        ? [
            {
              type: "TextBlock",
              text: "Varianten",
              weight: "Bolder",
              wrap: true,
            },
            {
              type: "FactSet",
              facts: variantFacts,
            },
          ]
        : []),
    ],
  };

  return card as IAdaptiveCard;
}

export function createShopErrorCard(message: string): IAdaptiveCard {
  const card = {
    type: "AdaptiveCard",
    $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
    body: [
      {
        type: "TextBlock",
        text: "Shop nicht verfuegbar",
        weight: "Bolder",
        size: "Large",
        wrap: true,
      },
      {
        type: "TextBlock",
        text: message,
        wrap: true,
      },
    ],
  };

  return card as IAdaptiveCard;
}
