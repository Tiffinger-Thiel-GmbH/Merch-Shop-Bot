import { BackgroundImage } from "@microsoft/teams.cards";
import {
  ProductVariantCategoriesDTO,
  ProductVariantListDTO,
} from "../api/merchApi";

type AdaptiveCardElement = Record<string, unknown>;

export function buildVariantsCard(
  productId: string,
  productName: string,
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
      text: `${productName}`,
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
        title: "Zurück zu Produkten",
        data: { action: "backToProducts" },
        style: "destructive",
      },
      {
        type: "Action.Submit",
        title: "Bestellen",
        style: "positive",
        data: {
          action: "submitProductSelection",
          productId,
          variantInputIds: categoryInputIds,
        },
      },
    ],
  };
}
