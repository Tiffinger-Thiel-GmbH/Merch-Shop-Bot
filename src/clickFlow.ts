import { TurnContext, CardFactory, MessageFactory } from "botbuilder";
import {
  productsControllerFindAll,
  productVariantCategoryControllerFindCategories,
  productVariantControllerFindVariants,
} from "./api/merchApi";
import { buildProductsCard, buildVariantsCard } from "./cardBuilder";

async function sendProductsCard(context: TurnContext, page = 0) {
  const data = await productsControllerFindAll();
  const card = buildProductsCard(data, page);
  await context.sendActivity(
    MessageFactory.attachment(CardFactory.adaptiveCard(card)),
  );
}

async function sendVariantsCard(
  context: TurnContext,
  productId: string,
  quantity: number,
  category?: string,
) {
  const categoriesData =
    await productVariantCategoryControllerFindCategories(productId);
  const selectedCategory = category ?? categoriesData.categories[0] ?? "";

  const variantsData = await productVariantControllerFindVariants(
    productId,
    selectedCategory ? { category: selectedCategory } : undefined,
  );

  const card = buildVariantsCard(
    productId,
    categoriesData,
    selectedCategory,
    variantsData,
    quantity,
  );
  await context.sendActivity(
    MessageFactory.attachment(CardFactory.adaptiveCard(card)),
  );
}

export async function onMessage(context: TurnContext) {
  const value = context.activity.value as
    | {
        action?: string;
        page?: number;
        productId?: string;
        category?: string; // kommt aus dem Input.ChoiceSet "category"
        productVariantId?: string;
        quantity?: number; // kommt aus dem Input.Number "quantity"
      }
    | undefined;

  switch (value?.action) {
    case "nextPage":
      return sendProductsCard(context, value.page ?? 0);

    case "selectProduct":
      if (value.productId) return sendVariantsCard(context, value.productId, 1);
      break;

    case "filterVariants":
      if (value.productId)
        return sendVariantsCard(
          context,
          value.productId,
          value.quantity ?? 1,
          value.category,
        );
      break;

    case "backToProducts":
      return sendProductsCard(context);

    case "selectVariant":
      if (value.productVariantId) {
        const quantity = value.quantity ?? 1;
        // z.B. in den Warenkorb / Order-Flow übergeben
        // await addToCart(context, value.productVariantId, quantity);
        await context.sendActivity(
          `Variante ${value.productVariantId} (Menge: ${quantity}) ausgewählt.`,
        );
      }
      break;
  }

  // Standardfall: keine Action -> Produkte laden
  return sendProductsCard(context);
}
