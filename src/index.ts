import { ManagedIdentityCredential } from "@azure/identity";
import { cardAttachment, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { IAdaptiveCard } from "@microsoft/teams.cards";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { DevtoolsPlugin } from "@microsoft/teams.dev";

import {
  createCard,
  createConversationMembersCard,
  createDummyCards,
  createLinkUnfurlCard,
  createMessageDetailsCard,
} from "./card";
import {
  orderControllerCreate,
  productsControllerFindAll,
  productsControllerFindOneById,
  productVariantCategoryControllerFindCategories,
  productVariantControllerFindVariants,
} from "./api/merchApi";
import { buildProductsCard } from "./cardBuilder/shopCardBuilder";
import { buildVariantsCard } from "./cardBuilder/variantCardBuilder";

const createTokenFactory = () => {
  return async (
    scope: string | string[],
    tenantId?: string,
  ): Promise<string> => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });

    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

// Use managed identity in cloud environment, otherwise use devtools plugin for local development
const options =
  process.env.BOT_TYPE === "UserAssignedMsi"
    ? { ...tokenCredentials }
    : { plugins: [new DevtoolsPlugin()] };

const app = new App({
  ...options,
  logger: new ConsoleLogger("testsimple", { level: "debug" }),
  skipAuth: !process.env.CLIENT_ID,
});

app.on("install.add", async ({ send }) => {
  const greeting = `
  Hi this app handles:<br>
    1. Basic message handling - echoing back what you say<br>
    2. Link unfurling - creating preview cards when you paste URLs<br>
    3. Message extension commands - handling card creation.
  `;
  await send(greeting);
});

// --- Activity handler ---

// Defining Card Actions
type CardActionData = {
  action?: string;
  page?: number;
  productId?: string;
  category?: string;
  variantInputIds?: VariantInputId[];
  [key: string]: unknown; // Inputs land here
};

type VariantInputId = {
  category: string;
  inputId: string;
};

type SendFunction = (activity: Parameters<typeof app.send>[1]) => Promise<any>;

// --- Funktionen für app.on("message") ---

// function for getting the action used by a button or click
function getCardAction(value: unknown): CardActionData | undefined {
  if (!value || typeof value !== "object") {
    return undefined;
  }

  const envelope = value as { action?: unknown };
  if (
    envelope.action &&
    typeof envelope.action === "object" &&
    "data" in envelope.action
  ) {
    return (envelope.action as { data?: CardActionData }).data;
  }

  return value as CardActionData;
}

// function used for sending the Product Card
async function sendProductsCard(send: SendFunction, page = 0) {
  const products = await productsControllerFindAll();
  const card = buildProductsCard(products, page) as IAdaptiveCard;
  await send({
    type: "message",
    attachments: [cardAttachment("adaptive", card)],
  });
}

// function used for sending the Variants Card
async function sendVariantsCard(
  send: SendFunction,
  productId: string,
  category?: string,
  quantity = 1,
) {
  // function for finding out the name of a product
  const product = await productsControllerFindOneById(productId);
  const categories =
    await productVariantCategoryControllerFindCategories(productId);
  const variants = await productVariantControllerFindVariants(
    productId,
    category ? { category } : undefined,
  );

  const card = buildVariantsCard(
    productId,
    product.name,
    categories,
    category ?? "",
    variants,
    quantity,
  ) as IAdaptiveCard;

  await send({
    type: "message",
    attachments: [cardAttachment("adaptive", card)],
  });
}

// helper: pulls the selected variant ids out of the submitted card data
function getSelectedVariantIds(data: CardActionData) {
  const inputIds = data.variantInputIds ?? [];
  const selected = inputIds
    .map(({ category, inputId }) => ({
      category,
      productVariantId: data[inputId],
    }))
    .filter(
      (
        item,
      ): item is {
        category: string;
        productVariantId: string;
      } =>
        typeof item.productVariantId === "string" &&
        item.productVariantId.length > 0,
    );

  return {
    selected,
    missing: inputIds
      .filter(
        ({ inputId }) =>
          typeof data[inputId] !== "string" || data[inputId] === "",
      )
      .map(({ category }) => category),
  };
}

// function used for submitting the product selection and creating the order
async function sendProductSelectionCard(
  send: SendFunction,
  data: CardActionData,
) {
  if (!data.productId) {
    await send(`productId fehlt.`);
    return;
  }

  const { selected, missing } = getSelectedVariantIds(data);
  if (missing.length > 0) {
    await send(`Bitte auswählen: ${missing.join(", ")}.`);
    return;
  }

  const productVariantIds = selected.map((item) => item.productVariantId);
  if (productVariantIds.length === 0) {
    await send(`Keine Varianten ausgewählt.`);
    return;
  }

  await orderControllerCreate({
    userId: process.env.ORDER_USER_ID,
    items: [
      {
        productId: data.productId,
        productVariantId: productVariantIds,
        quantity: 1,
      },
    ],
  });
}

// --- Message verwaltung ---

app.on("message", async ({ send, activity }) => {
  const data = getCardAction(activity.value);
  console.log(data);
  if (data?.action) {
    switch (data.action) {
      case "nextPage":
        return sendProductsCard(send, data.page ?? 0);

      case "selectProduct":
        if (data.productId) {
          return sendVariantsCard(send, data.productId);
        }
        break;

      case "filterVariants":
        if (data.productId) {
          return sendVariantsCard(send, data.productId, data.category);
        }
        break;

      case "submitProductSelection":
        return sendProductSelectionCard(send, data);

      case "backToProducts":
        return sendProductsCard(send);
    }
  }

  const text = activity.text?.trim().toLowerCase();
  if (text === "/shop") {
    return sendProductsCard(send);
  }
});

// :snippet-start: message-ext-query-link
app.on("message.ext.query-link", async ({ activity }) => {
  const { url } = activity.value;

  if (!url) {
    return { status: 400 };
  }

  const { card, thumbnail } = createLinkUnfurlCard(url);
  const attachment = {
    ...cardAttachment("adaptive", card), // expanded card in the compose box...
    preview: cardAttachment("thumbnail", thumbnail), //preview card in the compose box...
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
});
// :snippet-end: message-ext-query-link
// :snippet-start: message-ext-submit
app.on("message.ext.submit", async ({ activity }) => {
  const { commandId } = activity.value;
  let card: IAdaptiveCard;

  if (commandId === "createCard") {
    // activity.value.commandContext == "compose"
    card = createCard(activity.value.data);
  } else if (
    commandId === "getMessageDetails" &&
    activity.value.messagePayload
  ) {
    // activity.value.commandContext == "message"
    card = createMessageDetailsCard(activity.value.messagePayload);
  } else {
    throw new Error(`Unknown commandId: ${commandId}`);
  }

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [cardAttachment("adaptive", card)],
    },
  };
});
// :snippet-end: message-ext-submit

// :snippet-start: message-ext-open
app.on("message.ext.open", async ({ activity, api }) => {
  const conversationId = activity.conversation.id;
  const members = await api.conversations.members(conversationId).get();
  const card = createConversationMembersCard(members);

  return {
    task: {
      type: "continue",
      value: {
        title: "Conversation members",
        height: "small",
        width: "small",
        card: cardAttachment("adaptive", card),
      },
    },
  };
});
// :snippet-end: message-ext-open

// :snippet-start: message-ext-query
app.on("message.ext.query", async ({ activity }) => {
  const { commandId } = activity.value;
  const searchQuery = activity.value.parameters![0].value;

  if (commandId == "searchQuery") {
    const cards = await createDummyCards(searchQuery);
    const attachments = cards.map(({ card, thumbnail }) => {
      return {
        ...cardAttachment("adaptive", card), // expanded card in the compose box...
        preview: cardAttachment("thumbnail", thumbnail), // preview card in the compose box...
      };
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  return { status: 400 };
});
// :snippet-end: message-ext-query

(async () => {
  await app.start();
})();
