import { ManagedIdentityCredential } from "@azure/identity";
import { cardAttachment, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { IAdaptiveCard } from "@microsoft/teams.cards";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import welcomeCardJson from "./cards/welcomeCard.json";
import { buildProductsCard, buildVariantsCard } from "./cardBuilder";
import {
  createCard,
  createConversationMembersCard,
  createDummyCards,
  createLinkUnfurlCard,
  createMessageDetailsCard,
} from "./card";
import {
  orderControllerCreate,
  productVariantCategoryControllerFindCategories,
  productVariantControllerFindVariants,
  productsControllerFindAll,
} from "./api/merchApi";

const welcomeCard = welcomeCardJson as IAdaptiveCard;
const ORDER_USER_ID = "9aaca58e-4ea2-4008-bfc7-2007cd91c0f1";

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
  logger: new ConsoleLogger("Merch-Shop-Bot", { level: "debug" }),
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
  const card = welcomeCard;
  await send({
    type: "message",
    attachments: [cardAttachment("adaptive", card)],
  });
});

app.on("message", async ({ send, activity }) => {
  const data = getCardActionData(activity.value);
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

type CardActionData = {
  action?: string;
  page?: number;
  productId?: string;
  category?: string;
  variantInputIds?: VariantInputId[];
  [key: string]: unknown;
};

type VariantInputId = {
  category: string;
  inputId: string;
};

type SendFunction = (activity: Parameters<typeof app.send>[1]) => Promise<any>;

function getCardActionData(value: unknown): CardActionData | undefined {
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

function adaptiveCardResponse(card: object) {
  return {
    statusCode: 200 as const,
    type: "application/vnd.microsoft.card.adaptive" as const,
    value: card as IAdaptiveCard,
  };
}

function errorResponse(message: string, statusCode: 400 | 500 = 400) {
  return {
    statusCode,
    type: "application/vnd.microsoft.error" as const,
    value: {
      code: "BadRequest",
      message,
      innerHttpError: {
        statusCode,
        body: { message },
      },
    },
  };
}

async function buildProductsResponse(page = 0) {
  const products = await productsControllerFindAll();
  return adaptiveCardResponse(buildProductsCard(products, page));
}

async function sendProductsCard(send: SendFunction, page = 0) {
  const products = await productsControllerFindAll();
  const card = buildProductsCard(products, page) as IAdaptiveCard;
  await send({
    type: "message",
    attachments: [cardAttachment("adaptive", card)],
  });
}

async function buildVariantsResponse(productId: string, category?: string) {
  const categories =
    await productVariantCategoryControllerFindCategories(productId);
  const variants = await productVariantControllerFindVariants(
    productId,
    category ? { category } : undefined,
  );

  return adaptiveCardResponse(
    buildVariantsCard(productId, categories, category ?? "", variants),
  );
}

async function sendVariantsCard(
  send: SendFunction,
  productId: string,
  category?: string,
) {
  const response = await buildVariantsResponse(productId, category);
  await send({
    type: "message",
    attachments: [cardAttachment("adaptive", response.value)],
  });
}

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

async function buildProductSelectionResponse(data: CardActionData) {
  if (!data.productId) {
    return errorResponse("productId fehlt.");
  }

  const { selected, missing } = getSelectedVariantIds(data);
  if (missing.length > 0) {
    return errorResponse(`Bitte auswählen: ${missing.join(", ")}.`);
  }

  const productVariantIds = selected.map((item) => item.productVariantId);
  if (productVariantIds.length === 0) {
    return errorResponse("Keine Varianten ausgewählt.");
  }

  await orderControllerCreate({
    userId: ORDER_USER_ID,
    items: [
      {
        productId: data.productId,
        productVariantId: productVariantIds,
        quantity: 1,
      },
    ],
  });

  return null;
}

async function sendProductSelectionCard(
  send: SendFunction,
  data: CardActionData,
) {
  const response = await buildProductSelectionResponse(data);
  if (response) {
    await send(response.value.message);
  }
}

app.on("card.action.nextPage", async ({ activity }) => {
  const data = activity.value.action.data as CardActionData;
  return buildProductsResponse(data.page ?? 0);
});

app.on("card.action.selectProduct", async ({ activity }) => {
  const data = activity.value.action.data as CardActionData;
  if (!data.productId) {
    return errorResponse("productId fehlt.");
  }

  return buildVariantsResponse(data.productId);
});

app.on("card.action.filterVariants", async ({ activity }) => {
  const data = activity.value.action.data as CardActionData;
  if (!data.productId) {
    return errorResponse("productId fehlt.");
  }

  return buildVariantsResponse(data.productId, data.category);
});

app.on("card.action.submitProductSelection", async ({ activity }) => {
  const data = activity.value.action.data as CardActionData;
  const response = await buildProductSelectionResponse(data);
  return response ?? undefined;
});

app.on("card.action.backToProducts", async () => {
  return buildProductsResponse();
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
