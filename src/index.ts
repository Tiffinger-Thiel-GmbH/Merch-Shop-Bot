import { ManagedIdentityCredential } from "@azure/identity";
import { cardAttachment, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { IAdaptiveCard } from "@microsoft/teams.cards";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import welcomeCardJson from "./cards/welcomeCard.json";
import {
  createProductDetailCard,
  createShopCard,
  createShopErrorCard,
} from "./shop-cards";
import { getShopProduct, getShopProducts } from "./shop-api";
import {
  createCard,
  createConversationMembersCard,
  createDummyCards,
  createLinkUnfurlCard,
  createMessageDetailsCard,
} from "./card";

const welcomeCard = welcomeCardJson as IAdaptiveCard;

type ShopActionData = {
  action?: string;
  id?: string;
};

function getShopActionData(value: unknown): ShopActionData | undefined {
  if (!value || typeof value !== "object") {
    return undefined;
  }

  const actionData =
    "action" in value &&
    value.action &&
    typeof value.action === "object" &&
    "data" in value.action
      ? value.action.data
      : value;

  if (!actionData || typeof actionData !== "object") {
    return undefined;
  }

  return actionData as ShopActionData;
}

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
  const text = activity.text?.trim().toLowerCase();
  if (text === "/shop") {
    try {
      const products = await getShopProducts();
      const card = createShopCard(products);

      await send({
        type: "message",
        attachments: [cardAttachment("adaptive", card)],
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      const card = createShopErrorCard(message);

      await send({
        type: "message",
        attachments: [cardAttachment("adaptive", card)],
      });
    }

    return;
  }

  const actionData = getShopActionData(activity.value);
  if (actionData?.action === "openMessage") {
    try {
      if (!actionData.id) {
        throw new Error("Missing product id");
      }

      const product = await getShopProduct(actionData.id);
      const card = createProductDetailCard(product);

      await send({
        type: "message",
        attachments: [cardAttachment("adaptive", card)],
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      const card = createShopErrorCard(message);

      await send({
        type: "message",
        attachments: [cardAttachment("adaptive", card)],
      });
    }
  }
});

app.on("card.action.openMessage", async ({ activity }) => {
  const actionData = getShopActionData(activity.value);

  try {
    if (!actionData?.id) {
      throw new Error("Missing product id");
    }

    const product = await getShopProduct(actionData.id);

    return {
      statusCode: 200,
      type: "application/vnd.microsoft.card.adaptive",
      value: createProductDetailCard(product),
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown error";

    return {
      statusCode: 200,
      type: "application/vnd.microsoft.card.adaptive",
      value: createShopErrorCard(message),
    };
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
