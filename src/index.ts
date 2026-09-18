import { ManagedIdentityCredential } from "@azure/identity";
import orderResponseCardJson from "./cards/orderResponse.json";
import { IAdaptiveCard } from "@microsoft/teams.cards";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { buildProductsCard, buildVariantsCard } from "./cardBuilder";
import {
  orderControllerCreate,
  productsControllerFindAll,
  productVariantCategoryControllerFindCategories,
  productVariantControllerFindVariants,
  userControllerPutUser,
} from "./api/merchApi";
import { App, IBaseActivityContext } from "@microsoft/teams.apps";
import {
  ActivityLike,
  cardAttachment,
  IMessageActivity,
  IMessageActivityInput,
  MessageActivityInput,
  SentActivity,
  TokenCredentials,
} from "@microsoft/teams.api";
const orderResponseCard = orderResponseCardJson as IAdaptiveCard;

/** Save the activity reference of the message previously sent with sendWithRef as user.id -> activity.id mapping */
const PreviousMessageReferences: Record<string, string> = {};
/**
 * Replaces the previously tracked message that was sent with this function or sendWithRef.
 * @param activity new Activity to replace the previous message with
 * @param context message Context object
 * @returns promise to new MessageActivity
 */
const sendOrReplace = async (
  activity: ActivityLike,
  context: IBaseActivityContext<IMessageActivity, Record<string, any>>,
) => {
  const {
    activity: { conversation, from: messageSender },
    api,
  } = context;
  const senderId = messageSender.id;

  // checks Activity for supportet Types
  let sendActivity: IMessageActivityInput | IAdaptiveCard;
  if (typeof activity === "string") {
    sendActivity = {
      type: "message",
      text: activity,
    } satisfies IMessageActivityInput;
  } else if (activity.type == "message") {
    sendActivity = activity as IMessageActivityInput;
  } else if (activity.type == "AdaptiveCard") {
    sendActivity = new MessageActivityInput().addCard(
      "adaptive",
      activity as IAdaptiveCard,
    );
  } else {
    throw new Error("unsupported type:" + (activity?.type ?? "???"));
  }

  // Checks if PreviousMessageReferences has a senderId if it does it updates the Activity
  if (PreviousMessageReferences[senderId]) {
    const previousMessageId = PreviousMessageReferences[senderId];
    await api.conversations.updateActivity(
      conversation.id,
      previousMessageId,
      sendActivity,
    );
  } else {
    await sendWithRef(sendActivity, context);
  }
};
/**
 * Send a message in to the user and track the activity id for later reuse/manipulation
 * @param activity the activity to send
 * @param context the conversation
 * @returns
 */
async function sendWithRef(
  activity: ActivityLike,
  context: IBaseActivityContext<IMessageActivity, Record<string, any>>,
): Promise<SentActivity> {
  const {
    activity: { from: messageSender },
    send,
  } = context;
  const result = await send(activity);
  PreviousMessageReferences[messageSender.id] = result.id;
  return result;
}

const microsoftLoginTokenFactory = () => {
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
  token: microsoftLoginTokenFactory(),
};

// Use managed identity in cloud environment, otherwise use devtools plugin for local development
const options =
  process.env.BOT_TYPE === "UserAssignedMsi"
    ? { ...tokenCredentials }
    : { plugins: [] };

const app = new App({
  ...options,
  logger: new ConsoleLogger("MerchShop-Bot", { level: "debug" }),
  skipAuth: !process.env.CLIENT_ID,
});

app.on("install.add", async ({ send }) => {
  const greeting = `
  Wilkommen im Tiffinger & Thiel Merch-Shop!
  Stell dir bitte vor, es würde im Hintergrund die Ocarina of Time Shop Musik laufen :)
  Wenn du die Musik in deinem Kopf hören kannst, schreibe "shop", um Merch zu ordern.`;
  await send(greeting);
});

app.on("message", async (context): Promise<void> => {
  const { send, activity, api } = context;
  const data = getCardActionData(activity.value);
  // Define the ID(s) to filter out (e.g., your own previous messages)
  const filteredIds = ["message-id-to-ignore-1", "message-id-to-ignore-2"];

  // Check if the incoming message ID is in the filter list
  if (filteredIds.includes(activity.id)) {
    return; // Ignore this message
  }
  console.log(filteredIds);

  // Process other messages

  if (data?.action) {
    switch (data.action) {
      case "nextPage":
        return sendOrReplace(await makeProductsCard(data.page ?? 0), context);

      case "selectProduct":
        if (data.productId) {
          return sendOrReplace(await makeVariantsCard(data.productId), context);
        }
        break;

      case "filterVariants":
        if (data.productId) {
          return sendOrReplace(
            await makeVariantsCard(data.productId, data.category),
            context,
          );
        }
        break;

      case "submitProductSelection":
        // Find User Email
        const user = await api.conversations.getMemberById(
          activity.conversation.id,
          activity.from.id,
        );
        const name = user.name;
        const email = user.email;

        const putUser = await userControllerPutUser({
          userName: name!,
          userMail: email!,
        });
        console.log(putUser);

        try {
          const newActivity = await buildProductSelectionResponse(
            data,
            putUser.id,
          );
          await sendOrReplace(newActivity, context);
          delete PreviousMessageReferences[activity.from.id];
          return;
        } catch (error: unknown) {
          console.log(error);
          const card = errorResponse(
            error instanceof Error ? error.message : "Unbekannter Fehler",
          );
          send(card.card);
        }
        break;

      case "backToProducts":
        return sendOrReplace(await makeProductsCard(0), context);
    }
  }

  const text = activity.text?.trim().toLowerCase();
  if (text === "shop") {
    await sendWithRef(await makeProductsCard(), context);
  }
});

// Defining Card Actions
type CardActionData = {
  action?: string;
  page?: number;
  productId?: string;
  category?: string;
  quantity?: string; // teams inputs liefern nur strings
  variantInputIds?: VariantInputId[];
  [key: string]: unknown; // Inputs land here
};

type VariantInputId = {
  category: string;
  inputId: string;
};

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
  const card: IAdaptiveCard = {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: message,
        weight: "Bolder",
        color: "Attention",
        wrap: true,
      },
    ],
  };

  return { statusCode, card };
}

async function makeProductsCard(page: number = 0): Promise<ActivityLike> {
  const products = await productsControllerFindAll();
  const card = buildProductsCard(products, page) as IAdaptiveCard;
  return {
    type: "message",
    attachments: [cardAttachment("adaptive", card)],
  };
}

async function buildVariantsResponse(productId: string, category?: string) {
  const categories =
    await productVariantCategoryControllerFindCategories(productId);
  const variants = await productVariantControllerFindVariants(
    productId,
    category ? { category } : undefined,
  );

  return adaptiveCardResponse(
    buildVariantsCard(productId, categories, variants),
  );
}

async function makeVariantsCard(
  productId: string,
  category?: string,
): Promise<ActivityLike> {
  const response = await buildVariantsResponse(productId, category);
  return {
    type: "message",
    attachments: [cardAttachment("adaptive", response.value)],
  };
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

async function buildProductSelectionResponse(
  data: CardActionData,
  userId: string,
) {
  if (!data.productId) {
    throw new Error("productId fehlt.");
  }

  const { selected, missing } = getSelectedVariantIds(data);
  if (missing.length > 0) {
    throw new Error(`Bitte auswählen: ${missing.join(", ")}.`);
  }

  const productVariantIds = selected.map((item) => item.productVariantId);
  if (productVariantIds.length === 0) {
    throw new Error("Keine Varianten ausgewählt.");
  }

  try {
    await orderControllerCreate({
      userId,
      items: [
        {
          productId: data.productId,
          productVariantId: productVariantIds,
          quantity: Number(data.quantity) || 1,
        },
      ],
    });
  } catch (error: unknown) {
    const message =
      error instanceof Error
        ? error.message
        : "Unbekannter Fehler bei der Bestellung.";
    throw new Error(message);
  }

  return orderResponseCard;
}

(async () => {
  await app.start();
})();
