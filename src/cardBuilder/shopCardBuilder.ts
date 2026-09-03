import { ProductDTO, ProductListDTO } from "../api/merchApi";
import { buildEmptyStateCard } from "./cardBuilder";

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
