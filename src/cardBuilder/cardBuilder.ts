export function buildEmptyStateCard(message: string): object {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [{ type: "TextBlock", text: message, isSubtle: true, wrap: true }],
  };
}
