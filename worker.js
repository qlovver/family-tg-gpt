// Entry point for Cloudflare Workers (see wrangler.jsonc "main").
// The bot logic lives in src/: bot.js (Telegram + history), model.js
// (Tooken Club API), format.js (Markdown/LaTeX -> Telegram HTML).
export { default } from "./src/bot.js";
