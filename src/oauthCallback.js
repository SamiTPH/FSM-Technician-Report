const express = require("express");

const app = express();
const port = 3000;

app.get("/oauth/callback", (req, res) => {
  const code = req.query.code || "";

  console.log("Zoho authorization code:");
  console.log(code || "(no code received)");

  res.type("html").send(`
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>Zoho OAuth Callback</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.5; }
          pre { background: #f4f4f4; border: 1px solid #ddd; padding: 16px; white-space: pre-wrap; }
        </style>
      </head>
      <body>
        <h1>Zoho authorization received</h1>
        <p>Copy this code from here or from your terminal:</p>
        <pre>${escapeHtml(code || "No code received")}</pre>
      </body>
    </html>
  `);
});

app.listen(port, () => {
  console.log(`Listening on http://localhost:${port}/oauth/callback`);
});

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
