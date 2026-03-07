const http = require("http");

const port = Number(process.env.PORT) || 3000;

const server = http.createServer((req, res) => {
  res.writeHead(200, { "Content-Type": "application/json" });
  res.end(
    JSON.stringify({
      name: "tsifulator.ai",
      status: "ok",
      hasOpenAiKey: Boolean(process.env.OPENAI_API_KEY),
      timestamp: new Date().toISOString()
    })
  );
});

server.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
