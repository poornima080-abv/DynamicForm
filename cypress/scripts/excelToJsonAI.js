const http = require("http");

const prompt = `Return only this JSON: {"test": true}`;
const body   = JSON.stringify({ model: "llama3", prompt, stream: false });

const options = {
  hostname: "localhost",
  port:     11434,
  path:     "/api/generate",
  method:   "POST",
  headers:  {
    "Content-Type":   "application/json",
    "Content-Length": Buffer.byteLength(body),
  },
};

console.log("Sending request to Ollama...");

const req = http.request(options, (res) => {
  let data = "";
  res.on("data", chunk => { data += chunk; });
  res.on("end", () => {
    console.log("Response received");
    console.log("Raw:", data.slice(0, 500));
    try {
      const parsed = JSON.parse(data);
      console.log("Response text:", parsed.response);
    } catch (e) {
      console.log("Parse error:", e.message);
    }
  });
});

req.on("error", (e) => console.error("Error:", e.message));
req.write(body);
req.end();