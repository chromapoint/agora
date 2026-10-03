// api/claude.js — Story Forge relay + document generator
// Set ANTHROPIC_API_KEY in Vercel environment variables.

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: "ANTHROPIC_API_KEY not configured." });

  const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  if (!body) return res.status(400).json({ error: "Request body missing or too large." });

  // Mode 1: Generate both Story Compass documents from chart text
  if (body.mode === "generate") {
    try {
      const { run } = require('./pipeline.js');
      const result = await run(body.chartText, apiKey, "/tmp");
      const fs = require('fs');
      const treatment = fs.readFileSync(result.treatment).toString('base64');
      const questions = fs.readFileSync(result.questions).toString('base64');
      return res.status(200).json({
        title: result.title,
        treatment: { filename: `${result.title}-Treatment.docx`,  data: treatment },
        questions: { filename: `${result.title}-CoreQuestions.docx`, data: questions },
      });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // Mode 2: Direct Claude API relay
  try {
    const fetch = (await import('node-fetch')).default;
    const headers = {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    };
    const betas = body.betas ? [...body.betas] : [];

    const hasPdf = Array.isArray(body.messages) && body.messages.some(m =>
      Array.isArray(m.content) && m.content.some(c => c.type === "document" && c.source?.media_type === "application/pdf")
    );
    if (hasPdf && !betas.includes("pdfs-2024-09-25")) betas.push("pdfs-2024-09-25");

    let systemPayload = body.system;
    if (typeof body.system === "string" && body.system.length >= 4000) {
      systemPayload = [{ type: "text", text: body.system, cache_control: { type: "ephemeral" } }];
      if (!betas.includes("prompt-caching-2024-07-31")) betas.push("prompt-caching-2024-07-31");
    }

    if (betas.length) headers["anthropic-beta"] = betas.join(",");

    const useStream = body.stream === true;

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers,
      body: JSON.stringify({
        model: body.model || "claude-sonnet-4-6",
        max_tokens: body.max_tokens || 2000,
        system: systemPayload,
        messages: body.messages,
        stream: useStream,
      }),
    });

    if (useStream) {
      res.setHeader("Content-Type", "text/event-stream");
      res.setHeader("Cache-Control", "no-cache");
      res.setHeader("Connection", "keep-alive");
      res.status(response.status);
      await new Promise((resolve, reject) => {
        response.body.on("data", chunk => res.write(chunk));
        response.body.on("end", resolve);
        response.body.on("error", reject);
      });
      res.end();
      return;
    }

    const data = await response.json();
    return res.status(response.status).json(data);
  } catch (err) {
    return res.status(500).json({ error: "Relay failed: " + err.message });
  }
}
