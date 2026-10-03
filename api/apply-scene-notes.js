// api/apply-scene-notes.js — Compass: Apply Story Compass feedback to screenplay scenes
// Set ANTHROPIC_API_KEY in Vercel environment variables.

function analyzeChanges(original, revised) {
  const origLines = original.split("\n");
  const revLines = revised.split("\n");

  // Track changes by category
  const changes = {
    dialogueChanged: 0,
    actionChanged: 0,
    charactersRenamed: false,
    linesAdded: 0,
    linesRemoved: 0,
    linesModified: 0,
  };

  // Simple diff: identify changed/added/removed lines
  const origSet = new Set(origLines);
  const revSet = new Set(revLines);

  // Count added and removed lines
  for (const line of revLines) {
    if (!origSet.has(line.trim())) {
      changes.linesAdded++;
      // Try to categorize the change
      if (line.match(/^[A-Z\s]+$/) && line.trim().length > 0) {
        changes.dialogueChanged++;
      } else if (line.trim() && !line.match(/^[A-Z\s(]/)) {
        changes.actionChanged++;
      }
    }
  }

  for (const line of origLines) {
    if (!revSet.has(line.trim())) {
      changes.linesRemoved++;
    }
  }

  // Basic calculation of modified lines
  changes.linesModified = Math.min(
    origLines.length,
    revLines.length
  ) - Math.max(0, Math.floor(origLines.length * 0.7));

  // Build concise summary
  const summaryParts = [];

  if (changes.linesAdded > 0) {
    summaryParts.push(`Added ${changes.linesAdded} line${changes.linesAdded > 1 ? "s" : ""}`);
  }
  if (changes.linesRemoved > 0) {
    summaryParts.push(`Removed ${changes.linesRemoved} line${changes.linesRemoved > 1 ? "s" : ""}`);
  }
  if (changes.linesModified > 0) {
    summaryParts.push(`Revised dialogue and action`);
  }
  if (summaryParts.length === 0) {
    summaryParts.push("Minor refinements applied");
  }

  return summaryParts.join(" · ");
}

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

  const { scene, feedback } = body;

  // Validate both fields are non-empty strings
  if (!scene || typeof scene !== "string" || !scene.trim()) {
    return res.status(400).json({ error: "Scene text is required." });
  }
  if (!feedback || typeof feedback !== "string" || !feedback.trim()) {
    return res.status(400).json({ error: "Story Compass notes are required." });
  }

  // Check combined length (~60,000 character limit)
  const combined = scene + feedback;
  if (combined.length > 60000) {
    return res.status(400).json({ error: "Combined scene and notes are too long. Please shorten and try again." });
  }

  try {
    const fetch = (await import('node-fetch')).default;
    const headers = {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    };

    const systemPrompt = `You are a professional script editor. You will receive a screenplay scene and a written analysis of that scene from a script-analysis tool. Your job is to produce a revised version of the scene that applies the analysis's suggestions.

WHICH PARTS OF THE ANALYSIS TO ACT ON
- The analysis usually contains descriptive sections (Function and Purpose, Scene Goals, Character Analysis, Scene Structure, Dialogue and Subtext). These describe the scene. Do not treat them as instructions, but use them to understand what the scene must keep doing.
- Act on the "Weaknesses" and "Improvements" sections. The "Summary" tells you which changes matter most.
- The analysis may arrive with messy formatting (bullets run together, missing line breaks). Read through that carefully.
- If the analysis offers a specific replacement line (e.g. "Try: ..."), use it or something very close to it that fits the character's voice.
- If a suggestion is vague, make the smallest change that clearly delivers it.
- If two suggestions conflict, favour the one the Summary emphasises.
- If the analysis says something works well, keep it.

THE SHOW
"Poto me Pipi and the Mysteries of Tāpito". A 2D animated mystery series: 10 episodes of 5 minutes each, for tamariki aged 5 to 8 and whānau co-viewing. The finished show is fluent, 100% te reo Māori. In the fictional town of Tāpito, Poto and Pipi, inspired by the comic-book detective Matua Mātai, solve whimsical mysteries that show empathy, problem-solving and te ao Māori values.

CHARACTERS (keep every revision true to these)
- Pipi: an energetic, quick-witted ngeru (cat). Confident and bold, she tries to take charge and dives headfirst into mysteries. Her impulsiveness sometimes gets her and Poto into trouble, but in a fun way. She relies on Poto to think things through.
- Poto: a thoughtful, analytical kurī (dog) and Pipi's best buddy. The more cautious one of the pair. He is kind-hearted and steady and thinks problems through carefully, sometimes holding back out of fear. He comes up with ideas, and Pipi runs with them. He learns from Pipi's boldness.
- Koro Tama: Pipi's wise and wonderfully quirky koro. A storyteller who shares Matua Mātai comics with the tamariki. At storytime he sometimes casts himself as the hero, makes up the ending or dozes off. He's a passionate gardener who teaches life lessons through stories, nature and his own experience. His quirks are funny, but he is always warm, loving and close to the tamariki. He is never cold, uncaring or neglectful. If a note calls him "detached", show it as a comic distraction and not as a lack of love.

AUDIENCE: TAMARIKI AGED 5 TO 8
- Every change must suit 5–8-year-olds watching with their whānau.
- Episodes are only 5 minutes long, so keep scenes tight. Don't pad them out.
- Tamariki characters talk like real 5–8-year-olds: short sentences, simple everyday words, big feelings said plainly. No adult phrasing or witty adult banter coming from the tamariki.

LANGUAGE
- Write the revision in the same language as the scene. If the scene is in te reo Māori, keep all dialogue and action in te reo Māori, and translate any English wording from the notes (such as a suggested line) into natural te reo Māori that fits the characters. If the scene is in English with Māori words, keep that mix.
- Never turn te reo Māori into English.
- Humour should be warm and clear to kids: physical comedy, silly visuals, playful misunderstandings, catchphrases. Don't use sarcasm that belittles, innuendo, or jokes that only land for adults.
- Adults can be quirky, distracted or comic, but they stay loving and safe. Don't make them mean, dismissive in a hurtful way, or frightening.
- Mystery and tension should feel exciting, not scary. No peril, threat or content inappropriate for children.
- Favour visual storytelling (actions, expressions, glances) that young viewers can follow without dialogue.
- Keep positive values the scene already carries, like curiosity, teamwork and family warmth.

WHAT TO PRESERVE
- The scene's plot, story function, and every story beat the analysis calls essential.
- Character and place names, spellings and macrons as the show uses them: Pipi, Poto, Koro Tama, Matua Mātai, Tāpito. The analysis may misspell these (e.g. "Mātua Mātai"). Always follow the scene's and the show's spelling, not the analysis's. Never anglicise or remove macrons.
- Each character's voice, age-appropriateness, and the overall tone.
- Screenplay formatting conventions used in the original: scene headings, action lines, character cues in caps, parentheticals, dialogue, transitions, and montage formatting. Match the original's layout and line spacing.
- Anything the notes don't touch should stay as close to the original as possible.

WHAT NOT TO DO
- Don't add new characters, subplots or story events the notes don't call for.
- Don't lengthen the scene unless a note asks for something that requires it. Notes about tightening or trimming should make it shorter.
- Don't add any commentary, explanation, headings, notes, or a list of changes.

OUTPUT
Return only the full revised scene as plain screenplay text. Nothing before it, nothing after it, no code fences.`;

    const userMessage = `<scene>
${scene}
</scene>

<analysis>
${feedback}
</analysis>

Apply the analysis to the scene and return only the revised scene.`;

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers,
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 8000,
        temperature: 0.4,
        system: systemPrompt,
        messages: [{ role: "user", content: userMessage }],
      }),
    });

    const data = await response.json();

    if (response.status !== 200) {
      return res.status(502).json({ error: "Couldn't apply notes, please try again." });
    }

    if (!Array.isArray(data.content)) {
      return res.status(502).json({ error: "Couldn't apply notes, please try again." });
    }

    // Extract all text blocks, join, and trim
    let revisedScene = data.content
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("");

    // Strip leading/trailing code fence markers if present
    revisedScene = revisedScene.replace(/^```[\w]*\n?/, "").replace(/\n?```$/, "").trim();

    // Analyze changes between original and revised
    const changeSummary = analyzeChanges(scene, revisedScene);

    return res.status(200).json({ revisedScene, changeSummary });
  } catch (err) {
    return res.status(502).json({ error: "Couldn't apply notes, please try again." });
  }
}
