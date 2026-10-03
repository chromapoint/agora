/**
 * Story Forge — Complete Pipeline
 * 
 * Takes raw HartChart text, calls Claude to extract story structure,
 * then generates both Story Compass documents.
 * 
 * Called by the web app's /api/generate endpoint.
 * Also runnable directly: node pipeline.js chart.txt
 */

const fs   = require('fs');
const path = require('path');
const https = require('https');

const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, NumberFormat
} = require('docx');

// ── Extraction ─────────────────────────────────────────────────────
const EXTRACT_SYSTEM = `You are a story analyst trained in James V. Hart's HartChart method.

Read this HartChart export (INT. BEAT TITLE / description format) and extract the full story structure using Hart's vocabulary. Return ONLY a single-line JSON object — no markdown, no backticks, no newlines inside string values (use pipe | as separator).

Keys (null for anything not found):
{"title":null,"genre":null,"logline":null,"protagonist":null,"wound":null,"want":null,"need":null,"world":null,"newOpportunity":null,"vtg":null,"cinderellaMoment":null,"topOfMountain":null,"pointOfNoReturn":null,"planFallsApart":null,"resurrection":null,"resolution":null,"antagonist":null,"nemesisRole":null,"reflection":null,"stakes":null,"theme":null,"arc":null,"whyWeCare":null,"whyNow":null,"fear":null,"jeopardy":null,"satisfyingEnding":null,"rawNotes":null}

Definitions: want=material/ego-driven; need=inner/spiritual; vtg=Visible Tangible Goal (the Desire Line); cinderellaMoment=deserved success hit; topOfMountain=highest point, false security; pointOfNoReturn=irreversible action; planFallsApart=lowest point worse than start; resurrection=chance to recover; nemesisRole=what antagonist represents as protagonist's inner fear.`;

const TREATMENT_SYSTEM = `You are a professional story editor trained in James V. Hart's HartChart method.

Given structured story data, write the TREATMENT section content for each part of the Story Treatment document. Write in flowing prose, present tense, active voice, vivid but professional.

Return ONLY a JSON object with these string keys (no markdown, no backticks, single line):
{"beginning":null,"middle":null,"climaxAndEnd":null,"conflictAndOpposition":null,"stakesAndJeopardy":null,"themeStatement":null,"arcStatement":null}

Each value should be 2-4 rich paragraphs separated by \\n\\n. Do not invent details not in the data — write [PLACEHOLDER] for genuine gaps.`;

const QUESTIONS_SYSTEM = `You are a story analyst trained in James V. Hart's HartChart method.

Given structured story data, write the answers to Hart's Core and Bonus Questions for the main character. Each answer should be 2-4 sentences — specific, character-driven, and grounded in the actual story data provided.

Return ONLY a JSON object with these keys (no markdown, no backticks, single line):
{"q1":null,"q2":null,"q3":null,"q4":null,"q5":null,"q6":null,"q7":null,"q8":null,"q9":null,"q10":null}

q1=who is MC and what do they want; q2=obstacles; q3=do they get it (good/bad); q4=why we care; q5=what they need vs want; q6=why now; q7=wound and flaws; q8=jeopardy/stakes/price; q9=what they fear; q10=satisfying ending.
Write [PLACEHOLDER] for anything genuinely missing from the data.`;

function callClaude(apiKey, system, userMsg) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 2000,
      system,
      messages: [{ role:"user", content: userMsg }],
    });
    const req = https.request({
      hostname: "api.anthropic.com",
      path: "/v1/messages",
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "Content-Length": Buffer.byteLength(payload),
      }
    }, res => {
      let data = "";
      res.on("data", c => data += c);
      res.on("end", () => {
        try {
          const p = JSON.parse(data);
          if (p.error) return reject(new Error(p.error.message || JSON.stringify(p.error)));
          const text = p.content.filter(b=>b.type==="text").map(b=>b.text).join("");
          resolve(text);
        } catch(e) { reject(e); }
      });
    });
    req.on("error", reject);
    req.write(payload);
    req.end();
  });
}

function parseJson(raw) {
  if (!raw) return null;
  const s = raw.replace(/```json\s*/gi,"").replace(/```/g,"").trim();
  try { return JSON.parse(s); } catch {}
  const a = s.indexOf("{"), b = s.lastIndexOf("}");
  if (a !== -1 && b > a) { try { return JSON.parse(s.slice(a,b+1)); } catch {} }
  return null;
}

const g = (obj, key) => obj && obj[key] && obj[key] !== "null" ? obj[key] : null;

// ── Document builders ──────────────────────────────────────────────
const font = "Georgia";
const pageProps = {
  page: {
    size: { width:12240, height:15840 },
    margin: { top:1440, right:1440, bottom:1440, left:1440 },
    pageNumbers: { start:1, formatType: NumberFormat.DECIMAL },
  }
};

function titleBlock(title, sub1, sub2) {
  return [
    new Paragraph({ children:[new TextRun({text:(title||"UNTITLED").toUpperCase(),font,size:52,bold:true,color:"1a3a4a"})], alignment:AlignmentType.CENTER, spacing:{after:80} }),
    new Paragraph({ children:[new TextRun({text:sub1,font,size:26,italics:true,color:"555555"})], alignment:AlignmentType.CENTER, spacing:{after:60} }),
    new Paragraph({ children:[new TextRun({text:sub2,font,size:18,color:"888888"})], alignment:AlignmentType.CENTER, spacing:{after:440} }),
  ];
}
function secHead(t) { return new Paragraph({ children:[new TextRun({text:t.toUpperCase(),font,size:22,bold:true,color:"1a3a4a",allCaps:true})], spacing:{before:320,after:100}, border:{bottom:{style:BorderStyle.SINGLE,size:4,color:"2e7d9e",space:4}} }); }
function h1(t) { return new Paragraph({ children:[new TextRun({text:t,font,size:28,bold:true,color:"1a3a4a"})], spacing:{before:360,after:180}, border:{bottom:{style:BorderStyle.SINGLE,size:6,color:"1a3a4a",space:4}} }); }
function body(t) { return new Paragraph({ children:[new TextRun({text:t||"[PLACEHOLDER]",font,size:22,color:t?"000000":"999999",italics:!t})], spacing:{after:160}, alignment:AlignmentType.JUSTIFIED }); }
function qLabel(t) { return new Paragraph({ children:[new TextRun({text:t,font,size:19,italics:true,color:"2e7d9e"})], spacing:{before:240,after:60} }); }
function qAnswer(t) { return new Paragraph({ children:[new TextRun({text:t||"[PLACEHOLDER]",font,size:22,color:t?"000000":"999999",italics:!t})], spacing:{after:240}, alignment:AlignmentType.JUSTIFIED, indent:{left:360} }); }
function rule() { return new Paragraph({ children:[new TextRun("")], spacing:{before:60,after:60}, border:{bottom:{style:BorderStyle.SINGLE,size:1,color:"dddddd",space:1}} }); }
function gap() { return new Paragraph({ children:[new TextRun("")], spacing:{after:100} }); }
function boldNote(label, text) {
  return new Paragraph({ children:[new TextRun({text:label+" ",font,size:22,bold:true}),new TextRun({text:text||"[PLACEHOLDER]",font,size:22,color:text?"000000":"999999",italics:!text})], spacing:{after:140}, alignment:AlignmentType.JUSTIFIED });
}

// Prose paragraphs from a block of text (split on \n\n)
function proseBlock(text) {
  if (!text) return [body(null)];
  return text.split(/\n\n+/).filter(s=>s.trim()).map(s => body(s.trim()));
}

function buildTreatmentDoc(story, prose) {
  const title = g(story,"title") || "Untitled";
  const protagonist = g(story,"protagonist");
  const charName = protagonist ? protagonist.split(/[—–-]/)[0].trim() : "The Protagonist";

  return new Document({
    styles: { default: { document: { run: { font, size:22 } } } },
    sections: [{
      properties: pageProps,
      children: [
        ...titleBlock(title, "Story Treatment", [g(story,"genre"),"For Story Compass"].filter(Boolean).join("  ·  ")),
        secHead("Logline"),
        body(g(story,"logline")),
        gap(),
        secHead("Beginning"),
        ...proseBlock(g(prose,"beginning")),
        gap(),
        secHead("Middle"),
        ...proseBlock(g(prose,"middle")),
        gap(),
        secHead("Climax and End"),
        ...proseBlock(g(prose,"climaxAndEnd")),
        gap(),
        secHead("Conflict and Opposition"),
        ...proseBlock(g(prose,"conflictAndOpposition")),
        gap(),
        secHead("Stakes and Jeopardy"),
        ...proseBlock(g(prose,"stakesAndJeopardy")),
        gap(),
        secHead("Theme"),
        ...proseBlock(g(prose,"themeStatement") || g(story,"theme")),
        gap(),
        secHead("The Arc"),
        ...proseBlock(g(prose,"arcStatement") || g(story,"arc")),
        gap(),
        rule(), gap(),
        secHead("Story Compass Ready"),
        body("Three structural questions to resolve before uploading:"),
        gap(),
        boldNote("1.  Point of No Return:", `Is the moment ${charName} commits to the journey landing with full weight?`),
        boldNote("2.  The Nemesis:", g(story,"nemesisRole") ? "Does the Nemesis embody what the protagonist fears most about themselves?" : "[Identify who the Nemesis is and what inner fear they represent]"),
        boldNote("3.  Satisfying Ending:", "Is the ending earned and properly anticipated — not merely happy, but satisfying?"),
      ]
    }]
  });
}

function buildQuestionsDoc(story, answers) {
  const title = g(story,"title") || "Untitled";
  const protagonist = g(story,"protagonist");
  const charName = protagonist ? protagonist.split(/[—–-]/)[0].trim() : "[Character Name]";

  return new Document({
    styles: { default: { document: { run: { font, size:22 } } } },
    sections: [{
      properties: pageProps,
      children: [
        ...titleBlock(title, "Core and Bonus Questions", `Main Character: ${charName}`),
        h1("The Core Questions"),
        body("Francis Ford Coppola's three foundational questions, as developed in the HartChart method by James V. Hart."),
        gap(),
        qLabel("Core Question 1 — Who is the main character and what do they WANT?"),
        qAnswer(g(answers,"q1")), rule(),
        qLabel("Core Question 2 — Who and what are the OBSTACLES they must overcome?"),
        qAnswer(g(answers,"q2")), rule(),
        qLabel("Core Question 3 — In the end, do they get what they want? Is it good or bad?"),
        qAnswer(g(answers,"q3")),
        gap(),
        h1("The Bonus Questions"),
        body("Hart's seven additional questions that reveal deeper character and story understanding."),
        gap(),
        qLabel("Bonus Question 4 — Why do we care? Why does the audience empathize?"),
        qAnswer(g(answers,"q4")), rule(),
        qLabel("Bonus Question 5 — What does the main character NEED? How does this differ from what they want?"),
        qAnswer(g(answers,"q5")), rule(),
        qLabel("Bonus Question 6 — Why now? Why is this story happening today?"),
        qAnswer(g(answers,"q6")), rule(),
        qLabel("Bonus Question 7 — What is the main character's wound/ache? What are their flaws?"),
        qAnswer(g(answers,"q7")), rule(),
        qLabel("Bonus Question 8 — What is the jeopardy? The danger? The stakes? The price?"),
        qAnswer(g(answers,"q8")), rule(),
        qLabel("Bonus Question 9 — What does the character fear? What shuts them down emotionally?"),
        qAnswer(g(answers,"q9")), rule(),
        qLabel("Bonus Question 10 — Do you have a satisfying ending?"),
        qAnswer(g(answers,"q10")),
        gap(),
        ...(g(story,"reflection") ? [
          h1("Supporting Character Notes"),
          qLabel("Reflection Character — who mirrors the protagonist and gives them a voice?"),
          qAnswer(g(story,"reflection")), rule(),
        ] : []),
        ...(g(story,"nemesisRole") ? [
          qLabel("Nemesis — what inner fear does the antagonist embody?"),
          qAnswer(g(story,"nemesisRole")),
        ] : []),
      ]
    }]
  });
}

// ── Main pipeline ──────────────────────────────────────────────────
async function run(chartText, apiKey, outputDir) {
  outputDir = outputDir || "/mnt/user-data/outputs";

  // Step 1: Extract story structure
  process.stderr.write("1/3 Extracting story structure…\n");
  const rawExtract = await callClaude(apiKey, EXTRACT_SYSTEM, `HartChart export:\n\n${chartText.slice(0,12000)}`);
  const story = parseJson(rawExtract);
  if (!story) throw new Error("Could not parse story structure from chart");
  process.stderr.write(`    → ${story.title || "Untitled"} (${Object.values(story).filter(v=>v&&v!=="null").length} fields)\n`);

  const storyContext = Object.entries(story).filter(([,v])=>v&&v!=="null").map(([k,v])=>`${k}: ${v}`).join("\n");

  // Step 2: Generate prose for treatment + answers for questions (parallel)
  process.stderr.write("2/3 Writing treatment and questions…\n");
  const [rawProse, rawAnswers] = await Promise.all([
    callClaude(apiKey, TREATMENT_SYSTEM, `Story data:\n${storyContext}`),
    callClaude(apiKey, QUESTIONS_SYSTEM, `Story data:\n${storyContext}`),
  ]);
  const prose   = parseJson(rawProse)   || {};
  const answers = parseJson(rawAnswers) || {};

  // Step 3: Build documents
  process.stderr.write("3/3 Building documents…\n");
  const safeName = (story.title || "story").replace(/[^a-zA-Z0-9]/g, "-");
  const tPath = `${outputDir}/${safeName}-Treatment.docx`;
  const qPath = `${outputDir}/${safeName}-CoreQuestions.docx`;

  await Promise.all([
    Packer.toBuffer(buildTreatmentDoc(story, prose)).then(b => fs.writeFileSync(tPath, b)),
    Packer.toBuffer(buildQuestionsDoc(story, answers)).then(b => fs.writeFileSync(qPath, b)),
  ]);

  process.stderr.write(`    → ${tPath}\n    → ${qPath}\n`);
  return { treatment: tPath, questions: qPath, title: story.title };
}

// ── CLI entry point ────────────────────────────────────────────────
if (require.main === module) {
  const chartFile = process.argv[2];
  const key       = process.argv[3] || process.env.ANTHROPIC_API_KEY;

  if (!chartFile || !fs.existsSync(chartFile)) {
    console.error("Usage: node pipeline.js path/to/chart.txt [API_KEY]");
    process.exit(1);
  }
  if (!key) {
    console.error("No API key. Set ANTHROPIC_API_KEY or pass as second argument.");
    process.exit(1);
  }

  const text = fs.readFileSync(chartFile, "utf8");
  run(text, key).then(result => {
    console.log(`\nStory Forge complete.\n  "${result.title}"\n  Treatment:      ${result.treatment}\n  Core Questions: ${result.questions}`);
  }).catch(err => {
    console.error("Error:", err.message);
    process.exit(1);
  });
}

module.exports = { run };
