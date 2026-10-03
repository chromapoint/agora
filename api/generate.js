/**
 * Story Forge — Story Compass Document Generator
 * 
 * Generates two Word documents from extracted HartChart data:
 *   1. Story Treatment (1-3 pages)
 *   2. Core & Bonus Questions (Hart's 10 questions answered)
 * 
 * Usage: node generate.js '<json string>'
 *    or: node generate.js  (uses built-in Cinderella test data)
 */

const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, NumberFormat, LevelFormat
} = require('docx');
const fs = require('fs');

// ── Story data passed in as JSON arg, or use test data ─────────────
const story = process.argv[2] ? JSON.parse(process.argv[2]) : {
  title:            "Cinderella",
  genre:            "Fairy Tale / Revenge Drama",
  logline:          "A young woman enslaved by the stepmother who murdered her father seizes a royal ball as her one chance to reclaim her name, her inheritance, and her life — if she can stop waiting to be rescued and fight for herself.",
  protagonist:      "Cindy — daughter of a wealthy estate owner, now a servant in her own home, her identity erased one humiliation at a time.",
  wound:            "Her father was secretly poisoned and her home stolen. She has been stripped of safety, identity, and inheritance.",
  want:             "To reach the ball, expose her stepmother before the Prince, and reclaim her father's estate.",
  need:             "To trust that she deserves her own life and seize it through her own courage — not rescue.",
  world:            "A monarchy seeking its next princess. A stolen estate become a gilded prison. Stark thresholds: basement against ballroom, rags against gowns.",
  newOpportunity:   "The Prince invites every eligible woman to a ball. Cindy hears of it at the market after her invitation is hidden and burned.",
  vtg:              "Cindy commits to getting to the ball herself — sewing her own gown from stolen materials.",
  cinderellaMoment: "Cindy arrives transformed. The Prince is immediately captivated, casts the stepsisters aside, and dances with her all evening.",
  topOfMountain:    "The ball — Cindy at the center of the court, seen and valued, the Prince wholly devoted. The highest point.",
  pointOfNoReturn:  "Cindy stays past midnight — the irreversible choice that cannot be undone.",
  planFallsApart:   "The spell unravels in public. She flees in rags. She is re-imprisoned, humiliated, worse off than before. No hope.",
  resurrection:     "The Prince finds the glass slipper and searches the realm. There is still a path.",
  resolution:       "The slipper fits Cindy perfectly. She reclaims the estate, marries the Prince, and shows mercy to the stepfamily.",
  antagonist:       "The Stepmother — who murdered for the estate and embodies what Cindy most fears: that she is powerless.",
  nemesisRole:      "The Stepmother represents Cindy's deepest fear: that waiting and obedience are all that remain to her. Only by conquering that belief can Cindy neutralize the threat.",
  reflection:       "The Fairy Godmother — mirrors Cindy's wish to restore what she and her father had, and gives her the means to act.",
  stakes:           "Her home, her name, her freedom, her future. Whether a murderer profits indefinitely. Whether the rightful heir has any power at all.",
  theme:            "What is rightfully yours cannot be permanently stolen — but reclaiming it demands you stop waiting for rescue and fight for it yourself.",
  arc:              "Passive grief to active agency. Her magic only works once she has already chosen to act. She ends with power and chooses mercy.",
  whyWeCare:        "Cindy is innocent, wronged, and refuses to fully surrender. The moment she confronts her stepmother and demands her invitation, the audience is completely with her.",
  whyNow:           "The royal ball creates a unique, time-limited window — the one event in the realm powerful enough to override the stepmother's control. Without it, there is no pathway back.",
  fear:             "That the stepmother is right. That she is powerless. That effort is futile. That her father's murder goes unpunished forever.",
  jeopardy:         "Loss of home, name, freedom, and any future. The midnight deadline means even a partial victory can be revoked.",
  satisfyingEnding: "Yes — Cindy fights her way back rather than receiving her life as a gift. The mercy she shows is proof of transformation. The wry closing note (the Prince snores, chases the chambermaids) keeps it human rather than saccharine.",
  rawNotes:         "FGM appears in answer to a wish to restore the life Cindy shared with her father, not a wish for a prince.",
};

// ── Page setup ─────────────────────────────────────────────────────
const font = "Georgia";
const pageProps = {
  page: {
    size: { width: 12240, height: 15840 },
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
    pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
  }
};

// ── Paragraph helpers ──────────────────────────────────────────────
const titleBlock = (title, subtitle, sub2) => [
  new Paragraph({
    children: [new TextRun({ text: (title||"").toUpperCase(), font, size: 52, bold: true, color: "1a3a4a" })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
  }),
  new Paragraph({
    children: [new TextRun({ text: subtitle, font, size: 26, italics: true, color: "555555" })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
  }),
  new Paragraph({
    children: [new TextRun({ text: sub2, font, size: 18, color: "888888" })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 440 },
  }),
];

const sectionHead = (text) => new Paragraph({
  children: [new TextRun({ text: text.toUpperCase(), font, size: 22, bold: true, color: "1a3a4a", allCaps: true })],
  spacing: { before: 320, after: 100 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2e7d9e", space: 4 } },
});

const h1 = (text) => new Paragraph({
  children: [new TextRun({ text, font, size: 28, bold: true, color: "1a3a4a" })],
  spacing: { before: 360, after: 180 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1a3a4a", space: 4 } },
});

const body = (text) => new Paragraph({
  children: [new TextRun({ text: text || "[PLACEHOLDER]", font, size: 22, color: text ? "000000" : "999999", italics: !text })],
  spacing: { after: 160 },
  alignment: AlignmentType.JUSTIFIED,
});

const qLabel = (text) => new Paragraph({
  children: [new TextRun({ text, font, size: 19, italics: true, color: "2e7d9e" })],
  spacing: { before: 240, after: 60 },
});

const qAnswer = (text) => new Paragraph({
  children: [new TextRun({ text: text || "[PLACEHOLDER — answer this before uploading to Story Compass]", font, size: 22, color: text ? "000000" : "999999", italics: !text })],
  spacing: { after: 240 },
  alignment: AlignmentType.JUSTIFIED,
  indent: { left: 360 },
});

const rule = () => new Paragraph({
  children: [new TextRun({ text: "" })],
  spacing: { before: 60, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "dddddd", space: 1 } },
});

const gap = () => new Paragraph({ children: [new TextRun("")], spacing: { after: 100 } });

const boldNote = (label, text) => new Paragraph({
  children: [
    new TextRun({ text: label + " ", font, size: 22, bold: true }),
    new TextRun({ text: text || "[PLACEHOLDER]", font, size: 22, color: text ? "000000" : "999999", italics: !text }),
  ],
  spacing: { after: 140 },
  alignment: AlignmentType.JUSTIFIED,
});

// ── Safely get a field, return null if missing/null ─────────────────
const g = (key) => story[key] && story[key] !== "null" ? story[key] : null;

// ══════════════════════════════════════════════════════════════════
// DOCUMENT 1 — STORY TREATMENT
// ══════════════════════════════════════════════════════════════════
function buildTreatment() {
  return new Document({
    styles: { default: { document: { run: { font, size: 22 } } } },
    sections: [{
      properties: pageProps,
      children: [
        ...titleBlock(
          g("title") || "Untitled",
          "Story Treatment",
          [g("genre"), "For Story Compass"].filter(Boolean).join("  ·  ")
        ),

        sectionHead("Logline"),
        body(g("logline")),
        gap(),

        sectionHead("Beginning"),
        body(g("world")),
        gap(),
        body(`${g("protagonist")||"The protagonist"} ${g("wound") ? `carries a defining wound: ${g("wound")}` : ""}`),
        gap(),
        body(`What they want: ${g("want") || "[PLACEHOLDER]"}`),
        body(`What they need: ${g("need") || "[PLACEHOLDER]"}`),
        gap(),
        body(`The story begins with a new opportunity: ${g("newOpportunity") || "[PLACEHOLDER]"}`),
        gap(),

        sectionHead("Middle"),
        body(`The protagonist commits to a visible, tangible goal: ${g("vtg") || "[PLACEHOLDER]"}`),
        gap(),
        ...(g("cinderellaMoment") ? [
          body(`A deserved moment of success arrives — the Cinderella Moment: ${g("cinderellaMoment")}`),
          gap(),
        ] : []),
        body(`At the Top of the Mountain: ${g("topOfMountain") || "[PLACEHOLDER]"}`),
        gap(),
        body(`The Point of No Return: ${g("pointOfNoReturn") || "[PLACEHOLDER]"}`),
        gap(),
        body(`Then the Plan Falls Apart: ${g("planFallsApart") || "[PLACEHOLDER]"}`),
        gap(),

        sectionHead("Climax and End"),
        body(`The Resurrection Opportunity: ${g("resurrection") || "[PLACEHOLDER]"}`),
        gap(),
        body(`Conflict Resolution: ${g("resolution") || "[PLACEHOLDER]"}`),
        gap(),
        body(`Satisfying Ending: ${g("satisfyingEnding") || "[PLACEHOLDER]"}`),
        gap(),

        sectionHead("Conflict and Opposition"),
        body(`The Nemesis: ${g("antagonist") || "[PLACEHOLDER]"}`),
        ...(g("nemesisRole") ? [gap(), body(g("nemesisRole"))] : []),
        gap(),

        sectionHead("Stakes and Jeopardy"),
        body(g("stakes")),
        gap(),

        sectionHead("Theme"),
        body(g("theme")),
        gap(),

        sectionHead("The Arc"),
        body(g("arc")),
        gap(),

        ...(g("rawNotes") ? [sectionHead("Additional Notes"), body(g("rawNotes")), gap()] : []),

        rule(),
        gap(),
        sectionHead("Story Compass Ready"),
        body("Three structural questions to resolve before uploading:"),
        gap(),
        boldNote("1.  Point of No Return:", `Is the moment ${g("protagonist") ? "the protagonist" : "your character"} commits to the journey landing with full weight — or is it overshadowed by earlier plot?`),
        boldNote("2.  The Nemesis:", g("nemesisRole") ? "Does the Nemesis embody what the protagonist fears most about themselves, not just oppose them externally?" : "[Identify who the Nemesis is and what inner fear they embody]"),
        boldNote("3.  Satisfying Ending:", "Is the ending earned and properly anticipated — not merely happy, but satisfying?"),
      ]
    }]
  });
}

// ══════════════════════════════════════════════════════════════════
// DOCUMENT 2 — CORE & BONUS QUESTIONS
// ══════════════════════════════════════════════════════════════════
function buildQuestions() {
  return new Document({
    styles: { default: { document: { run: { font, size: 22 } } } },
    sections: [{
      properties: pageProps,
      children: [
        ...titleBlock(
          g("title") || "Untitled",
          "Core and Bonus Questions",
          `Main Character: ${g("protagonist") ? g("protagonist").split("—")[0].trim() : "[Character Name]"}`
        ),

        h1("The Core Questions"),
        body("Francis Ford Coppola's three foundational questions, as developed in the HartChart method by James V. Hart."),
        gap(),

        qLabel("Core Question 1 — Who is the main character and what do they WANT?"),
        qAnswer(g("want") ? `${g("protagonist") || "The protagonist"} wants ${g("want")}` : null),
        rule(),

        qLabel("Core Question 2 — Who and what are the OBSTACLES they must overcome?"),
        qAnswer(g("antagonist") || null),
        rule(),

        qLabel("Core Question 3 — In the end, do they get what they want? Is it good or bad?"),
        qAnswer(g("resolution") ? `${g("resolution")} ${g("satisfyingEnding") ? g("satisfyingEnding") : ""}`.trim() : null),
        gap(),

        h1("The Bonus Questions"),
        body("Hart's seven additional questions that reveal deeper character and story understanding."),
        gap(),

        qLabel("Bonus Question 4 — Why do we care? Why does the audience empathize?"),
        qAnswer(g("whyWeCare")),
        rule(),

        qLabel("Bonus Question 5 — What does the main character NEED? How does this differ from what they want?"),
        qAnswer(g("need") ? `They WANT: ${g("want") || "[PLACEHOLDER]"}. They NEED: ${g("need")}. ${g("arc") ? g("arc") : ""}`.trim() : null),
        rule(),

        qLabel("Bonus Question 6 — Why now? Why is this story happening to your character today?"),
        qAnswer(g("whyNow")),
        rule(),

        qLabel("Bonus Question 7 — What is the main character's wound/ache? What are their flaws?"),
        qAnswer(g("wound")),
        rule(),

        qLabel("Bonus Question 8 — What is the jeopardy? The danger? The stakes? The price?"),
        qAnswer(g("jeopardy") || g("stakes")),
        rule(),

        qLabel("Bonus Question 9 — What does the character fear? What shuts them down emotionally?"),
        qAnswer(g("fear")),
        rule(),

        qLabel("Bonus Question 10 — Do you have a satisfying ending?"),
        qAnswer(g("satisfyingEnding")),
        gap(),

        ...(g("reflection") ? [
          h1("Supporting Character Notes"),
          qLabel("Reflection Character"),
          qAnswer(g("reflection")),
          rule(),
        ] : []),

        ...(g("nemesisRole") ? [
          qLabel("Nemesis (Inner Fear Mirror)"),
          qAnswer(g("nemesisRole")),
        ] : []),
      ]
    }]
  });
}

// ── Build and write both ───────────────────────────────────────────
const safeName = (g("title") || "story").replace(/[^a-zA-Z0-9]/g, "-");
const outTreatment  = `/home/claude/${safeName}-Treatment.docx`;
const outQuestions  = `/home/claude/${safeName}-CoreQuestions.docx`;

Promise.all([
  Packer.toBuffer(buildTreatment()).then(b => { fs.writeFileSync(outTreatment, b); return outTreatment; }),
  Packer.toBuffer(buildQuestions()).then(b => { fs.writeFileSync(outQuestions, b); return outQuestions; }),
]).then(([t, q]) => {
  console.log(JSON.stringify({ treatment: t, questions: q }));
}).catch(err => {
  console.error("Error:", err.message);
  process.exit(1);
});
