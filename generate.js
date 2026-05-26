const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Hamilton Training – Day 1";

// Color palette: bold navy + gold (nod to Hamilton musical & professional feel)
const NAVY   = "1A2744";
const GOLD   = "C9A84C";
const CREAM  = "F5F0E8";
const WHITE  = "FFFFFF";
const LGRAY  = "E8E4DC";
const DGRAY  = "4A4A4A";

// ─── SLIDE 1: Title ───────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  // Gold accent bar left
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.35, h: 5.625, fill: { color: GOLD }, line: { color: GOLD } });

  // Gold thin line accent
  s.addShape(pres.shapes.RECTANGLE, { x: 0.55, y: 2.45, w: 5.2, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("HAMILTON", {
    x: 0.55, y: 0.75, w: 8, h: 0.9,
    fontSize: 52, bold: true, color: GOLD,
    fontFace: "Georgia", margin: 0, charSpacing: 8
  });

  s.addText("Training Program", {
    x: 0.55, y: 1.65, w: 8, h: 0.65,
    fontSize: 26, color: WHITE, fontFace: "Georgia", italic: true, margin: 0
  });

  s.addText("Day 1 — Welcome & Orientation", {
    x: 0.55, y: 2.65, w: 7, h: 0.55,
    fontSize: 18, color: CREAM, fontFace: "Calibri", margin: 0
  });

  s.addText("Rise to the occasion.", {
    x: 0.55, y: 4.8, w: 9, h: 0.4,
    fontSize: 12, color: GOLD, italic: true, fontFace: "Georgia", margin: 0
  });
}

// ─── SLIDE 2: Today's Agenda ──────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: CREAM };

  // Top bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("TODAY'S AGENDA", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 20, bold: true, color: GOLD, fontFace: "Georgia", valign: "middle", margin: 0, charSpacing: 4
  });

  const items = [
    ["01", "Welcome & Introductions",      "Meet your team and fellow trainees"],
    ["02", "Program Overview",             "What the Hamilton training covers"],
    ["03", "Roles & Expectations",         "Your responsibilities and goals"],
    ["04", "Tools & Resources",            "Systems, platforms, and where to find help"],
    ["05", "Q&A & Next Steps",             "Open floor — ask anything"],
  ];

  items.forEach(([num, title, sub], i) => {
    const y = 0.95 + i * 0.88;

    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.3, y: y + 0.02, w: 0.56, h: 0.56,
      fill: { color: NAVY }, line: { color: NAVY }
    });
    s.addText(num, {
      x: 0.3, y: y + 0.02, w: 0.56, h: 0.56,
      fontSize: 13, bold: true, color: GOLD, fontFace: "Georgia",
      align: "center", valign: "middle", margin: 0
    });

    // Title
    s.addText(title, {
      x: 1.05, y: y, w: 5.5, h: 0.35,
      fontSize: 15, bold: true, color: NAVY, fontFace: "Calibri", margin: 0
    });
    // Sub
    s.addText(sub, {
      x: 1.05, y: y + 0.34, w: 5.5, h: 0.3,
      fontSize: 12, color: DGRAY, fontFace: "Calibri", margin: 0
    });

    // Divider (except last)
    if (i < items.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: 0.3, y: y + 0.76, w: 9.4, h: 0,
        line: { color: LGRAY, width: 0.8 }
      });
    }
  });
}

// ─── SLIDE 3: About This Program ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  // Left panel
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.8, h: 5.625, fill: { color: NAVY }, line: { color: NAVY } });

  s.addText("About\nHamilton", {
    x: 0.2, y: 1.2, w: 3.3, h: 1.8,
    fontSize: 34, bold: true, color: GOLD, fontFace: "Georgia",
    align: "center", valign: "middle"
  });

  s.addText("A structured program designed\nto set you up for success\nfrom day one.", {
    x: 0.2, y: 3.1, w: 3.3, h: 1.5,
    fontSize: 13, color: CREAM, fontFace: "Calibri",
    align: "center", valign: "top"
  });

  // Right: 3 key facts
  const facts = [
    ["Immersive",   "Hands-on learning with real scenarios and live feedback."],
    ["Collaborative", "Work alongside peers, mentors, and cross-functional teams."],
    ["Goal-Driven", "Clear milestones track your growth from Day 1 to completion."],
  ];

  facts.forEach(([head, body], i) => {
    const y = 0.7 + i * 1.5;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 4.1, y, w: 5.6, h: 1.2,
      fill: { color: LGRAY }, line: { color: LGRAY },
      shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 135, opacity: 0.08 }
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.1, y, w: 0.1, h: 1.2, fill: { color: GOLD }, line: { color: GOLD } });
    s.addText(head, {
      x: 4.35, y: y + 0.1, w: 5.2, h: 0.35,
      fontSize: 15, bold: true, color: NAVY, fontFace: "Calibri", margin: 0
    });
    s.addText(body, {
      x: 4.35, y: y + 0.45, w: 5.15, h: 0.65,
      fontSize: 12, color: DGRAY, fontFace: "Calibri", margin: 0
    });
  });
}

// ─── SLIDE 4: What to Expect ──────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: CREAM };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.75, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("WHAT TO EXPECT", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 20, bold: true, color: GOLD, fontFace: "Georgia", valign: "middle", margin: 0, charSpacing: 4
  });

  // 2x2 grid of cards
  const cards = [
    ["📅", "Structured Sessions",  "Daily workshops, briefings,\nand skill-building exercises."],
    ["🤝", "Mentorship",           "Paired with experienced staff\nwho guide your development."],
    ["📊", "Progress Tracking",    "Regular check-ins and\nclear performance benchmarks."],
    ["🎯", "Real Assignments",     "From day one, you'll work\non meaningful, real tasks."],
  ];

  cards.forEach(([emoji, title, body], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.85;
    const y = 0.95 + row * 2.2;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.4, h: 1.95,
      fill: { color: WHITE }, line: { color: LGRAY, width: 1 },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.07 }
    });

    s.addText(emoji, { x: x + 0.15, y: y + 0.15, w: 0.7, h: 0.6, fontSize: 26, margin: 0 });

    s.addText(title, {
      x: x + 0.15, y: y + 0.65, w: 4.1, h: 0.38,
      fontSize: 14, bold: true, color: NAVY, fontFace: "Calibri", margin: 0
    });
    s.addText(body, {
      x: x + 0.15, y: y + 1.02, w: 4.1, h: 0.75,
      fontSize: 12, color: DGRAY, fontFace: "Calibri", margin: 0
    });
  });
}

// ─── SLIDE 5: Closing / Let's Get Started ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.35, h: 5.625, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText('"History has its\neyes on you."', {
    x: 0.65, y: 0.8, w: 8.8, h: 2.2,
    fontSize: 38, bold: true, italic: true, color: WHITE,
    fontFace: "Georgia", valign: "middle", margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.65, y: 3.1, w: 4, h: 0.04,
    fill: { color: GOLD }, line: { color: GOLD }
  });

  s.addText("Welcome to the Hamilton program.\nWe're glad you're here — let's make it count.", {
    x: 0.65, y: 3.3, w: 8.5, h: 0.85,
    fontSize: 16, color: CREAM, fontFace: "Calibri", margin: 0
  });

  s.addText("Day 1  ·  Hamilton Training", {
    x: 0.65, y: 4.9, w: 8, h: 0.4,
    fontSize: 11, color: GOLD, fontFace: "Calibri", italic: true, margin: 0
  });
}

pres.writeFile({ fileName: "/mnt/user-data/outputs/Hamilton_Day1.pptx" })
  .then(() => console.log("Done"))
  .catch(e => console.error(e));
