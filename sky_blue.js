const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

const SKY   = "1A78C2";
const DEEP  = "0D3B6E";
const WHITE = "FFFFFF";
const GOLD  = "F5C842";
const LGRAY = "EAF4FB";
const GRAY  = "555555";

// ── Slide 1: Title ──────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: DEEP };

  // Big circle as "sun"
  s.addShape(pres.shapes.OVAL, { x: 7.2, y: -1.0, w: 4.5, h: 4.5, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("Why is the", {
    x: 0.5, y: 1.0, w: 7, h: 0.8,
    fontSize: 32, color: WHITE, fontFace: "Calibri", bold: false, margin: 0
  });
  s.addText("Sky Blue?", {
    x: 0.5, y: 1.75, w: 7, h: 1.3,
    fontSize: 64, bold: true, color: GOLD, fontFace: "Georgia", margin: 0
  });
  s.addText("A 2-minute answer to a question\nyou've wondered about your whole life.", {
    x: 0.5, y: 3.2, w: 6.5, h: 0.9,
    fontSize: 16, color: WHITE, fontFace: "Calibri", italic: true, margin: 0
  });
}

// ── Slide 2: Sunlight is all colors ────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText("Sunlight isn't just white", {
    x: 0.5, y: 0.3, w: 9, h: 0.8,
    fontSize: 36, bold: true, color: DEEP, fontFace: "Georgia", margin: 0
  });
  s.addText("It's actually a mix of ALL colors of the rainbow.", {
    x: 0.5, y: 1.1, w: 9, h: 0.5,
    fontSize: 18, color: GRAY, fontFace: "Calibri", margin: 0
  });

  // Rainbow bars
  const colors = ["E63946","F4A261","E9C46A","2A9D8F","457B9D","5C4D7D","9B5DE5"];
  const labels = ["Red","Orange","Yellow","Green","Blue","Indigo","Violet"];
  colors.forEach((c, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5 + i * 1.27, y: 1.85, w: 1.15, h: 2.2,
      fill: { color: c }, line: { color: c }
    });
    s.addText(labels[i], {
      x: 0.5 + i * 1.27, y: 4.1, w: 1.15, h: 0.4,
      fontSize: 12, color: GRAY, fontFace: "Calibri", align: "center", margin: 0
    });
  });

  s.addText("All these wavelengths travel together from the Sun to Earth.", {
    x: 0.5, y: 4.7, w: 9, h: 0.5,
    fontSize: 14, color: GRAY, fontFace: "Calibri", italic: true, margin: 0
  });
}

// ── Slide 3: Rayleigh Scattering ────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LGRAY };

  s.addText("Meet: Rayleigh Scattering", {
    x: 0.5, y: 0.3, w: 9, h: 0.8,
    fontSize: 36, bold: true, color: DEEP, fontFace: "Georgia", margin: 0
  });

  // Left: explanation
  s.addText("When sunlight hits Earth's atmosphere, it collides with tiny gas molecules (mostly nitrogen and oxygen).\n\nThese molecules scatter light — but they scatter SHORT wavelengths much more than long ones.\n\nBlue light has a short wavelength → it gets scattered in ALL directions across the sky.\n\nRed & orange have long wavelengths → they pass straight through.", {
    x: 0.5, y: 1.25, w: 5.3, h: 3.8,
    fontSize: 15, color: GRAY, fontFace: "Calibri", margin: 0
  });

  // Right: simple visual — molecule scattering blue dots
  s.addShape(pres.shapes.OVAL, { x: 6.5, y: 1.5, w: 0.5, h: 0.5, fill: { color: SKY }, line: { color: SKY } });
  s.addShape(pres.shapes.OVAL, { x: 7.4, y: 1.2, w: 0.5, h: 0.5, fill: { color: SKY }, line: { color: SKY } });
  s.addShape(pres.shapes.OVAL, { x: 8.2, y: 1.7, w: 0.5, h: 0.5, fill: { color: SKY }, line: { color: SKY } });
  s.addShape(pres.shapes.OVAL, { x: 6.8, y: 2.5, w: 0.5, h: 0.5, fill: { color: SKY }, line: { color: SKY } });
  s.addShape(pres.shapes.OVAL, { x: 8.0, y: 2.8, w: 0.5, h: 0.5, fill: { color: SKY }, line: { color: SKY } });
  s.addShape(pres.shapes.OVAL, { x: 7.2, y: 3.3, w: 0.5, h: 0.5, fill: { color: SKY }, line: { color: SKY } });
  // Central molecule
  s.addShape(pres.shapes.OVAL, { x: 7.15, y: 2.1, w: 0.7, h: 0.7, fill: { color: DEEP }, line: { color: DEEP } });
  s.addText("molecule", { x: 6.85, y: 2.85, w: 1.3, h: 0.3, fontSize: 10, color: GRAY, fontFace: "Calibri", align: "center", margin: 0 });
  s.addText("🔵 blue light\nscattered everywhere", { x: 6.4, y: 3.7, w: 2.8, h: 0.6, fontSize: 12, color: SKY, fontFace: "Calibri", bold: true, align: "center", margin: 0 });
}

// ── Slide 4: Bonus — sunsets ────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: "1C1C2E" };

  // Gradient-ish sunset with rectangles
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.5, w: 10, h: 3.125, fill: { color: "F4845F" }, line: { color: "F4845F" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.5, w: 10, h: 1.2,   fill: { color: "F7B267" }, line: { color: "F7B267" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.8, w: 10, h: 0.9,   fill: { color: "F25C54" }, line: { color: "F25C54" } });

  s.addText("Bonus: Why are sunsets red?", {
    x: 0.5, y: 0.05, w: 9, h: 0.7,
    fontSize: 30, bold: true, color: GOLD, fontFace: "Georgia", margin: 0
  });

  s.addText("At sunset, sunlight travels through MUCH more atmosphere to reach your eyes.\n\nAll the blue light gets scattered away long before it arrives.\n\nOnly the long-wavelength reds and oranges make it through → stunning sunsets!", {
    x: 0.5, y: 2.65, w: 9, h: 2.5,
    fontSize: 15, color: WHITE, fontFace: "Calibri", margin: 0
  });
}

// ── Slide 5: Summary ────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: DEEP };

  s.addText("So... why IS the sky blue?", {
    x: 0.5, y: 0.35, w: 9, h: 0.8,
    fontSize: 36, bold: true, color: GOLD, fontFace: "Georgia", margin: 0
  });

  const steps = [
    ["☀️", "Sunlight carries all colors"],
    ["💨", "Atmosphere scatters short wavelengths most"],
    ["🔵", "Blue scatters in every direction → fills the sky"],
    ["🌅", "At sunset, only red/orange survive the longer path"],
  ];

  steps.forEach(([icon, text], i) => {
    const y = 1.3 + i * 0.92;
    s.addShape(pres.shapes.OVAL, { x: 0.4, y: y + 0.05, w: 0.65, h: 0.65, fill: { color: SKY }, line: { color: SKY } });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.65, h: 0.65, fontSize: 18, align: "center", valign: "middle", margin: 0 });
    s.addText(text, { x: 1.25, y: y + 0.1, w: 8, h: 0.55, fontSize: 18, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0 });
  });

  s.addText("It's not magic — it's physics. Pretty cool, right?", {
    x: 0.5, y: 5.0, w: 9, h: 0.4,
    fontSize: 14, color: GOLD, fontFace: "Calibri", italic: true, margin: 0
  });
}

pres.writeFile({ fileName: "/home/claude/WhyIsSkyBlue.pptx" })
  .then(() => console.log("Done"))
  .catch(e => console.error(e));
