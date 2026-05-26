const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

const SKY  = "1A78C2";
const DEEP = "0D3B6E";
const WHITE = "FFFFFF";
const GOLD  = "F5C842";
const GRAY  = "444444";
const LGRAY = "F0F4F8";

// ── Slide 1: Title ──────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: DEEP };

  s.addShape(pres.shapes.OVAL, { x: 7.5, y: -1.2, w: 4.5, h: 4.5, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("Why is the sky blue?", {
    x: 0.5, y: 1.6, w: 7, h: 1.5,
    fontSize: 52, bold: true, color: WHITE, fontFace: "Georgia", margin: 0
  });
  s.addText("A simple 2-minute explanation.", {
    x: 0.5, y: 3.3, w: 6, h: 0.6,
    fontSize: 20, color: GOLD, fontFace: "Calibri", italic: true, margin: 0
  });
}

// ── Slide 2: The Spectrum with wavelengths ──────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText("Light comes in different wavelengths", {
    x: 0.5, y: 0.25, w: 9, h: 0.7,
    fontSize: 28, bold: true, color: DEEP, fontFace: "Georgia", margin: 0
  });
  s.addText("Each color = a different wavelength, measured in nanometres (nm).", {
    x: 0.5, y: 0.82, w: 9, h: 0.38,
    fontSize: 14, color: GRAY, fontFace: "Calibri", margin: 0
  });

  // Spectrum bars with colors, names, and wavelength ranges
  const bands = [
    { color: "7B00D4", label: "Violet",  nm: "380–450 nm" },
    { color: "3B4FE0", label: "Blue",    nm: "450–495 nm" },
    { color: "1AAF5D", label: "Green",   nm: "495–570 nm" },
    { color: "F5D800", label: "Yellow",  nm: "570–590 nm" },
    { color: "FF8C00", label: "Orange",  nm: "590–620 nm" },
    { color: "E8000D", label: "Red",     nm: "620–750 nm" },
  ];

  const barY = 1.35;
  const barH = 1.8;
  const startX = 0.4;
  const totalW = 9.2;
  const bw = totalW / bands.length;

  bands.forEach((b, i) => {
    const x = startX + i * bw;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: barY, w: bw, h: barH,
      fill: { color: b.color }, line: { color: b.color }
    });
    // Color name
    s.addText(b.label, {
      x, y: barY + barH + 0.08, w: bw, h: 0.35,
      fontSize: 13, bold: true, color: DEEP, fontFace: "Calibri",
      align: "center", margin: 0
    });
    // Wavelength
    s.addText(b.nm, {
      x, y: barY + barH + 0.42, w: bw, h: 0.3,
      fontSize: 11, color: GRAY, fontFace: "Calibri",
      align: "center", margin: 0
    });
  });

  // Arrow + labels: SHORT <-----> LONG
  s.addShape(pres.shapes.LINE, {
    x: startX, y: 4.55, w: totalW, h: 0,
    line: { color: GRAY, width: 1.5 }
  });
  s.addText("◄ Short wavelength", {
    x: startX, y: 4.65, w: 4, h: 0.35,
    fontSize: 12, color: "7B00D4", fontFace: "Calibri", bold: true, margin: 0
  });
  s.addText("Long wavelength ►", {
    x: 5.6, y: 4.65, w: 4, h: 0.35,
    fontSize: 12, color: "E8000D", fontFace: "Calibri", bold: true,
    align: "right", margin: 0
  });

  // Highlight box around blue
  s.addShape(pres.shapes.RECTANGLE, {
    x: startX + bw - 0.05, y: barY - 0.1, w: bw + 0.1, h: barH + 0.2,
    fill: { type: "none" },
    line: { color: "FFFFFF", width: 3 }
  });
  s.addText("⬆ This is the key one", {
    x: startX + bw - 0.1, y: barY + barH + 0.9, w: bw + 0.3, h: 0.38,
    fontSize: 11, color: DEEP, fontFace: "Calibri", bold: true,
    align: "center", margin: 0
  });
}

// ── Slide 3: Why blue scatters ───────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LGRAY };

  s.addText("Short wavelengths scatter more", {
    x: 0.5, y: 0.25, w: 9, h: 0.7,
    fontSize: 28, bold: true, color: DEEP, fontFace: "Georgia", margin: 0
  });

  // Simple visual: sun ray hits atmosphere, blue bounces, red goes through
  // Sun
  s.addShape(pres.shapes.OVAL, { x: 0.3, y: 1.3, w: 1.0, h: 1.0, fill: { color: GOLD }, line: { color: GOLD } });
  s.addText("SUN", { x: 0.3, y: 1.3, w: 1.0, h: 1.0, fontSize: 10, bold: true, color: DEEP, align: "center", valign: "middle", margin: 0 });

  // White ray going right
  s.addShape(pres.shapes.LINE, { x: 1.35, y: 1.8, w: 1.5, h: 0, line: { color: "AAAAAA", width: 3 } });

  // Atmosphere blob
  s.addShape(pres.shapes.OVAL, { x: 2.9, y: 1.35, w: 1.5, h: 0.9, fill: { color: "D0E8FF" }, line: { color: SKY, width: 2 } });
  s.addText("atmosphere\nmolecule", { x: 2.9, y: 1.35, w: 1.5, h: 0.9, fontSize: 9, color: DEEP, align: "center", valign: "middle", margin: 0 });

  // Blue scatters up/down
  s.addShape(pres.shapes.LINE, { x: 3.65, y: 1.35, w: 0, h: -0.9, line: { color: "3B4FE0", width: 3 } });
  s.addShape(pres.shapes.LINE, { x: 3.65, y: 2.25, w: 0, h:  0.9, line: { color: "3B4FE0", width: 3 } });
  s.addText("Blue (450 nm)\nscatters away →\nfills the whole sky", { x: 2.7, y: 3.25, w: 2.5, h: 0.8, fontSize: 12, color: "3B4FE0", bold: true, fontFace: "Calibri", align: "center", margin: 0 });

  // Red passes through
  s.addShape(pres.shapes.LINE, { x: 4.45, y: 1.8, w: 2.8, h: 0, line: { color: "E8000D", width: 3 } });
  s.addText("Red (700 nm)\npasses straight through", { x: 5.1, y: 0.85, w: 2.5, h: 0.45, fontSize: 12, color: "E8000D", bold: true, fontFace: "Calibri", align: "center", margin: 0 });

  // Eye at right
  s.addShape(pres.shapes.OVAL, { x: 7.4, y: 1.5, w: 0.7, h: 0.6, fill: { color: "FFFFFF" }, line: { color: GRAY, width: 1 } });
  s.addShape(pres.shapes.OVAL, { x: 7.6, y: 1.62, w: 0.3, h: 0.3, fill: { color: DEEP }, line: { color: DEEP } });
  s.addText("your eye", { x: 7.2, y: 2.15, w: 1.2, h: 0.3, fontSize: 10, color: GRAY, fontFace: "Calibri", align: "center", margin: 0 });

  // Summary box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.3, w: 9, h: 0.95, fill: { color: DEEP }, line: { color: DEEP } });
  s.addText("Blue light (short wavelength) bounces off molecules in all directions → your eyes see blue everywhere you look up.", {
    x: 0.65, y: 4.35, w: 8.7, h: 0.85,
    fontSize: 14, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });
}

// ── Slide 4: Bonus sunset + summary ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: "1C1C2E" };

  s.addText("Bonus: Why are sunsets orange & red?", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 28, bold: true, color: GOLD, fontFace: "Georgia", margin: 0
  });

  // Sunset gradient bands
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.05, w: 10, h: 0.65, fill: { color: "F25C54" }, line: { color: "F25C54" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.7,  w: 10, h: 0.65, fill: { color: "F4845F" }, line: { color: "F4845F" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.35, w: 10, h: 0.65, fill: { color: "F7B267" }, line: { color: "F7B267" } });

  s.addText("At sunset, sunlight travels through much MORE atmosphere to reach you.\nAll the blue (short wavelength) gets scattered away before it arrives.\nOnly red & orange (long wavelength, 620–750 nm) survive the long trip.", {
    x: 0.5, y: 3.15, w: 9, h: 1.2,
    fontSize: 15, color: WHITE, fontFace: "Calibri", margin: 0
  });

  s.addText("Same physics. Longer path. Different color.", {
    x: 0.5, y: 4.5, w: 9, h: 0.5,
    fontSize: 16, color: GOLD, fontFace: "Georgia", italic: true, bold: true, margin: 0
  });
}

pres.writeFile({ fileName: "/home/claude/WhyIsSkyBlue2.pptx" })
  .then(() => console.log("Done"))
  .catch(e => console.error(e));
