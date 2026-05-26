# Why is the Sky Blue?

A simple 2-minute presentation for the Hamilton training class, Day 1.

## Slides

1. **Title** — Why is the Sky Blue?
2. **Light comes in different wavelengths** — Full color spectrum with nm ranges (380–750 nm) and short → long wavelength arrow
3. **Bonus: Why are sunsets orange & red?** — Same physics, longer path

## Files

- `WhyIsSkyBlue.pptx` — The presentation (open in PowerPoint or Google Slides)
- `sky_blue.js` — Node.js script to regenerate the deck

---

## How to regenerate the presentation

The `.pptx` file is already included in this repo — you only need to follow these steps if you want to edit and rebuild it yourself.

### Step 1 — Install Node.js

Node.js is the program that runs the script. If you type `node` in your terminal and it says "not recognized", you need to install it.

1. Go to https://nodejs.org
2. Download the **LTS** version (recommended for most users)
3. Run the installer and follow the steps
4. Once done, **close and reopen your terminal** (Command Prompt or PowerShell on Windows)
5. Check it worked by typing: `node -v` — it should print a version number like `v20.x.x`

npm comes included with Node.js, so after this step you'll have both.

### Step 2 — Install pptxgenjs

Once Node.js is installed, run this in your terminal:

```bash
npm install -g pptxgenjs
```

### Step 3 — Run the script

```bash
node sky_blue.js
```

This will generate a fresh `WhyIsSkyBlue.pptx` in the same folder.
