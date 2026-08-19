const rawInput = document.getElementById("raw-file");
const templateInput = document.getElementById("template-file");
const convInput = document.getElementById("conv-file");
const registerInput = document.getElementById("register-files");
const monthOverride = document.getElementById("month-override");
const convCostOverride = document.getElementById("conv-cost-override");
const outputNameInput = document.getElementById("output-name");
const convertBtn = document.getElementById("convert-btn");
const logEl = document.getElementById("log");
const advancedToggle = document.getElementById("advanced-toggle");
const advancedPanel = document.getElementById("advanced-panel");

advancedToggle.addEventListener("click", () => {
  advancedPanel.classList.toggle("hidden");
});

function log(msg, cls) {
  logEl.classList.add("visible");
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = msg;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

function resetLog() {
  logEl.innerHTML = "";
  logEl.classList.remove("visible");
}

convertBtn.addEventListener("click", async () => {
  resetLog();

  const rawFile = rawInput.files[0];
  const templateFile = templateInput.files[0];
  const convFile = convInput.files[0];

  if (!rawFile || !templateFile || !convFile) {
    log("Please choose all three files before converting.", "error");
    return;
  }

  convertBtn.disabled = true;
  convertBtn.textContent = "Converting...";

  try {
    const { loadWorkbook, loadRawRows, detectMonth, monthLabel, detectConvCostBase, buildOutput, downloadWorkbook } =
      window.MRMConverter;

    log("Reading raw file...");
    const rawWb = await loadWorkbook(rawFile);
    const rawRows = loadRawRows(rawWb);

    let monthObj;
    if (monthOverride.value) {
      const [y, m] = monthOverride.value.split("-").map(Number);
      monthObj = { year: y, month: m - 1 };
    } else {
      monthObj = detectMonth(rawRows);
    }
    log(`Month: ${monthLabel(monthObj)}`);

    log("Reading last month's MRM file (for the INFO sheet)...");
    const templateWb = await loadWorkbook(templateFile);

    let convCostBase = convCostOverride.value ? parseFloat(convCostOverride.value) : null;
    if (convCostBase === null) {
      log("Reading GP Conversion Cost file...");
      const convWb = await loadWorkbook(convFile);
      convCostBase = detectConvCostBase(convWb, monthObj);
    }
    if (convCostBase === null || convCostBase === undefined) {
      throw new Error(
        `Couldn't auto-detect the GP conversion cost for ${monthLabel(monthObj)}. ` +
          "Open Advanced overrides and enter it manually."
      );
    }
    log(`Conv Cost base: ${convCostBase}`);

    const registerFiles = Array.from(registerInput.files || []);
    if (registerFiles.length) {
      log(`Third-party register files: ${registerFiles.map((f) => f.name).join(", ")}`);
    }

    log(`Found ${rawRows.length} rows to convert...`);
    const outWb = await buildOutput(rawRows, templateWb, convCostBase, monthObj, registerFiles, log);

    const filename =
      outputNameInput.value.trim() ||
      `MRM-${monthLabel(monthObj)}.xlsx`;

    log(`Preparing download: ${filename}`);
    await downloadWorkbook(outWb, filename);
    log("Done.", "success");
  } catch (err) {
    console.error(err);
    log("Error: " + err.message, "error");
  } finally {
    convertBtn.disabled = false;
    convertBtn.textContent = "Convert";
  }
});
