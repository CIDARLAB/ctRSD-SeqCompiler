const worker = new Worker("./js/worker.js", { type: "module" });
const status = document.querySelector("#status");
const result = document.querySelector("#result");

window.onload = function() {
  hidePageBtns();
};

worker.onmessage = async ({ data }) => {
  if (data.type === "ready") {
    status.textContent = "Ready To Compile!";
    showPageBtns();
  
  } else if (data.type === "result") {
    await renderResults(data.result, data.compileOption, data.name);
    status.textContent = "Complete";
  
  } else if (data.type === "error") {
    status.textContent = `Error: ${data.message}`;
  }
};

worker.onerror = (event) => {
  status.textContent = `Worker error: ${event.message}`;
};

async function renderResults(data, option, singleName) {
  result.replaceChildren();

  const details = document.createElement("details");
  const summary = document.createElement("summary");
  summary.className = "btn btn-primary full-result-toggle";
  summary.textContent = "See Full result";
  const raw = document.createElement("pre");
  raw.textContent = JSON.stringify(data, null, 2);
  const copyJson = document.createElement("button");
  copyJson.type = "button";
  copyJson.className = "btn btn-primary";
  copyJson.textContent = "Copy JSON";
  copyJson.onclick = async () => {
    try {
      await navigator.clipboard.writeText(raw.textContent);
      status.textContent = "JSON copied.";
    } catch (error) {
      status.textContent = `Copy failed: ${error.message}`;
    }
  };

  details.append(summary, copyJson, raw);
  
  result.appendChild(details);

  function addComponent(name, dna, rna) {
    const card = document.createElement("article");
    const heading = document.createElement("h2");
    heading.textContent = name;
    card.appendChild(heading);

    for (const [label, sequence] of [["DNA", dna], ["RNA", rna]]) {
      const title = document.createElement("h3");
      title.textContent = label;

      const text = document.createElement("pre");
      text.textContent = sequence;

      const copy = document.createElement("button");
      copy.type = "button";
      copy.textContent = `Copy ${label}`;
      copy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(sequence);
          status.textContent = `${label} copied.`;
        } catch (error) {
          status.textContent = `Copy failed: ${error.message}`;
        }
      };

      card.append(title, text, copy);
    }
    result.appendChild(card);
  }

  if (option === "workbook") {
    const exportButton = document.createElement("button");
    exportButton.type = "button";
    exportButton.className = "btn btn-primary";
    exportButton.textContent = "Download Sequences CSV";

    exportButton.onclick = () => {
      try {
        downloadSequencesCsv(data);
      } catch (error) {
        status.textContent = `CSV export error: ${error.message}`;
      }
    };

    result.appendChild(exportButton);
  }

  const genbankButton = document.createElement("button");
  genbankButton.type = "button";
  genbankButton.className = "btn btn-primary";
  genbankButton.textContent = "Download GenBank";

  genbankButton.onclick = async () => {
    try {
      if (option === "workbook") {
        await downloadGenbankFilesZip(data);
      } else {
        const text = data.rna === 1
          ? data.genbank_rna
          : data.genbank_dna;

        if (typeof text !== "string" || !text) {
          throw new Error("No GenBank file was returned.");
        }
        downloadGenbankFile(text, singleName);
      }
    } catch (error) {
      status.textContent = `Download error: ${error.message}`;
    }
  };

  result.appendChild(genbankButton);

  if (option === "workbook") {
    data.name.forEach((names, row) => {
      names.forEach((name, item) => {
        addComponent(name, data.dna_template[row][item], data.rna_template[row][item]);
      });
    });
  } else {
    addComponent(singleName, data.dna_template, data.rna_template);
  }
}

const BtnIDs = {
  seqCompiler: "#seqcompiler-btn"
};

let seqCompilerBtn = $('#seqcompiler-btn');
seqCompilerBtn.tooltipster({
    content: $('#seqcompiler-tooltip'),
    side: 'top',
    interactive: true,
    theme: 'tooltipster-noir'
});

function bindTooltipToButton(buttonSelector, tooltipSelector) {
    $(tooltipSelector).trigger('click');

    $(buttonSelector).on('click', function (e) {
    e.preventDefault();
    e.stopPropagation();
    $(tooltipSelector).trigger('click');
    });
}

bindTooltipToButton(BtnIDs.seqCompiler, '#seqcompiler-tooltip');
$('#seqcompiler-tooltip').click(() => {
    let div = document.createElement('div');
    div.style.maxHeight = "60vh";
    div.style.overflowY = "auto";
    div.style.overflowX = "hidden";
    div.style.paddingRight = "8px";

    // Compile Single Component vs Workbook
    let compileOptionDiv = document.createElement('div');
    let compileOptionDropdown = makeCompileOptionDropdown();
    makeDiv(compileOptionDiv, compileOptionDropdown, 'Compile Option: ');

    // Domains Workbook div
    let excelWorkbookDiv = document.createElement('div');
    let excelWorkbookInput = document.createElement('input');
    excelWorkbookInput.type = 'file';
    excelWorkbookInput.accept = '.xls';
    makeDiv(excelWorkbookDiv, excelWorkbookInput, 'Domains Workbook (.xls - Optional): ', "Path to the Domains workbook.");
    const downloadDomains = document.createElement("a");
    downloadDomains.href = "./data/ctRSD_domains_list_300.xls";
    downloadDomains.download = "ctRSD_domains_list_300.xls";
    downloadDomains.textContent = "Download default domains";
    excelWorkbookDiv.appendChild(downloadDomains);

    // Components Workbook div
    let componentsWorkbookDiv = document.createElement('div');
    let componentsWorkbookInput = document.createElement('input');
    componentsWorkbookInput.type = 'file';
    componentsWorkbookInput.accept = '.xlsx';
    makeDiv(componentsWorkbookDiv, componentsWorkbookInput, 'Components Workbook (.xlsx - Required): ', "Path to the Components workbook.");
    const downloadComponents = document.createElement("a");
    downloadComponents.href = "./data/ctRSD_SeqCompile_template.xlsx";
    downloadComponents.download = "ctRSD_SeqCompile_template.xlsx";
    downloadComponents.textContent = "Download template";
    componentsWorkbookDiv.appendChild(downloadComponents);

    // Name Div
    let nameDiv = document.createElement('div');
    let nameInput = document.createElement('input');
    makeDiv(nameDiv, nameInput, 'Name: ');

    // Rz Div
    let RzDiv = document.createElement('div');
    let RzInput = document.createElement('input');
    RzInput.setAttribute("value", "Ro");
    makeDiv(RzDiv, RzInput, 'Rz: ');

    // L Div
    let LDiv = document.createElement('div');
    let LInput = document.createElement('input');
    LInput.setAttribute("value", "L");
    makeDiv(LDiv, LInput, 'L: ');

    // term Div
    let termDiv = document.createElement('div');
    let termInput = document.createElement('input');
    termInput.setAttribute("value", "T7t");
    makeDiv(termDiv, termInput, 'term: ');

    // hp5 Div
    let hp5Div = document.createElement('div');
    let hp5Input = document.createElement('input');
    hp5Input.setAttribute("value", "5hp");
    makeDiv(hp5Div, hp5Input, 'hp5: ');

    // prom Div
    let promDiv = document.createElement('div');
    let promInput = document.createElement('input');
    promInput.setAttribute("value", "T7p");
    makeDiv(promDiv, promInput, 'prom: ');

    // eI Div
    let eIDiv = document.createElement('div');
    let eIInput = document.createElement('input');
    makeDiv(eIDiv, eIInput, 'eI: ');

    // eO Div
    let eODiv = document.createElement('div');
    let eOInput = document.createElement('input');
    makeDiv(eODiv, eOInput, 'eO: ');

    // s Div
    let sDiv = document.createElement('div');
    let sInput = document.createElement('input');
    makeDiv(sDiv, sInput, 's: ');

    // invert Div
    let invertDiv = document.createElement('div');
    let invertInput = document.createElement('input');
    invertInput.setAttribute("type", "number");
    invertInput.setAttribute("step", "1");
    invertInput.setAttribute("value", "0");
    invertInput.setAttribute("min", "0");
    invertInput.setAttribute("max", "1");
    makeDiv(invertDiv, invertInput, 'invert: ');

    // invL Div
    let invLDiv = document.createElement('div');
    let invLInput = document.createElement('input');
    invLInput.setAttribute("value", "A");
    makeDiv(invLDiv, invLInput, 'invL: ');

    // agL Div
    let agLDiv = document.createElement('div');
    let agLInput = document.createElement('input');
    agLInput.setAttribute("value", "TA");
    makeDiv(agLDiv, agLInput, 'agL: ');

    // AGiloop Div
    let AGiloopDiv = document.createElement('div');
    let AGiloopInput = document.createElement('input');
    AGiloopInput.setAttribute("type", "number");
    AGiloopInput.setAttribute("step", "1");
    AGiloopInput.setAttribute("value", "5");
    makeDiv(AGiloopDiv, AGiloopInput, 'AGiloop: ');

    // otype Div
    let otypeDiv = document.createElement('div');
    let otypeInput = document.createElement('input');
    otypeInput.setAttribute("type", "number");
    otypeInput.setAttribute("step", "1");
    otypeInput.setAttribute("value", "1");
    makeDiv(otypeDiv, otypeInput, 'otype: ');

    // rna Div
    let rnaDiv = document.createElement('div');
    let rnaInput = document.createElement('input');
    rnaInput.setAttribute("type", "number");
    rnaInput.setAttribute("step", "1");
    rnaInput.setAttribute("value", "0");
    rnaInput.setAttribute("min", "0");
    rnaInput.setAttribute("max", "1");
    makeDiv(rnaDiv, rnaInput, 'rna: ');

    // us Div
    let usDiv = document.createElement('div');
    let usInput = document.createElement('input');
    makeDiv(usDiv, usInput, 'us (comma-separated): ');

    // ds Div
    let dsDiv = document.createElement('div');
    let dsInput = document.createElement('input');
    makeDiv(dsDiv, dsInput, 'ds (comma-separated): ');

    // temp_len Div
    let temp_lenDiv = document.createElement('div');
    let temp_lenInput = document.createElement('input');
    temp_lenInput.setAttribute("type", "number");
    temp_lenInput.setAttribute("step", "1");
    temp_lenInput.setAttribute("value", "0");
    makeDiv(temp_lenDiv, temp_lenInput, 'temp_len: ');

    // cp Div
    let cpDiv = document.createElement('div'); 
    let cpInput = document.createElement('input');
    makeDiv(cpDiv, cpInput, 'cp: ');

    // n Div
    let nDiv = document.createElement('div');
    let nInput = document.createElement('input');
    makeDiv(nDiv, nInput, 'n: ');

    // c Div
    let cDiv = document.createElement('div');
    let cInput = document.createElement('input');
    cInput.setAttribute("type", "number");
    cInput.setAttribute("step", "1");
    cInput.setAttribute("value", "0");
    cInput.setAttribute("min", "0");
    makeDiv(cDiv, cInput, 'c: ');

    // d Div
    let dDiv = document.createElement('div');
    let dInput = document.createElement('input');
    makeDiv(dDiv, dInput, 'd: ');

    // CDS Div
    let CDSDiv = document.createElement('div');
    let CDSInput = document.createElement('input');
    makeDiv(CDSDiv, CDSInput, 'CDS: ');

    // rflap Div
    let rflapDiv = document.createElement('div');
    let rflapInput = document.createElement('input');
    makeDiv(rflapDiv, rflapInput, 'rflap: ');

    const exampleButton = document.createElement("button");
    exampleButton.type = "button";
    exampleButton.className = "btn btn-default";
    exampleButton.textContent = "Fill Example";

    exampleButton.addEventListener("click", () => {
        nameInput.value = "G{u1,rbs1-2r}";
        RzInput.value = "R3";
        LInput.value = "L";
        termInput.value = "T7t0";
        hp5Input.value = "5hp";
        promInput.value = "T7p";
        eIInput.value = "";
        eOInput.value = "";
        sInput.value = "s4";
        invertInput.value = "0";
        invLInput.value = "A";
        agLInput.value = "TA";
        AGiloopInput.value = "5";
        otypeInput.value = "1";
        rnaInput.value = "0";
        usInput.value = "";
        dsInput.value = "";
        temp_lenInput.value = "0";
        cpInput.value = "c06";
        nInput.value = "THE_L";
        cInput.value = "0";
        dInput.value = "";
        CDSInput.value = "sCFP3A_tta";
        rflapInput.value = "";
    });

    // Append all divs to the main div
    div.appendChild(compileOptionDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(document.createElement('br'));
    div.appendChild(exampleButton);
    div.appendChild(document.createElement('br'));
    div.appendChild(document.createElement('br'));
    div.appendChild(excelWorkbookDiv);
    div.appendChild(document.createElement("br"));
    div.appendChild(document.createElement("br"));
    div.appendChild(componentsWorkbookDiv);
    div.appendChild(document.createElement("br"));
    div.appendChild(document.createElement("br"));
    div.appendChild(nameDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(RzDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(LDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(termDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(hp5Div);
    div.appendChild(document.createElement('br'));
    div.appendChild(promDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(eIDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(eODiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(sDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(invertDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(invLDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(agLDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(AGiloopDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(otypeDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(rnaDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(usDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(dsDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(temp_lenDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(cpDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(nDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(cDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(dDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(CDSDiv);
    div.appendChild(document.createElement('br'));
    div.appendChild(rflapDiv);
    div.appendChild(document.createElement('br'));

    setRowVisible(componentsWorkbookDiv, false);

    $(compileOptionDropdown).change(function() {
      if (this.value === "single_component"){
        // Show single component related fields
        setRowVisible(exampleButton, true)
        setRowVisible(nameDiv, true)
        setRowVisible(RzDiv, true)
        setRowVisible(LDiv, true)
        setRowVisible(termDiv, true)
        setRowVisible(hp5Div, true)
        setRowVisible(promDiv, true)
        setRowVisible(eIDiv, true)
        setRowVisible(eODiv, true)
        setRowVisible(sDiv, true)
        setRowVisible(invertDiv, true)
        setRowVisible(invLDiv, true)
        setRowVisible(agLDiv, true)
        setRowVisible(AGiloopDiv, true)
        setRowVisible(otypeDiv, true)
        setRowVisible(rnaDiv, true)
        setRowVisible(usDiv, true)
        setRowVisible(dsDiv, true)
        setRowVisible(temp_lenDiv, true)
        setRowVisible(cpDiv, true)
        setRowVisible(nDiv, true)
        setRowVisible(cDiv, true)
        setRowVisible(dDiv, true)
        setRowVisible(CDSDiv, true)
        setRowVisible(rflapDiv, true)
        setRowVisible(componentsWorkbookDiv, false);
      } else if (this.value === "workbook"){
        // Show workbook related fields
        setRowVisible(exampleButton, false)
        setRowVisible(nameDiv, false)
        setRowVisible(RzDiv, false)
        setRowVisible(LDiv, false)
        setRowVisible(termDiv, false)
        setRowVisible(hp5Div, false)
        setRowVisible(promDiv, false)
        setRowVisible(eIDiv, false)
        setRowVisible(eODiv, false)
        setRowVisible(sDiv, false)
        setRowVisible(invertDiv, false)
        setRowVisible(invLDiv, false)
        setRowVisible(agLDiv, false)
        setRowVisible(AGiloopDiv, false)
        setRowVisible(otypeDiv, false)
        setRowVisible(rnaDiv, false)
        setRowVisible(usDiv, false)
        setRowVisible(dsDiv, false)
        setRowVisible(temp_lenDiv, false)
        setRowVisible(cpDiv, false)
        setRowVisible(nDiv, false)
        setRowVisible(cDiv, false)
        setRowVisible(dDiv, false)
        setRowVisible(CDSDiv, false)
        setRowVisible(rflapDiv, false)
        setRowVisible(componentsWorkbookDiv, true)
      }
    });

    swal({
        title: "ctRSD Sequence Compiler",
        buttons: true,
        content: div
    }).then(async (confirm) => {
        if (!confirm) return;

        const compileOption = compileOptionDropdown.value;
        const name = nameInput.value;
        const Rz = RzInput.value;
        const L = LInput.value;
        const term = termInput.value;
        const hp5 = hp5Input.value;
        const prom = promInput.value;
        const eI = eIInput.value.trim();
        const eO = eOInput.value.trim();
        const s = sInput.value.trim();
        const invert = Number(invertInput.value);
        const invL = invLInput.value;
        const agL = agLInput.value;
        const AGiloop = Number(AGiloopInput.value);
        const otype = Number(otypeInput.value);
        const rna = Number(rnaInput.value);
        const us = (usInput.value.trim() !== "") ? usInput.value.split(',').map(s => s.trim()).filter(s => s.length > 0) : [];
        const ds = (dsInput.value.trim() !== "") ? dsInput.value.split(',').map(s => s.trim()).filter(s => s.length > 0) : [];
        const temp_len = Number(temp_lenInput.value);
        const cp = cpInput.value.trim();
        const n = nInput.value.trim();
        const c = Number(cInput.value);
        const d = dInput.value.trim();
        const CDS = CDSInput.value.trim();
        const rflap = rflapInput.value.trim();


        if (compileOption === "single_component" && !name) {
            swalError("Name is required."); 
            return; 
        }

        if (compileOption === "workbook" && !componentsWorkbookInput.files[0]) {
            swalError("Please select a component workbook file."); 
            return; 
        }

        // Call
        try {
            const file = excelWorkbookInput.files[0];
            let domains;

            const components = componentsWorkbookInput.files[0];

            if (file) {
                domains = await file.arrayBuffer();
            } else {
                const response = await fetch("./data/ctRSD_domains_list_300.xls");
                if (!response.ok) {
                throw new Error(`Could not load default domains: ${response.status}`);
                }
                domains = await response.arrayBuffer();
            }
            const message = {
              compileOption, name, domains,
              Rz, L, term, hp5, prom, eI, eO, s, invert, invL, agL,
              AGiloop, otype, rna, us, ds, temp_len, cp, n, c, d, CDS, rflap
            };
            const transfers = [domains];

            if (compileOption === "workbook") {
              message.components = await componentsWorkbookInput.files[0].arrayBuffer();
              transfers.push(message.components);
            }

            worker.postMessage(message, transfers);

        } catch (error) {
            swalError(`Could not load domains: ${error.message}`);
        }

    });

});

function makeDiv(div, input, title, infoText = ""){
  div.appendChild(document.createTextNode(title));
  div.appendChild(input);

  if (infoText) {
    const infoBtn = document.createElement("button");
    infoBtn.type = "button";
    infoBtn.className = "btn btn-link btn-xs";
    infoBtn.innerHTML = '<i class="fa fa-info-circle" aria-hidden="true"></i>';
    infoBtn.setAttribute("aria-label", "More info");
    infoBtn.style.padding = "0 6px";
    infoBtn.style.verticalAlign = "middle";

    div.appendChild(infoBtn);

    $(infoBtn).tooltipster({
      content: infoText,
      side: "right",
      theme: "tooltipster-noir",
      trigger: "hover",
      interactive: true,
      delay: 100
    });
  }
}

function makeCompileOptionDropdown(){
  let compileOptionDropdown = document.createElement('select');
  compileOptionDropdown.setAttribute("id", "compile-option-dropdown");
  compileOptionDropdown.appendChild(new Option("Single Component", "single_component", true, true));
  compileOptionDropdown.appendChild(new Option("Workbook", "workbook"));
  return compileOptionDropdown;
}

function setRowVisible(rowDiv, show) {
  rowDiv.style.display = show ? "" : "none";

  // Also hide/show the spacer <br> inserted right after the row
  const spacer = rowDiv.nextElementSibling;
  if (spacer && spacer.tagName === "BR") {
    spacer.style.display = show ? "" : "none";
  }
}

function swalError(errorMsg){
  swal({
    title: "Error!",
    text: errorMsg,
    icon: "error"
  });
}

function swalSuccess(msg = ""){
  swal({
    title: "Success!",
    text: msg,
    icon: "success"
  });
}

function downloadGenbankFile(genbankText, name) {
    const blob = new Blob([genbankText], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = (name || "compiled_sequence") + ".gbk";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

async function downloadGenbankFilesZip(result) {
    const zip = new JSZip();
    let count = 0;

    for (let row = 0; row < result.name.length; row++) {
      for (let item = 0; item < result.name[row].length; item++) {
        const name = result.name[row][item].replace(/[^\w.-]/g, "_");
        const prefix = `${row + 1}_${item + 1}_${name}`;

        for (const [type, records] of [
          ["DNA", result.genbank_dna],
          ["RNA", result.genbank_rna]
        ]) {
          const text = records?.[row]?.[item];
          if (typeof text !== "string" || !text) {
            throw new Error(`Missing ${type} GenBank data for ${name}`);
          }
          zip.file(`${prefix}_${type}.gbk`, text);
          count++;
        }
      }
    }

    if (!count) throw new Error("No GenBank files were returned.");

    const blob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "compiled_sequences.zip";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadSequencesCsv(data) {
  const rows = [["name", "DNA", "RNA"]];

  for (let row = 0; row < data.name.length; row++) {
    for (let item = 0; item < data.name[row].length; item++) {
      const name = data.name[row][item];
      const dna = data.dna_template?.[row]?.[item];
      const rna = data.rna_template?.[row]?.[item];

      if (typeof name !== "string" || typeof dna !== "string" || typeof rna !== "string") {
        throw new Error(`Missing sequence data for workbook row ${row + 1}, item ${item + 1}.`);
      }

      rows.push([name, dna, rna]);
    }
  }

  const csv = "\uFEFF" + rows
    .map(row => row.map(value => `"${value.replace(/"/g, '""')}"`).join(","))
    .join("\r\n");
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = "compiled_sequences.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function hidePageBtns() {
  Object.keys(BtnIDs).map((id, _) => {
    $(BtnIDs[id]).hide();
  });
}

function showPageBtns() {
  Object.keys(BtnIDs).map((id, _) => {
    $(BtnIDs[id]).show();
  });
}

