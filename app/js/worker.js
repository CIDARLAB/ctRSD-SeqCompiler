import { loadPyodide } from "https://cdn.jsdelivr.net/pyodide/v314.0.8/full/pyodide.mjs";

async function initialize() {
  const pyodide = await loadPyodide();
  await pyodide.loadPackage(["xlrd", "micropip"]);
  
  await pyodide.runPythonAsync(`

import micropip
await micropip.install("openpyxl")
  
  `);

  const response = await fetch(new URL("../scripts/ctRSD_simulator_210.py", import.meta.url), { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Could not load Python file: ${response.status}`);
  }

  pyodide.FS.writeFile("ctRSD_simulator_210.py", await response.text());
  pyodide.runPython("from ctRSD_simulator_210 import RSD_sim");
  pyodide.runPython("from ctRSD_simulator_210 import process_workbook");
  pyodide.runPython("from ctRSD_simulator_210 import compile_workbook_data");

  self.postMessage({ type: "ready" });
  return pyodide;
}

const ready = initialize();

self.onmessage = async (event) => {
    try {
        const pyodide = await ready;
        const { compileOption, name, domains, components, Rz, L, term, hp5, prom, eI, eO, s, invert, invL, agL, AGiloop, otype, rna, us, ds, temp_len, cp, n, c, d, CDS, rflap } = event.data;

        pyodide.globals.set("compileOption", compileOption);
        pyodide.FS.writeFile("domains.xls", new Uint8Array(domains));
        if (compileOption === "workbook") {
            pyodide.FS.writeFile("components.xlsx", new Uint8Array(components));
        }
        pyodide.globals.set("name", name);
        pyodide.globals.set("Rz", Rz);
        pyodide.globals.set("L", L);
        pyodide.globals.set("term", term);
        pyodide.globals.set("hp5", hp5);
        pyodide.globals.set("prom", prom);
        pyodide.globals.set("eI", eI);
        pyodide.globals.set("eO", eO);
        pyodide.globals.set("s", s);
        pyodide.globals.set("invert", invert);
        pyodide.globals.set("invL", invL);
        pyodide.globals.set("agL", agL);
        pyodide.globals.set("AGiloop", AGiloop);
        pyodide.globals.set("otype", otype);
        pyodide.globals.set("rna", rna);
        pyodide.globals.set("us", us);
        pyodide.globals.set("ds", ds);
        pyodide.globals.set("temp_len", temp_len);
        pyodide.globals.set("cp", cp);
        pyodide.globals.set("n", n);
        pyodide.globals.set("c", c);
        pyodide.globals.set("d", d);
        pyodide.globals.set("CDS", CDS);
        pyodide.globals.set("rflap", rflap);

        const result = pyodide.runPython(`

import json

if compileOption == "single_component":
    model = RSD_sim()
    model.load_domains("domains.xls")
    values = model.ctRSD_seq_compile(
        name=name,
        Rz=Rz,
        L=L,
        term=term,
        hp5=hp5,
        prom=prom,
        eI=eI,
        eO=eO,
        s=s,
        invert=invert,
        invL=invL,
        agL=agL,
        AGiloop=AGiloop,
        otype=otype,
        rna=rna,
        us=us,
        ds=ds,
        temp_len=temp_len,
        cp=cp,
        n=n,
        c=c,
        d=d,
        CDS=CDS,
        rflap=rflap
    )

    headers = [
        "dna_template", "rna_template",
        "dna_part_IDs", "rna_part_IDs",
        "dna_part_types", "rna_part_types",
        "dna_part_orientations", "rna_part_orientations",
        "genbank_dna", "genbank_rna",
    ]

elif compileOption == "workbook":
    data, num_rows = process_workbook("components.xlsx")
    results = compile_workbook_data(data, num_rows, domains_file="domains.xls")
    headers = results.keys()
    values = results.values()

else:
    headers = ["error"]
    values = [f"Invalid compile option: {compileOption}"]

json.dumps(dict(zip(headers, values)))
        
        `);

        self.postMessage({ type: "result", result: JSON.parse(result), compileOption, name, rna });
    } catch (error) {
        self.postMessage({ type: "error", message: String(error) });
    }
};

ready.catch((error) => {
  self.postMessage({ type: "error", message: String(error) });
});

