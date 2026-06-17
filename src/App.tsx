import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";

type Row = Record<string, unknown>;

const OUTPUT_COLUMNS = [
  "Job ID",
  "Customer Name",
  "Customer Phone",
  "Address",
  "Installation Date",
  "CoES",
] as const;

type OutputCol = (typeof OUTPUT_COLUMNS)[number];

const norm = (s: string) =>
  String(s).toLowerCase().replace(/[\s_-]+/g, "");

function formatCell(v: unknown) {
  if (v === undefined || v === null || v === "") return "";
  if (v instanceof Date) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(
      v.getDate()
    ).padStart(2, "0")}`;
  }
  return String(v);
}

function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [rows, setRows] = useState<Row[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);

  const handleFile = async (file: File) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });

    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json<Row>(sheet, { defval: "" });

    const cols = Object.keys(json[0] || {});
    setHeaders(cols);
    setRows(json);
    setFileName(file.name);
  };

  // ✅ GROUP JOBS (FIXED DETECTION)
  const jobGroups = useMemo(() => {
    const map = new Map<string, Row[]>();

    for (const r of rows) {
      const id =
        String(
          r["Job ID"] ??
          r["JobId"] ??
          r["jobid"] ??
          r["job id"] ??
          ""
        ).trim();

      if (!id) continue;

      if (!map.has(id)) map.set(id, []);
      map.get(id)!.push(r);
    }

    return map;
  }, [rows]);

  // ✅ FINAL OUTPUT (YOUR CORE LOGIC KEPT)
  const output = useMemo(() => {
    const result: any[] = [];

    jobGroups.forEach((jobRows, jobId) => {
      const first = jobRows[0] || {};

      const get = (keys: string[]) => {
        for (const k of keys) {
          if (first[k]) return first[k];
        }
        return "";
      };

      const firstName = get(["First Name", "firstname"]);
      const lastName = get(["Last Name", "lastname"]);
      const phone = get(["Phone", "Mobile"]);
      const address = get(["Address"]);
      const suburb = get(["Suburb"]);
      const install = get(["Install Date"]);

      const coes = jobRows
        .map((r) => {
          const product = r["Product"] || "";
          const model = r["Model"] || "";
          const serial = r["Serial"] || "";
          const type = String(r["Type"] || "").toLowerCase();

          return `${product} ${model}\n${
            type.includes("outdoor") ? "Outdoor" : "Indoor"
          }: ${serial}`;
        })
        .join("\n");

      result.push({
        "Job ID": jobId,
        "Customer Name": `${firstName} ${lastName}`.trim(),
        "Customer Phone": phone,
        Address: `${address} ${suburb}`,
        "Installation Date": install,
        CoES: coes,
      });
    });

    return result;
  }, [jobGroups]);

  const download = () => {
    const ws = XLSX.utils.json_to_sheet(output);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Jobs");
    XLSX.writeFile(wb, "formatted_output.xlsx");
  };

  return (
    <div style={{ padding: 20, fontFamily: "sans-serif" }}>
      <h1>Excel Job Formatter</h1>

      <input
        type="file"
        ref={fileInputRef}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleFile(f);
        }}
      />

      {fileName && (
        <div style={{ marginTop: 20 }}>
          <p>File: {fileName}</p>
          <p>Rows: {rows.length}</p>
          <p>Jobs: {jobGroups.size}</p>

          <button
            onClick={download}
            style={{
              padding: "10px 15px",
              background: "green",
              color: "white",
              borderRadius: 6,
              border: 0,
              marginTop: 10,
            }}
          >
            Download Excel
          </button>

          <h3 style={{ marginTop: 20 }}>Preview</h3>
          <pre style={{ fontSize: 12, background: "#f5f5f5", padding: 10 }}>
            {JSON.stringify(output.slice(0, 3), null, 2)}
          </pre>
        </div>
      )}
    </div>
  );
}

export default App;
