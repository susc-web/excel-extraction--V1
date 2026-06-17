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
    return v.toISOString().split("T")[0];
  }
  return String(v);
}

function formatPhone(v: unknown) {
  if (!v) return "";
  let s = String(v).replace(/[\s\-().+]/g, "");
  if (s.startsWith("61") && s.length === 11) s = "0" + s.slice(2);
  if (/^\d{9}$/.test(s)) s = "0" + s;
  return s;
}

function App() {
  const fileRef = useRef<HTMLInputElement | null>(null);

  const [rows, setRows] = useState<Row[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState("");

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

  // GROUP BY JOB ID
  const jobGroups = useMemo(() => {
    const map = new Map<string, Row[]>();

    for (const r of rows) {
      const id = String(r["Job ID"] || r["job id"] || r["jobid"] || "");
      if (!id) continue;

      if (!map.has(id)) map.set(id, []);
      map.get(id)!.push(r);
    }

    return map;
  }, [rows]);

  // BUILD FINAL OUTPUT (THIS IS YOUR CORE LOGIC)
  const output = useMemo(() => {
    const result: any[] = [];

    jobGroups.forEach((jobRows, jobId) => {
      const first = jobRows[0] || {};

      const firstName = first["First Name"] || first["firstname"] || "";
      const lastName = first["Last Name"] || first["lastname"] || "";
      const phone = first["Phone"] || first["mobile"] || "";
      const address = first["Address"] || "";
      const suburb = first["Suburb"] || "";
      const postcode = first["Postcode"] || "";
      const state = first["State"] || "";
      const install = first["Install Date"] || first["Installation Date"] || "";

      // COES LOGIC (KEEP FULL DETAIL)
      const coes = jobRows
        .map((r) => {
          const product = r["Product"] || "";
          const model = r["Model"] || "";
          const serial = r["Serial"] || "";
          const type = String(r["Type"] || "").toLowerCase();

          return `${product} ${model}\n${type.includes("outdoor") ? "Outdoor" : "Indoor"}: ${serial}`;
        })
        .join("\n");

      result.push({
        "Job ID": jobId,
        "Customer Name": `${firstName} ${lastName}`.trim(),
        "Customer Phone": formatPhone(phone),
        "Address": `${address}, ${suburb}, ${postcode}, ${state}`,
        "Installation Date": install,
        CoES: coes,
      });
    });

    return result;
  }, [jobGroups]);

  const download = () => {
    const ws = XLSX.utils.json_to_sheet(output, {
      header: [...OUTPUT_COLUMNS],
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Jobs");
    XLSX.writeFile(wb, "formatted_output.xlsx");
  };

  return (
    <div style={{ padding: 20, fontFamily: "sans-serif" }}>
      <h1>Excel Job Formatter</h1>

      <input
        type="file"
        ref={fileRef}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleFile(f);
        }}
      />

      {fileName && (
        <>
          <p>File: {fileName}</p>
          <p>Total Rows: {rows.length}</p>
          <p>Jobs Found: {jobGroups.size}</p>

          <button
            onClick={download}
            style={{
              marginTop: 10,
              padding: "10px 15px",
              background: "green",
              color: "white",
              borderRadius: 6,
              border: 0,
            }}
          >
            Download Excel
          </button>

          <h3>Preview Output</h3>
          <pre style={{ fontSize: 12, background: "#f5f5f5", padding: 10 }}>
            {JSON.stringify(output.slice(0, 2), null, 2)}
          </pre>
        </>
      )}
    </div>
  );
}

export default App;
