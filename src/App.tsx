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

type InputKey =
  | "jobId"
  | "firstName"
  | "lastName"
  | "phone"
  | "address"
  | "suburb"
  | "postcode"
  | "state"
  | "installDate"
  | "product"
  | "productModel"
  | "serial"
  | "questionCol";

const INPUT_FIELDS: {
  key: InputKey;
  label: string;
  candidates: string[];
}[] = [
  { key: "jobId", label: "Job ID", candidates: ["jobid", "job"] },
  { key: "firstName", label: "First Name", candidates: ["firstname", "first"] },
  { key: "lastName", label: "Last Name", candidates: ["lastname", "last"] },
  { key: "phone", label: "Phone", candidates: ["phone", "mobile"] },
  { key: "address", label: "Address", candidates: ["address"] },
  { key: "suburb", label: "Suburb", candidates: ["suburb"] },
  { key: "postcode", label: "Postcode", candidates: ["postcode", "zip"] },
  { key: "state", label: "State", candidates: ["state"] },
  { key: "installDate", label: "Install Date", candidates: ["date", "install"] },
  { key: "product", label: "Product", candidates: ["product", "brand"] },
  { key: "productModel", label: "Model", candidates: ["model"] },
  { key: "serial", label: "Serial", candidates: ["serial"] },
  { key: "questionCol", label: "Type Column", candidates: ["type", "question"] },
];

const norm = (s: string) => s.toLowerCase().replace(/[\s_-]+/g, "");

function findColumn(headers: string[], candidates: string[]) {
  const targets = candidates.map(norm);
  return headers.find((h) => targets.includes(norm(h))) || "";
}

function autoDetect(headers: string[]) {
  const out = {} as Record<InputKey, string>;
  for (const f of INPUT_FIELDS) {
    out[f.key] = findColumn(headers, f.candidates);
  }
  return out;
}

function formatCell(v: unknown) {
  if (!v) return "";
  if (v instanceof Date) return v.toISOString().split("T")[0];
  return String(v);
}

function App() {
  const fileRef = useRef<HTMLInputElement | null>(null);

  const [rows, setRows] = useState<Row[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState("");
  const [mapping, setMapping] = useState<Record<InputKey, string>>(
    autoDetect([])
  );

  const handleFile = async (file: File) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json<Row>(sheet, { defval: "" });

    const cols = Object.keys(json[0] || {});
    setHeaders(cols);
    setRows(json);
    setFileName(file.name);
    setMapping(autoDetect(cols));
  };

  const jobGroups = useMemo(() => {
    const col = mapping.jobId;
    if (!col) return new Map<string, Row[]>();

    const map = new Map<string, Row[]>();

    for (const r of rows) {
      const id = String(r[col] || "");
      if (!id) continue;
      if (!map.has(id)) map.set(id, []);
      map.get(id)!.push(r);
    }

    return map;
  }, [rows, mapping]);

  const output = useMemo(() => {
    const result: any[] = [];

    jobGroups.forEach((jobRows, jobId) => {
      const first = jobRows[0];

      result.push({
        "Job ID": jobId,
        "Customer Name": `${first[mapping.firstName] || ""} ${
          first[mapping.lastName] || ""
        }`,
        "Customer Phone": first[mapping.phone] || "",
        Address: `${first[mapping.address] || ""} ${first[mapping.suburb] || ""}`,
        "Installation Date": first[mapping.installDate] || "",
        CoES: jobRows.map((r) => JSON.stringify(r)).join("\n"),
      });
    });

    return result;
  }, [jobGroups, mapping]);

  const download = () => {
    const ws = XLSX.utils.json_to_sheet(output);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Jobs");
    XLSX.writeFile(wb, "output.xlsx");
  };

  return (
    <div style={{ padding: 20, fontFamily: "sans-serif" }}>
      <h1>Excel Merger</h1>

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
          <p>Loaded: {fileName}</p>
          <p>Rows: {rows.length}</p>

          <button onClick={download}>Download Excel</button>

          <h3>Preview</h3>
          <pre style={{ fontSize: 12 }}>
            {JSON.stringify(output.slice(0, 3), null, 2)}
          </pre>
        </>
      )}
    </div>
  );
}

export default App;
