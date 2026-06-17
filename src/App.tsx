import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import {
  Button,
  Card,
  CardContent,
  CardHeader,
  CardTitle,
  CardDescription,
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
  Badge,
  Toaster,
} from "@/components/ui";
import { useToast } from "@/hooks/use-toast";
import { FileSpreadsheet, Upload, X, ListChecks } from "lucide-react";

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

const INPUT_FIELDS: { key: InputKey; label: string; candidates: string[] }[] =
  [
    { key: "jobId", label: "Job ID", candidates: ["jobid", "job id", "job"] },
    { key: "firstName", label: "First Name", candidates: ["firstname"] },
    { key: "lastName", label: "Last Name", candidates: ["lastname"] },
    { key: "phone", label: "Phone", candidates: ["phone", "mobile"] },
    { key: "address", label: "Address", candidates: ["address"] },
    { key: "suburb", label: "Suburb", candidates: ["suburb"] },
    { key: "postcode", label: "Postcode", candidates: ["postcode", "zip"] },
    { key: "state", label: "State", candidates: ["state"] },
    {
      key: "installDate",
      label: "Install Date",
      candidates: ["date", "install"],
    },
    { key: "product", label: "Product", candidates: ["product", "brand"] },
    { key: "productModel", label: "Model", candidates: ["model"] },
    { key: "serial", label: "Serial", candidates: ["serial"] },
    {
      key: "questionCol",
      label: "Type Column",
      candidates: ["type", "question"],
    },
  ];

const norm = (s: string) =>
  String(s).toLowerCase().replace(/[^a-z0-9]/g, "");

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
  if (v instanceof Date)
    return v.toISOString().split("T")[0];
  return String(v);
}

function App() {
  const { toast } = useToast();
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

  const buildRowForJob = (jobId: string, jobRows: Row[]) => {
    const first = jobRows[0];

    return {
      "Job ID": jobId,
      "Customer Name": `${first[mapping.firstName] || ""} ${
        first[mapping.lastName] || ""
      }`,
      "Customer Phone": first[mapping.phone] || "",
      Address: `${first[mapping.address] || ""} ${
        first[mapping.suburb] || ""
      }`,
      "Installation Date": first[mapping.installDate] || "",
      CoES: jobRows.map((r) => JSON.stringify(r)).join("\n"),
    };
  };

  // ✅ FIXED PREVIEW (IMPORTANT)
  const allPreviewRows = useMemo(() => {
    const col =
      mapping.jobId ||
      headers.find((h) => h.toLowerCase().includes("job")) ||
      "";

    if (!col || rows.length === 0) return [];

    const groups = new Map<string, Row[]>();

    for (const r of rows) {
      const v = r[col as keyof Row];
      if (!v) continue;

      const key = String(v);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push(r);
    }

    return Array.from(groups.entries()).map(([id, jobRows]) =>
      buildRowForJob(id, jobRows)
    );
  }, [rows, mapping, headers]);

  return (
    <div className="p-6">
      <h1 className="text-xl font-bold">Excel Job Formatter</h1>

      <input type="file" ref={fileRef} onChange={(e) => {
        const f = e.target.files?.[0];
        if (f) handleFile(f);
      }} />

      {fileName && (
        <>
          <p>{fileName}</p>

          <h2 className="mt-4 font-bold">Preview</h2>

          {/* ✅ THIS WILL NOW WORK */}
          <pre className="text-xs">
            {JSON.stringify(allPreviewRows, null, 2)}
          </pre>
        </>
      )}
    </div>
  );
}

export default App;
