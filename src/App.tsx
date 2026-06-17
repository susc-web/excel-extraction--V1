import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Toaster } from "@/components/ui/toaster";
import { useToast } from "@/hooks/use-toast";
import { FileSpreadsheet, Upload, X, ListChecks } from "lucide-react";

type Row = Record<string, unknown>;

const OUTPUT_COLUMNS = ["Job ID","Customer Name","Customer Phone","Address","Installation Date","CoES"] as const;
type OutputCol = (typeof OUTPUT_COLUMNS)[number];

type InputKey = "jobId"|"firstName"|"lastName"|"phone"|"address"|"suburb"|"postcode"|"state"|"installDate"|"product"|"productModel"|"serial"|"questionCol";

const INPUT_FIELDS: { key: InputKey; label: string; candidates: string[] }[] = [
  { key: "jobId", label: "Job ID", candidates: ["jobid","job id","job"] },
  { key: "firstName", label: "Customer First Name", candidates: ["customerfirstname","customer first name","firstname","first name","first","fname","givenname","given name"] },
  { key: "lastName", label: "Customer Last Name", candidates: ["customerlastname","customer last name","customersurname","customer surname","lastname","last name","last","lname","surname","familyname","family name"] },
  { key: "phone", label: "Customer Phone Number", candidates: ["customerphonenumber","customer phone number","customerphone","customer phone","phonenumber","phone number","phone","customermobilenumber","customer mobile number","customermobile","customer mobile","mobilenumber","mobile number","mobile","mobileno","mobile no","contactnumber","contact number","contact","tel","telephone"] },
  { key: "address", label: "Customer Address", candidates: ["customeraddress","address","street address","street"] },
  { key: "suburb", label: "Customer Suburb", candidates: ["customersuburb","suburb","city","town"] },
  { key: "postcode", label: "Customer Postcode", candidates: ["customerpostcode","postcode","postal code","zip","zipcode","zip code"] },
  { key: "state", label: "Customer State", candidates: ["customerstate","state","region","province"] },
  { key: "installDate", label: "Installation Date", candidates: ["dateschedu","datescheduled","date schedu","date scheduled","installationdate","installation date","installdate","install date","scheduleddate","scheduled date"] },
  { key: "product", label: "Primary Product Brand", candidates: ["primaryproduct1brand","primaryproductbrand","productbrand","primary product 1 brand","primary product brand","product brand","brand","primaryproduct","product","primary product"] },
  { key: "productModel", label: "Primary Product Model", candidates: ["primaryproduct1model","primaryproductmodel","productmodel","primary product 1 model","primary product model","product model","model"] },
  { key: "serial", label: "Serial", candidates: ["serial","serialnumber","serial number","sn"] },
  { key: "questionCol", label: "Question / Type column (for outdoor detection)", candidates: ["question","questions","type","unittype","unit type","scantype","scan type","description"] },
];

const norm = (s: string) => s.toLowerCase().replace(/[\s_-]+/g, "");

function findColumn(headers: string[], candidates: string[]): string | null {
  const targets = candidates.map(norm);
  for (const h of headers) { if (targets.includes(norm(h))) return h; }
  for (const h of headers) {
    const hn = norm(h);
    if (candidates.some((c) => hn.startsWith(norm(c)) || norm(c).startsWith(hn))) return h;
  }
  return null;
}

function autoDetect(headers: string[]): Record<InputKey, string> {
  const out = {} as Record<InputKey, string>;
  for (const f of INPUT_FIELDS) out[f.key] = findColumn(headers, f.candidates) ?? "";
  return out;
}

function formatCell(v: unknown): string {
  if (v === undefined || v === null || v === "") return "";
  if (v instanceof Date) {
    return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
  }
  return String(v);
}

function formatPhone(v: unknown): string {
  if (v === undefined || v === null || v === "") return "";
  let s = String(v).replace(/[\s\-().+]/g, "");
  if (s === "") return "";
  if (s.startsWith("61") && s.length === 11) s = "0" + s.slice(2);
  if (/^\d{9}$/.test(s)) s = "0" + s;
  return s;
}

function joinNonEmpty(parts: string[], sep: string): string {
  return parts.map((p) => p.trim()).filter((p) => p.length > 0).join(sep);
}

function App() {
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [rows, setRows] = useState<Row[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<Record<InputKey, string>>(() => autoDetect([]));

  const setCol = (key: InputKey, value: string) => setMapping((m) => ({ ...m, [key]: value }));

  const reset = () => {
    setFileName(null); setRows([]); setHeaders([]); setMapping(autoDetect([]));
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const handleFile = async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const firstSheetName = wb.SheetNames[0];
      if (!firstSheetName) throw new Error("Workbook has no sheets");
      const ws = wb.Sheets[firstSheetName];
      if (!ws) throw new Error("Sheet not found");
      const json = XLSX.utils.sheet_to_json<Row>(ws, { defval: "" });
      if (json.length === 0) {
        toast({ title: "Empty file", description: "No rows found in the first sheet.", variant: "destructive" });
        return;
      }
      const cols = Object.keys(json[0] ?? {});
      const lastValues: Record<string, unknown> = {};
      const filled = json.map((r) => {
        const out: Row = {};
        for (const col of cols) {
          const v = r[col];
          const isEmpty = v === undefined || v === null || v === "";
          out[col] = isEmpty ? (lastValues[col] ?? "") : v;
          if (!isEmpty) lastValues[col] = v;
        }
        return out;
      });
      setFileName(file.name);
      setRows(filled);
      setHeaders(cols);
      setMapping(autoDetect(cols));
      toast({ title: "File loaded", description: `${json.length.toLocaleString()} rows from "${firstSheetName}".` });
    } catch (err) {
      toast({ title: "Could not read file", description: err instanceof Error ? err.message : "Unknown error", variant: "destructive" });
    }
  };

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => { const f = e.target.files?.[0]; if (f) void handleFile(f); };
  const onDrop = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) void handleFile(f); };

  const jobCount = useMemo(() => {
    const col = mapping.jobId;
    if (!col || rows.length === 0) return 0;
    const set = new Set<string>();
    for (const r of rows) { const v = r[col]; if (v !== undefined && v !== null && v !== "") set.add(String(v)); }
    return set.size;
  }, [rows, mapping.jobId]);

  const buildRowForJob = (jobId: string, jobRows: Row[]): Record<OutputCol, string> => {
    const first = jobRows[0] ?? {};
    const get = (key: InputKey) => { const col = mapping[key]; return col ? formatCell(first[col]) : ""; };
    const customerName = joinNonEmpty([get("firstName"), get("lastName")], " ");
    const customerPhone = mapping.phone ? formatPhone(first[mapping.phone]) : "";
    const address = joinNonEmpty([get("address"), get("suburb"), get("postcode"), get("state")], ", ");
    const installationDate = get("installDate");

    const productCol = mapping.product, modelCol = mapping.productModel;
    const serialCol = mapping.serial, questionCol = mapping.questionCol;

    type Group = { product: string; model: string; outdoorSerials: string[]; otherSerials: string[] };
    const groups = new Map<string, Group>();
    const groupOrder: string[] = [];

    for (const r of jobRows) {
      const p = productCol ? formatCell(r[productCol]) : "";
      const m = modelCol ? formatCell(r[modelCol]) : "";
      const s = serialCol ? formatCell(r[serialCol]) : "";
      if (!p && !m && !s) continue;
      const key = `${p}\u0000${m}`;
      if (!groups.has(key)) { groups.set(key, { product: p, model: m, outdoorSerials: [], otherSerials: [] }); groupOrder.push(key); }
      const grp = groups.get(key)!;
      if (s) {
        const question = questionCol ? String(r[questionCol] ?? "").toLowerCase() : "";
        const isOutdoor = question.includes("outdoor");
        if (isOutdoor) { if (!grp.outdoorSerials.includes(s)) grp.outdoorSerials.push(s); }
        else { if (!grp.otherSerials.includes(s)) grp.otherSerials.push(s); }
      }
    }

    const coesLines: string[] = [];
    for (const key of groupOrder) {
      const g = groups.get(key)!;
      const header = joinNonEmpty([g.product, g.model], " - ");
      if (header) coesLines.push(header);
      g.outdoorSerials.forEach((s, i) => coesLines.push(i === 0 ? `Serial Number: Outdoor: ${s}` : `Outdoor: ${s}`));
      g.otherSerials.forEach((s, i) => coesLines.push(i === 0 ? `Indoor: ${s}` : s));
    }

    return { "Job ID": jobId, "Customer Name": customerName, "Customer Phone": customerPhone, Address: address, "Installation Date": installationDate, CoES: coesLines.join("\n") };
  };

  const writeWorkbook = (data: Record<OutputCol, string>[], sheetName: string, fileName: string) => {
    const ws = XLSX.utils.json_to_sheet(data, { header: [...OUTPUT_COLUMNS] });
    ws["!cols"] = OUTPUT_COLUMNS.map((c) => c === "CoES" ? { wch: 40 } : c === "Address" ? { wch: 50 } : c === "Customer Name" ? { wch: 28 } : { wch: 18 });
    const headerStyle = { font: { bold: true }, alignment: { vertical: "top", wrapText: true } };
    const bodyStyle = { alignment: { vertical: "top", wrapText: true } };
    const rowHeights: { hpt: number }[] = [{ hpt: 18 }];
    for (let r = 0; r < data.length; r++) rowHeights.push({ hpt: Math.min(15 * Math.max((data[r]?.["CoES"] ?? "").split("\n").length, 1) + 4, 409) });
    ws["!rows"] = rowHeights;
    for (let R = 0; R < data.length + 1; R++) for (let C = 0; C < OUTPUT_COLUMNS.length; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell) cell.s = R === 0 ? headerStyle : bodyStyle;
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);
  };

  const downloadAll = () => {
    const col = mapping.jobId;
    if (!col || rows.length === 0) return;
    const groups = new Map<string, Row[]>();
    for (const r of rows) {
      const j = r[col];
      if (j === undefined || j === null || j === "") continue;
      const jk = String(j);
      if (!groups.has(jk)) groups.set(jk, []);
      groups.get(jk)!.push(r);
    }
    const out = Array.from(groups.entries())
      .sort((a, b) => { const na = Number(a[0]), nb = Number(b[0]); return !isNaN(na) && !isNaN(nb) ? na - nb : a[0].localeCompare(b[0]); })
      .map(([id, jobRows]) => buildRowForJob(id, jobRows));
    writeWorkbook(out, "All Jobs", "formatted_output.xlsx");
    toast({ title: "Downloaded", description: `formatted_output.xlsx (${out.length} jobs)` });
  };

  const allPreviewRows = useMemo(() => {
    const col = mapping.jobId;
    if (!col || rows.length === 0) return [];
    const groups = new Map<string, Row[]>();
    const order: string[] = [];
    for (const r of rows) {
      const v = r[col];
      if (v === undefined || v === null || v === "") continue;
      const jk = String(v);
      if (!groups.has(jk)) { groups.set(jk, []); order.push(jk); }
      groups.get(jk)!.push(r);
    }
    return order
      .sort((a, b) => { const na = Number(a), nb = Number(b); return !isNaN(na) && !isNaN(nb) ? na - nb : a.localeCompare(b); })
      .map((id) => buildRowForJob(id, groups.get(id)!));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows, mapping]);

  return (
    <div className="min-h-screen w-full bg-background text-foreground">
      <div className="mx-auto max-w-5xl px-4 py-10 sm:py-14">
        <header className="mb-8">
          <div className="inline-flex items-center gap-2 rounded-full border border-border bg-card px-3 py-1 text-xs font-medium text-muted-foreground shadow-sm">
            <FileSpreadsheet className="h-3.5 w-3.5 text-primary" />
            Excel Job Formatter
          </div>
          <h1 className="mt-3 text-3xl font-semibold tracking-tight sm:text-4xl">Format jobs into one row each</h1>
          <p className="mt-2 max-w-2xl text-sm text-muted-foreground">
            Upload an Excel file, and download a clean sheet with one row per Job ID — Customer Name, Address, Installation Date, and all CoES combined.
          </p>
        </header>

        {!fileName ? (
          <Card className="border-dashed">
            <CardContent className="p-0">
              <label htmlFor="file-upload" className="flex cursor-pointer flex-col items-center justify-center gap-3 px-6 py-16 text-center" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
                <div className="flex h-14 w-14 items-center justify-center rounded-full bg-primary/10 text-primary"><Upload className="h-6 w-6" /></div>
                <div>
                  <p className="text-base font-medium">Drop your Excel file here, or click to browse</p>
                  <p className="mt-1 text-sm text-muted-foreground">Supports .xlsx, .xls, and .csv — processed in your browser, nothing is uploaded.</p>
                </div>
                <input id="file-upload" ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={onFileChange} />
              </label>
            </CardContent>
          </Card>
        ) : (
          <div className="grid gap-6">
            <Card>
              <CardHeader className="flex flex-row items-center justify-between gap-4">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 items-center justify-center rounded-md bg-primary/10 text-primary"><FileSpreadsheet className="h-5 w-5" /></div>
                  <div>
                    <CardTitle className="text-base">{fileName}</CardTitle>
                    <CardDescription>{rows.length.toLocaleString()} rows · {headers.length} columns</CardDescription>
                  </div>
                </div>
                <Button variant="ghost" size="sm" onClick={reset}><X className="mr-1 h-4 w-4" />Remove</Button>
              </CardHeader>
              <CardContent>
                <p className="mb-3 text-sm text-muted-foreground">Match each piece of the output to a column in your file. We've made our best guesses — adjust if anything looks off.</p>
                <div className="grid gap-4 sm:grid-cols-2">
                  {INPUT_FIELDS.map((f) => (
                    <div key={f.key} className="space-y-2">
                      <label className="text-sm font-medium">{f.label}</label>
                      <Select value={mapping[f.key]} onValueChange={(v) => setCol(f.key, v)}>
                        <SelectTrigger><SelectValue placeholder="Select column" /></SelectTrigger>
                        <SelectContent>{headers.map((h) => <SelectItem key={h} value={h}>{h}</SelectItem>)}</SelectContent>
                      </Select>
                    </div>
                  ))}
                </div>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="text-base">Export all jobs</CardTitle>
                <CardDescription>{jobCount.toLocaleString()} unique Job IDs found — one row per job.</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <Button size="lg" onClick={downloadAll} disabled={!mapping.jobId || rows.length === 0} className="w-full sm:w-auto">
                  <ListChecks className="mr-2 h-4 w-4" />Download formatted_output.xlsx
                </Button>
                {allPreviewRows.length > 0 && (
                  <div className="rounded-lg border border-border bg-muted/30 p-4">
                    <div className="mb-3 flex flex-wrap items-center gap-2">
                      <Badge variant="secondary">Preview</Badge>
                      <Badge variant="outline">{allPreviewRows.length} jobs</Badge>
                      <span className="text-xs text-muted-foreground">This is exactly what will be downloaded.</span>
                    </div>
                    <div className="overflow-x-auto rounded-md border border-border bg-background">
                      <Table>
                        <TableHeader>
                          <TableRow>{OUTPUT_COLUMNS.map((c) => <TableHead key={c} className="whitespace-nowrap text-xs">{c}</TableHead>)}</TableRow>
                        </TableHeader>
                        <TableBody>
                          {allPreviewRows.map((row) => (
                            <TableRow key={row["Job ID"]}>
                              {OUTPUT_COLUMNS.map((c) => (
                                <TableCell key={c} className={c === "CoES" ? "max-w-md whitespace-pre-line break-words align-top font-mono text-xs" : c === "Address" ? "max-w-md break-words align-top font-mono text-xs" : "whitespace-nowrap align-top font-mono text-xs"}>
                                  {row[c] ?? ""}
                                </TableCell>
                              ))}
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="text-base">Source data preview</CardTitle>
                <CardDescription>First 10 rows of your file</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="overflow-x-auto rounded-md border border-border">
                  <Table>
                    <TableHeader><TableRow>{headers.map((h) => <TableHead key={h} className="whitespace-nowrap">{h}</TableHead>)}</TableRow></TableHeader>
                    <TableBody>
                      {rows.slice(0, 10).map((r, i) => (
                        <TableRow key={i}>{headers.map((h) => <TableCell key={h} className="whitespace-nowrap font-mono text-xs">{formatCell(r[h])}</TableCell>)}</TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </div>
        )}

        <footer className="mt-10 text-center text-xs text-muted-foreground">Files are processed entirely in your browser — nothing is uploaded.</footer>
      </div>
      <Toaster />
    </div>
  );
}

export default App;
