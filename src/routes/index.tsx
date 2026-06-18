import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import cfaLogo from "@/assets/cfa-logo.png.asset.json";
import cfaText from "@/assets/cfa-text.png.asset.json";
import bgWallpaper from "@/assets/bg-wallpaper.jpg";
import windowsBuild from "@/assets/cfa-windows.zip.asset.json";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, HeadingLevel, PageOrientation, ImageRun,
} from "docx";
import ExcelJS from "exceljs";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Christian Faith Assembly — Tithe Register" },
      { name: "description", content: "Tithe Register Form for Christian Faith Assembly. Log member tithes in OMR with live dashboard totals." },
      { property: "og:title", content: "Christian Faith Assembly — Tithe Register" },
      { property: "og:description", content: "Tithe Register Form for Christian Faith Assembly." },
    ],
  }),
  component: Index,
});

type Entry = {
  id: string;
  name: string;
  countryCode: "+968";
  phone: string;
  currency: "OMR";
  amount: number;
  category: string;
  method: string;
  note: string;
  title: "Brother" | "Sister";
  date: string; // ISO
};

const CATEGORIES = ["Tithe", "Offering", "Missions", "Building Fund", "Thanksgiving"];
const METHODS = ["Cash", "Bank Transfer", "Card", "Cheque"];
const TITLES: Array<"Brother" | "Sister"> = ["Brother", "Sister"];
const STORAGE_KEY = "cfa-tithe-entries-v1";

function formatAmount(amount: number) {
  return `${amount.toFixed(3)} OMR`;
}

// Convert a number into English words (supports OMR with 3-decimal baisa).
const ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
  "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
const TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
function twoDigits(n: number): string {
  if (n < 20) return ONES[n];
  const t = Math.floor(n / 10), o = n % 10;
  return TENS[t] + (o ? " " + ONES[o] : "");
}
function threeDigits(n: number): string {
  const h = Math.floor(n / 100), r = n % 100;
  const parts: string[] = [];
  if (h) parts.push(ONES[h] + " Hundred");
  if (r) parts.push(twoDigits(r));
  return parts.join(" ");
}
function intToWords(n: number): string {
  if (n === 0) return "Zero";
  const units = ["", "Thousand", "Million", "Billion"];
  let i = 0;
  const chunks: string[] = [];
  while (n > 0) {
    const c = n % 1000;
    if (c) chunks.unshift(threeDigits(c) + (units[i] ? " " + units[i] : ""));
    n = Math.floor(n / 1000);
    i++;
  }
  return chunks.join(" ");
}
function amountInWords(amount: number): string {
  if (!isFinite(amount) || amount <= 0) return "";
  const rial = Math.floor(amount);
  const baisa = Math.round((amount - rial) * 1000);
  const parts: string[] = [];
  if (rial > 0) parts.push(`${intToWords(rial)} Rial${rial === 1 ? "" : "s"}`);
  if (baisa > 0) parts.push(`${intToWords(baisa)} Baisa`);
  if (!parts.length) return "Zero Rials";
  return parts.join(" and ") + " Only";
}

function Index() {
  const [entries, setEntries] = useState<Entry[]>(() => {
    if (typeof window === "undefined") return [];
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? (JSON.parse(raw) as Entry[]) : [];
    } catch {
      return [];
    }
  });
  const [name, setName] = useState("");
  const [countryCode, setCountryCode] = useState<"+968">("+968");
  const [phone, setPhone] = useState("");
  const [currency, setCurrency] = useState<"OMR">("OMR");
  const [amount, setAmount] = useState("");
  const [category, setCategory] = useState(CATEGORIES[0]);
  const [method, setMethod] = useState(METHODS[0]);
  const [note, setNote] = useState("");
  const [title, setTitle] = useState<"Brother" | "Sister">("Brother");
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [query, setQuery] = useState("");
  const [filterCurrency, setFilterCurrency] = useState<"ALL" | "OMR">("ALL");
  const [editingId, setEditingId] = useState<string | null>(null);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [darkMode, setDarkMode] = useState(() => {
    if (typeof window === "undefined") return false;
    return localStorage.getItem("cfa-theme") === "dark";
  });

  useEffect(() => {
    if (typeof window !== "undefined") {
      localStorage.setItem("cfa-theme", darkMode ? "dark" : "light");
    }
  }, [darkMode]);

  const isDark = darkMode;

  const phoneDigits = 8;

  // Persist to localStorage on every change
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
    } catch {
      // ignore quota errors
    }
  }, [entries]);

  const totals = useMemo(() => {
    let omr = 0;
    for (const e of entries) {
      omr += e.amount;
    }
    return { omr, count: entries.length };
  }, [entries]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return entries.filter((e) => {
      if (filterCurrency !== "ALL" && e.currency !== filterCurrency) return false;
      if (!q) return true;
      const fullPhone = `${e.countryCode} ${e.phone}`;
      return (
        e.name.toLowerCase().includes(q) ||
        fullPhone.toLowerCase().includes(q) ||
        e.phone.toLowerCase().includes(q) ||
        e.category.toLowerCase().includes(q)
      );
    });
  }, [entries, query, filterCurrency]);

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    const amt = parseFloat(amount);
    if (!name || !phone || !amt || amt <= 0) return;
    if (phone.length !== phoneDigits) {
      window.alert(`Phone number must be exactly ${phoneDigits} digits for ${countryCode}.`);
      return;
    }
    if (editingId) {
      setEntries((prev) => prev.map((en) =>
        en.id === editingId
          ? { ...en, name, countryCode, phone, currency, amount: amt, category, method, note, title, date }
          : en
      ));
      setEditingId(null);
    } else {
      setEntries((prev) => [
        {
          id: crypto.randomUUID(),
          name, countryCode, phone, currency, amount: amt, category, method, note, title, date,
        },
        ...prev,
      ]);
    }
    setName(""); setPhone(""); setAmount(""); setNote("");
  }

  function startEdit(en: Entry) {
    setEditingId(en.id);
    setName(en.name);
    setCountryCode(en.countryCode);
    setPhone(en.phone);
    setCurrency(en.currency);
    setAmount(String(en.amount));
    setCategory(en.category);
    setMethod(en.method);
    setNote(en.note);
    setTitle(en.title ?? "Brother");
    setDate(en.date);
    if (typeof window !== "undefined") window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function cancelEdit() {
    setEditingId(null);
    setName(""); setPhone(""); setAmount(""); setNote("");
  }

  function deleteEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id));
    if (editingId === id) cancelEdit();
  }

  function deleteAllEntries() {
    if (!entries.length) return;
    const ok = window.confirm(
      `Delete ALL ${entries.length} record(s)? This cannot be undone.`,
    );
    if (!ok) return;
    setEntries([]);
  }

  function deleteAllEntriesWithDoubleConfirm() {
    if (!entries.length) return;
    const first = window.confirm(
      `Delete ALL ${entries.length} record(s)? This cannot be undone.`,
    );
    if (!first) return;
    const second = window.confirm(
      `Are you absolutely sure? All ${entries.length} record(s) will be permanently erased.`,
    );
    if (!second) return;
    setEntries([]);
    if (editingId) cancelEdit();
  }

  async function fetchImageBuffer(url: string): Promise<ArrayBuffer> {
    const res = await fetch(url);
    return await res.arrayBuffer();
  }

  async function exportWord() {
    if (!entries.length) return;
    const HEADERS = ["Date", "Title", "Name", "Phone", "Category", "Method", "Amount (OMR)", "Amount In Words", "Note"];
    const COL_COLORS = ["#dbeafe", "#fef3c7", "#dcfce7", "#e0e7ff", "#fce7f3", "#ffedd5", "#d1fae5", "#ede9fe", "#f1f5f9"];

    const border = { style: BorderStyle.SINGLE, size: 6, color: "94a3b8" };
    const borders = { top: border, bottom: border, left: border, right: border };

    const headerRow = new TableRow({
      tableHeader: true,
      children: HEADERS.map((h) => new TableCell({
        borders,
        shading: { fill: "0F172A", type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 22 })],
        })],
      })),
    });

    const dataRows = entries.map((e, idx) => {
      const cells = [
        e.date,
        e.title ?? "",
        e.name,
        `${e.countryCode} ${e.phone}`,
        e.category,
        e.method,
        `${e.amount.toFixed(3)} OMR`,
        amountInWords(e.amount),
        e.note || "",
      ];
      const zebra = idx % 2 === 0 ? "FFFFFF" : "F8FAFC";
      return new TableRow({
        children: cells.map((val, ci) => new TableCell({
          borders,
          shading: { fill: zebra, type: ShadingType.CLEAR, color: "auto" },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({
              text: val,
              size: 20,
              bold: ci === 2 || ci === 6,
              color: ci === 6 ? "059669" : "0F172A",
            })],
          })],
        })),
      });
    });
    void COL_COLORS;

    const totalOmr = entries.reduce((s, e) => s + e.amount, 0);

    const [logoBuf, textBuf] = await Promise.all([
      fetchImageBuffer(cfaLogo.url),
      fetchImageBuffer(cfaText.url),
    ]);

    const doc = new Document({
      styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
      sections: [{
        properties: {
          page: {
            size: { width: 16838, height: 11906, orientation: PageOrientation.LANDSCAPE },
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
          },
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new ImageRun({
              type: "png",
              data: logoBuf,
              transformation: { width: 90, height: 90 },
            })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 120 },
            children: [new ImageRun({
              type: "png",
              data: textBuf,
              transformation: { width: 420, height: 70 },
            })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            heading: HeadingLevel.HEADING_2,
            children: [new TextRun({ text: "Tithe Register", bold: true, size: 28, color: "64748B" })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 240 },
            children: [new TextRun({ text: `Generated on ${new Date().toLocaleDateString()}`, italics: true, size: 18, color: "94A3B8" })],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [headerRow, ...dataRows],
          }),
          new Paragraph({ spacing: { before: 240 }, children: [new TextRun("")] }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ text: "Total Entries: ", bold: true, size: 22 }),
              new TextRun({ text: `${entries.length}    `, size: 22 }),
              new TextRun({ text: "Total: ", bold: true, size: 22 }),
              new TextRun({ text: `${totalOmr.toFixed(3)} OMR`, bold: true, size: 24, color: "059669" }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ text: "In Words: ", bold: true, size: 20 }),
              new TextRun({ text: amountInWords(totalOmr), italics: true, size: 20, color: "475569" }),
            ],
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `tithe-register-${new Date().toISOString().slice(0, 10)}.docx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function exportExcel() {
    if (!entries.length) return;
    const wb = new ExcelJS.Workbook();
    wb.creator = "Christian Faith Assembly";
    wb.created = new Date();
    const ws = wb.addWorksheet("Tithe Register", {
      pageSetup: { orientation: "landscape", horizontalCentered: true, fitToPage: true, fitToWidth: 1, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } },
      views: [{ state: "normal", showGridLines: false }],
    });

    const HEADERS = ["Date", "Title", "Name", "Phone", "Category", "Method", "Amount (OMR)", "Amount In Words", "Note"];
    const widths = [14, 10, 22, 18, 16, 16, 18, 42, 26];
    ws.columns = HEADERS.map((h, i) => ({ header: h, key: h, width: widths[i] }));

    // Embed logo + text image at the top
    const [logoBuf, textBuf] = await Promise.all([
      fetchImageBuffer(cfaLogo.url),
      fetchImageBuffer(cfaText.url),
    ]);
    const logoId = wb.addImage({ buffer: logoBuf, extension: "png" });
    const textId = wb.addImage({ buffer: textBuf, extension: "png" });

    // Reserve top rows for branding (rows 1-5)
    for (let r = 1; r <= 5; r++) ws.getRow(r).height = 22;

    ws.addImage(logoId, { tl: { col: 0.2, row: 0.2 }, ext: { width: 90, height: 90 } });
    ws.addImage(textId, { tl: { col: 2.5, row: 0.5 }, ext: { width: 380, height: 70 } });

    // Subtitle row
    ws.mergeCells(6, 1, 6, HEADERS.length);
    const subtitle = ws.getCell(6, 1);
    subtitle.value = "Tithe Register";
    subtitle.alignment = { horizontal: "center", vertical: "middle" };
    subtitle.font = { name: "Calibri", size: 18, bold: true, color: { argb: "FF334155" } };
    ws.getRow(6).height = 28;

    ws.mergeCells(7, 1, 7, HEADERS.length);
    const gen = ws.getCell(7, 1);
    gen.value = `Generated on ${new Date().toLocaleDateString()}`;
    gen.alignment = { horizontal: "center" };
    gen.font = { italic: true, color: { argb: "FF94A3B8" }, size: 11 };
    ws.getRow(7).height = 18;

    // Header row at row 9
    const headerRowIdx = 9;
    const headerRow = ws.getRow(headerRowIdx);
    HEADERS.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 12 };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin", color: { argb: "FF64748B" } },
        bottom: { style: "thin", color: { argb: "FF64748B" } },
        left: { style: "thin", color: { argb: "FF64748B" } },
        right: { style: "thin", color: { argb: "FF64748B" } },
      };
    });
    headerRow.height = 26;

    // Data rows
    entries.forEach((e, idx) => {
      const r = headerRowIdx + 1 + idx;
      const row = ws.getRow(r);
      const values = [
        e.date,
        e.title ?? "",
        e.name,
        `${e.countryCode} ${e.phone}`,
        e.category,
        e.method,
        e.amount,
        amountInWords(e.amount),
        e.note || "",
      ];
      values.forEach((v, i) => {
        const cell = row.getCell(i + 1);
        cell.value = v as string | number;
        cell.alignment = { vertical: "middle", wrapText: true, horizontal: i === 6 ? "right" : "left" };
        cell.font = { name: "Calibri", size: 11, bold: i === 2 || i === 6, color: { argb: i === 6 ? "FF059669" : "FF0F172A" } };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: idx % 2 === 0 ? "FFFFFFFF" : "FFF1F5F9" },
        };
        cell.border = {
          top: { style: "hair", color: { argb: "FFCBD5E1" } },
          bottom: { style: "hair", color: { argb: "FFCBD5E1" } },
          left: { style: "hair", color: { argb: "FFCBD5E1" } },
          right: { style: "hair", color: { argb: "FFCBD5E1" } },
        };
        if (i === 6) cell.numFmt = "0.000\" OMR\"";
      });
      row.height = 22;
    });

    // Totals row
    const totalOmr = entries.reduce((s, e) => s + e.amount, 0);
    const totalRowIdx = headerRowIdx + 1 + entries.length + 1;
    ws.mergeCells(totalRowIdx, 1, totalRowIdx, 6);
    const lbl = ws.getCell(totalRowIdx, 1);
    lbl.value = `Total Entries: ${entries.length}    Grand Total`;
    lbl.alignment = { horizontal: "right", vertical: "middle" };
    lbl.font = { bold: true, size: 12, color: { argb: "FF0F172A" } };
    const tot = ws.getCell(totalRowIdx, 7);
    tot.value = totalOmr;
    tot.numFmt = "0.000\" OMR\"";
    tot.font = { bold: true, size: 13, color: { argb: "FF059669" } };
    tot.alignment = { horizontal: "right", vertical: "middle" };
    ws.mergeCells(totalRowIdx, 8, totalRowIdx, 9);
    const wordsCell = ws.getCell(totalRowIdx, 8);
    wordsCell.value = amountInWords(totalOmr);
    wordsCell.font = { italic: true, color: { argb: "FF475569" }, size: 11 };
    wordsCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    ws.getRow(totalRowIdx).height = 26;

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `tithe-register-${new Date().toISOString().slice(0, 10)}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  return (
    <>
      <style>{`
        .btn-glow { transition: all 0.2s ease; }
        .btn-glow:hover:not(:disabled) {
          filter: brightness(1.15);
          transform: translateY(-2px);
          box-shadow: 0 10px 25px -5px rgba(0,0,0,0.25);
        }
        [data-theme="dark"] .theme-container {
          background: #1e293b !important;
          color: #e2e8f0 !important;
        }
        [data-theme="dark"] .theme-form {
          background: #334155 !important;
          border-color: #475569 !important;
        }
        [data-theme="dark"] .theme-input {
          background: #1e293b !important;
          border-color: #475569 !important;
          color: #e2e8f0 !important;
        }
        [data-theme="dark"] .theme-table-wrap {
          border-color: #475569 !important;
        }
        [data-theme="dark"] .theme-table th {
          background: #0f172a !important;
        }
        [data-theme="dark"] .theme-table td {
          border-bottom-color: #475569 !important;
          color: #e2e8f0 !important;
        }
        [data-theme="dark"] .theme-table tr:hover {
          background: #334155 !important;
        }
        [data-theme="dark"] .theme-stat-box {
          background: #334155 !important;
        }
        [data-theme="dark"] .theme-stat-label {
          color: #94a3b8 !important;
        }
        [data-theme="dark"] .theme-h2 {
          color: #e2e8f0 !important;
        }
        [data-theme="dark"] .theme-empty {
          color: #94a3b8 !important;
        }
        [data-theme="dark"] .theme-subtitle {
          color: #94a3b8 !important;
        }
        [data-theme="dark"] .theme-label {
          color: #cbd5e1 !important;
        }
        [data-theme="dark"] .theme-phone-prefix {
          background: #1e293b !important;
          color: #e2e8f0 !important;
          border-color: #475569 !important;
        }
        [data-theme="dark"] .theme-in-words {
          background: #312e81 !important;
          border-color: #6B5BFF !important;
          color: #e2e8f0 !important;
        }
        [data-theme="dark"] .theme-logo-img {
          filter: brightness(0) invert(1);
        }
        [data-theme="dark"] .theme-text-img {
          filter: brightness(0) invert(1);
        }
        [data-theme="dark"] .theme-subtitle {
          color: #ffffff !important;
        }
        [data-theme="dark"] .theme-branch {
          color: #ffffff !important;
        }
      `}</style>
      <div style={styles.page} data-theme={darkMode ? "dark" : "light"}>
        <div style={styles.container} className="theme-container">
          <header style={styles.header}>
            <div style={styles.headerLeft}>
              <img src={cfaLogo.url} alt="CFA Logo" style={styles.logoImg} className="theme-logo-img" />
            </div>
            <div style={styles.headerCenter}>
              <img src={cfaText.url} alt="Christian Faith Assembly" style={styles.textImg} className="theme-text-img" />
              <p style={styles.subtitle} className="theme-subtitle">Tithe Registration</p>
              <p style={styles.branch} className="theme-branch">Mabela-Branch</p>
            </div>
            <div style={styles.headerRight}>
              <button
                type="button"
                onClick={() => setDarkMode((d) => !d)}
                className="btn-glow"
                style={{
                  background: darkMode ? "#f59e0b" : "#4A3F9F",
                  color: "white",
                  border: "none",
                  width: 42,
                  height: 42,
                  borderRadius: 12,
                  cursor: "pointer",
                  fontSize: 20,
                  fontWeight: 700,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  boxShadow: darkMode ? "0 4px 12px -2px rgba(245,158,11,0.4)" : "0 4px 12px -2px rgba(74,63,159,0.4)",
                }}
                title={darkMode ? "Switch to Light Mode" : "Switch to Dark Mode"}
              >
                {darkMode ? "☀" : "☾"}
              </button>
            </div>
          </header>

        <section style={styles.dashboard}>
          <StatCard label="Total Omani Rial (OMR)" value={`${totals.omr.toFixed(3)} OMR`} accent="#6B9EFF" />
          <StatCard label="Total Entries" value={String(totals.count)} accent="#4A3F9F" />
        </section>

        <form onSubmit={handleSubmit} style={styles.form} className="theme-form">
          <div style={styles.formGrid}>
            <Field label="Date">
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={styles.input} className="theme-input" required />
            </Field>
            <Field label="Person's Name">
              <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Full name" style={styles.input} className="theme-input" required />
            </Field>
            <Field label="Phone Number">
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ ...styles.input, width: 90, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 600, background: "#f1f5f9" }} className="theme-phone-prefix">
                  +968
                </div>
                <input
                  type="tel"
                  value={phone}
                  onChange={(e) => {
                    const digits = e.target.value.replace(/\D/g, "").slice(0, phoneDigits);
                    setPhone(digits);
                  }}
                  placeholder={`${phoneDigits} digits`}
                  style={styles.input}
                  className="theme-input"
                  required
                />
              </div>
            </Field>
            <Field label="Category">
              <select value={category} onChange={(e) => setCategory(e.target.value)} style={styles.input} className="theme-input">
                {CATEGORIES.map((c) => <option key={c}>{c}</option>)}
              </select>
            </Field>
          </div>

          <div style={styles.thickDivider} />

          <div style={styles.formGrid}>
            <Field label="Payment Method">
              <select value={method} onChange={(e) => setMethod(e.target.value)} style={styles.input} className="theme-input">
                {METHODS.map((m) => <option key={m}>{m}</option>)}
              </select>
            </Field>
            <Field label="Amount">
              <input type="number" min="0.001" step="any" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0.00" style={styles.input} className="theme-input" required />
            </Field>
            <Field label="In Words (auto)">
              <div style={{ ...styles.input, display: "flex", alignItems: "center", fontWeight: 600, color: "#0f172a", background: "#E8E4FF", border: "1px solid #6B5BFF", minHeight: 40 }} className="theme-in-words">
                {amountInWords(parseFloat(amount) || 0) || <span style={{ color: "#94a3b8", fontWeight: 400 }}>e.g. Ten Rials Only</span>}
              </div>
            </Field>
            <Field label="Note (optional)">
              <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Reference / remark" style={styles.input} className="theme-input" />
            </Field>
            <Field label="Title">
              <select value={title} onChange={(e) => setTitle(e.target.value as "Brother" | "Sister")} style={styles.input} className="theme-input">
                {TITLES.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </Field>
          </div>

          <div style={styles.actions}>
            <button type="submit" className="btn-glow" style={styles.primaryBtn}>{editingId ? "✓ Update Record" : "＋ Save Record"}</button>
            {editingId && (
              <button type="button" onClick={cancelEdit} className="btn-glow" style={styles.secondaryBtn}>Cancel Edit</button>
            )}
            <div style={{ position: "relative" }}>
              <button
                type="button"
                onClick={() => setShowExportMenu((s) => !s)}
                style={styles.secondaryBtn}
                className="btn-glow"
                disabled={!entries.length}
              >
                ⤓ Exports ▼
              </button>
              {showExportMenu && (
                <div style={styles.dropdown}>
                  <button type="button" onClick={() => { setShowExportMenu(false); exportWord(); }} className="btn-glow" style={styles.dropdownItem}>WORD</button>
                  <button type="button" onClick={() => { setShowExportMenu(false); exportExcel(); }} className="btn-glow" style={styles.dropdownItem}>EXCEL</button>
                </div>
              )}
            </div>
            <button type="button" onClick={deleteAllEntries} className="btn-glow" style={styles.dangerBtn} disabled={!entries.length}>🗑 Delete All</button>
          </div>
        </form>

        <div style={styles.tableHeader}>
          <h2 style={styles.h2} className="theme-h2">Logged Transactions</h2>
          <div style={styles.tableTools}>
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search name, phone, category…" style={{ ...styles.input, width: 240 }} className="theme-input" />
            <select value={filterCurrency} onChange={(e) => setFilterCurrency(e.target.value as "ALL" | "OMR")} style={{ ...styles.input, width: 140 }} className="theme-input">
              <option value="ALL">All records</option>
              <option value="OMR">OMR only</option>
            </select>
          </div>
        </div>

        <div style={styles.tableWrap} className="theme-table-wrap">
          <table style={styles.table} className="theme-table">
            <thead>
              <tr>
                <th style={styles.th}>Date</th>
                <th style={styles.th}>Title</th>
                <th style={styles.th}>Name</th>
                <th style={{ ...styles.th, ...styles.splitCol }}>Phone</th>
                <th style={styles.th}>Category</th>
                <th style={styles.th}>Method</th>
                <th style={styles.th}>Currency</th>
                <th style={styles.th}>Amount</th>
                <th style={styles.th}></th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={9} style={styles.empty} className="theme-empty">No records yet — log your first entry above.</td></tr>
              ) : filtered.map((e) => (
                <tr key={e.id} style={{ ...styles.row, background: editingId === e.id ? "#fef3c7" : undefined }}>
                  <td style={styles.td}>{e.date}</td>
                  <td style={styles.td}>{e.title ?? "—"}</td>
                  <td style={{ ...styles.td, fontWeight: 600 }}>{e.name}</td>
                  <td style={{ ...styles.td, ...styles.splitCol }}>{e.countryCode} {e.phone}</td>
                  <td style={styles.td}><span style={styles.pill}>{e.category}</span></td>
                  <td style={styles.td}>{e.method}</td>
                  <td style={styles.td}>{e.currency}</td>
                  <td style={{ ...styles.td, fontWeight: 700, color: "#059669" }}>{formatAmount(e.amount)}</td>
                  <td style={styles.td}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button onClick={() => startEdit(e)} className="btn-glow" style={styles.editBtn} aria-label="Edit">✎</button>
                      <button onClick={() => deleteEntry(e.id)} className="btn-glow" style={styles.deleteBtn} aria-label="Delete">✕</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={styles.bottomBar}>
          <button type="button" onClick={deleteAllEntriesWithDoubleConfirm} className="btn-glow" style={styles.dangerBtn} disabled={!entries.length}>
            🗑 Delete All Records
          </button>
        </div>
      </div>
      <a
        href={windowsBuild.url}
        download="CFA-Register-Windows.zip"
        className="btn-glow"
        style={styles.winDownload}
        title="Download CFA Register for Windows 11 (.zip) — unzip and run CFA-Register.exe"
      >
        <span style={styles.winIcon} aria-hidden>⊞</span>
        <span style={{ display: "flex", flexDirection: "column", lineHeight: 1.1 }}>
          <span style={{ fontSize: 11, fontWeight: 600, opacity: 0.85 }}>Download for</span>
          <span style={{ fontSize: 15, fontWeight: 800 }}>Windows 11</span>
        </span>
      </a>
    </div>
    </>
  );
}

function StatCard({ label, value, accent }: { label: string; value: string; accent: string }) {
  return (
    <div style={{ ...styles.statBox, borderTop: `4px solid ${accent}` }} className="theme-stat-box">
      <div style={styles.statLabel} className="theme-stat-label">{label}</div>
      <div style={{ ...styles.statValue, color: accent }}>{value}</div>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      <label style={styles.label} className="theme-label">{label}</label>
      {children}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    backgroundImage: `url(${bgWallpaper})`,
    backgroundSize: "cover",
    backgroundPosition: "center",
    backgroundAttachment: "fixed",
    padding: "32px 16px",
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
    color: "#0f172a",
  },
  container: {
    maxWidth: 1180,
    margin: "0 auto",
    background: "#ffffff",
    borderRadius: 16,
    padding: 36,
    boxShadow: "0 25px 60px -15px rgba(0,0,0,0.5)",
  },
  header: {
    display: "grid",
    gridTemplateColumns: "120px 1fr 120px",
    alignItems: "center",
    gap: 10,
    paddingBottom: 24,
    borderBottom: "2px solid #e2e8f0",
    marginBottom: 28,
  },
  headerLeft: { display: "flex", alignItems: "center" },
  headerCenter: { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" },
  headerRight: { display: "flex", alignItems: "center", justifyContent: "flex-end" },
  logoImg: { width: 100, height: "auto", objectFit: "contain" },
  textImg: { width: 520, height: "auto", objectFit: "contain", maxWidth: "100%" },
  h1: { margin: 0, fontSize: 34, fontWeight: 800, letterSpacing: "-0.02em", color: "#0f172a" },
  subtitle: { margin: "6px 0 0", fontSize: 13, fontWeight: 700, color: "#64748b", letterSpacing: "0.1em", textTransform: "uppercase" },
  dashboard: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16, marginBottom: 28 },
  statBox: {
    background: "#f8fafc", borderRadius: 12, padding: "18px 20px",
    boxShadow: "0 1px 3px rgba(0,0,0,0.04)",
  },
  statLabel: { fontSize: 12, fontWeight: 600, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.05em" },
  statValue: { fontSize: 26, fontWeight: 800, marginTop: 6 },
  form: {
    background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 14,
    padding: 24, marginBottom: 32,
  },
  formGrid: { display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16 },
  thickDivider: {
    height: 3, background: "linear-gradient(90deg, transparent, #6B5BFF, transparent)",
    margin: "20px 0", borderRadius: 2,
  },
  label: { fontSize: 12, fontWeight: 700, color: "#475569", letterSpacing: "0.04em", textTransform: "uppercase" },
  input: {
    width: "100%", padding: "10px 12px", border: "1px solid #cbd5e1",
    borderRadius: 8, fontSize: 14, background: "white", boxSizing: "border-box",
    outline: "none",
  },
  actions: { display: "flex", gap: 12, marginTop: 20, justifyContent: "flex-end" },
  primaryBtn: {
    background: "#6B9EFF", color: "white",
    border: "none", padding: "12px 24px", borderRadius: 8, fontWeight: 700,
    cursor: "pointer", fontSize: 14, boxShadow: "0 6px 14px -4px rgba(107,158,255,0.55)",
  },
  secondaryBtn: {
    background: "#4A3F9F", color: "white", border: "none",
    padding: "12px 20px", borderRadius: 8, fontWeight: 600, cursor: "pointer", fontSize: 14,
    boxShadow: "0 6px 14px -4px rgba(74,63,159,0.45)",
  },
  dangerBtn: {
    background: "#F06B6B", color: "white",
    border: "none", padding: "12px 20px", borderRadius: 8, fontWeight: 700,
    cursor: "pointer", fontSize: 14, boxShadow: "0 6px 14px -4px rgba(240,107,107,0.45)",
  },
  bottomBar: { display: "flex", justifyContent: "center", marginTop: 24 },
  tableHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" },
  h2: { margin: 0, fontSize: 20, fontWeight: 700, color: "#0f172a" },
  tableTools: { display: "flex", gap: 10 },
  tableWrap: { overflowX: "auto", borderRadius: 12, border: "1px solid #e2e8f0" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 14 },
  th: {
    background: "#0f172a", color: "white", textAlign: "left",
    padding: "12px 14px", fontWeight: 600, fontSize: 12, textTransform: "uppercase", letterSpacing: "0.05em",
  },
  td: { padding: "12px 14px", borderBottom: "1px solid #e2e8f0" },
  splitCol: { borderRight: "3px solid #6B5BFF" },
  row: { transition: "background 0.15s" },
  pill: {
    display: "inline-block", padding: "3px 10px", borderRadius: 999,
    background: "#E8E4FF", color: "#4A3F9F", fontSize: 12, fontWeight: 600,
  },
  deleteBtn: {
    background: "#F06B6B", color: "white", border: "none",
    width: 30, height: 30, borderRadius: 8, cursor: "pointer", fontWeight: 700,
  },
  editBtn: {
    background: "#6B5BFF", color: "white", border: "none",
    width: 30, height: 30, borderRadius: 8, cursor: "pointer", fontWeight: 700,
  },
  empty: { padding: 36, textAlign: "center", color: "#94a3b8", fontStyle: "italic" },
  dropdown: {
    position: "absolute", top: "calc(100% + 6px)", right: 0,
    background: "white", border: "1px solid #cbd5e1", borderRadius: 8,
    boxShadow: "0 10px 25px -5px rgba(0,0,0,0.15)", zIndex: 50,
    display: "flex", flexDirection: "column", minWidth: 140,
    overflow: "hidden",
  },
  dropdownItem: {
    background: "#6B9EFF", color: "white", border: "none",
    padding: "12px 16px", fontWeight: 600, cursor: "pointer", fontSize: 14,
    textAlign: "left",
  },
  winDownload: {
    position: "fixed",
    bottom: 24,
    right: 24,
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "12px 18px",
    background: "#4A3F9F",
    color: "white",
    textDecoration: "none",
    borderRadius: 12,
    boxShadow: "0 10px 25px -5px rgba(74, 63, 159, 0.55), 0 4px 10px rgba(0,0,0,0.2)",
    fontFamily: "'Segoe UI', Tahoma, sans-serif",
    zIndex: 100,
    border: "1px solid rgba(255,255,255,0.2)",
  },
  winIcon: {
    fontSize: 24,
    fontWeight: 900,
    lineHeight: 1,
  },
};
