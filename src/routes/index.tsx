import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Christian Faith Assembly — Tithe Register" },
      { name: "description", content: "Tithe Register Form for Christian Faith Assembly. Log member tithes in INR and OMR with live dashboard totals." },
      { property: "og:title", content: "Christian Faith Assembly — Tithe Register" },
      { property: "og:description", content: "Tithe Register Form for Christian Faith Assembly." },
    ],
  }),
  component: Index,
});

type Entry = {
  id: string;
  name: string;
  countryCode: "+91" | "+968";
  phone: string;
  currency: "INR" | "OMR";
  amount: number;
  category: string;
  method: string;
  note: string;
  title: "Brother" | "Sister";
  date: string; // ISO
};

const CATEGORIES = ["Tithe", "Offering", "Missions", "Building Fund", "Thanksgiving"];
const METHODS = ["Cash", "Bank Transfer", "Card", "UPI", "Cheque"];
const TITLES: Array<"Brother" | "Sister"> = ["Brother", "Sister"];
const STORAGE_KEY = "cfa-tithe-entries-v1";

function formatAmount(currency: "INR" | "OMR", amount: number) {
  return currency === "INR" ? `₹${amount.toFixed(2)}` : `${amount.toFixed(3)} OMR`;
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
  const [countryCode, setCountryCode] = useState<"+91" | "+968">("+968");
  const [phone, setPhone] = useState("");
  const [currency, setCurrency] = useState<"INR" | "OMR">("INR");
  const [amount, setAmount] = useState("");
  const [category, setCategory] = useState(CATEGORIES[0]);
  const [method, setMethod] = useState(METHODS[0]);
  const [note, setNote] = useState("");
  const [title, setTitle] = useState<"Brother" | "Sister">("Brother");
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [query, setQuery] = useState("");
  const [filterCurrency, setFilterCurrency] = useState<"ALL" | "INR" | "OMR">("ALL");

  const phoneDigits = countryCode === "+91" ? 10 : 8;

  // Persist to localStorage on every change
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
    } catch {
      // ignore quota errors
    }
  }, [entries]);

  const totals = useMemo(() => {
    let inr = 0, omr = 0;
    for (const e of entries) {
      if (e.currency === "INR") inr += e.amount;
      else omr += e.amount;
    }
    return { inr, omr, count: entries.length };
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
    setEntries((prev) => [
      {
        id: crypto.randomUUID(),
        name, countryCode, phone, currency, amount: amt, category, method, note, title, date,
      },
      ...prev,
    ]);
    setName(""); setPhone(""); setAmount(""); setNote("");
  }

  function deleteEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id));
  }

  function deleteAllEntries() {
    if (!entries.length) return;
    const ok = window.confirm(
      `Delete ALL ${entries.length} record(s)? This cannot be undone.`,
    );
    if (!ok) return;
    setEntries([]);
  }

  function exportCSV() {
    const header = ["Date", "Title", "Name", "Phone", "Category", "Method", "Currency", "Amount", "Note"];
    const rows = entries.map((e) => [e.date, e.title ?? "", e.name, `${e.countryCode} ${e.phone}`, e.category, e.method, e.currency, e.amount, e.note]);
    const csv = [header, ...rows].map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `tithe-register-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click(); URL.revokeObjectURL(url);
  }

  return (
    <div style={styles.page}>
      <div style={styles.container}>
        <header style={styles.header}>
          <div style={styles.cross}>✟</div>
          <div>
            <h1 style={styles.h1}>Christian Faith Assembly</h1>
            <p style={styles.subtitle}>Tithe Register Form</p>
          </div>
        </header>

        <section style={styles.dashboard}>
          <StatCard label="Total Indian Rupee (INR)" value={`₹${totals.inr.toFixed(2)}`} accent="#10b981" />
          <StatCard label="Total Omani Rial (OMR)" value={`${totals.omr.toFixed(3)} OMR`} accent="#0ea5e9" />
          <StatCard label="Total Entries" value={String(totals.count)} accent="#6366f1" />
        </section>

        <form onSubmit={handleSubmit} style={styles.form}>
          <div style={styles.formGrid}>
            <Field label="Date">
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={styles.input} required />
            </Field>
            <Field label="Person's Name">
              <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Full name" style={styles.input} required />
            </Field>
            <Field label="Phone Number">
              <div style={{ display: "flex", gap: 8 }}>
                <select
                  value={countryCode}
                  onChange={(e) => { setCountryCode(e.target.value as "+91" | "+968"); setPhone(""); }}
                  style={{ ...styles.input, width: 90, flexShrink: 0 }}
                >
                  <option value="+968">+968</option>
                  <option value="+91">+91</option>
                </select>
                <input
                  type="tel"
                  value={phone}
                  onChange={(e) => {
                    const digits = e.target.value.replace(/\D/g, "").slice(0, phoneDigits);
                    setPhone(digits);
                  }}
                  placeholder={`${phoneDigits} digits`}
                  style={styles.input}
                  required
                />
              </div>
            </Field>
            <Field label="Category">
              <select value={category} onChange={(e) => setCategory(e.target.value)} style={styles.input}>
                {CATEGORIES.map((c) => <option key={c}>{c}</option>)}
              </select>
            </Field>
          </div>

          <div style={styles.thickDivider} />

          <div style={styles.formGrid}>
            <Field label="Payment Method">
              <select value={method} onChange={(e) => setMethod(e.target.value)} style={styles.input}>
                {METHODS.map((m) => <option key={m}>{m}</option>)}
              </select>
            </Field>
            <Field label="Currency">
              <select value={currency} onChange={(e) => setCurrency(e.target.value as "INR" | "OMR")} style={styles.input}>
                <option value="INR">Rupee (₹)</option>
                <option value="OMR">Rial (OMR)</option>
              </select>
            </Field>
            <Field label="Amount">
              <input type="number" min="0.001" step="any" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0.00" style={styles.input} required />
            </Field>
            <Field label="Note (optional)">
              <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Reference / remark" style={styles.input} />
            </Field>
            <Field label="Title">
              <select value={title} onChange={(e) => setTitle(e.target.value as "Brother" | "Sister")} style={styles.input}>
                {TITLES.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </Field>
          </div>

          <div style={styles.actions}>
            <button type="submit" style={styles.primaryBtn}>＋ Save Record</button>
            <button type="button" onClick={exportCSV} style={styles.secondaryBtn} disabled={!entries.length}>⤓ Export CSV</button>
            <button type="button" onClick={deleteAllEntries} style={styles.dangerBtn} disabled={!entries.length}>🗑 Delete All</button>
          </div>
        </form>

        <div style={styles.tableHeader}>
          <h2 style={styles.h2}>Logged Transactions</h2>
          <div style={styles.tableTools}>
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search name, phone, category…" style={{ ...styles.input, width: 240 }} />
            <select value={filterCurrency} onChange={(e) => setFilterCurrency(e.target.value as "ALL" | "INR" | "OMR")} style={{ ...styles.input, width: 140 }}>
              <option value="ALL">All currencies</option>
              <option value="INR">INR only</option>
              <option value="OMR">OMR only</option>
            </select>
          </div>
        </div>

        <div style={styles.tableWrap}>
          <table style={styles.table}>
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
                <tr><td colSpan={9} style={styles.empty}>No records yet — log your first entry above.</td></tr>
              ) : filtered.map((e) => (
                <tr key={e.id} style={styles.row}>
                  <td style={styles.td}>{e.date}</td>
                  <td style={styles.td}>{e.title ?? "—"}</td>
                  <td style={{ ...styles.td, fontWeight: 600 }}>{e.name}</td>
                  <td style={{ ...styles.td, ...styles.splitCol }}>{e.phone}</td>
                  <td style={styles.td}><span style={styles.pill}>{e.category}</span></td>
                  <td style={styles.td}>{e.method}</td>
                  <td style={styles.td}>{e.currency}</td>
                  <td style={{ ...styles.td, fontWeight: 700, color: "#059669" }}>{formatAmount(e.currency, e.amount)}</td>
                  <td style={styles.td}>
                    <button onClick={() => deleteEntry(e.id)} style={styles.deleteBtn} aria-label="Delete">✕</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={styles.bottomBar}>
          <button type="button" onClick={deleteAllEntries} style={styles.dangerBtn} disabled={!entries.length}>
            🗑 Delete All Records
          </button>
        </div>
      </div>
    </div>
  );
}

function StatCard({ label, value, accent }: { label: string; value: string; accent: string }) {
  return (
    <div style={{ ...styles.statBox, borderTop: `4px solid ${accent}` }}>
      <div style={styles.statLabel}>{label}</div>
      <div style={{ ...styles.statValue, color: accent }}>{value}</div>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      <label style={styles.label}>{label}</label>
      {children}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    background: "linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%)",
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
    display: "flex",
    alignItems: "center",
    gap: 18,
    paddingBottom: 24,
    borderBottom: "2px solid #e2e8f0",
    marginBottom: 28,
  },
  cross: {
    width: 60, height: 60, borderRadius: 14,
    background: "linear-gradient(135deg, #6366f1, #8b5cf6)",
    color: "white", display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: 32, fontWeight: 700, boxShadow: "0 8px 20px -8px #6366f1",
  },
  h1: { margin: 0, fontSize: 34, fontWeight: 800, letterSpacing: "-0.02em", color: "#0f172a" },
  subtitle: { margin: "4px 0 0", fontSize: 14, fontWeight: 700, color: "#64748b", letterSpacing: "0.08em", textTransform: "uppercase" },
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
    height: 3, background: "linear-gradient(90deg, transparent, #6366f1, transparent)",
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
    background: "linear-gradient(135deg, #10b981, #059669)", color: "white",
    border: "none", padding: "12px 24px", borderRadius: 8, fontWeight: 700,
    cursor: "pointer", fontSize: 14, boxShadow: "0 6px 14px -4px #10b981",
  },
  secondaryBtn: {
    background: "white", color: "#0f172a", border: "1px solid #cbd5e1",
    padding: "12px 20px", borderRadius: 8, fontWeight: 600, cursor: "pointer", fontSize: 14,
  },
  dangerBtn: {
    background: "linear-gradient(135deg, #ef4444, #b91c1c)", color: "white",
    border: "none", padding: "12px 20px", borderRadius: 8, fontWeight: 700,
    cursor: "pointer", fontSize: 14, boxShadow: "0 6px 14px -4px #ef4444",
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
  splitCol: { borderRight: "3px solid #6366f1" },
  row: { transition: "background 0.15s" },
  pill: {
    display: "inline-block", padding: "3px 10px", borderRadius: 999,
    background: "#ede9fe", color: "#6d28d9", fontSize: 12, fontWeight: 600,
  },
  deleteBtn: {
    background: "#fee2e2", color: "#dc2626", border: "none",
    width: 30, height: 30, borderRadius: 8, cursor: "pointer", fontWeight: 700,
  },
  empty: { padding: 36, textAlign: "center", color: "#94a3b8", fontStyle: "italic" },
};
