"use client";

import { ChangeEvent, useEffect, useMemo, useState } from "react";
import { exportStyledWorkbook, generateTransactions, GenerationConstraintError } from "@/lib/generator";
import { exportStatementDocx, exportStatementPdf } from "@/lib/statement-export";
import { SpecialTransactionInput, TransactionRow } from "@/lib/types";

const today = new Date().toISOString().slice(0, 10);
const suggestionsStorageKey = "tg-form-suggestions";
const historyStorageKey = "tg-statement-history";

type SuggestionsStore = {
  customerNames: string[];
  salaryCompanies: string[];
  counterparties: string[];
  suffixes: string[];
};

type FormSnapshot = {
  customerName: string;
  startDate: string;
  closingDate: string;
  openingBalance: number;
  targetClosingBalance: number;
  minimumBalance: number;
  maximumBalance: number;
  minIncomingAmount: number;
  maxIncomingAmount: number;
  yorubaNameRatio: number;
  igboNameRatio: number;
  hausaNameRatio: number;
  otherNameRatio: number;
  maxNameUses: number;
  minDaysBeforeNameReuse: number;
  repeatableNameCount: number;
  includeSalary: boolean;
  salaryAmount: number;
  salaryDay: number;
  salaryCompanyName: string;
  minTransactionsPerMonth: number;
  maxTransactionsPerMonth: number;
  specialTransactions: SpecialTransactionInput[];
};

type HistoryRecord = {
  id: string;
  savedAt: string;
  form: FormSnapshot;
  rows: TransactionRow[];
};

type ConstraintModalState = {
  title: string;
  message: string;
  suggestions: string[];
};

type SavePickerAcceptType = {
  description?: string;
  accept: Record<string, string[]>;
};

type SaveFilePickerOptions = {
  suggestedName?: string;
  types?: SavePickerAcceptType[];
};

type SaveFilePickerHandle = {
  createWritable: () => Promise<{
    write: (data: Blob) => Promise<void>;
    close: () => Promise<void>;
  }>;
};

const emptySuggestions: SuggestionsStore = {
  customerNames: [],
  salaryCompanies: [],
  counterparties: [],
  suffixes: [],
};

function readStoredJson<T>(key: string, fallback: T): T {
  if (typeof window === "undefined") {
    return fallback;
  }

  const rawValue = window.localStorage.getItem(key);
  if (!rawValue) {
    return fallback;
  }

  try {
    return JSON.parse(rawValue) as T;
  } catch {
    return fallback;
  }
}

function persistHistoryRecords(records: HistoryRecord[]): HistoryRecord[] {
  if (typeof window === "undefined") {
    return records;
  }

  for (let size = records.length; size >= 0; size -= 1) {
    const nextRecords = records.slice(0, size);

    try {
      window.localStorage.setItem(historyStorageKey, JSON.stringify(nextRecords));
      return nextRecords;
    } catch {
      continue;
    }
  }

  return [];
}

function formatNumberWithCommas(value: number) {
  return value.toLocaleString("en-NG", {
    minimumFractionDigits: Number.isInteger(value) ? 0 : 2,
    maximumFractionDigits: 2,
  });
}

function parseFormattedNumber(value: string) {
  const sanitized = value
    .replace(/,/g, "")
    .replace(/[^\d.]/g, "")
    .replace(/(\..*)\./g, "$1");
  return sanitized && sanitized !== "." ? Number(sanitized) : 0;
}

function normalizeNumericInput(value: string) {
  const sanitized = value
    .replace(/,/g, "")
    .replace(/[^\d.]/g, "")
    .replace(/(\..*)\./g, "$1");

  if (!sanitized) {
    return "";
  }

  const [wholePart, decimalPart] = sanitized.split(".");
  const normalizedWhole = wholePart ? Number(wholePart).toLocaleString("en-NG") : "0";

  if (sanitized.endsWith(".")) {
    return `${normalizedWhole}.`;
  }

  if (decimalPart !== undefined) {
    return `${normalizedWhole}.${decimalPart.slice(0, 2)}`;
  }

  return normalizedWhole;
}

function uniqueRecent(items: string[], limit = 12) {
  return Array.from(new Set(items.map((item) => item.trim()).filter(Boolean))).slice(0, limit);
}

function sanitizeFilenamePart(value: string) {
  const cleaned = value
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  return cleaned || "statement";
}

function buildExportBaseName(customerName: string) {
  return sanitizeFilenamePart(customerName) || "statement";
}

function NumericInput({
  value,
  onValueChange,
  placeholder,
}: {
  value: number;
  onValueChange: (value: number) => void;
  placeholder?: string;
}) {
  const [displayValue, setDisplayValue] = useState(() => formatNumberWithCommas(value));

  useEffect(() => {
    setDisplayValue(formatNumberWithCommas(value));
  }, [value]);

  return (
    <input
      type="text"
      inputMode="decimal"
      value={displayValue}
      placeholder={placeholder}
      onChange={(event) => {
        const nextDisplay = normalizeNumericInput(event.target.value);
        setDisplayValue(nextDisplay);
        onValueChange(parseFormattedNumber(nextDisplay));
      }}
      onBlur={() => setDisplayValue(formatNumberWithCommas(parseFormattedNumber(displayValue)))}
    />
  );
}

function createSpecialTransaction(): SpecialTransactionInput {
  return {
    id: crypto.randomUUID(),
    suffix: "",
    amount: 0,
    kind: "debit",
    mode: "transfer_out",
    counterpartyName: "",
    date: today,
  };
}

function cloneSpecialTransactions(items: SpecialTransactionInput[]): SpecialTransactionInput[] {
  return items.map((item) => ({ ...item }));
}

async function saveBlob(
  blob: Blob,
  filename: string,
  fileType?: { description: string; mimeType: string; extension: string },
) {
  const pickerWindow = window as Window & {
    showSaveFilePicker?: (options?: SaveFilePickerOptions) => Promise<SaveFilePickerHandle>;
  };

  if (pickerWindow.showSaveFilePicker && fileType) {
    try {
      const handle = await pickerWindow.showSaveFilePicker({
        suggestedName: filename,
        types: [
          {
            description: fileType.description,
            accept: {
              [fileType.mimeType]: [fileType.extension],
            },
          },
        ],
      });

      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return;
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        return;
      }
    }
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

export default function Home() {
  const [customerName, setCustomerName] = useState("");
  const [startDate, setStartDate] = useState(today);
  const [closingDate, setClosingDate] = useState(today);
  const [openingBalance, setOpeningBalance] = useState(250000);
  const [targetClosingBalance, setTargetClosingBalance] = useState(600000);
  const [minimumBalance, setMinimumBalance] = useState(200000);
  const [maximumBalance, setMaximumBalance] = useState(800000);
  const [minIncomingAmount, setMinIncomingAmount] = useState(12000);
  const [maxIncomingAmount, setMaxIncomingAmount] = useState(220000);
  const [yorubaNameRatio, setYorubaNameRatio] = useState(60);
  const [igboNameRatio, setIgboNameRatio] = useState(20);
  const [hausaNameRatio, setHausaNameRatio] = useState(10);
  const [otherNameRatio, setOtherNameRatio] = useState(10);
  const [maxNameUses, setMaxNameUses] = useState(2);
  const [minDaysBeforeNameReuse, setMinDaysBeforeNameReuse] = useState(7);
  const [repeatableNameCount, setRepeatableNameCount] = useState(4);
  const [includeSalary, setIncludeSalary] = useState(true);
  const [salaryAmount, setSalaryAmount] = useState(320000);
  const [salaryDay, setSalaryDay] = useState(28);
  const [salaryCompanyName, setSalaryCompanyName] = useState("Davikosi Nigeria Limited");
  const [minTransactionsPerMonth, setMinTransactionsPerMonth] = useState(12);
  const [maxTransactionsPerMonth, setMaxTransactionsPerMonth] = useState(22);
  const [specialTransactions, setSpecialTransactions] = useState<SpecialTransactionInput[]>([createSpecialTransaction()]);
  const [exportFileName, setExportFileName] = useState("");
  const [rows, setRows] = useState<TransactionRow[]>([]);
  const [generationError, setGenerationError] = useState("");
  const [constraintModal, setConstraintModal] = useState<ConstraintModalState | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [isExportingDocx, setIsExportingDocx] = useState(false);
  const [isExportingExcel, setIsExportingExcel] = useState(false);
  const [suggestions, setSuggestions] = useState<SuggestionsStore>(emptySuggestions);
  const [history, setHistory] = useState<HistoryRecord[]>([]);
  const [hasLoadedStorage, setHasLoadedStorage] = useState(false);
  const [isHistoryModalOpen, setIsHistoryModalOpen] = useState(false);
  const [historySearch, setHistorySearch] = useState("");

  const months = useMemo(() => {
    const start = new Date(`${startDate}T00:00:00`);
    const end = new Date(`${closingDate}T00:00:00`);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end < start) {
      return 1;
    }

    return Math.max(
      1,
      (end.getFullYear() - start.getFullYear()) * 12 +
        (end.getMonth() - start.getMonth()) +
        1,
    );
  }, [closingDate, startDate]);

  const filteredHistory = useMemo(() => {
    const searchTerm = historySearch.trim().toLowerCase();
    if (!searchTerm) {
      return history;
    }

    return history.filter((record) => {
      const haystack = [
        record.form.customerName,
        record.form.salaryCompanyName,
        record.form.startDate,
        record.form.closingDate,
        ...record.form.specialTransactions.map((item) => item.counterpartyName),
        ...record.form.specialTransactions.map((item) => item.suffix),
      ]
        .join(" ")
        .toLowerCase();

      return haystack.includes(searchTerm);
    });
  }, [history, historySearch]);

  useEffect(() => {
    setSuggestions({
      ...emptySuggestions,
      ...readStoredJson<SuggestionsStore>(suggestionsStorageKey, emptySuggestions),
    });
    setHistory(readStoredJson<HistoryRecord[]>(historyStorageKey, []));
    setHasLoadedStorage(true);
  }, []);

  useEffect(() => {
    if (!hasLoadedStorage) {
      return;
    }

    try {
      window.localStorage.setItem(suggestionsStorageKey, JSON.stringify(suggestions));
    } catch {
      // Ignore storage write failures so the UI remains usable.
    }
  }, [hasLoadedStorage, suggestions]);

  useEffect(() => {
    if (!hasLoadedStorage) {
      return;
    }

    const persistedRecords = persistHistoryRecords(history);
    if (persistedRecords.length !== history.length) {
      setHistory(persistedRecords);
    }
  }, [hasLoadedStorage, history]);

  function createFormSnapshot(): FormSnapshot {
    return {
      customerName,
      startDate,
      closingDate,
      openingBalance,
      targetClosingBalance,
      minimumBalance,
      maximumBalance,
      minIncomingAmount,
      maxIncomingAmount,
      yorubaNameRatio,
      igboNameRatio,
      hausaNameRatio,
      otherNameRatio,
      maxNameUses,
      minDaysBeforeNameReuse,
      repeatableNameCount,
      includeSalary,
      salaryAmount,
      salaryDay,
      salaryCompanyName,
      minTransactionsPerMonth,
      maxTransactionsPerMonth,
      specialTransactions: cloneSpecialTransactions(specialTransactions),
    };
  }

  function buildGeneratorInput(snapshot: FormSnapshot) {
    const snapshotStart = new Date(`${snapshot.startDate}T00:00:00`);
    const snapshotEnd = new Date(`${snapshot.closingDate}T00:00:00`);
    const snapshotMonths =
      Number.isNaN(snapshotStart.getTime()) ||
      Number.isNaN(snapshotEnd.getTime()) ||
      snapshotEnd < snapshotStart
        ? 1
        : Math.max(
          1,
          (snapshotEnd.getFullYear() - snapshotStart.getFullYear()) * 12 +
            (snapshotEnd.getMonth() - snapshotStart.getMonth()) +
            1,
        );

    return {
      customerName: snapshot.customerName.trim() || "Customer",
      months: snapshotMonths,
      startDate: snapshot.startDate,
      closingDate: snapshot.closingDate,
      namePool: [],
      yorubaNameRatio: snapshot.yorubaNameRatio,
      igboNameRatio: snapshot.igboNameRatio,
      hausaNameRatio: snapshot.hausaNameRatio,
      otherNameRatio: snapshot.otherNameRatio,
      openingBalance: snapshot.openingBalance,
      targetClosingBalance: snapshot.targetClosingBalance,
      minimumBalance: snapshot.minimumBalance,
      maximumBalance: snapshot.maximumBalance,
      minIncomingAmount: snapshot.minIncomingAmount,
      maxIncomingAmount: snapshot.maxIncomingAmount,
      maxNameUses: snapshot.maxNameUses,
      minDaysBeforeNameReuse: snapshot.minDaysBeforeNameReuse,
      repeatableNameCount: snapshot.repeatableNameCount,
      includeSalary: snapshot.includeSalary,
      salaryAmount: snapshot.salaryAmount,
      salaryDay: snapshot.salaryDay,
      salaryCompanyName: snapshot.salaryCompanyName,
      minTransactionsPerMonth: snapshot.minTransactionsPerMonth,
      maxTransactionsPerMonth: snapshot.maxTransactionsPerMonth,
      specialTransactions: cloneSpecialTransactions(snapshot.specialTransactions),
    };
  }

  function loadSnapshot(snapshot: FormSnapshot, nextRows: TransactionRow[]) {
    setCustomerName(snapshot.customerName);
    setStartDate(snapshot.startDate);
    setClosingDate(snapshot.closingDate);
    setOpeningBalance(snapshot.openingBalance);
    setTargetClosingBalance(snapshot.targetClosingBalance);
    setMinimumBalance(snapshot.minimumBalance);
    setMaximumBalance(snapshot.maximumBalance);
    setMinIncomingAmount(snapshot.minIncomingAmount);
    setMaxIncomingAmount(snapshot.maxIncomingAmount);
    setYorubaNameRatio(snapshot.yorubaNameRatio);
    setIgboNameRatio(snapshot.igboNameRatio);
    setHausaNameRatio(snapshot.hausaNameRatio);
    setOtherNameRatio(snapshot.otherNameRatio);
    setMaxNameUses(snapshot.maxNameUses);
    setMinDaysBeforeNameReuse(snapshot.minDaysBeforeNameReuse);
    setRepeatableNameCount(snapshot.repeatableNameCount);
    setIncludeSalary(snapshot.includeSalary);
    setSalaryAmount(snapshot.salaryAmount);
    setSalaryDay(snapshot.salaryDay);
    setSalaryCompanyName(snapshot.salaryCompanyName);
    setMinTransactionsPerMonth(snapshot.minTransactionsPerMonth);
    setMaxTransactionsPerMonth(snapshot.maxTransactionsPerMonth);
    setSpecialTransactions(snapshot.specialTransactions.length > 0 ? snapshot.specialTransactions : [createSpecialTransaction()]);
    setRows(nextRows);
  }

  function rememberSuggestions(snapshot: FormSnapshot) {
    setSuggestions((current) => ({
      customerNames: uniqueRecent([snapshot.customerName, ...current.customerNames]),
      salaryCompanies: uniqueRecent([snapshot.salaryCompanyName, ...current.salaryCompanies]),
      counterparties: uniqueRecent([
        ...snapshot.specialTransactions.map((item) => item.counterpartyName),
        ...current.counterparties,
      ]),
      suffixes: uniqueRecent([
        ...snapshot.specialTransactions.map((item) => item.suffix),
        ...current.suffixes,
      ]),
    }));
  }

  function updateSpecialTransaction(id: string, field: keyof SpecialTransactionInput, value: string) {
    setSpecialTransactions((current) =>
      current.map((item) => {
        if (item.id !== id) {
          return item;
        }

        if (field === "amount") {
          return { ...item, amount: Number(value) };
        }

        if (field === "kind") {
          return { ...item, kind: value as SpecialTransactionInput["kind"] };
        }

        if (field === "mode") {
          const mode = value as SpecialTransactionInput["mode"];
          return {
            ...item,
            mode,
            kind: mode === "transfer_in" || mode === "salary" ? "credit" : "debit",
          };
        }

        return { ...item, [field]: value };
      }),
    );
  }

  function handleGenerate() {
    setIsGenerating(true);
    setGenerationError("");
    setConstraintModal(null);
    try {
      const snapshot = createFormSnapshot();
      const generated = generateTransactions(buildGeneratorInput(snapshot));
      const nextRecord: HistoryRecord = {
        id: crypto.randomUUID(),
        savedAt: new Date().toISOString(),
        form: snapshot,
        rows: generated,
      };

      setRows(generated);
      rememberSuggestions(snapshot);
      setHistory((current) => [nextRecord, ...current].slice(0, 50));
    } catch (error) {
      if (error instanceof GenerationConstraintError) {
        setGenerationError(error.message);
        setConstraintModal({
          title: error.title,
          message: error.message,
          suggestions: error.suggestions,
        });
      } else {
        setGenerationError("Unable to generate transactions with the current configuration.");
      }
    } finally {
      setIsGenerating(false);
    }
  }

  function handleNumberChange(setter: (value: number) => void) {
    return (event: ChangeEvent<HTMLInputElement>) => {
      setter(Number(event.target.value));
    };
  }

  function loadHistoryRecord(record: HistoryRecord) {
    loadSnapshot(record.form, record.rows);
    rememberSuggestions(record.form);
    setIsHistoryModalOpen(false);
    setHistorySearch("");
  }

  async function handleExportPdf() {
    if (rows.length === 0) {
      return;
    }

    setIsExportingPdf(true);
    try {
      const blob = await exportStatementPdf(rows, {
        customerName,
        startDate,
        closingDate,
      });
      await saveBlob(blob, `${sanitizeFilenamePart(exportFileName || buildExportBaseName(customerName))}.pdf`, {
        description: "PDF Document",
        mimeType: "application/pdf",
        extension: ".pdf",
      });
    } finally {
      setIsExportingPdf(false);
    }
  }

  async function handleExportDocx() {
    if (rows.length === 0) {
      return;
    }

    setIsExportingDocx(true);
    try {
      const blob = await exportStatementDocx(rows, {
        customerName,
        startDate,
        closingDate,
      });
      await saveBlob(blob, `${sanitizeFilenamePart(exportFileName || buildExportBaseName(customerName))}.docx`, {
        description: "Word Document",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        extension: ".docx",
      });
    } finally {
      setIsExportingDocx(false);
    }
  }

  async function handleExportExcel() {
    if (rows.length === 0) {
      return;
    }

    setIsExportingExcel(true);
    try {
      const blob = await exportStyledWorkbook(rows);
      await saveBlob(blob, `${sanitizeFilenamePart(exportFileName || buildExportBaseName(customerName))}.xlsx`, {
        description: "Excel Workbook",
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        extension: ".xlsx",
      });
    } finally {
      setIsExportingExcel(false);
    }
  }

  return (
    <main className="page-shell">
      <section className="hero-card">
        <div className="brand-mark" aria-label="Transaction Generator logo">
          <span className="brand-badge">TG</span>
          <div className="brand-copy">
            <strong>Transaction Generator</strong>
            <span>Statement Builder</span>
          </div>
        </div>
        <div className="hero-metrics">
          <div>
            <span>Months</span>
            <strong>{months}</strong>
          </div>
          <div>
            <span>Rows Generated</span>
            <strong>{rows.length}</strong>
          </div>
        </div>
      </section>

      <section className="workspace-grid">
        <section className="panel form-panel">
          <h2>Generator Form</h2>

          <div className="field-grid two-up">
            <label>
              <span>Customer Name</span>
              <input list="customer-name-suggestions" value={customerName} onChange={(event) => setCustomerName(event.target.value)} placeholder="Enter customer name" />
            </label>
            <label>
              <span>Start Date</span>
              <input type="date" value={startDate} onChange={(event) => setStartDate(event.target.value)} />
            </label>
            <label>
              <span>Closing Date</span>
              <input type="date" value={closingDate} min={startDate} onChange={(event) => setClosingDate(event.target.value)} />
            </label>
            <label>
              <span>Months Covered</span>
              <input type="number" value={months} readOnly />
            </label>
            <label>
              <span>Opening Balance</span>
              <NumericInput value={openingBalance} onValueChange={setOpeningBalance} />
            </label>
            <label>
              <span>Target Closing Balance</span>
              <NumericInput value={targetClosingBalance} onValueChange={setTargetClosingBalance} />
            </label>
            <label>
              <span>Minimum Balance Allowed</span>
              <NumericInput value={minimumBalance} onValueChange={setMinimumBalance} />
            </label>
            <label>
              <span>Maximum Balance Allowed</span>
              <NumericInput value={maximumBalance} onValueChange={setMaximumBalance} />
            </label>
            <label>
              <span>Minimum Incoming Amount</span>
              <NumericInput value={minIncomingAmount} onValueChange={setMinIncomingAmount} />
            </label>
            <label>
              <span>Maximum Transaction Amount</span>
              <NumericInput value={maxIncomingAmount} onValueChange={setMaxIncomingAmount} />
              <p className="helper-text">Applies to every system-generated debit and credit transaction.</p>
            </label>
            <label>
              <span>Min Transactions Per Month</span>
              <input type="number" min="1" max="50" value={minTransactionsPerMonth} onChange={handleNumberChange(setMinTransactionsPerMonth)} />
            </label>
            <label>
              <span>Max Transactions Per Month</span>
              <input type="number" min="1" max="60" value={maxTransactionsPerMonth} onChange={handleNumberChange(setMaxTransactionsPerMonth)} />
            </label>
          </div>

          <div className="field-grid three-up">
            <label>
              <span>Yoruba Name Ratio</span>
              <input type="number" min="0" max="100" value={yorubaNameRatio} onChange={handleNumberChange(setYorubaNameRatio)} />
            </label>
            <label>
              <span>Igbo Name Ratio</span>
              <input type="number" min="0" max="100" value={igboNameRatio} onChange={handleNumberChange(setIgboNameRatio)} />
            </label>
            <label>
              <span>Hausa Name Ratio</span>
              <input type="number" min="0" max="100" value={hausaNameRatio} onChange={handleNumberChange(setHausaNameRatio)} />
            </label>
          </div>

          <div className="field-grid three-up">
            <label>
              <span>Other Name Ratio</span>
              <input type="number" min="0" max="100" value={otherNameRatio} onChange={handleNumberChange(setOtherNameRatio)} />
            </label>
            <label>
              <span>Maximum Uses Per Name</span>
              <input type="number" min="1" max="5" value={maxNameUses} onChange={handleNumberChange(setMaxNameUses)} />
            </label>
            <label>
              <span>Days Before Name Reuse</span>
              <input type="number" min="0" max="60" value={minDaysBeforeNameReuse} onChange={handleNumberChange(setMinDaysBeforeNameReuse)} />
            </label>
            <label>
              <span>Repeatable Names Count</span>
              <input type="number" min="0" max="16" value={repeatableNameCount} onChange={handleNumberChange(setRepeatableNameCount)} />
            </label>
          </div>

          <div className="toggle-row">
            <label className="checkbox-row">
              <input type="checkbox" checked={includeSalary} onChange={(event) => setIncludeSalary(event.target.checked)} />
              <span>Make salary appear every month</span>
            </label>
          </div>

          {includeSalary ? (
            <div className="field-grid three-up">
              <label>
                <span>Salary Amount</span>
                <NumericInput value={salaryAmount} onValueChange={setSalaryAmount} />
              </label>
              <label>
                <span>Salary Company Name</span>
                <input list="salary-company-suggestions" value={salaryCompanyName} onChange={(event) => setSalaryCompanyName(event.target.value)} placeholder="Davikosi Nigeria Limited" />
              </label>
              <label>
                <span>Salary Day Of Month</span>
                <input type="number" min="1" max="31" value={salaryDay} onChange={handleNumberChange(setSalaryDay)} />
              </label>
            </div>
          ) : null}

          <div className="specials-header">
            <div>
              <h3>Special Transactions</h3>
              <p>Add one-off transactions and choose the exact dates they should appear.</p>
            </div>
            <button type="button" className="secondary-button" onClick={() => setSpecialTransactions((current) => [...current, createSpecialTransaction()])}>
              Add Transaction
            </button>
          </div>

          <div className="specials-list">
            {specialTransactions.map((item, index) => (
              <article className="special-card" key={item.id}>
                <div className="special-card-top">
                  <strong>Special #{index + 1}</strong>
                  <button
                    type="button"
                    className="text-button"
                    onClick={() => setSpecialTransactions((current) => current.filter((entry) => entry.id !== item.id))}
                    disabled={specialTransactions.length === 1}
                  >
                    Remove
                  </button>
                </div>

                <div className="field-grid three-up">
                  <label>
                    <span>Mode</span>
                    <select value={item.mode} onChange={(event) => updateSpecialTransaction(item.id, "mode", event.target.value)}>
                      <option value="transfer_out">Transfer Out</option>
                      <option value="transfer_in">Deposit / Transfer In</option>
                      <option value="cash_withdrawal">Cash Withdrawal</option>
                      <option value="salary">Salary</option>
                    </select>
                  </label>
                  <label>
                    <span>Amount</span>
                    <NumericInput value={item.amount} onValueChange={(value) => updateSpecialTransaction(item.id, "amount", String(value))} />
                  </label>
                  <label>
                    <span>Date</span>
                    <input type="date" value={item.date} onChange={(event) => updateSpecialTransaction(item.id, "date", event.target.value)} />
                  </label>
                  <label>
                    <span>Other Party Name</span>
                    <input list="counterparty-suggestions" value={item.counterpartyName} onChange={(event) => updateSpecialTransaction(item.id, "counterpartyName", event.target.value)} placeholder="Recipient or sender name" />
                  </label>
                  <label>
                    <span>Type</span>
                    <select value={item.kind} onChange={(event) => updateSpecialTransaction(item.id, "kind", event.target.value)}>
                      <option value="debit">Debit</option>
                      <option value="credit">Credit</option>
                    </select>
                  </label>
                  <label>
                    <span>Suffix</span>
                    <input list="suffix-suggestions" value={item.suffix} onChange={(event) => updateSpecialTransaction(item.id, "suffix", event.target.value)} placeholder="Example: expense" />
                  </label>
                </div>
              </article>
            ))}
          </div>

          <datalist id="customer-name-suggestions">
            {suggestions.customerNames.map((item) => <option key={item} value={item} />)}
          </datalist>
          <datalist id="salary-company-suggestions">
            {suggestions.salaryCompanies.map((item) => <option key={item} value={item} />)}
          </datalist>
          <datalist id="counterparty-suggestions">
            {suggestions.counterparties.map((item) => <option key={item} value={item} />)}
          </datalist>
          <datalist id="suffix-suggestions">
            {suggestions.suffixes.map((item) => <option key={item} value={item} />)}
          </datalist>

          <div className="field-grid two-up">
            <label>
              <span>Export File Name</span>
              <input
                value={exportFileName}
                onChange={(event) => setExportFileName(event.target.value)}
                placeholder={buildExportBaseName(customerName)}
              />
              <p className="helper-text">You can still change the folder and final filename in the save dialog.</p>
            </label>
          </div>

          <div className="action-row">
            <button type="button" className="primary-button" onClick={handleGenerate} disabled={isGenerating}>
              {isGenerating ? (
                <span className="button-loader">
                  <span className="spinner" />
                  Generating...
                </span>
              ) : "Generate Transactions"}
            </button>
            <button type="button" className="secondary-button" onClick={handleExportPdf} disabled={rows.length === 0 || isExportingPdf}>
              {isExportingPdf ? "Preparing PDF..." : "Export PDF"}
            </button>
            <button type="button" className="secondary-button" onClick={handleExportDocx} disabled={rows.length === 0 || isExportingDocx}>
              {isExportingDocx ? "Preparing DOCX..." : "Export DOCX"}
            </button>
            <button type="button" className="secondary-button" onClick={handleExportExcel} disabled={rows.length === 0 || isExportingExcel}>
              {isExportingExcel ? "Preparing Excel..." : "Export Excel"}
            </button>
          </div>

          {generationError ? (
            <p className="empty-state">{generationError}</p>
          ) : null}
        </section>

        <section className="panel preview-panel">
          <div className="preview-header">
            <div>
              <h2>Generated Preview</h2>
              <p>{customerName || "Customer"} statement preview based on your form inputs.</p>
            </div>
            <button type="button" className="secondary-button" onClick={() => setIsHistoryModalOpen(true)}>
              Recent Records
            </button>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Transaction Details</th>
                  <th>Debit Amount</th>
                  <th>Credit Amount</th>
                  <th>Balance</th>
                </tr>
              </thead>
              <tbody>
                {rows.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="empty-state">Generate transactions to see the statement preview here.</td>
                  </tr>
                ) : (
                  rows.map((row) => (
                    <tr key={row.id}>
                      <td>{row.date}</td>
                      <td>{row.description}</td>
                      <td className="debit-cell">{row.debit.toLocaleString("en-NG", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                      <td className="credit-cell">{row.credit.toLocaleString("en-NG", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                      <td className="balance-cell">{row.balance.toLocaleString("en-NG", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>
      </section>

      {isHistoryModalOpen ? (
        <div className="modal-backdrop" onClick={() => setIsHistoryModalOpen(false)}>
          <div className="modal-card" onClick={(event) => event.stopPropagation()}>
            <div className="modal-header">
              <div>
                <h3>Recent Records</h3>
                <p>Search and reload saved statements without retyping.</p>
              </div>
              <button type="button" className="text-button" onClick={() => setIsHistoryModalOpen(false)}>
                Close
              </button>
            </div>

            <input
              className="modal-search"
              value={historySearch}
              onChange={(event) => setHistorySearch(event.target.value)}
              placeholder="Search by customer, company, date, counterparty, or suffix"
            />

            {filteredHistory.length > 0 ? (
              <div className="history-list modal-history-list">
                {filteredHistory.map((record) => (
                  <article className="history-card" key={record.id}>
                    <div>
                      <strong>{record.form.customerName || "Customer"}</strong>
                      <p>{record.form.startDate} to {record.form.closingDate}</p>
                    </div>
                    <button type="button" className="secondary-button" onClick={() => loadHistoryRecord(record)}>
                      Load Record
                    </button>
                  </article>
                ))}
              </div>
            ) : (
              <p className="empty-history">
                {history.length === 0 ? "No saved records yet. Generate one statement and it will appear here." : "No record matches your search."}
              </p>
            )}
          </div>
        </div>
      ) : null}

      {constraintModal ? (
        <div className="modal-backdrop" onClick={() => setConstraintModal(null)}>
          <div className="modal-card" onClick={(event) => event.stopPropagation()}>
            <div className="modal-header">
              <div>
                <h3>{constraintModal.title}</h3>
                <p>{constraintModal.message}</p>
              </div>
              <button type="button" className="text-button" onClick={() => setConstraintModal(null)}>
                Close
              </button>
            </div>

            {constraintModal.suggestions.length > 0 ? (
              <div className="history-list">
                {constraintModal.suggestions.map((suggestion) => (
                  <article className="history-card" key={suggestion}>
                    <p>{suggestion}</p>
                  </article>
                ))}
              </div>
            ) : null}
          </div>
        </div>
      ) : null}
    </main>
  );
}
