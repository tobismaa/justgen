"use client";

import { ChangeEvent, useMemo, useState } from "react";
import { exportStyledWorkbook, generateTransactions } from "@/lib/generator";
import { SpecialTransactionInput, TransactionRow } from "@/lib/types";

const today = new Date().toISOString().slice(0, 10);

function createSpecialTransaction(): SpecialTransactionInput {
  return {
    id: crypto.randomUUID(),
    description: "",
    amount: 0,
    kind: "debit",
    mode: "transfer_out",
    counterpartyName: "",
    date: today,
  };
}

async function downloadWorkbook(rows: TransactionRow[]) {
  const blob = await exportStyledWorkbook(rows);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "generated-transactions.xlsx";
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
  const [rows, setRows] = useState<TransactionRow[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);

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

    window.setTimeout(() => {
      const generated = generateTransactions({
        customerName: customerName.trim() || "Customer",
        months,
        startDate,
        closingDate,
        namePool: [],
        openingBalance,
        targetClosingBalance,
        minimumBalance,
        maximumBalance,
        minIncomingAmount,
        maxIncomingAmount,
        maxNameUses,
        minDaysBeforeNameReuse,
        repeatableNameCount,
        includeSalary,
        salaryAmount,
        salaryDay,
        salaryCompanyName,
        minTransactionsPerMonth,
        maxTransactionsPerMonth,
        specialTransactions,
      });

      setRows(generated);
      setIsGenerating(false);
    }, 150);
  }

  function handleNumberChange(setter: (value: number) => void) {
    return (event: ChangeEvent<HTMLInputElement>) => {
      setter(Number(event.target.value));
    };
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
              <input value={customerName} onChange={(event) => setCustomerName(event.target.value)} placeholder="Enter customer name" />
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
              <input type="number" min="0" value={openingBalance} onChange={handleNumberChange(setOpeningBalance)} />
            </label>
            <label>
              <span>Target Closing Balance</span>
              <input type="number" min="0" value={targetClosingBalance} onChange={handleNumberChange(setTargetClosingBalance)} />
            </label>
            <label>
              <span>Minimum Balance Allowed</span>
              <input type="number" min="0" value={minimumBalance} onChange={handleNumberChange(setMinimumBalance)} />
            </label>
            <label>
              <span>Maximum Balance Allowed</span>
              <input type="number" min="0" value={maximumBalance} onChange={handleNumberChange(setMaximumBalance)} />
            </label>
            <label>
              <span>Minimum Incoming Amount</span>
              <input type="number" min="0" value={minIncomingAmount} onChange={handleNumberChange(setMinIncomingAmount)} />
            </label>
            <label>
              <span>Maximum Incoming Amount</span>
              <input type="number" min="0" value={maxIncomingAmount} onChange={handleNumberChange(setMaxIncomingAmount)} />
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
                <input type="number" min="0" value={salaryAmount} onChange={handleNumberChange(setSalaryAmount)} />
              </label>
              <label>
                <span>Salary Company Name</span>
                <input value={salaryCompanyName} onChange={(event) => setSalaryCompanyName(event.target.value)} placeholder="Davikosi Nigeria Limited" />
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
                    <input type="number" min="0" value={item.amount || ""} onChange={(event) => updateSpecialTransaction(item.id, "amount", event.target.value)} />
                  </label>
                  <label>
                    <span>Date</span>
                    <input type="date" value={item.date} onChange={(event) => updateSpecialTransaction(item.id, "date", event.target.value)} />
                  </label>
                  <label>
                    <span>Other Party Name</span>
                    <input value={item.counterpartyName} onChange={(event) => updateSpecialTransaction(item.id, "counterpartyName", event.target.value)} placeholder="Recipient or sender name" />
                  </label>
                  <label>
                    <span>Type</span>
                    <select value={item.kind} onChange={(event) => updateSpecialTransaction(item.id, "kind", event.target.value)}>
                      <option value="debit">Debit</option>
                      <option value="credit">Credit</option>
                    </select>
                  </label>
                  <label>
                    <span>Custom Narration</span>
                    <input value={item.description} onChange={(event) => updateSpecialTransaction(item.id, "description", event.target.value)} placeholder="Leave blank to auto-use sample narration" />
                  </label>
                </div>
              </article>
            ))}
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
            <button type="button" className="secondary-button" onClick={() => downloadWorkbook(rows)} disabled={rows.length === 0}>Export Excel</button>
          </div>
        </section>

        <section className="panel preview-panel">
          <div className="preview-header">
            <div>
              <h2>Generated Preview</h2>
              <p>{customerName || "Customer"} statement preview based on your form inputs.</p>
            </div>
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
    </main>
  );
}
