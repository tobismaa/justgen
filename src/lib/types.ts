export type EntryKind = "credit" | "debit";
export type TransactionMode = "transfer_out" | "transfer_in" | "cash_withdrawal" | "salary";

export type SpecialTransactionInput = {
  id: string;
  description: string;
  amount: number;
  kind: EntryKind;
  mode: TransactionMode;
  counterpartyName: string;
  date: string;
};

export type GeneratorInput = {
  customerName: string;
  months: number;
  startDate: string;
  closingDate: string;
  namePool: string[];
  openingBalance: number;
  targetClosingBalance: number;
  minimumBalance: number;
  maximumBalance: number;
  minIncomingAmount: number;
  maxIncomingAmount: number;
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

export type TransactionRow = {
  id: string;
  date: string;
  description: string;
  debit: number;
  credit: number;
  balance: number;
};
