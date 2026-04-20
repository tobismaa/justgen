import ExcelJS from "exceljs";
import { GeneratorInput, SpecialTransactionInput, TransactionMode, TransactionRow } from "./types";

type SeedRow = {
  id: string;
  date: string;
  description: string;
  debit: number;
  credit: number;
  group: number;
  order: number;
};

type NameUsageMeta = {
  lastUsedDayKey: number | null;
};

type GeneratedNameState = {
  usedFullNames: Set<string>;
  usageMap: Map<string, NameUsageMeta>;
};

function uniqueValues(items: string[]): string[] {
  return Array.from(new Set(items.map((item) => item.trim()).filter(Boolean)));
}

function buildCombinations(starts: string[], ends: string[]): string[] {
  const items: string[] = [];

  for (const start of starts) {
    for (const end of ends) {
      items.push(`${start}${end}`);
    }
  }

  return items;
}

const yorubaFirstNames = uniqueValues([
  "Adebayo",
  "Adebola",
  "Adekunle",
  "Adenike",
  "Aderemi",
  "Aderonke",
  "Adewale",
  "Adewunmi",
  "Afolabi",
  "Akinade",
  "Akinlabi",
  "Akinwale",
  "Akinyemi",
  "Anjola",
  "Ayobami",
  "Ayodeji",
  "Ayomide",
  "Ayomikun",
  "Ayotola",
  "Bimpe",
  "Bisola",
  "Bolanle",
  "Damilare",
  "Damilola",
  "Dapo",
  "Dare",
  "Dele",
  "Dupe",
  "Eniola",
  "Femi",
  "Folake",
  "Folasade",
  "Funke",
  "Gbemi",
  "Idowu",
  "Jide",
  "Jumoke",
  "Kehinde",
  "Kemi",
  "Korede",
  "Kunle",
  "Lekan",
  "Mide",
  "Mobolaji",
  "Mofe",
  "Mojisola",
  "Morenike",
  "Moyinoluwa",
  "Olabisi",
  "Oladapo",
  "Oladele",
  "Oladimeji",
  "Oladipo",
  "Oladunni",
  "Olaide",
  "Olamide",
  "Olamilekan",
  "Olanike",
  "Olanrewaju",
  "Olasunkanmi",
  "Olatayo",
  "Olatunbosun",
  "Olatunde",
  "Olawale",
  "Olayemi",
  "Oluwabamise",
  "Oluwabusayo",
  "Oluwadarasimi",
  "Oluwadamilola",
  "Oluwafemi",
  "Oluwafunmilayo",
  "Oluwajomiloju",
  "Oluwakemi",
  "Oluwakorede",
  "Oluwaleke",
  "Oluwamide",
  "Oluwamodupe",
  "Oluwapelumi",
  "Oluwapelunmi",
  "Oluwapemi",
  "Oluwasegun",
  "Oluwaseun",
  "Oluwashola",
  "Oluwatosin",
  "Oluwatobi",
  "Oluwatoyin",
  "Oreoluwa",
  "Pelumi",
  "Sade",
  "Seun",
  "Seyi",
  "Shola",
  "Simisola",
  "Taiwo",
  "Temidayo",
  "Temilade",
  "Temiloluwa",
  "Temitope",
  "Tife",
  "Titilayo",
  "Tobi",
  "Tolani",
  "Tolulope",
  "Tolu",
  "Tomi",
  "Tosin",
  "Wale",
  "Yetunde",
  "Yewande",
  "Yinka",
  ...buildCombinations(
    ["Ade", "Ayo", "Ola", "Olu", "Oluwa", "Temi", "Tobi", "Femi", "Kemi", "Seyi", "Yemi", "Tunde", "Dami", "Mide", "Yinka", "Bola", "Kore", "Remi", "Ore", "Dayo", "Tolu", "Tomi", "Sade", "Jide", "Akin", "Folu", "Dele", "Funmi", "Shola", "Kehinde"],
    ["bayo", "yemi", "tope", "sola", "tunde", "dapo", "kanmi", "niyi", "dara", "bunmi", "jide", "deji", "dayo", "tayo", "ranti", "wumi", "timi", "tunji", "koya", "femi", "kemi", "mide", "seun", "lola"],
  ),
]);

const igboFirstNames = uniqueValues([
  "Adaeze",
  "Adanna",
  "Amaka",
  "Amarachi",
  "Anuli",
  "Chiamaka",
  "Chibueze",
  "Chibuzo",
  "Chichi",
  "Chidera",
  "Chidinma",
  "Chidubem",
  "Chiemeka",
  "Chiemerie",
  "Chijioke",
  "Chika",
  "Chikamso",
  "Chikodi",
  "Chima",
  "Chimamanda",
  "Chinaza",
  "Chinenye",
  "Chinonso",
  "Chinwe",
  "Chioma",
  "Chisom",
  "Ebube",
  "Emeka",
  "Ezinne",
  "Ifeanyi",
  "Ikem",
  "Kelechi",
  "Kenechukwu",
  "Kosisochukwu",
  "Ndidiamaka",
  "Nkechi",
  "Nkem",
  "Nkiru",
  "Nnamdi",
  "Nneamaka",
  "Nonso",
  "Obinna",
  "Ogechi",
  "Oluchi",
  "Somto",
  "Uche",
  "Ugochukwu",
  "Ujunwa",
  ...buildCombinations(
    ["Chi", "Chukwu", "Ife", "Uche", "Ngozi", "Nke", "Ada", "Obi", "Kene", "Som", "Ama", "Ugo"],
    ["nedu", "dinma", "kodi", "nna", "eze", "oma", "amaka", "eka", "jika", "dimma", "madu", "yelu", "dalu", "sochukwu"],
  ),
]);

const hausaFirstNames = uniqueValues([
  "Aisha",
  "Aliyu",
  "Amina",
  "Aminu",
  "Bello",
  "Bilkisu",
  "Fatima",
  "Hadiza",
  "Hafsat",
  "Hamza",
  "Hauwa",
  "Ibrahim",
  "Jamila",
  "Kabiru",
  "Khadija",
  "Ladan",
  "Lawal",
  "Maryam",
  "Mubarak",
  "Murtala",
  "Nafisa",
  "Nana",
  "Nasiru",
  "Rabi",
  "Rashida",
  "Sadiya",
  "Sadiq",
  "Safiya",
  "Salisu",
  "Sani",
  "Shehu",
  "Umar",
  "Usman",
  "Yahaya",
  "Yakubu",
  "Zainab",
  "Zulaiha",
]);

const otherNigerianFirstNames = uniqueValues([
  "Akpan",
  "Asuquo",
  "Bassey",
  "Briggs",
  "Duke",
  "Ebi",
  "Ebiere",
  "Efemena",
  "Effiong",
  "Ekanem",
  "Ekpere",
  "Eno",
  "Etim",
  "George",
  "Ima",
  "Ime",
  "Ivie",
  "Izuchi",
  "Pere",
  "Preye",
  "Tamara",
  "Tonye",
  "Tuoyo",
  "Uduak",
  "Wariso",
]);

const nigerianSurnames = uniqueValues([
  "Abiola",
  "Adebayo",
  "Adeyemi",
  "Afolabi",
  "Agbaje",
  "Aigbe",
  "Aina",
  "Akinola",
  "Akinyemi",
  "Akpan",
  "Akubue",
  "Aliyu",
  "Asuquo",
  "Attah",
  "Awoniyi",
  "Baba",
  "Balogun",
  "Bassey",
  "Bello",
  "Briggs",
  "Danjuma",
  "Dike",
  "Duru",
  "Ekanem",
  "Ekong",
  "Elechi",
  "Emenike",
  "Etim",
  "Eze",
  "Ezeani",
  "Ezeh",
  "George",
  "Ibe",
  "Ibrahim",
  "Ifejika",
  "Ijeoma",
  "Inyang",
  "Jibril",
  "Lawal",
  "Madu",
  "Mamman",
  "Mohammed",
  "Musa",
  "Nnaji",
  "Nnamani",
  "Nnanna",
  "Nwafor",
  "Nwagwu",
  "Nwankwo",
  "Nwanne",
  "Nwosu",
  "Obasi",
  "Obi",
  "Obiakor",
  "Obinna",
  "Odia",
  "Odili",
  "Ojo",
  "Okafor",
  "Okeke",
  "Okoli",
  "Okonkwo",
  "Okoro",
  "Okoye",
  "Olawale",
  "Olowe",
  "Omotayo",
  "Onyeka",
  "Onoh",
  "Oshodi",
  "Owolabi",
  "Peters",
  "Sani",
  "Shehu",
  "Sule",
  "Udo",
  "Udoh",
  "Umeh",
  "Usman",
  ...buildCombinations(
    ["Ade", "Akin", "Ala", "Ari", "Bola", "Duro", "Fola", "Ibi", "Lani", "Ola", "Olowo", "Olu", "Ore", "Oye", "Ogun", "Ojo", "Oke", "Ader", "Atoy", "Ayen", "Ayan", "Awo"],
    ["bayo", "dele", "jide", "tunde", "yemi", "kunle", "wale", "sanya", "niyi", "kanmi", "ranti", "dipo", "femi", "bunmi", "teju", "toye", "shola", "dapo", "sola", "mide"],
  ),
  ...buildCombinations(
    ["Oka", "Nwa", "Ume", "Eze", "Obi", "Ogbu", "Ano", "Ilo", "Nkwọ", "Ozo", "Onye", "Uzo"],
    ["for", "chukwu", "nna", "madu", "kafor", "nnamdi", "dika", "emena", "obi", "onye", "ike", "eze", "adi", "kwu"],
  ).map((item) => item.replace("Nkwọ", "Nkwo")),
]);

const generatedNameBuckets = {
  yoruba: yorubaFirstNames,
  igbo: igboFirstNames,
  hausa: hausaFirstNames,
  other: otherNigerianFirstNames,
};

const generatedNameUniverseSize =
  (yorubaFirstNames.length + igboFirstNames.length + hausaFirstNames.length + otherNigerianFirstNames.length) *
  nigerianSurnames.length;

export function getGeneratedNigerianNamePoolSize(): number {
  return generatedNameUniverseSize;
}

function randomFrom<T>(items: T[]): T {
  return items[Math.floor(Math.random() * items.length)];
}

function randomInt(min: number, max: number): number {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function roundToStep(value: number, step: number): number {
  return Math.round(value / step) * step;
}

function mostlyRoundAmount(min: number, max: number): number {
  const raw = randomInt(min, max);
  const roll = Math.random();

  if (roll < 0.6) {
    return roundToStep(raw, 1000);
  }

  if (roll < 0.9) {
    return roundToStep(raw, 100);
  }

  return raw;
}

function randomAmount(mode: TransactionMode, input: GeneratorInput): number {
  if (mode === "salary") {
    return mostlyRoundAmount(180000, 950000);
  }

  if (mode === "transfer_in") {
    const minIncomingAmount = Math.max(1000, input.minIncomingAmount);
    const maxIncomingAmount = Math.max(minIncomingAmount, input.maxIncomingAmount);
    return mostlyRoundAmount(minIncomingAmount, maxIncomingAmount);
  }

  if (mode === "cash_withdrawal") {
    return randomInt(10, 250) * 1000;
  }

  return mostlyRoundAmount(1500, 180000);
}

function clampDay(year: number, monthIndex: number, day: number): number {
  const lastDay = new Date(year, monthIndex + 1, 0).getDate();
  return Math.min(day, lastDay);
}

function buildTellerNumber(): string {
  return String(randomInt(10000, 99999999)).padStart(7, "0");
}

function formatNairaAmount(value: number): string {
  return Math.round(value).toString();
}

function getDayKey(value: string): number {
  const date = new Date(value);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

function formatStatementDate(value: string): string {
  const date = new Date(value);
  const day = String(date.getDate()).padStart(2, "0");
  const month = date.toLocaleString("en-GB", { month: "short" });
  const year = String(date.getFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
}

function buildChargeCode(mode: TransactionMode): string {
  if (mode === "cash_withdrawal") {
    return `WDR004/${randomInt(100000, 999999)}`;
  }

  if (mode === "salary") {
    return `MDB005/${randomInt(1000, 9999)}`;
  }

  if (mode === "transfer_in") {
    return `TRF004/${randomInt(100000, 9999999)}`;
  }

  return `TRF003/${randomInt(10000, 99999)}`;
}

function buildSmsNarration(chargeCode: string): string {
  return `Charges On SMS Alert For : ${chargeCode}`;
}

function buildTransferCharge(amount: number): number {
  if (amount < 10000) {
    return 0;
  }

  if (amount <= 50000) {
    return 26.88;
  }

  return 53.75;
}

function buildNarration(
  mode: TransactionMode,
  customerName: string,
  amount: number,
  counterpartyName: string,
  salaryCompanyName: string,
): string {
  if (mode === "salary") {
    return `App To Coop Mortgage Bank From ${salaryCompanyName} - Salary`;
  }

  if (mode === "cash_withdrawal") {
    return `A Withdrawal of ${formatNairaAmount(amount)} Naira, By ${customerName} - Cash. Teller No: ${buildTellerNumber()}`;
  }

  if (mode === "transfer_in") {
    return Math.random() > 0.5
      ? `App To Coop Mortgage Bank from ${counterpartyName}`
      : `Fund Transfer from ${counterpartyName} to ${customerName}`;
  }

  return `Funds Transfer from ${customerName} to ${counterpartyName}`;
}

function normalizeSpecialTransactions(items: SpecialTransactionInput[]): SpecialTransactionInput[] {
  return items.filter((item) => item.amount > 0 && item.date);
}

function pickMode(): TransactionMode {
  const pool: TransactionMode[] = [
    "transfer_out",
    "transfer_out",
    "transfer_out",
    "transfer_in",
    "transfer_in",
    "cash_withdrawal",
  ];
  return randomFrom(pool);
}

function getAvailableNames(customerName: string, namePool: string[]): string[] {
  const cleanedPool = namePool
    .map((item) => item.trim())
    .filter(Boolean);
  const customerKey = customerName.trim().toLowerCase();
  const uniqueNames = Array.from(new Set(cleanedPool));

  return uniqueNames.filter((item) => item.toLowerCase() !== customerKey);
}

function pickWeightedBucket(): string[] {
  const roll = Math.random();

  if (roll < 0.6) {
    return generatedNameBuckets.yoruba;
  }

  if (roll < 0.8) {
    return generatedNameBuckets.igbo;
  }

  if (roll < 0.9) {
    return generatedNameBuckets.hausa;
  }

  return generatedNameBuckets.other;
}

function pickGeneratedCounterparty(
  customerName: string,
  date: Date,
  generatedState: GeneratedNameState,
): string {
  if (generatedState.usedFullNames.size >= generatedNameUniverseSize) {
    generatedState.usedFullNames.clear();
    generatedState.usageMap.clear();
  }

  const customerKey = customerName.trim().toLowerCase();
  const dayKey = getDayKey(date.toISOString());

  for (let attempt = 0; attempt < 1200; attempt += 1) {
    const firstName = randomFrom(pickWeightedBucket());
    const surname = randomFrom(nigerianSurnames);
    const fullName = `${firstName} ${surname}`.trim();
    const normalized = fullName.toLowerCase();

    if (normalized === customerKey || generatedState.usedFullNames.has(normalized)) {
      continue;
    }

    generatedState.usedFullNames.add(normalized);
    generatedState.usageMap.set(normalized, { lastUsedDayKey: dayKey });
    return fullName;
  }

  for (const firstName of [
    ...generatedNameBuckets.yoruba,
    ...generatedNameBuckets.igbo,
    ...generatedNameBuckets.hausa,
    ...generatedNameBuckets.other,
  ]) {
    for (const surname of nigerianSurnames) {
      const fullName = `${firstName} ${surname}`.trim();
      const normalized = fullName.toLowerCase();

      if (normalized === customerKey || generatedState.usedFullNames.has(normalized)) {
        continue;
      }

      generatedState.usedFullNames.add(normalized);
      generatedState.usageMap.set(normalized, { lastUsedDayKey: dayKey });
      return fullName;
    }
  }

  generatedState.usedFullNames.clear();
  return `${randomFrom(generatedNameBuckets.yoruba)} ${randomFrom(nigerianSurnames)}`;
}

function pickCounterpartyForDate(
  customerName: string,
  namePool: string[],
  date: Date,
  usageMap: Map<string, NameUsageMeta>,
  generatedState: GeneratedNameState,
  input: GeneratorInput,
): string {
  if (namePool.length === 0) {
    return pickGeneratedCounterparty(customerName, date, generatedState);
  }

  const dayKey = getDayKey(date.toISOString());
  const baseNames = getAvailableNames(customerName, namePool);

  const candidates = baseNames.filter((name) => {
    const meta = usageMap.get(name) ?? { lastUsedDayKey: null };
    const respectsGap =
      meta.lastUsedDayKey === null ||
      dayKey - meta.lastUsedDayKey >= Math.max(0, input.minDaysBeforeNameReuse) * 24 * 60 * 60 * 1000;
    return respectsGap;
  });

  const selected = randomFrom(
    candidates.length > 0 ? candidates : baseNames,
  );

  usageMap.set(selected, {
    lastUsedDayKey: dayKey,
  });

  return selected;
}

function applyCellBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: "thin", color: { argb: "FFD9D9D9" } },
    left: { style: "thin", color: { argb: "FFD9D9D9" } },
    bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
    right: { style: "thin", color: { argb: "FFD9D9D9" } },
  };
}

function updateNarrationAmount(description: string, amount: number): string {
  if (description.startsWith("A Withdrawal of ")) {
    return description.replace(/A Withdrawal of \d+ Naira/, `A Withdrawal of ${formatNairaAmount(amount)} Naira`);
  }

  if (/^(Fund|Funds) Transfer/i.test(description) || /^MT - Transfer/i.test(description) || /^Ft o /i.test(description)) {
    return description;
  }

  return description;
}

function ensureBalanceRange(rows: SeedRow[], openingBalance: number, minimumBalance: number, maximumBalance: number): SeedRow[] {
  let runningBalance = openingBalance;
  const skippedGroups = new Set<number>();

  return rows.flatMap((row) => {
    if (skippedGroups.has(row.group)) {
      return [];
    }

    if (row.order === 0) {
      runningBalance = row.credit - row.debit;
      return [row];
    }

    let nextDebit = row.debit;
    let nextCredit = row.credit;

    if (row.order === 1 && nextDebit > runningBalance) {
      if (runningBalance <= minimumBalance + 5) {
        skippedGroups.add(row.group);
        return [];
      }

      nextDebit = Math.max(1000, Math.floor((runningBalance - minimumBalance - 5) / 100) * 100);
      if (nextDebit <= 0) {
        skippedGroups.add(row.group);
        return [];
      }
    }

    if (row.order === 1 && nextDebit > 0 && runningBalance - nextDebit < minimumBalance) {
      const allowedDebit = Math.floor((runningBalance - minimumBalance) / 100) * 100;
      if (allowedDebit <= 0) {
        skippedGroups.add(row.group);
        return [];
      }
      nextDebit = allowedDebit;
    }

    if (row.order === 1 && nextCredit > 0 && runningBalance + nextCredit > maximumBalance) {
      const allowedCredit = Math.floor((maximumBalance - runningBalance) / 100) * 100;
      if (allowedCredit <= 0) {
        skippedGroups.add(row.group);
        return [];
      }
      nextCredit = allowedCredit;
    }

    runningBalance += nextCredit;
    runningBalance -= nextDebit;

    return [{
      ...row,
      debit: nextDebit,
      credit: nextCredit,
      description: updateNarrationAmount(row.description, nextDebit),
    }];
  });
}

function buildClosingAdjustmentAmount(requiredChange: number): { mode: TransactionMode; amount: number } | null {
  if (requiredChange === 0) {
    return null;
  }

  if (requiredChange > 0) {
    const withoutStamp = Math.round(requiredChange + 5);
    if (withoutStamp < 10000) {
      return { mode: "transfer_in", amount: Math.max(1000, roundToStep(withoutStamp, 100)) };
    }

    return { mode: "transfer_in", amount: Math.max(10000, roundToStep(Math.round(requiredChange + 55), 100)) };
  }

  const decreaseNeeded = Math.abs(requiredChange);
  const debitAmount = Math.max(0, Math.round(decreaseNeeded - 5));
  return { mode: "transfer_out", amount: Math.max(1000, roundToStep(debitAmount, 100)) };
}

function pushPrimaryTransaction(
  rows: SeedRow[],
  payload: {
    date: Date;
    description: string;
    debit: number;
    credit: number;
    mode: TransactionMode;
  },
) {
  const group = rows.length + 1;
  const chargeCode = buildChargeCode(payload.mode);
  const baseTime = payload.date.getTime();
  const transferCharge = payload.mode === "transfer_out" ? buildTransferCharge(payload.debit) : 0;
  const hasStampDuty = payload.credit >= 10000;

  rows.push({
    id: crypto.randomUUID(),
    date: new Date(baseTime).toISOString(),
    description: payload.description,
    debit: payload.debit,
    credit: payload.credit,
    group,
    order: 1,
  });

  if (transferCharge > 0) {
    rows.push({
      id: crypto.randomUUID(),
      date: new Date(baseTime + 1000).toISOString(),
      description: "Transfer Charge",
      debit: transferCharge,
      credit: 0,
      group,
      order: 2,
    });
  }

  if (hasStampDuty && payload.mode !== "salary") {
    rows.push({
      id: crypto.randomUUID(),
      date: new Date(baseTime + (transferCharge > 0 ? 2000 : 1000)).toISOString(),
      description: "Stamp Duty Charge",
      debit: 50,
      credit: 0,
      group,
      order: transferCharge > 0 ? 3 : 2,
    });
  }

  rows.push({
    id: crypto.randomUUID(),
    date: new Date(baseTime + (transferCharge > 0 ? 3000 : hasStampDuty && payload.mode !== "salary" ? 2000 : 1000)).toISOString(),
    description: buildSmsNarration(chargeCode),
    debit: 5,
    credit: 0,
    group,
    order: transferCharge > 0 ? 4 : hasStampDuty && payload.mode !== "salary" ? 3 : 2,
  });

  if (hasStampDuty && payload.mode === "salary") {
    rows.push({
      id: crypto.randomUUID(),
      date: new Date(baseTime + (transferCharge > 0 ? 4000 : 2000)).toISOString(),
      description: "Stamp Duty Charge",
      debit: 50,
      credit: 0,
      group,
      order: transferCharge > 0 ? 5 : 3,
    });
  }
}

export function generateTransactions(input: GeneratorInput): TransactionRow[] {
  const startBoundary = new Date(`${input.startDate}T00:00:00`);
  const endBoundary = new Date(`${input.closingDate}T00:00:00`);
  const safeEndBoundary = endBoundary >= startBoundary ? endBoundary : startBoundary;
  const months = Math.max(
    1,
    (safeEndBoundary.getFullYear() - startBoundary.getFullYear()) * 12 +
      (safeEndBoundary.getMonth() - startBoundary.getMonth()) +
      1,
  );
  const minMonthly = Math.max(1, input.minTransactionsPerMonth);
  const maxMonthly = Math.max(minMonthly, input.maxTransactionsPerMonth);
  const rows: SeedRow[] = [];
  const specials = normalizeSpecialTransactions(input.specialTransactions);
  const salaryCompanyName = input.salaryCompanyName.trim() || "Company Name";
  const nameUsageMap = new Map<string, NameUsageMeta>();
  const generatedNameState: GeneratedNameState = {
    usedFullNames: new Set<string>(),
    usageMap: new Map<string, NameUsageMeta>(),
  };

  rows.push({
    id: crypto.randomUUID(),
    date: new Date(`${input.startDate}T00:00:00`).toISOString(),
    description: "Opening Balance",
    debit: 0,
    credit: input.openingBalance,
    group: 0,
    order: 0,
  });

  for (let monthOffset = 0; monthOffset < months; monthOffset += 1) {
    const monthDate = new Date(startBoundary.getFullYear(), startBoundary.getMonth() + monthOffset, 1);
    const year = monthDate.getFullYear();
    const monthIndex = monthDate.getMonth();
    const totalDays = new Date(year, monthIndex + 1, 0).getDate();
    const totalTransactions = randomInt(minMonthly, maxMonthly);

    if (input.includeSalary && input.salaryAmount > 0) {
      const salaryDay = clampDay(year, monthIndex, input.salaryDay);
      const salaryDate = new Date(year, monthIndex, salaryDay);

      if (salaryDate >= startBoundary && salaryDate <= safeEndBoundary) {
        pushPrimaryTransaction(rows, {
          date: salaryDate,
          description: buildNarration("salary", input.customerName, input.salaryAmount, "", salaryCompanyName),
          debit: 0,
          credit: input.salaryAmount,
          mode: "salary",
        });
      }
    }

    for (let index = 0; index < totalTransactions; index += 1) {
      const mode = pickMode();
      const minDay = monthOffset === 0 ? startBoundary.getDate() : 1;
      const maxDay = monthOffset === months - 1 ? safeEndBoundary.getDate() : totalDays;
      if (maxDay < minDay) {
        continue;
      }

      const day = randomInt(minDay, maxDay);
      const date = new Date(year, monthIndex, day);
      const amount = randomAmount(mode, input);
      const counterpartyName = pickCounterpartyForDate(
        input.customerName,
        input.namePool,
        date,
        nameUsageMap,
        generatedNameState,
        input,
      );
      const isCredit = mode === "transfer_in";

      pushPrimaryTransaction(rows, {
        date,
        description: buildNarration(mode, input.customerName, amount, counterpartyName, salaryCompanyName),
        debit: isCredit ? 0 : amount,
        credit: isCredit ? amount : 0,
        mode,
      });
    }
  }

  for (const item of specials) {
    const specialDate = new Date(`${item.date}T00:00:00`);
    if (specialDate < startBoundary || specialDate > safeEndBoundary) {
      continue;
    }

    const counterpartyName = item.counterpartyName.trim() || pickCounterpartyForDate(
      input.customerName,
      input.namePool,
      specialDate,
      nameUsageMap,
      generatedNameState,
      input,
    );
    const narration = item.description.trim() || buildNarration(
      item.mode,
      input.customerName,
      item.amount,
      counterpartyName,
      salaryCompanyName,
    );

    pushPrimaryTransaction(rows, {
      date: specialDate,
      description: narration,
      debit: item.kind === "debit" ? item.amount : 0,
      credit: item.kind === "credit" ? item.amount : 0,
      mode: item.mode,
    });
  }

  rows.sort((first, second) => {
    const dayDiff = getDayKey(first.date) - getDayKey(second.date);
    if (dayDiff !== 0) {
      return dayDiff;
    }

    const groupDiff = first.group - second.group;
    if (groupDiff !== 0) {
      return groupDiff;
    }

    return first.order - second.order;
  });

  const boundedRows = ensureBalanceRange(
    rows,
    input.openingBalance,
    input.minimumBalance,
    input.maximumBalance,
  );

  let provisionalBalance = input.openingBalance;
  for (const row of boundedRows.slice(1)) {
    provisionalBalance += row.credit;
    provisionalBalance -= row.debit;
  }

  const closingServiceCharge = 2500;
  const closingVatCharge = 250;
  const adjustmentNeeded = Number((input.targetClosingBalance - (provisionalBalance - closingServiceCharge - closingVatCharge)).toFixed(2));
  const adjustment = buildClosingAdjustmentAmount(adjustmentNeeded);

  if (adjustment && adjustment.amount > 0) {
    const counterpartyName = pickCounterpartyForDate(
      input.customerName,
      input.namePool,
      safeEndBoundary,
      nameUsageMap,
      generatedNameState,
      input,
    );
    pushPrimaryTransaction(boundedRows, {
      date: safeEndBoundary,
      description: buildNarration(
        adjustment.mode,
        input.customerName,
        adjustment.amount,
        counterpartyName,
        salaryCompanyName,
      ),
      debit: adjustment.mode === "transfer_out" ? adjustment.amount : 0,
      credit: adjustment.mode === "transfer_in" ? adjustment.amount : 0,
      mode: adjustment.mode,
    });
  }

  boundedRows.push({
    id: crypto.randomUUID(),
    date: safeEndBoundary.toISOString(),
    description: "Charge for Embassy Reference Letter and Account Statement",
    debit: closingServiceCharge,
    credit: 0,
    group: Number.MAX_SAFE_INTEGER - 1,
    order: 4,
  });

  boundedRows.push({
    id: crypto.randomUUID(),
    date: safeEndBoundary.toISOString(),
    description: "VAT fee for Embassy Reference Letter and Account Statement Charge",
    debit: closingVatCharge,
    credit: 0,
    group: Number.MAX_SAFE_INTEGER,
    order: 5,
  });

  boundedRows.sort((first, second) => {
    const dayDiff = getDayKey(first.date) - getDayKey(second.date);
    if (dayDiff !== 0) {
      return dayDiff;
    }

    const groupDiff = first.group - second.group;
    if (groupDiff !== 0) {
      return groupDiff;
    }

    return first.order - second.order;
  });

  const finalRows = ensureBalanceRange(
    boundedRows,
    input.openingBalance,
    input.minimumBalance,
    input.maximumBalance,
  );

  let runningBalance = input.openingBalance;

  return finalRows.map((row, index) => {
    if (index === 0) {
      runningBalance = row.credit - row.debit;
    } else {
      runningBalance += row.credit;
      runningBalance -= row.debit;
    }

    return {
      ...row,
      date: formatStatementDate(row.date),
      balance: Number(runningBalance.toFixed(2)),
    };
  });
}

export async function exportStyledWorkbook(rows: TransactionRow[]) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Statement");

  worksheet.columns = [
    { header: "Date", key: "date", width: 12 },
    { header: "Transaction Details", key: "description", width: 58 },
    { header: "Debit Amount", key: "debit", width: 16 },
    { header: "Credit Amount", key: "credit", width: 16 },
    { header: "Balance", key: "balance", width: 16 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.height = 15;
  headerRow.font = { name: "Arial", size: 8, color: { argb: "FF00007B" } };
  headerRow.alignment = { vertical: "middle", horizontal: "left" };
  headerRow.eachCell((cell) => {
    applyCellBorder(cell);
  });

  rows.forEach((row) => {
    const newRow = worksheet.addRow({
      date: row.date,
      description: row.description,
      debit: row.debit,
      credit: row.credit,
      balance: row.balance,
    });

    newRow.height = 15;
    newRow.eachCell((cell, colNumber) => {
      let color = "FF000000";
      if (colNumber === 3) {
        color = "FF7C0000";
      }
      if (colNumber === 4) {
        color = "FF007C00";
      }
      if (colNumber === 5) {
        color = "FF00007B";
      }

      cell.font = { name: "Arial", size: 8, color: { argb: color } };
      cell.alignment = { vertical: "middle", horizontal: "left" };
      applyCellBorder(cell);

      if (colNumber >= 3) {
        cell.numFmt = "#,##0.00";
      }
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
