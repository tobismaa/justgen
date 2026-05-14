import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { TransactionRow } from "./types";

const balancedTableColumns = [860, 5900, 1160, 1480, 1760] as const;

export type StatementExportMeta = {
  customerName: string;
  startDate: string;
  closingDate: string;
};

function formatMoney(value: number): string {
  return value.toLocaleString("en-NG", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function safeCustomerName(value: string): string {
  return value.trim() || "Customer";
}

function chunkStatementRows(rows: TransactionRow[]): TransactionRow[][] {
  if (rows.length === 0) {
    return [[]];
  }

  const pages: TransactionRow[][] = [];
  const firstPageCount = 18;
  const laterPageCount = 40;

  pages.push(rows.slice(0, firstPageCount));

  for (let index = firstPageCount; index < rows.length; index += laterPageCount) {
    pages.push(rows.slice(index, index + laterPageCount));
  }

  return pages;
}

function buildHeaderCapsule(text: string): TableCell {
  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    shading: { fill: "C0C0C0", type: ShadingType.CLEAR, color: "auto" },
    margins: { top: 40, bottom: 40, left: 70, right: 70 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text,
            bold: true,
            font: "Arial",
            size: 16,
            color: "000000",
          }),
        ],
      }),
    ],
  });
}

function buildInfoTable(meta: StatementExportMeta): Table {
  const labelStyle = {
    bold: true,
    size: 18,
    font: "Arial",
    color: "26190F",
  } as const;

  const valueStyle = {
    size: 18,
    font: "Arial",
    color: "26190F",
  } as const;

  const makeCell = (label: string, value: string) =>
    new TableCell({
      width: { size: 50, type: WidthType.PERCENTAGE },
      margins: { top: 90, bottom: 90, left: 120, right: 120 },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "D9D9D9" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "D9D9D9" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "D9D9D9" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "D9D9D9" },
      },
      children: [
        new Paragraph({
          spacing: { after: 40 },
          children: [new TextRun({ text: label, ...labelStyle })],
        }),
        new Paragraph({
          children: [new TextRun({ text: value, ...valueStyle })],
        }),
      ],
    });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        children: [
          makeCell("Customer Name", safeCustomerName(meta.customerName)),
          makeCell("Statement Period", `${meta.startDate} to ${meta.closingDate}`),
        ],
      }),
    ],
  });
}


function buildHeaderRow(): TableRow {
  const headerCell = (
    text: string,
    width: number,
    leftIndent: number,
    align: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT,
  ) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      borders: {
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      },
      children: [
        new Paragraph({
          alignment: align,
          indent: { left: leftIndent },
          spacing: { before: 0, line: 179, lineRule: "exact" },
          children: [
            new TextRun({
              text,
              bold: true,
              font: "Arial",
              size: 16,
              color: "00007B",
            }),
          ],
        }),
      ],
    });

  return new TableRow({
    children: [
      headerCell("Date", balancedTableColumns[0], 96),
      headerCell("Transaction Details", balancedTableColumns[1], 320),
      headerCell("Debit Amount", balancedTableColumns[2], 12),
      headerCell("Credit Amount", balancedTableColumns[3], 148),
      headerCell("Balance", balancedTableColumns[4], 290),
    ],
  });
}

function buildBodyRow(row: TransactionRow): TableRow {
  const makeTextCell = (
    text: string,
    width: number,
    leftIndent: number,
    spacingBefore: number,
    borderTop: ((typeof BorderStyle)[keyof typeof BorderStyle]) | null,
    align: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT,
  ) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      borders: {
        top: borderTop ? { style: borderTop, size: 2, color: "000000" } : { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      },
      children: [
        new Paragraph({
          alignment: align,
          indent: { left: leftIndent },
          spacing: spacingBefore > 0 ? { before: spacingBefore } : undefined,
          children: [new TextRun({ text, font: "Arial", size: 16, color: "000000" })],
        }),
      ],
    });

  const makeAmountCell = (
    value: number,
    width: number,
    color: string,
    leftIndent: number,
    spacingBefore: number,
    topStyle: (typeof BorderStyle)[keyof typeof BorderStyle],
    bottomStyle: (typeof BorderStyle)[keyof typeof BorderStyle],
  ) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      borders: {
        top: { style: topStyle, size: 2, color: "000000" },
        bottom: { style: bottomStyle, size: 2, color: "000000" },
        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      },
      children: [
        new Paragraph({
          indent: { left: leftIndent },
          spacing: spacingBefore > 0 ? { before: spacingBefore } : undefined,
          children: [
            new TextRun({
              text: formatMoney(value),
              font: "Arial",
              size: 16,
              bold: true,
              color,
            }),
          ],
        }),
      ],
    });

  const isOpeningBalance = row.description === "Opening Balance";
  const isDescriptionCharge =
    row.description === "Transfer Charge" ||
    row.description === "Stamp Duty Charge" ||
    row.description.startsWith("Charges On SMS Alert For");
  const rowSpacing = isOpeningBalance ? 114 : isDescriptionCharge ? 102 : 106;

  return new TableRow({
    children: [
      makeTextCell(
        row.date,
        balancedTableColumns[0],
        6,
        isOpeningBalance ? 116 : rowSpacing,
        isOpeningBalance ? BorderStyle.SINGLE : null,
        AlignmentType.CENTER,
      ),
      makeTextCell(
        row.description,
        balancedTableColumns[1],
        56,
        isOpeningBalance ? 116 : rowSpacing,
        isOpeningBalance ? BorderStyle.SINGLE : null,
      ),
      makeAmountCell(
        row.debit,
        balancedTableColumns[2],
        "790000",
        14,
        rowSpacing,
        isOpeningBalance ? BorderStyle.SINGLE : BorderStyle.DASH_SMALL_GAP,
        BorderStyle.DASH_SMALL_GAP,
      ),
      makeAmountCell(
        row.credit,
        balancedTableColumns[3],
        "007900",
        148,
        rowSpacing,
        BorderStyle.SINGLE,
        BorderStyle.SINGLE,
      ),
      makeAmountCell(
        row.balance,
        balancedTableColumns[4],
        "000080",
        255,
        rowSpacing,
        BorderStyle.SINGLE,
        BorderStyle.SINGLE,
      ),
    ],
  });
}

export async function exportStatementDocx(rows: TransactionRow[], meta: StatementExportMeta) {
  const customerName = safeCustomerName(meta.customerName);
  const pagedRows = chunkStatementRows(rows);
  const tables = pagedRows.map((pageRows, index) =>
    new Table({
      width: { size: 11160, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      alignment: AlignmentType.CENTER,
      columnWidths: [...balancedTableColumns],
      rows: [
        ...(index === 0 ? [buildHeaderRow()] : []),
        ...pageRows.map((row) => buildBodyRow(row)),
      ],
    }),
  );

  const contentChildren = [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({
          children: [
            buildHeaderCapsule("STATEMENT OF ACCOUNT"),
            buildHeaderCapsule(customerName),
          ],
        }),
      ],
    }),
    new Paragraph({ spacing: { after: 0, before: 0, line: 1, lineRule: "exact" } }),
    tables[0],
  ];

  for (let index = 1; index < tables.length; index += 1) {
    contentChildren.push(
      new Paragraph({
        pageBreakBefore: true,
        spacing: {
          before: 0,
          after: 0,
          line: 1,
          lineRule: "exact",
        },
      }),
      tables[index],
    );
  }

  const document = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              width: 12240,
              height: 15840,
            },
            margin: {
              top: 420,
              right: 0,
              bottom: 1500,
              left: 0,
              header: 0,
              footer: 0,
              gutter: 0,
            },
          },
        },
        children: contentChildren,
      },
    ],
  });

  return Packer.toBlob(document);
}

export async function exportStatementPdf(rows: TransactionRow[], meta: StatementExportMeta) {
  const doc = new jsPDF({
    orientation: "portrait",
    unit: "pt",
    format: "a4",
  });

  const customerName = safeCustomerName(meta.customerName);

  doc.setFillColor(182, 94, 48);
  doc.roundedRect(40, 32, 72, 72, 16, 16, "F");
  doc.setTextColor(255, 248, 242);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(24);
  doc.text("TG", 60, 78);

  doc.setFillColor(192, 192, 192);
  doc.setDrawColor(0, 0, 0);
  doc.roundedRect(130, 34, 176, 22, 10, 10, "FD");
  doc.roundedRect(316, 34, 220, 22, 10, 10, "FD");
  doc.setFontSize(10);
  doc.setTextColor(0, 0, 0);
  doc.text("STATEMENT OF ACCOUNT", 149, 49);
  doc.text(customerName.toUpperCase().slice(0, 28), 334, 49);

  doc.setTextColor(38, 25, 15);
  doc.setFontSize(20);
  doc.text("Statement of Account", 130, 84);
  doc.setFontSize(12);
  doc.setFont("helvetica", "normal");
  doc.text(customerName, 130, 106);
  doc.text(`Period: ${meta.startDate} to ${meta.closingDate}`, 130, 124);

  autoTable(doc, {
    startY: 154,
    margin: { left: 40, right: 40, bottom: 34, top: 34 },
    theme: "grid",
    head: [["Date", "Transaction Details", "Debit Amount", "Credit Amount", "Balance"]],
    body: rows.map((row) => [
      row.date,
      row.description,
      formatMoney(row.debit),
      formatMoney(row.credit),
      formatMoney(row.balance),
    ]),
    styles: {
      font: "helvetica",
      fontSize: 8,
      textColor: [0, 0, 0],
      lineColor: [0, 0, 0],
      lineWidth: 0.4,
      cellPadding: { top: 4, right: 6, bottom: 4, left: 6 },
      overflow: "linebreak",
    },
    headStyles: {
      fillColor: [238, 232, 222],
      textColor: [0, 0, 128],
      fontStyle: "bold",
    },
    columnStyles: {
      0: { halign: "center", cellWidth: 62 },
      1: { halign: "left", cellWidth: 248 },
      2: { halign: "right", cellWidth: 72, textColor: [124, 0, 0], fontStyle: "bold" },
      3: { halign: "right", cellWidth: 72, textColor: [0, 124, 0], fontStyle: "bold" },
      4: { halign: "right", cellWidth: 88, textColor: [0, 0, 128], fontStyle: "bold" },
    },
    didDrawPage: () => {
      const pageHeight = doc.internal.pageSize.getHeight();
      doc.setFontSize(9);
      doc.setTextColor(119, 98, 81);
      doc.text("© Vino Banking", 40, pageHeight - 18);
      doc.text("Printed By: UMEZURIKEC", 132, pageHeight - 18);
      doc.text(
        `Page ${doc.getCurrentPageInfo().pageNumber}`,
        doc.internal.pageSize.getWidth() - 70,
        pageHeight - 18,
      );
    },
  });

  return doc.output("blob");
}
