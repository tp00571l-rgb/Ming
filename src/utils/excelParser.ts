import * as XLSX from 'xlsx';

export interface ParsedSheetData {
  sheetName: string;
  headers: string[];
  rows: Record<string, unknown>[];
  emptyCellCount: number;
}

export interface ParsedWorkbookResult {
  fileName: string;
  fileSizeBytes: number;
  sheetNames: string[];
  sheets: Record<string, ParsedSheetData>;
}

const ALLOWED_EXTENSIONS = ['.xlsx', '.xls', '.csv'];

export function validateExcelFile(file: File): string | null {
  const lowerName = file.name.toLowerCase();
  const hasValidExt = ALLOWED_EXTENSIONS.some((ext) => lowerName.endsWith(ext));
  if (!hasValidExt) {
    return 'รองรับเฉพาะไฟล์นามสกุล .xlsx, .xls หรือ .csv เท่านั้น';
  }
  if (file.size === 0) {
    return 'ไฟล์ที่เลือกมีขนาด 0 ไบต์ ไม่สามารถอ่านข้อมูลได้';
  }
  return null;
}

export function extractSheetData(
  worksheet: XLSX.WorkSheet,
  sheetName: string,
  useFormattedValues: boolean
): ParsedSheetData {
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
    defval: '',
    raw: !useFormattedValues,
  });

  const headerSet = new Set<string>();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      headerSet.add(key);
    }
  }
  const headers = Array.from(headerSet);

  let emptyCellCount = 0;
  for (const row of rows) {
    for (const header of headers) {
      const value = row[header];
      if (value === '' || value === null || value === undefined) {
        emptyCellCount += 1;
      }
    }
  }

  return {
    sheetName,
    headers,
    rows,
    emptyCellCount,
  };
}

export function parseWorkbookBuffer(
  buffer: ArrayBuffer,
  fileName: string,
  fileSizeBytes: number,
  useFormattedValues: boolean
): { workbook: XLSX.WorkBook; result: ParsedWorkbookResult } {
  const data = new Uint8Array(buffer);
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('ไม่พบแผ่นงาน (Sheet) ภายในไฟล์นี้');
  }

  const sheets: Record<string, ParsedSheetData> = {};
  for (const name of workbook.SheetNames) {
    const worksheet = workbook.Sheets[name];
    if (worksheet) {
      sheets[name] = extractSheetData(worksheet, name, useFormattedValues);
    }
  }

  return {
    workbook,
    result: {
      fileName,
      fileSizeBytes,
      sheetNames: workbook.SheetNames,
      sheets,
    },
  };
}

export function formatCellValue(value: unknown): { text: string; isEmpty: boolean; isNumeric: boolean } {
  if (value === '' || value === null || value === undefined) {
    return { text: '-', isEmpty: true, isNumeric: false };
  }
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return { text: '-', isEmpty: true, isNumeric: false };
    }
    return {
      text: value.toLocaleDateString('th-TH', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
      }),
      isEmpty: false,
      isNumeric: true,
    };
  }
  if (typeof value === 'number') {
    return {
      text: String(value),
      isEmpty: false,
      isNumeric: true,
    };
  }
  if (typeof value === 'boolean') {
    return {
      text: value ? 'true' : 'false',
      isEmpty: false,
      isNumeric: false,
    };
  }
  const str = String(value);
  const isNumericString = /^-?\d[\d,]*(\.\d+)?$/.test(str.trim());
  return {
    text: str,
    isEmpty: false,
    isNumeric: isNumericString,
  };
}

export function buildSampleWorkbookBuffer(): ArrayBuffer {
  const workOrders = [
    {
      รหัสใบงาน: 'WO-2026-001',
      วันที่แจ้ง: '2026-11-10',
      แผนก: 'สายการผลิต A',
      อุปกรณ์: 'ปั๊มไฮดรอลิก P-104',
      ผู้รับผิดชอบ: 'สมชาย ใจดี',
      ความเร่งด่วน: 'สูง',
      จำนวนชั่วโมง: 4.5,
      สถานะ: 'ดำเนินการเสร็จสิ้น',
    },
    {
      รหัสใบงาน: 'WO-2026-002',
      วันที่แจ้ง: '2026-11-12',
      แผนก: 'คลังสินค้า',
      อุปกรณ์: 'โฟล์คลิฟท์ไฟฟ้า FL-03',
      ผู้รับผิดชอบ: 'วิชัย รักงาน',
      ความเร่งด่วน: 'ปานกลาง',
      จำนวนชั่วโมง: 2.0,
      สถานะ: 'กำลังดำเนินการ',
    },
    {
      รหัสใบงาน: 'WO-2026-003',
      วันที่แจ้ง: '2026-11-14',
      แผนก: 'ระบบไฟฟ้า',
      อุปกรณ์: 'ตู้ควบคุม MDB-02',
      ผู้รับผิดชอบ: 'นภา สว่างวงศ์',
      ความเร่งด่วน: 'สูง',
      จำนวนชั่วโมง: 0,
      สถานะ: 'รอดำเนินการ',
    },
    {
      รหัสใบงาน: 'WO-2026-004',
      วันที่แจ้ง: '2026-11-15',
      แผนก: 'สายการผลิต B',
      อุปกรณ์: 'สายพานลำเลียง CV-08',
      ผู้รับผิดชอบ: '',
      ความเร่งด่วน: 'ต่ำ',
      จำนวนชั่วโมง: 1.5,
      สถานะ: 'รอมอบหมาย',
    },
  ];

  const summaryRows = [
    { หมวดหมู่: 'รวมใบงานทั้งหมด', จำนวน: 4, หมายเหตุ: 'ข้อมูลทดสอบระบบ' },
    { หมวดหมู่: 'ดำเนินการเสร็จสิ้น', จำนวน: 1, หมายเหตุ: 'ปิดงานแล้ว' },
    { หมวดหมู่: 'อยู่ระหว่างดำเนินการ', จำนวน: 3, หมายเหตุ: 'ติดตามต่อ' },
  ];

  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.json_to_sheet(workOrders);
  const ws2 = XLSX.utils.json_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, ws1, 'ใบงานซ่อมบำรุง');
  XLSX.utils.book_append_sheet(wb, ws2, 'สรุปรายการ');

  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
  return out;
}

export function exportSheetToCsv(sheet: ParsedSheetData): string {
  const ws = XLSX.utils.json_to_sheet(sheet.rows, { header: sheet.headers });
  return XLSX.utils.sheet_to_csv(ws);
}
