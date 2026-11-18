import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  AlertCircle,
  Download,
  FileSpreadsheet,
  RotateCcw,
  Search,
  Upload,
} from 'lucide-react';
import {
  buildSampleWorkbookBuffer,
  exportSheetToCsv,
  extractSheetData,
  formatCellValue,
  ParsedWorkbookResult,
  parseWorkbookBuffer,
  validateExcelFile,
} from './utils/excelParser';

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const workbookRef = useRef<XLSX.WorkBook | null>(null);

  const [parsedResult, setParsedResult] = useState<ParsedWorkbookResult | null>(null);
  const [activeSheet, setActiveSheet] = useState<string>('');
  const [useFormattedValues, setUseFormattedValues] = useState<boolean>(true);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const handleLoadBuffer = (
    buffer: ArrayBuffer,
    fileName: string,
    fileSizeBytes: number,
    formatted: boolean
  ) => {
    try {
      const { workbook, result } = parseWorkbookBuffer(
        buffer,
        fileName,
        fileSizeBytes,
        formatted
      );
      workbookRef.current = workbook;
      setParsedResult(result);
      setActiveSheet(result.sheetNames[0]);
      setSearchQuery('');
      setErrorMessage(null);
    } catch (err) {
      workbookRef.current = null;
      setParsedResult(null);
      setActiveSheet('');
      setErrorMessage(
        err instanceof Error
          ? `ไม่สามารถอ่านไฟล์ได้: ${err.message}`
          : 'เกิดข้อผิดพลาดในการอ่านไฟล์ Excel กรุณาตรวจสอบว่าไฟล์ไม่เสียหายหรือไม่ได้ตั้งรหัสผ่าน'
      );
    } finally {
      setIsLoading(false);
    }
  };

  const processFile = (file: File) => {
    const validationError = validateExcelFile(file);
    if (validationError) {
      setErrorMessage(validationError);
      return;
    }

    setIsLoading(true);
    setErrorMessage(null);

    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result;
      if (!(result instanceof ArrayBuffer)) {
        setIsLoading(false);
        setErrorMessage('ไม่สามารถอ่านเนื้อหาไฟล์เป็นข้อมูลไบนารีได้');
        return;
      }
      handleLoadBuffer(result, file.name, file.size, useFormattedValues);
    };
    reader.onerror = () => {
      setIsLoading(false);
      setErrorMessage('เกิดข้อผิดพลาดระหว่างการเปิดอ่านไฟล์จากอุปกรณ์ของคุณ');
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      processFile(file);
    }
    // Reset input value so selecting the same file again triggers onChange
    event.target.value = '';
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const handleLoadSample = () => {
    setIsLoading(true);
    setErrorMessage(null);
    const sampleBuffer = buildSampleWorkbookBuffer();
    handleLoadBuffer(
      sampleBuffer,
      'sample_work_orders.xlsx',
      sampleBuffer.byteLength,
      useFormattedValues
    );
  };

  const handleToggleFormatted = (nextFormatted: boolean) => {
    setUseFormattedValues(nextFormatted);
    const wb = workbookRef.current;
    if (!wb || !parsedResult) return;

    const updatedSheets: ParsedWorkbookResult['sheets'] = {};
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (ws) {
        updatedSheets[name] = extractSheetData(ws, name, nextFormatted);
      }
    }
    setParsedResult({
      ...parsedResult,
      sheets: updatedSheets,
    });
  };

  const handleReset = () => {
    workbookRef.current = null;
    setParsedResult(null);
    setActiveSheet('');
    setSearchQuery('');
    setErrorMessage(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const currentSheetData = useMemo(() => {
    if (!parsedResult || !activeSheet) return null;
    return parsedResult.sheets[activeSheet] ?? null;
  }, [parsedResult, activeSheet]);

  const filteredRows = useMemo(() => {
    if (!currentSheetData) return [];
    const trimmed = searchQuery.trim().toLowerCase();
    if (!trimmed) return currentSheetData.rows;

    return currentSheetData.rows.filter((row) =>
      currentSheetData.headers.some((header) => {
        const { text } = formatCellValue(row[header]);
        return text.toLowerCase().includes(trimmed);
      })
    );
  }, [currentSheetData, searchQuery]);

  const handleDownloadCsv = () => {
    if (!currentSheetData || !parsedResult) return;
    const csv = exportSheetToCsv(currentSheetData);
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const baseName = parsedResult.fileName.replace(/\.[^/.]+$/, '');
    link.download = `${baseName}_${currentSheetData.sheetName}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* Top Navigation Bar: 3-zone contract */}
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-8">
          <a
            href="#top"
            className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 rounded"
          >
            ระบบจัดการข้อมูลใบงาน
          </a>

          <nav
            aria-label="ส่วนหลัก"
            className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600"
          >
            <a
              href="#import-section"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0"
            >
              นำเข้าไฟล์
            </a>
            <a
              href="#data-grid-section"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0"
            >
              ตารางข้อมูลดิบ
            </a>
          </nav>

          <div className="flex items-center gap-3 shrink-0">
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="inline-flex items-center justify-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 focus-visible:ring-offset-2"
            >
              <Upload className="w-4 h-4" aria-hidden="true" />
              <span>เลือกไฟล์ Excel</span>
            </button>
          </div>
        </div>
      </header>

      <main id="top" className="max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6 flex-1">
        {/* Page Heading + Unboxed Metadata Summary */}
        <section className="flex flex-col md:flex-row md:items-end justify-between gap-4">
          <div className="space-y-1">
            <h1 className="text-2xl font-bold tracking-tight text-slate-900 [text-wrap:balance]">
              นำเข้าและตรวจสอบข้อมูลใบงาน
            </h1>
            <p className="text-sm text-slate-600">
              อ่านไฟล์ .xlsx, .xls หรือ .csv ในเบราว์เซอร์โดยตรง พร้อมสลับแผ่นงานและค้นหาข้อมูลทันที
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
            <span>สถานะ:</span>
            <span className="font-medium text-slate-800">
              {isLoading
                ? 'กำลังประมวลผลไฟล์...'
                : parsedResult
                  ? `โหลดสำเร็จ (${parsedResult.fileName})`
                  : 'รอการนำเข้าไฟล์'}
            </span>
            {parsedResult && (
              <>
                <span aria-hidden="true">·</span>
                <span className="font-mono-tabular">
                  {formatBytes(parsedResult.fileSizeBytes)}
                </span>
                <span aria-hidden="true">·</span>
                <span className="font-mono-tabular">
                  {parsedResult.sheetNames.length} ชีต
                </span>
              </>
            )}
          </div>
        </section>

        {/* Error Alert (Explicit icon + text, no color-only signaling) */}
        {errorMessage && (
          <div
            role="alert"
            className="bg-red-50 border border-red-200 rounded-xl p-4 flex items-start gap-3 text-sm text-red-800"
          >
            <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" aria-hidden="true" />
            <div className="flex-1">
              <p className="font-semibold">เกิดข้อผิดพลาดในการอ่านไฟล์</p>
              <p className="mt-0.5 text-red-700">{errorMessage}</p>
            </div>
            <button
              type="button"
              onClick={() => setErrorMessage(null)}
              className="px-3 py-1.5 text-xs font-medium text-red-800 hover:bg-red-100 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              ปิด
            </button>
          </div>
        )}

        {/* File Import Dropzone */}
        <section
          id="import-section"
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
          className={`bg-white rounded-xl border transition-colors p-6 ${
            isDragging
              ? 'border-blue-500 bg-blue-50/40'
              : 'border-slate-200'
          }`}
        >
          <input
            ref={fileInputRef}
            id="excelFileInput"
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileChange}
            className="sr-only"
          />

          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-start sm:items-center gap-4">
              <div className="w-11 h-11 rounded-lg bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-700 shrink-0">
                <FileSpreadsheet className="w-5 h-5" aria-hidden="true" />
              </div>
              <div>
                <label
                  htmlFor="excelFileInput"
                  className="block text-sm font-semibold text-slate-900 cursor-pointer"
                >
                  เลือกไฟล์ Excel (.xlsx, .xls) หรือ CSV
                </label>
                <p className="text-xs text-slate-500 mt-0.5">
                  คลิกปุ่มด้านขวา หรือลากไฟล์มาวางในกรอบนี้เพื่อเริ่มอ่านข้อมูล
                </p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2 shrink-0">
              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 text-sm font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-900 focus-visible:ring-offset-2"
              >
                เลือกจากเครื่อง
              </button>
              <button
                type="button"
                onClick={handleLoadSample}
                className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
              >
                ทดสอบด้วยไฟล์ตัวอย่าง
              </button>
              {parsedResult && (
                <button
                  type="button"
                  onClick={handleReset}
                  className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <RotateCcw className="w-4 h-4" aria-hidden="true" />
                  <span>ล้างข้อมูล</span>
                </button>
              )}
            </div>
          </div>
        </section>

        {/* Action Toolbar & Sheet Controls (shown when workbook is loaded) */}
        {parsedResult && currentSheetData && (
          <section
            aria-label="ตัวกรองและการจัดการข้อมูล"
            className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-4"
          >
            {/* Sheet Switcher */}
            <div className="flex flex-wrap items-center gap-3">
              <span className="text-xs font-semibold text-slate-500 whitespace-nowrap">
                แผ่นงาน:
              </span>
              <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg overflow-x-auto max-w-full">
                {parsedResult.sheetNames.map((name) => {
                  const isActive = name === activeSheet;
                  return (
                    <button
                      key={name}
                      type="button"
                      onClick={() => {
                        setActiveSheet(name);
                        setSearchQuery('');
                      }}
                      className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        isActive
                          ? 'bg-white text-slate-900 shadow-xs'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {name}
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Search, Value Mode Toggle & Export */}
            <div className="flex flex-wrap items-center gap-3">
              <div className="relative min-w-[220px] flex-1 sm:flex-initial">
                <Search
                  className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none"
                  aria-hidden="true"
                />
                <input
                  type="search"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="ค้นหาในตาราง..."
                  aria-label="ค้นหาข้อมูลในตาราง"
                  className="w-full pl-9 pr-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-600 focus:bg-white"
                />
              </div>

              <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
                <button
                  type="button"
                  onClick={() => handleToggleFormatted(true)}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    useFormattedValues
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  ค่าตามรูปแบบ
                </button>
                <button
                  type="button"
                  onClick={() => handleToggleFormatted(false)}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    !useFormattedValues
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  ค่าดิบ (Raw)
                </button>
              </div>

              <button
                type="button"
                onClick={handleDownloadCsv}
                disabled={currentSheetData.rows.length === 0}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" aria-hidden="true" />
                <span>ส่งออก CSV</span>
              </button>
            </div>
          </section>
        )}

        {/* Raw Data Table Section */}
        <section
          id="data-grid-section"
          className="bg-white rounded-xl border border-slate-200 overflow-hidden"
        >
          <div className="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-2">
            <h2 className="font-bold text-slate-900 text-sm">ตารางข้อมูลจากไฟล์</h2>
            <div className="flex items-center gap-2 text-xs text-slate-600 font-mono-tabular">
              {currentSheetData ? (
                <>
                  <span>
                    แสดง {filteredRows.length.toLocaleString('th-TH')} /{' '}
                    {currentSheetData.rows.length.toLocaleString('th-TH')} แถว
                  </span>
                  <span aria-hidden="true">·</span>
                  <span>{currentSheetData.headers.length} คอลัมน์</span>
                  {currentSheetData.emptyCellCount > 0 && (
                    <>
                      <span aria-hidden="true">·</span>
                      <span className="text-amber-700">
                        ช่องว่าง {currentSheetData.emptyCellCount} ช่อง
                      </span>
                    </>
                  )}
                </>
              ) : (
                <span>ทั้งหมด 0 แถว</span>
              )}
            </div>
          </div>

          <div className="overflow-x-auto max-h-[600px]">
            {isLoading ? (
              <div className="p-6 space-y-3" aria-busy="true">
                {Array.from({ length: 5 }).map((_, idx) => (
                  <div
                    key={idx}
                    className="h-10 bg-slate-100 rounded-md animate-pulse"
                  />
                ))}
              </div>
            ) : !currentSheetData ? (
              <div className="py-16 px-6 text-center space-y-3">
                <p className="text-sm font-medium text-slate-700">
                  ยังไม่มีข้อมูลที่นำเข้า
                </p>
                <p className="text-xs text-slate-500 max-w-md mx-auto">
                  เลือกไฟล์ Excel (.xlsx, .xls) หรือ CSV ด้านบนเพื่อแสดงข้อมูลดิบทั้งหมด หรือลองโหลดไฟล์ตัวอย่างเพื่อทดสอบตาราง
                </p>
                <div className="pt-2">
                  <button
                    type="button"
                    onClick={handleLoadSample}
                    className="px-4 py-2 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg transition-colors cursor-pointer"
                  >
                    โหลดข้อมูลใบงานตัวอย่าง
                  </button>
                </div>
              </div>
            ) : currentSheetData.rows.length === 0 ? (
              <div className="py-14 px-6 text-center text-sm text-slate-500">
                ไม่พบแถวข้อมูลในแผ่นงาน “{currentSheetData.sheetName}”
              </div>
            ) : filteredRows.length === 0 ? (
              <div className="py-14 px-6 text-center space-y-2">
                <p className="text-sm text-slate-600">
                  ไม่พบข้อมูลที่ตรงกับคำค้นหา “{searchQuery}”
                </p>
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  className="text-xs font-medium text-blue-600 hover:underline cursor-pointer"
                >
                  ล้างคำค้นหา
                </button>
              </div>
            ) : (
              <table className="w-full text-left border-collapse text-sm">
                <thead className="sticky top-0 z-10 bg-slate-100 border-b border-slate-200">
                  <tr>
                    <th
                      scope="col"
                      className="px-4 py-3 whitespace-nowrap bg-slate-200/80 w-14 text-center text-xs font-bold text-slate-700 font-mono-tabular"
                    >
                      #
                    </th>
                    {currentSheetData.headers.map((header) => (
                      <th
                        key={header}
                        scope="col"
                        className="px-4 py-3 whitespace-nowrap text-xs font-bold text-slate-700"
                      >
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 text-slate-700">
                  {filteredRows.map((row, rowIndex) => (
                    <tr
                      key={rowIndex}
                      className="hover:bg-blue-50/50 transition-colors"
                    >
                      <td className="px-4 py-2.5 text-center text-xs text-slate-500 font-medium bg-slate-50/70 font-mono-tabular">
                        {rowIndex + 1}
                      </td>
                      {currentSheetData.headers.map((header) => {
                        const cell = formatCellValue(row[header]);
                        return (
                          <td
                            key={header}
                            className={`px-4 py-2.5 whitespace-nowrap ${
                              cell.isEmpty
                                ? 'text-slate-400'
                                : cell.isNumeric
                                  ? 'font-mono-tabular text-slate-800'
                                  : 'text-slate-800'
                            }`}
                          >
                            {cell.text}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </section>
      </main>
    </div>
  );
}
