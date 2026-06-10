/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, type ChangeEvent } from "react";
import { motion, AnimatePresence } from "motion/react";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import * as XLSX from "xlsx";
import { INITIAL_DIARY_STATE, type CounselingDiary } from "./types";
import { generateDefaultDocxBlob } from "./utils/docxGenerator";

// Helper functions for Excel parsing
function parseExcelDate(val: any): string {
  if (!val) return "";

  if (val instanceof Date) {
    const day = String(val.getUTCDate()).padStart(2, "0");
    const month = String(val.getUTCMonth() + 1).padStart(2, "0");
    const year = val.getUTCFullYear();
    return `${day}/${month}/${year}`;
  }

  if (typeof val === "number") {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const msPerDay = 24 * 60 * 60 * 1000;
    const date = new Date(epoch.getTime() + val * msPerDay);
    const day = String(date.getUTCDate()).padStart(2, "0");
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    const year = date.getUTCFullYear();
    return `${day}/${month}/${year}`;
  }

  const str = String(val).trim();
  const yyyymmddRegex = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/;
  const matchYmd = str.match(yyyymmddRegex);
  if (matchYmd) {
    return `${matchYmd[3].padStart(2, "0")}/${matchYmd[2].padStart(2, "0")}/${matchYmd[1]}`;
  }

  const ddmmyyyyRegex = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/;
  const matchDmy = str.match(ddmmyyyyRegex);
  if (matchDmy) {
    return `${matchDmy[1].padStart(2, "0")}/${matchDmy[2].padStart(2, "0")}/${matchDmy[3]}`;
  }

  return str;
}

function parseGender(val: any): { gt_nam: string; gt_nu: string; gt_kb: string } {
  const str = String(val || "").trim().toLowerCase();
  if (str === "nam" || str === "m" || str === "male" || str === "boy") {
    return { gt_nam: "X", gt_nu: "", gt_kb: "" };
  }
  if (str === "nữ" || str === "nu" || str === "f" || str === "female" || str === "girl") {
    return { gt_nam: "", gt_nu: "X", gt_kb: "" };
  }
  return { gt_nam: "", gt_nu: "", gt_kb: "X" };
}

export default function App() {
  const [diary, setDiary] = useState<CounselingDiary>({
    ...INITIAL_DIARY_STATE,
  });
  const [customTemplateArrayBuffer, setCustomTemplateArrayBuffer] =
    useState<ArrayBuffer | null>(null);
  const [uploadedTemplateName, setUploadedTemplateName] = useState<
    string | null
  >(null);
  const [statusMessage, setStatusMessage] = useState<{
    text: string;
    type: "success" | "error" | "info" | null;
  }>({
    text: "",
    type: null,
  });

  // Keep track of the active section for the sidebar scrollspy
  const [activeTab, setActiveTab] = useState<number>(1);

  // References to section elements for smooth scrolling
  const section1Ref = useRef<HTMLDivElement>(null);
  const section2Ref = useRef<HTMLDivElement>(null);
  const section3Ref = useRef<HTMLDivElement>(null);
  const section4Ref = useRef<HTMLDivElement>(null);

  const scrollToSection = (
    sectionIndex: number,
    ref: React.RefObject<HTMLDivElement | null>,
  ) => {
    setActiveTab(sectionIndex);
    ref.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const handleInputChange = (
    e: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => {
    const { name, value } = e.target;
    setDiary((prev) => ({
      ...prev,
      [name]: value,
    }));
  };

  // Helper to handle priority number input validation [1-6]
  const handlePriorityChange = (name: keyof CounselingDiary, value: string) => {
    // Standardize mapping: can only be numbers 1 to 6 or empty
    if (value === "" || /^[1-6]$/.test(value)) {
      setDiary((prev) => ({
        ...prev,
        [name]: value,
      }));
    }
  };

  // Gender toggles (strictly sets "X" to selected, empty to others)
  const selectGender = (gender: "Nam" | "Nữ" | "Không biết") => {
    setDiary((prev) => ({
      ...prev,
      gt_nam: gender === "Nam" ? "X" : "",
      gt_nu: gender === "Nữ" ? "X" : "",
      gt_kb: gender === "Không biết" ? "X" : "",
    }));
  };

  // Follow-up/Ending toggles
  const selectEndingMethod = (
    method: "dung_theo_doi" | "len_ke_hoach" | "thuc_hien_chuyen",
  ) => {
    setDiary((prev) => ({
      ...prev,
      kt_dung_theo_doi: method === "dung_theo_doi" ? "X" : "",
      kt_len_ke_hoach: method === "len_ke_hoach" ? "X" : "",
      kt_thuc_hien_chuyen: method === "thuc_hien_chuyen" ? "X" : "",
    }));
  };

  // Handle uploading custom MS Word template file (.docx)
  const handleDocxUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    if (file.name.split(".").pop()?.toLowerCase() !== "docx") {
      showStatus(
        "File tải lên phải đúng định dạng Microsoft Word (.docx)",
        "error",
      );
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      if (event.target?.result instanceof ArrayBuffer) {
        setCustomTemplateArrayBuffer(event.target.result);
        setUploadedTemplateName(file.name);
        showStatus(`Đã tải tệp mẫu của bạn: "${file.name}"`, "success");
      }
    };
    reader.onerror = () => {
      showStatus("Không thể đọc tệp tin đã chọn. Vui lòng thử lại.", "error");
    };
    reader.readAsArrayBuffer(file);
  };

  // Handle importing data from student profile Excel sheet (.xlsx)
  const handleExcelUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    const extension = file.name.split(".").pop()?.toLowerCase();
    if (extension !== "xlsx" && extension !== "xls") {
      showStatus(
        "File tải lên phải đúng định dạng Excel (.xlsx hoặc .xls)",
        "error",
      );
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        if (!event.target || !event.target.result) {
          throw new Error("Không thể đọc nội dung tệp");
        }

        const data = new Uint8Array(event.target.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });

        // Get the first worksheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Parse sheet as 2D Array
        const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });

        // Extracted variables
        let hoTen = "";
        let ngaySinh = "";
        let gtNam = "";
        let gtNu = "";
        let gtKb = "";
        let hoTenCha = "";
        let ngheNghiepCha = "";
        let hoTenMe = "";
        let ngheNghiepMe = "";

        // Iterate through rows and columns to find key-value pairs
        for (let r = 0; r < rows.length; r++) {
          const row = rows[r];
          if (!row || !Array.isArray(row)) continue;

          for (let c = 0; c < row.length; c++) {
            const cellValue = row[c];
            if (cellValue === undefined || cellValue === null) continue;

            const normalizedLabel = String(cellValue)
              .trim()
              .toLowerCase()
              .replace(/:$/, "")
              .trim();

            const nextCellValue = row[c + 1];

            if (normalizedLabel === "họ và tên") {
              hoTen = String(nextCellValue || "").trim();
            } else if (normalizedLabel === "ngày sinh") {
              ngaySinh = parseExcelDate(nextCellValue);
            } else if (normalizedLabel === "giới tính") {
              const parsedGender = parseGender(nextCellValue);
              gtNam = parsedGender.gt_nam;
              gtNu = parsedGender.gt_nu;
              gtKb = parsedGender.gt_kb;
            } else if (normalizedLabel === "họ tên cha") {
              hoTenCha = String(nextCellValue || "").trim();
            } else if (normalizedLabel === "nghề nghiệp cha") {
              ngheNghiepCha = String(nextCellValue || "").trim();
            } else if (normalizedLabel === "họ tên mẹ") {
              hoTenMe = String(nextCellValue || "").trim();
            } else if (normalizedLabel === "nghề nghiệp mẹ") {
              ngheNghiepMe = String(nextCellValue || "").trim();
            }
          }
        }

        // Apply to state (only modify fields that exist in template and were found)
        setDiary((prev) => ({
          ...prev,
          ho_ten: hoTen || prev.ho_ten,
          ngay_sinh: ngaySinh || prev.ngay_sinh,
          // Update gender if found
          ...(gtNam || gtNu || gtKb ? { gt_nam: gtNam, gt_nu: gtNu, gt_kb: gtKb } : {}),
          ho_ten_cha: hoTenCha || prev.ho_ten_cha,
          nghe_nghiep_cha: ngheNghiepCha || prev.nghe_nghiep_cha,
          ho_ten_me: hoTenMe || prev.ho_ten_me,
          nghe_nghiep_me: ngheNghiepMe || prev.nghe_nghiep_me,
        }));

        showStatus(
          `Đã nhập thành công hồ sơ học sinh: ${hoTen || "Ẩn danh"}`,
          "success",
        );
      } catch (err: any) {
        console.error(err);
        showStatus("Lỗi khi đọc file Excel: " + err.message, "error");
      }
    };

    reader.onerror = () => {
      showStatus("Lỗi đọc file. Vui lòng thử lại.", "error");
    };

    reader.readAsArrayBuffer(file);
  };

  // Status popups helper
  const showStatus = (text: string, type: "success" | "error" | "info") => {
    setStatusMessage({ text, type });
    setTimeout(() => {
      setStatusMessage({ text: "", type: null });
    }, 5000);
  };

  // Reset all forms
  const resetForm = () => {
    if (
      window.confirm("Bạn có chắc chắn muốn nhập lại toàn bộ dữ liệu từ đầu?")
    ) {
      setDiary({ ...INITIAL_DIARY_STATE });
      showStatus("Đã thiết lập lại toàn bộ biểu mẫu dữ liệu.", "info");
    }
  };

  // Downloads programmatically compiled Word template
  const downloadSampleTemplateFile = () => {
    try {
      const blob = generateDefaultDocxBlob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "Bieu_mau_Nhat_ky_Tu_van_Hoc_duong.docx";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
      showStatus(
        "Đã tải tệp Word mẫu chuẩn hệ thống về máy tính của bạn!",
        "success",
      );
    } catch (err: any) {
      showStatus("Lỗi khi sinh tệp mẫu: " + err.message, "error");
    }
  };

  // Generates complete data-filled file and triggers browser download
  const handleExportWord = async () => {
    try {
      showStatus("Đang biên dịch và xuất tài liệu...", "info");

      let docxTemplateArrayBuffer: ArrayBuffer;

      if (customTemplateArrayBuffer) {
        docxTemplateArrayBuffer = customTemplateArrayBuffer;
      } else {
        // Fallback to programmatic template blob
        const blob = generateDefaultDocxBlob();
        docxTemplateArrayBuffer = await blob.arrayBuffer();
      }

      const zip = new PizZip(docxTemplateArrayBuffer);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // Prepare mapped variables from state
      // Ensure empty strings are handled cleanly
      const valuesToRender: Record<string, string> = {};
      Object.keys(diary).forEach((key) => {
        const item = diary[key as keyof CounselingDiary];
        valuesToRender[key] = item || "";
      });

      // Execute render
      doc.render(valuesToRender);

      // Extract generated model file blob
      const out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      // Output filename safe string
      const sanitizedName = diary.ho_ten
        ? diary.ho_ten.trim().replace(/\s+/g, "_")
        : "Hoc_Sinh";
      const filename = `Nhat_ky_Tu_van_Hoc_duong_${sanitizedName}.docx`;

      // Trigger standard browser download
      const url = window.URL.createObjectURL(out);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      showStatus(
        "Chúc mừng! Đã xuất và tải tài liệu nhật ký thành công (.docx)!",
        "success",
      );
    } catch (error: any) {
      console.error(error);
      showStatus(
        `Có lỗi phát sinh khi xuất file Word: ${error.message}`,
        "error",
      );
    }
  };

  // Helper utility to dynamically calculate textarea auto-growth row size
  const calculateRows = (text: string) => {
    if (!text) return 3;
    const lines = text.split("\n").length;
    return Math.max(3, lines);
  };

  return (
    <div className="min-h-screen pb-32 bg-[#F8F9FA] relative flex flex-col selection:bg-[#EE6C4D]/20">
      {/* Dynamic Status Toast Indicator */}
      <AnimatePresence>
        {statusMessage.type && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className={`fixed top-4 left-1/2 -translate-x-1/2 z-50 px-6 py-4 rounded-xl text-sm font-semibold tracking-wide shadow-soft max-w-lg text-center ${
              statusMessage.type === "success"
                ? "bg-emerald-600 text-white"
                : statusMessage.type === "error"
                  ? "bg-[#D90429] text-white"
                  : "navy-bg text-white"
            }`}
          >
            {statusMessage.text}
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header Panel (Cố định hoặc ở trên cùng trang) */}
      <header className="navy-bg text-white py-6 px-8 text-center shrink-0 shadow-soft border-b-4 border-[#EE6C4D]">
        <div className="max-w-5xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4 text-center md:text-left opacity-90">
          <div className="space-y-1">
            <p className="text-[10px] sm:text-xs font-bold text-slate-300 uppercase tracking-widest font-mono">
              Ủy ban Nhân dân Phường 3 Bảo Lộc
            </p>
            <h2 className="text-sm sm:text-base font-extrabold tracking-tight">
              TRƯỜNG THCS PHAN CHU TRINH
            </h2>
          </div>
          <div className="text-center md:text-right hidden sm:block">
            <p className="font-bold text-xs sm:text-sm tracking-wide">
              CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
            </p>
            <p className="text-[10px] sm:text-xs font-semibold text-slate-300 border-t border-white/20 pt-1 mt-1 font-mono tracking-wider">
              Độc lập - Tự do - Hạnh phúc
            </p>
          </div>
        </div>
        <div className="max-w-3xl mx-auto mt-6 text-center px-4">
          <h1 className="text-base sm:text-lg md:text-xl font-bold uppercase tracking-tight leading-relaxed">
            Nhật ký thực hiện công tác tư vấn học đường và công tác xã hội cho
            người học
          </h1>
        </div>
      </header>

      {/* Main Body Content Frame */}
      <main className="max-w-5xl mx-auto mt-6 px-4 w-full flex flex-col lg:flex-row gap-6 items-start">
        {/* Navigation Flow Indicators (Left Sidebar Scrollspy Desktop-Only) */}
        <nav className="w-full lg:w-64 shrink-0 bg-white p-4 rounded-xl shadow-soft border border-slate-100 lg:sticky lg:top-6 space-y-1">
          <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-3 px-2">
            Mục lục nhật ký công tác
          </p>
          <button
            onClick={() => scrollToSection(1, section1Ref)}
            className={`w-full text-left px-3 py-2.5 rounded-lg text-xs font-bold transition-all flex items-center justify-between ${
              activeTab === 1
                ? "bg-[#0B2545]/5 text-[#0B2545] border-l-4 border-[#EE6C4D]"
                : "text-slate-600 hover:bg-slate-50"
            }`}
          >
            <span>1. Thông tin người học</span>
            {activeTab === 1 && (
              <span className="w-1.5 h-1.5 rounded-full bg-[#EE6C4D]"></span>
            )}
          </button>

          <button
            onClick={() => scrollToSection(2, section2Ref)}
            className={`w-full text-left px-3 py-2.5 rounded-lg text-xs font-bold transition-all flex items-center justify-between ${
              activeTab === 2
                ? "bg-[#0B2545]/5 text-[#0B2545] border-l-4 border-[#EE6C4D]"
                : "text-slate-600 hover:bg-slate-50"
            }`}
          >
            <span>2. Nội dung &amp; Hình thức</span>
            {activeTab === 2 && (
              <span className="w-1.5 h-1.5 rounded-full bg-[#EE6C4D]"></span>
            )}
          </button>

          <button
            onClick={() => scrollToSection(3, section3Ref)}
            className={`w-full text-left px-3 py-2.5 rounded-lg text-xs font-bold transition-all flex items-center justify-between ${
              activeTab === 3
                ? "bg-[#0B2545]/5 text-[#0B2545] border-l-4 border-[#EE6C4D]"
                : "text-slate-600 hover:bg-slate-50"
            }`}
          >
            <span>3. Nhật ký tự luận</span>
            {activeTab === 3 && (
              <span className="w-1.5 h-1.5 rounded-full bg-[#EE6C4D]"></span>
            )}
          </button>

          <button
            onClick={() => scrollToSection(4, section4Ref)}
            className={`w-full text-left px-3 py-2.5 rounded-lg text-xs font-bold transition-all flex items-center justify-between ${
              activeTab === 4
                ? "bg-[#0B2545]/5 text-[#0B2545] border-l-4 border-[#EE6C4D]"
                : "text-slate-600 hover:bg-slate-50"
            }`}
          >
            <span>4. Kết thúc &amp; Ký tên</span>
            {activeTab === 4 && (
              <span className="w-1.5 h-1.5 rounded-full bg-[#EE6C4D]"></span>
            )}
          </button>

          <div className="border-t border-slate-100 pt-4 mt-4 px-2 space-y-2">
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
              Tệp tin mẫu hiện tại:
            </p>
            <div className="bg-[#F8F9FA] p-2 rounded-lg border border-slate-200">
              <p className="text-[11px] font-medium text-slate-700 truncate">
                {uploadedTemplateName || "Mẫu chuẩn của hệ thống"}
              </p>
              <p className="text-[9px] text-slate-400 mt-0.5 font-mono">
                {uploadedTemplateName
                  ? "Tệp tùy cấu trúc người dùng"
                  : "Mặc định Times New Roman"}
              </p>
            </div>

            <button
              onClick={downloadSampleTemplateFile}
              className="w-full text-center text-[#0B2545] bg-[#0B2545]/5 hover:bg-[#0B2545]/10 border border-[#0B2545]/20 py-1 px-2 rounded-md text-[10px] font-bold transition"
            >
              Tải tệp tin Word (.docx) mẫu
            </button>
          </div>
        </nav>

        {/* Content Panel Frame (Contains Form blocks scrollable smoothly) */}
        <div className="flex-1 w-full space-y-8">
          {/* EXCEL IMPORT UTILITY PANEL */}
          <div className="bg-gradient-to-r from-[#0B2545]/5 via-[#EE6C4D]/5 to-transparent rounded-2xl border border-dashed border-[#0B2545]/20 p-5 flex flex-col md:flex-row items-center justify-between gap-4 shadow-soft hover:shadow-md transition duration-300">
            <div className="space-y-1 text-center md:text-left">
              <h4 className="text-xs font-bold text-[#0B2545] uppercase tracking-wider flex items-center justify-center md:justify-start gap-2">
                <span>⚡ Nhập liệu nhanh từ Excel</span>
                <span className="bg-[#EE6C4D] text-white text-[9px] font-extrabold px-1.5 py-0.5 rounded-full uppercase tracking-normal">Mới</span>
              </h4>
              <p className="text-[11px] text-slate-500 font-medium">
                Tải lên hồ sơ học sinh (.xlsx) để tự động điền các thông tin cơ bản của người học và cha mẹ.
              </p>
            </div>
            
            <label className="w-full md:w-auto relative cursor-pointer group shrink-0">
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleExcelUpload}
                className="hidden"
              />
              <div className="w-full select-none text-center bg-[#0B2545] hover:bg-[#0B2545]/90 text-white py-2.5 px-5 rounded-xl text-xs font-extrabold tracking-wider transition-all transform active:scale-95 shadow-soft flex items-center justify-center gap-2 border border-transparent">
                📊 Import từ Excel (.xlsx)
              </div>
            </label>
          </div>

          {/* BLOCK 1: Thông tin người học */}
          <div
            id="khoi-1"
            ref={section1Ref}
            className="bg-white rounded-2xl shadow-soft border-l-4 border-[#0B2545] p-6 lg:p-8 space-y-6"
          >
            <div className="border-b border-slate-100 pb-4">
              <span className="text-[10px] font-mono font-bold text-[#0B2545] uppercase tracking-widest block mb-1">
                KHỐI 1
              </span>
              <h3 className="text-base font-extrabold text-[#0B2545] uppercase tracking-wide">
                Thông tin người học
              </h3>
            </div>

            {/* Sub fields container */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  Họ và tên học sinh <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  name="ho_ten"
                  placeholder="Ví dụ: Nguyễn Văn A"
                  value={diary.ho_ten}
                  onChange={handleInputChange}
                  className="input-base"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  Ngày tháng năm sinh <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  name="ngay_sinh"
                  placeholder="Ví dụ: 15/08/2012 hoặc 2012-08-15"
                  value={diary.ngay_sinh}
                  onChange={handleInputChange}
                  className="input-base"
                />
              </div>

              <div className="md:col-span-2 space-y-2">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  Giới tính <span className="text-red-500">*</span>
                </label>
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => selectGender("Nam")}
                    className={`flex-1 py-2 rounded border text-sm transition-colors ${
                      diary.gt_nam === "X"
                        ? "bg-[#0B2545] text-white border-[#0B2545]"
                        : "bg-white text-slate-700 border-gray-200 hover:bg-[#0B2545] hover:text-white"
                    }`}
                  >
                    Nam
                  </button>
                  <button
                    type="button"
                    onClick={() => selectGender("Nữ")}
                    className={`flex-1 py-2 rounded border text-sm transition-colors ${
                      diary.gt_nu === "X"
                        ? "bg-[#0B2545] text-white border-[#0B2545]"
                        : "bg-white text-slate-700 border-gray-200 hover:bg-[#0B2545] hover:text-white"
                    }`}
                  >
                    Nữ
                  </button>
                  <button
                    type="button"
                    onClick={() => selectGender("Không biết")}
                    className={`flex-1 py-2 rounded border text-sm transition-colors ${
                      diary.gt_kb === "X"
                        ? "bg-[#0B2545] text-white border-[#0B2545]"
                        : "bg-white text-slate-700 border-gray-200 hover:bg-[#0B2545] hover:text-white"
                    }`}
                  >
                    Khác / Không rõ
                  </button>
                </div>
              </div>

              {/* Parents rows */}
              <div className="border border-slate-100 bg-[#F8F9FA] rounded-xl p-4 md:col-span-2 space-y-4">
                <div>
                  <p className="text-[11px] font-bold text-[#0B2545] uppercase tracking-wider">
                    Thông tin gia đình liên hệ (Cha)
                  </p>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 uppercase">
                      Họ tên Cha
                    </label>
                    <input
                      type="text"
                      name="ho_ten_cha"
                      placeholder="Họ tên Cha"
                      value={diary.ho_ten_cha}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 uppercase">
                      Tuổi
                    </label>
                    <input
                      type="text"
                      name="tuoi_cha"
                      placeholder="Ví dụ: 40"
                      value={diary.tuoi_cha}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 uppercase">
                      Nghề nghiệp
                    </label>
                    <input
                      type="text"
                      name="nghe_nghiep_cha"
                      placeholder="Nghề nghiệp"
                      value={diary.nghe_nghiep_cha}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>
                </div>
              </div>

              <div className="border border-slate-100 bg-[#F8F9FA] rounded-xl p-4 md:col-span-2 space-y-4">
                <div>
                  <p className="text-[11px] font-bold text-[#0B2545] uppercase tracking-wider">
                    Thông tin gia đình liên hệ (Mẹ)
                  </p>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 uppercase">
                      Họ tên Mẹ
                    </label>
                    <input
                      type="text"
                      name="ho_ten_me"
                      placeholder="Họ tên Mẹ"
                      value={diary.ho_ten_me}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 uppercase">
                      Tuổi
                    </label>
                    <input
                      type="text"
                      name="tuoi_me"
                      placeholder="Ví dụ: 38"
                      value={diary.tuoi_me}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 uppercase">
                      Nghề nghiệp
                    </label>
                    <input
                      type="text"
                      name="nghe_nghiep_me"
                      placeholder="Nghề nghiệp"
                      value={diary.nghe_nghiep_me}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>
                </div>
              </div>

              <div className="md:col-span-2 space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  Người chăm sóc hiện tại (nếu biết)
                </label>
                <input
                  type="text"
                  name="nguoi_cham_soc"
                  placeholder="Ví dụ: Cha mẹ, Ông bà nội, Cô ruột..."
                  value={diary.nguoi_cham_soc}
                  onChange={handleInputChange}
                  className="input-base"
                />
              </div>

              <div className="md:col-span-2 space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  Hoàn cảnh gia đình
                </label>
                <textarea
                  name="hoan_canh_gia_dinh"
                  rows={3}
                  placeholder="Mô tả hoàn cảnh khó khăn hoặc hoàn cảnh gia đình đặc biệt của học sinh (nếu có)..."
                  value={diary.hoan_canh_gia_dinh}
                  onChange={handleInputChange}
                  className="input-base resize-none"
                />
              </div>
            </div>
          </div>

          {/* BLOCK 2: Nội dung & Hình thức tư vấn */}
          <div
            id="khoi-2"
            ref={section2Ref}
            className="bg-white rounded-2xl shadow-soft border-l-4 border-[#EE6C4D] p-6 lg:p-8 space-y-8"
          >
            <div>
              <span className="text-[10px] font-mono font-bold text-[#EE6C4D] uppercase tracking-widest block mb-1">
                KHỐI 2
              </span>
              <h3 className="text-base font-extrabold text-[#0B2545] uppercase tracking-wide">
                Nội dung &amp; Hình thức tư vấn
              </h3>
            </div>

            {/* 2. Nội dung tư vấn xếp hạng */}
            <div className="space-y-3">
              <div className="p-3 bg-[#FFF7F2] rounded-lg text-xs leading-relaxed text-[#0B2545] font-semibold border border-[#EE6C4D]/10">
                Nếu chỉ có 1 nội dung tư vấn thì điền số 1 vào dòng tương ứng.
                Nếu có nhiều nội dung thì đánh số thứ tự từ 1 đến 6 (trên 7 nội
                dung) theo thứ tự ưu tiên của bạn.
              </div>

              {/* Grid content and entries */}
              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-soft">
                <div className="grid grid-cols-12 bg-[#F8F9FA] border-b border-slate-200 font-bold text-[10px] sm:text-xs uppercase text-slate-600 py-3 px-4 font-mono">
                  <div className="col-span-2 text-center">STT</div>
                  <div className="col-span-7">
                    Nội dung tư vấn học đường &amp; công tác xã hội
                  </div>
                  <div className="col-span-3 text-center">Thứ tự ưu tiên</div>
                </div>

                <div className="divide-y divide-slate-100">
                  {/* Row 1 */}
                  <div className="grid grid-cols-12 items-center py-2.5 px-4 text-xs font-semibold text-slate-700">
                    <div className="col-span-2 text-center font-mono">1</div>
                    <div className="col-span-7 text-[13px]">
                      Học tập (Kết quả, thái độ, kỳ vọng...)
                    </div>
                    <div className="col-span-3 flex justify-center">
                      <input
                        type="text"
                        maxLength={1}
                        value={diary.ut_hoc_tap}
                        onChange={(e) =>
                          handlePriorityChange("ut_hoc_tap", e.target.value)
                        }
                        placeholder="--"
                        className="w-12 text-center py-1.5 rounded border border-[#EE6C4D] bg-[#FFF7F2] orange-text font-bold focus:ring-2 focus:ring-[#EE6C4D]/10 outline-none shadow-sm text-sm"
                      />
                    </div>
                  </div>

                  {/* Row 2 */}
                  <div className="grid grid-cols-12 items-center py-2.5 px-4 text-xs font-semibold text-slate-700">
                    <div className="col-span-2 text-center font-mono">2</div>
                    <div className="col-span-7 text-[13px]">
                      Quan hệ xã hội (Thầy cô, bạn bè, gia đình...)
                    </div>
                    <div className="col-span-3 flex justify-center">
                      <input
                        type="text"
                        maxLength={1}
                        value={diary.ut_quan_he_xa_hoi}
                        onChange={(e) =>
                          handlePriorityChange(
                            "ut_quan_he_xa_hoi",
                            e.target.value,
                          )
                        }
                        placeholder="--"
                        className="w-12 text-center py-1.5 rounded border border-[#EE6C4D] bg-[#FFF7F2] orange-text font-bold focus:ring-2 focus:ring-[#EE6C4D]/10 outline-none shadow-sm text-sm"
                      />
                    </div>
                  </div>

                  {/* Row 3 */}
                  <div className="grid grid-cols-12 items-center py-2.5 px-4 text-xs font-semibold text-slate-700">
                    <div className="col-span-2 text-center font-mono">3</div>
                    <div className="col-span-7 text-[13px]">
                      Tâm lý (Cảm xúc, hành vi, lo âu...)
                    </div>
                    <div className="col-span-3 flex justify-center">
                      <input
                        type="text"
                        maxLength={1}
                        value={diary.ut_tam_ly}
                        onChange={(e) =>
                          handlePriorityChange("ut_tam_ly", e.target.value)
                        }
                        placeholder="--"
                        className="w-12 text-center py-1.5 rounded border border-[#EE6C4D] bg-[#FFF7F2] orange-text font-bold focus:ring-2 focus:ring-[#EE6C4D]/10 outline-none shadow-sm text-sm"
                      />
                    </div>
                  </div>

                  {/* Row 4 */}
                  <div className="grid grid-cols-12 items-center py-2.5 px-4 text-xs font-semibold text-slate-700">
                    <div className="col-span-2 text-center font-mono">4</div>
                    <div className="col-span-7 text-[13px]">
                      Kỹ năng sống &amp; Giá trị sống
                    </div>
                    <div className="col-span-3 flex justify-center">
                      <input
                        type="text"
                        maxLength={1}
                        value={diary.ut_ky_nang_song}
                        onChange={(e) =>
                          handlePriorityChange(
                            "ut_ky_nang_song",
                            e.target.value,
                          )
                        }
                        placeholder="--"
                        className="w-12 text-center py-1.5 rounded border border-[#EE6C4D] bg-[#FFF7F2] orange-text font-bold focus:ring-2 focus:ring-[#EE6C4D]/10 outline-none shadow-sm text-sm"
                      />
                    </div>
                  </div>

                  {/* Row 5 */}
                  <div className="grid grid-cols-12 items-center py-2.5 px-4 text-xs font-semibold text-slate-700">
                    <div className="col-span-2 text-center font-mono">5</div>
                    <div className="col-span-7 text-[13px]">
                      Hướng nghiệp, định hướng nghề nghiệp
                    </div>
                    <div className="col-span-3 flex justify-center">
                      <input
                        type="text"
                        maxLength={1}
                        value={diary.ut_huong_nghiep}
                        onChange={(e) =>
                          handlePriorityChange(
                            "ut_huong_nghiep",
                            e.target.value,
                          )
                        }
                        placeholder="--"
                        className="w-12 text-center py-1.5 rounded border border-[#EE6C4D] bg-[#FFF7F2] orange-text font-bold focus:ring-2 focus:ring-[#EE6C4D]/10 outline-none shadow-sm text-sm"
                      />
                    </div>
                  </div>

                  {/* Row 6 */}
                  <div className="grid grid-cols-12 items-center py-2.5 px-4 text-xs font-semibold text-slate-700">
                    <div className="col-span-2 text-center font-mono">6</div>
                    <div className="col-span-7 text-[13px]">
                      Chính sách, pháp luật về giáo dục &amp; trẻ em
                    </div>
                    <div className="col-span-3 flex justify-center">
                      <input
                        type="text"
                        maxLength={1}
                        value={diary.ut_chinh_sach}
                        onChange={(e) =>
                          handlePriorityChange("ut_chinh_sach", e.target.value)
                        }
                        placeholder="--"
                        className="w-12 text-center py-1.5 rounded border border-[#EE6C4D] bg-[#FFF7F2] orange-text font-bold focus:ring-2 focus:ring-[#EE6C4D]/10 outline-none shadow-sm text-sm"
                      />
                    </div>
                  </div>

                  {/* Row 7 */}
                  <div className="grid grid-cols-12 items-center py-2.5 px-4 text-xs font-semibold text-slate-700">
                    <div className="col-span-2 text-center font-mono">7</div>
                    <div className="col-span-7 text-[13px]">
                      Dịch vụ công tác xã hội cho người học
                    </div>
                    <div className="col-span-3 flex justify-center">
                      <input
                        type="text"
                        maxLength={1}
                        value={diary.ut_dich_vu_ctxh}
                        onChange={(e) =>
                          handlePriorityChange(
                            "ut_dich_vu_ctxh",
                            e.target.value,
                          )
                        }
                        placeholder="--"
                        className="w-12 text-center py-1.5 rounded border border-[#EE6C4D] bg-[#FFF7F2] orange-text font-bold focus:ring-2 focus:ring-[#EE6C4D]/10 outline-none shadow-sm text-sm"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* 3. Hình thức tư vấn */}
            <div className="space-y-4 pt-4 border-t border-slate-100">
              <h4 className="text-[11px] font-bold text-[#0B2545] uppercase tracking-wider block mb-1">
                3. Hình thức hỗ trợ tư vấn học sinh
              </h4>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Trực tiếp Group */}
                <div className="border border-emerald-100 rounded-xl p-5 bg-[#E8F5E9]/30 space-y-3.5 shadow-soft">
                  <span className="inline-block bg-emerald-600 text-white text-[10px] uppercase font-mono font-extrabold px-2 py-0.5 rounded">
                    Hình thức: Trực tiếp
                  </span>

                  <div className="space-y-1">
                    <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                      Địa điểm hỗ trợ
                    </label>
                    <input
                      type="text"
                      name="tt_dia_diem"
                      placeholder="Địa điểm"
                      value={diary.tt_dia_diem}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div className="space-y-1">
                      <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                        Thời gian
                      </label>
                      <input
                        type="text"
                        name="tt_thoi_gian"
                        placeholder="Thời gian"
                        value={diary.tt_thoi_gian}
                        onChange={handleInputChange}
                        className="input-base bg-white"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                        Phút
                      </label>
                      <input
                        type="number"
                        name="tt_thoi_luong"
                        placeholder="Phút"
                        value={diary.tt_thoi_luong}
                        onChange={handleInputChange}
                        className="input-base bg-white"
                      />
                    </div>
                  </div>
                </div>

                {/* Trực tuyến Group */}
                <div className="border border-emerald-100 rounded-xl p-5 bg-[#E8F5E9]/30 space-y-3.5 shadow-soft">
                  <span className="inline-block bg-[#0B2545] text-white text-[10px] uppercase font-mono font-extrabold px-2 py-0.5 rounded">
                    Hình thức: Trực tuyến
                  </span>

                  <div className="space-y-1">
                    <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                      Kênh (Zalo, Meet)
                    </label>
                    <input
                      type="text"
                      name="on_kenh"
                      placeholder="Kênh (Zalo, Meet)"
                      value={diary.on_kenh}
                      onChange={handleInputChange}
                      className="input-base bg-white"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div className="space-y-1">
                      <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                        Thời gian
                      </label>
                      <input
                        type="text"
                        name="on_thoi_gian"
                        placeholder="Thời gian"
                        value={diary.on_thoi_gian}
                        onChange={handleInputChange}
                        className="input-base bg-white"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                        Phút
                      </label>
                      <input
                        type="number"
                        name="on_thoi_luong"
                        placeholder="Phút"
                        value={diary.on_thoi_luong}
                        onChange={handleInputChange}
                        className="input-base bg-white"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* BLOCK 3: Nhật ký tự luận */}
          <div
            id="khoi-3"
            ref={section3Ref}
            className="bg-white rounded-2xl shadow-soft border-l-4 border-emerald-600 p-6 lg:p-8 space-y-6"
          >
            <div>
              <span className="text-[10px] font-mono font-bold text-emerald-700 uppercase tracking-widest block mb-1">
                KHỐI 3
              </span>
              <h3 className="text-base font-extrabold text-[#0B2545] uppercase tracking-wide">
                Nhật ký tự luận chi tiết
              </h3>
            </div>

            <div className="space-y-5">
              {/* Question 4 */}
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  4. Khó khăn và nhu cầu của người học{" "}
                  <span className="text-red-500">*</span>
                </label>
                <textarea
                  name="muc_4_kho_khan_nhu_cau"
                  rows={calculateRows(diary.muc_4_kho_khan_nhu_cau)}
                  value={diary.muc_4_kho_khan_nhu_cau}
                  onChange={handleInputChange}
                  placeholder="Ghi chép các vấn đề khó khăn mà học sinh đang phải đối mặt, cùng các nguyện vọng mong muốn được hỗ trợ..."
                  className="input-base green-bg focus:bg-white resize-y min-h-[80px]"
                />
              </div>

              {/* Question 5 */}
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  5. Tóm tắt thông tin về người học{" "}
                  <span className="text-red-500">*</span>
                </label>
                <textarea
                  name="muc_5_tom_tat_thong_tin"
                  rows={calculateRows(diary.muc_5_tom_tat_thong_tin)}
                  value={diary.muc_5_tom_tat_thong_tin}
                  onChange={handleInputChange}
                  placeholder="Tổng quan thông tin thu thập được qua trao đổi (Ghi chép ngắn gọn các điểm tác động trực tiếp/gián tiếp đến em)..."
                  className="input-base green-bg focus:bg-white resize-y min-h-[80px]"
                />
              </div>

              {/* Question 6 */}
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  6. Nhận định sơ bộ của cán bộ tư vấn{" "}
                  <span className="text-red-500">*</span>
                </label>
                <textarea
                  name="muc_6_nhan_dinh_so_bo"
                  rows={calculateRows(diary.muc_6_nhan_dinh_so_bo)}
                  value={diary.muc_6_nhan_dinh_so_bo}
                  onChange={handleInputChange}
                  placeholder="Ý kiến chuyên môn sơ bộ của bạn sau khi quan sát thái độ, nét mặt, lời bình của học sinh..."
                  className="input-base green-bg focus:bg-white resize-y min-h-[80px]"
                />
              </div>

              {/* Question 7 */}
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  7. Các hình thức tư vấn đã áp dụng{" "}
                  <span className="text-red-500">*</span>
                </label>
                <textarea
                  name="muc_7_hinh_thuc_da_ap_dung"
                  rows={calculateRows(diary.muc_7_hinh_thuc_da_ap_dung)}
                  value={diary.muc_7_hinh_thuc_da_ap_dung}
                  onChange={handleInputChange}
                  placeholder="Ghi nhận các phương án trực tiếp, trực tuyến, làm việc nhóm hoặc hỗ trợ vật lý đã dùng..."
                  className="input-base green-bg focus:bg-white resize-y min-h-[80px]"
                />
              </div>

              {/* Question 8 */}
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  8. Đánh giá hiệu quả sau buổi tư vấn{" "}
                  <span className="text-red-500">*</span>
                </label>
                <textarea
                  name="muc_8_danh_gia_hieu_qua"
                  rows={calculateRows(diary.muc_8_danh_gia_hieu_qua)}
                  value={diary.muc_8_danh_gia_hieu_qua}
                  onChange={handleInputChange}
                  placeholder="Mức độ cải thiện tâm trạng, hiểu biết, tinh thần cải thiện thế nào sau tiến trình tư vấn?..."
                  className="input-base green-bg focus:bg-white resize-y min-h-[80px]"
                />
              </div>
            </div>
          </div>

          {/* BLOCK 4: Kết thúc và Ký tên */}
          <div
            id="khoi-4"
            ref={section4Ref}
            className="bg-white rounded-2xl shadow-soft border-l-4 border-[#D90429] p-6 lg:p-8 space-y-6"
          >
            <div>
              <span className="text-[10px] font-mono font-bold text-[#D90429] uppercase tracking-widest block mb-1">
                KHỐI 4
              </span>
              <h3 className="text-base font-extrabold text-[#0B2545] uppercase tracking-wide">
                Kết thúc &amp; Ký tên xác thực
              </h3>
            </div>

            {/* Part 9: Kết thúc hành trình */}
            <div className="space-y-3.5 border-b border-slate-100 pb-5">
              <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                9. Phương thức kết thúc tư vấn{" "}
                <span className="text-red-500">*</span>
              </label>

              <div className="space-y-2">
                <button
                  type="button"
                  onClick={() => selectEndingMethod("dung_theo_doi")}
                  className={`w-full text-left p-3.5 rounded-lg text-xs font-bold transition flex items-center gap-3 border ${
                    diary.kt_dung_theo_doi === "X"
                      ? "bg-red-50/75 text-[#D90429] border-[#D90429]/55 shadow-soft"
                      : "bg-[#F8F9FA] text-slate-600 border-slate-200 hover:bg-slate-50"
                  }`}
                >
                  <span
                    className={`w-4 h-4 rounded-full border-2 flex items-center justify-center transition ${
                      diary.kt_dung_theo_doi === "X"
                        ? "border-[#D90429] bg-[#D90429]"
                        : "border-slate-300"
                    }`}
                  >
                    {diary.kt_dung_theo_doi === "X" && (
                      <span className="w-1.5 h-1.5 rounded-full bg-white"></span>
                    )}
                  </span>
                  <span>
                    Dừng hoạt động tư vấn học đường và công tác xã hội và chuyển
                    sang theo dõi
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => selectEndingMethod("len_ke_hoach")}
                  className={`w-full text-left p-3.5 rounded-lg text-xs font-bold transition flex items-center gap-3 border ${
                    diary.kt_len_ke_hoach === "X"
                      ? "bg-red-50/75 text-[#D90429] border-[#D90429]/55 shadow-soft"
                      : "bg-[#F8F9FA] text-slate-600 border-slate-200 hover:bg-slate-50"
                  }`}
                >
                  <span
                    className={`w-4 h-4 rounded-full border-2 flex items-center justify-center transition ${
                      diary.kt_len_ke_hoach === "X"
                        ? "border-[#D90429] bg-[#D90429]"
                        : "border-slate-300"
                    }`}
                  >
                    {diary.kt_len_ke_hoach === "X" && (
                      <span className="w-1.5 h-1.5 rounded-full bg-white"></span>
                    )}
                  </span>
                  <span>Xây dựng kế hoạch các đợt tư vấn tiếp theo</span>
                </button>

                <div className="space-y-2">
                  <button
                    type="button"
                    onClick={() => selectEndingMethod("thuc_hien_chuyen")}
                    className={`w-full text-left p-3.5 rounded-lg text-xs font-bold transition flex items-center gap-3 border ${
                      diary.kt_thuc_hien_chuyen === "X"
                        ? "bg-red-50/75 text-[#D90429] border-[#D90429]/55 shadow-soft"
                        : "bg-[#F8F9FA] text-slate-600 border-slate-200 hover:bg-slate-50"
                    }`}
                  >
                    <span
                      className={`w-4 h-4 rounded-full border-2 flex items-center justify-center transition ${
                        diary.kt_thuc_hien_chuyen === "X"
                          ? "border-[#D90429] bg-[#D90429]"
                          : "border-slate-300"
                      }`}
                    >
                      {diary.kt_thuc_hien_chuyen === "X" && (
                        <span className="w-1.5 h-1.5 rounded-full bg-white"></span>
                      )}
                    </span>
                    <span>Thực hiện chuyển gửi người học</span>
                  </button>

                  <AnimatePresence>
                    {diary.kt_thuc_hien_chuyen === "X" && (
                      <motion.div
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: "auto" }}
                        exit={{ opacity: 0, height: 0 }}
                        className="overflow-hidden pl-7 pt-1"
                      >
                        <div className="space-y-1 bg-red-50/50 p-3 rounded-lg border border-red-200/55">
                          <label className="text-[11px] font-bold text-red-800 uppercase tracking-wider block">
                            Nơi nhận chuyển gửi của học sinh{" "}
                            <span className="text-red-500">*</span>
                          </label>
                          <input
                            type="text"
                            name="kt_chuyen_gui_noi"
                            placeholder="Ví dụ: Trung tâm Y tế Bảo Lộc hoặc Viện sức khỏe tâm thần..."
                            value={diary.kt_chuyen_gui_noi}
                            onChange={handleInputChange}
                            className="input-base"
                          />
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              </div>
            </div>

            {/* Part 10 & Counselor signature */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5 pt-1">
              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  10. Lịch hẹn gặp tiếp theo (nếu có)
                </label>
                <input
                  type="text"
                  name="lich_hen_tiep_theo"
                  placeholder="Ví dụ: Thứ Hai tuần sau, ngày 15/06"
                  value={diary.lich_hen_tiep_theo}
                  onChange={handleInputChange}
                  className="input-base"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[11px] font-bold text-[#0B2545] uppercase mb-1 block">
                  Cán bộ tư vấn ký tên <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  name="ten_nguoi_thuc_hien"
                  placeholder="Họ tên đầy đủ người ký"
                  value={diary.ten_nguoi_thuc_hien}
                  onChange={handleInputChange}
                  className="input-base font-bold text-[#0B2545]"
                />
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Sticky Bottom Toolbelt Action Control Section */}
      <footer className="fixed bottom-0 left-0 right-0 bg-[#0B2545] text-white border-t-2 border-[#EE6C4D] py-4 px-6 z-40 shadow-soft">
        <div className="max-w-5xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4">
          {/* File Template Upload Interface tool */}
          <div className="flex flex-col sm:flex-row items-center gap-3 w-full md:w-auto">
            <span className="text-xs font-bold text-slate-300 uppercase tracking-widest hidden sm:inline">
              Mẫu DOCX:
            </span>
            <label className="w-full sm:w-auto relative cursor-pointer group">
              <input
                type="file"
                accept=".docx"
                onChange={handleDocxUpload}
                className="hidden"
              />
              <div className="w-full select-none text-center bg-white/10 hover:bg-white/15 border border-white/20 text-white py-2.5 px-4 rounded-lg text-xs font-bold transition flex items-center justify-center gap-2">
                📂 Tải Lên Mẫu File Word Riêng (.docx)
              </div>
            </label>
            {uploadedTemplateName && (
              <button
                type="button"
                onClick={() => {
                  setCustomTemplateArrayBuffer(null);
                  setUploadedTemplateName(null);
                  showStatus(
                    "Đã quay lại sử dụng mẫu mặc định của trường Phan Chu Trinh",
                    "info",
                  );
                }}
                className="text-xs text-red-400 hover:text-red-300 underline underline-offset-4 cursor-pointer select-none font-bold"
              >
                Gỡ tệp custom X
              </button>
            )}
          </div>

          {/* Core Controls execution buttons */}
          <div className="flex items-center gap-3 w-full md:w-auto justify-end">
            <button
              onClick={resetForm}
              className="px-4 py-3 bg-transparent hover:bg-white/10 text-slate-200 border border-slate-500 rounded-lg text-xs font-bold transition font-mono uppercase tracking-wider"
            >
              Nhập Lại Dữ Liệu
            </button>

            <button
              onClick={handleExportWord}
              className="flex-1 sm:flex-initial text-center bg-[#D90429] hover:bg-red-600 active:scale-95 text-white py-3 px-6 rounded-lg text-xs font-extrabold tracking-widest transition shadow-soft font-mono uppercase"
            >
              XUẤT FILE WORD (.DOCX)
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
}
