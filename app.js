/* ==========================================================
   Eastern Oak College of Nursing Sciences
   Hostel Attendance Report Builder
   Everything runs locally in the browser — no file is ever
   uploaded anywhere. Output (.xlsx and .pdf) matches the
   college's standard "Hostel Absentees" report template:
   Name/Date | Hostel | <day columns as 16th, 17th, ...> | Total Fine
   ========================================================== */

(() => {
  "use strict";

  const RATE_DEFAULT = 15000;

  const REQUIRED_COLUMNS = [
    "Department",
    "Name",
    "No.",
    "Date/Time",
    "Status",
    "Location ID",
  ];

  const MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ];

  // ---------- DOM ----------
  const fileInput = document.getElementById("fileInput");
  const dropzone = document.getElementById("dropzone");
  const filecheck = document.getElementById("filecheck");
  const generateBtn = document.getElementById("generateBtn");
  const statusEl = document.getElementById("status");
  const resultsEl = document.getElementById("results");
  const summaryEl = document.getElementById("summary");
  const previewTable = document.getElementById("previewTable");

  const startDateInput = document.getElementById("startDate");
  const endDateInput = document.getElementById("endDate");
  const rateInput = document.getElementById("ratePerDay");
  const hostel1Input = document.getElementById("hostel1Name");
  const hostel2Input = document.getElementById("hostel2Name");

  let parsedRows = null; // raw sheet_to_json rows for the chosen file

  // Default the date range to the current month, so first-time users
  // already see something sensible rather than a blank field.
  (function seedDefaultDates() {
    const today = new Date();
    const first = new Date(today.getFullYear(), today.getMonth(), 1);
    startDateInput.value = toInputDate(first);
    endDateInput.value = toInputDate(today);
  })();

  // ---------- File selection ----------

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("is-dragover");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("is-dragover"));
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("is-dragover");
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      fileInput.files = e.dataTransfer.files;
      handleFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files && fileInput.files[0]) {
      handleFile(fileInput.files[0]);
    }
  });

  function handleFile(file) {
    setStatus("", "");
    resultsEl.hidden = true;
    showFilecheck("Reading the file…", "ok");

    const reader = new FileReader();
    reader.onerror = () => showFilecheck("We couldn't open that file. Please try again.", "err");
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });

        const validation = validateHeaders(rows);
        if (!validation.ok) {
          parsedRows = null;
          showFilecheck(validation.message, "err");
          return;
        }

        parsedRows = rows;
        const studentCount = countLikelyStudents(rows, validation.headerIndex);
        showFilecheck(
          `Looking good — "${escapeHtml(file.name)}" is ready. Found about ${studentCount} attendance records.`,
          "ok"
        );
      } catch (err) {
        console.error(err);
        parsedRows = null;
        showFilecheck(
          "We couldn't read that file. Please make sure it's the .xls or .xlsx file exported from the gate machine.",
          "err"
        );
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function validateHeaders(rows) {
    if (!rows.length) {
      return { ok: false, message: "That file looks empty. Please choose a different file." };
    }
    const headerRow = rows[0];
    const headerIndex = {};
    headerRow.forEach((value, i) => {
      const key = String(value).trim();
      if (key) headerIndex[key] = i;
    });
    const missing = REQUIRED_COLUMNS.filter((c) => !(c in headerIndex));
    if (missing.length) {
      return {
        ok: false,
        message:
          "This file is missing some columns we need: " +
          missing.map((m) => `"${m}"`).join(", ") +
          ". Please upload the original file exported from the gate machine, unedited.",
      };
    }
    return { ok: true, headerIndex };
  }

  function countLikelyStudents(rows, headerIndex) {
    const nameCol = headerIndex["Name"];
    const noCol = headerIndex["No."];
    const seen = new Set();
    for (let r = 1; r < rows.length; r++) {
      const name = rows[r][nameCol];
      const no = rows[r][noCol];
      if (name && no !== "") seen.add(String(no).trim());
    }
    return seen.size;
  }

  function showFilecheck(message, kind) {
    filecheck.hidden = false;
    filecheck.className = "filecheck " + kind;
    filecheck.innerHTML = message;
  }

  // ---------- Generate ----------

  generateBtn.addEventListener("click", async () => {
    setStatus("", "");

    if (!parsedRows) {
      setStatus("Please choose a valid attendance file first.", "is-error");
      return;
    }
    const startDate = parseInputDate(startDateInput.value);
    const endDate = parseInputDate(endDateInput.value);
    if (!startDate || !endDate) {
      setStatus("Please set both a start date and an end date.", "is-error");
      return;
    }
    if (endDate < startDate) {
      setStatus("The end date can't be before the start date.", "is-error");
      return;
    }
    const rate = Number(rateInput.value);
    if (!Number.isFinite(rate) || rate < 0) {
      setStatus("Please enter a valid rate per absent day.", "is-error");
      return;
    }
    const hostel1Name = (hostel1Input.value || "Hostel 1").trim();
    const hostel2Name = (hostel2Input.value || "Hostel 2").trim();

    generateBtn.disabled = true;
    setStatus("Working on it…", "");

    try {
      const { students, attendance, dates } = computeAttendance(
        parsedRows,
        startDate,
        endDate,
        hostel1Name,
        hostel2Name
      );

      if (Object.keys(students).length === 0) {
        setStatus(
          "No students were found in this date range. Double-check the dates and the file.",
          "is-error"
        );
        return;
      }

      const dayHeaders = buildDayHeaders(dates);
      renderPreview(students, attendance, dates, dayHeaders, rate);

      const baseName = buildOutputBaseName(startDate, endDate);
      await downloadWorkbook(students, attendance, dates, dayHeaders, rate, baseName + ".xlsx");
      downloadPdf(students, attendance, dates, dayHeaders, rate, baseName + ".pdf");

      setStatus(`Done — "${baseName}.xlsx" and "${baseName}.pdf" have been downloaded.`, "is-ok");
    } catch (err) {
      console.error(err);
      setStatus("Something went wrong while building the report. Please try again.", "is-error");
    } finally {
      generateBtn.disabled = false;
    }
  });

  // ---------- Core computation (mirrors the Python script) ----------

  function computeAttendance(rows, startDate, endDate, hostel1Name, hostel2Name) {
    const headerRow = rows[0];
    const headerIndex = {};
    headerRow.forEach((value, i) => {
      const key = String(value).trim();
      if (key) headerIndex[key] = i;
    });

    const nameCol = headerIndex["Name"];
    const noCol = headerIndex["No."];
    const dateCol = headerIndex["Date/Time"];
    const statusCol = headerIndex["Status"];
    const locationCol = headerIndex["Location ID"];

    const students = {}; // id -> { name, hostel }
    const attendance = {}; // id -> Set of "YYYY-MM-DD"

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      const rawName = row[nameCol];
      const studentIdRaw = row[noCol];
      if (!rawName || studentIdRaw === "" || studentIdRaw === undefined || studentIdRaw === null) {
        continue;
      }

      const name = String(rawName).trim();
      const studentId = normalizeStudentId(studentIdRaw);

      const locationRaw = String(row[locationCol] || "").trim();
      let hostel;
      if (locationRaw === "1") hostel = hostel1Name;
      else if (locationRaw === "2") hostel = hostel2Name;
      else hostel = "Unknown";

      if (!(studentId in students)) {
        students[studentId] = { name, hostel };
        attendance[studentId] = new Set();
      }

      const rawDate = row[dateCol];
      if (rawDate === "" || rawDate === undefined || rawDate === null) continue;

      const attendanceDate = parseAttendanceDate(rawDate);
      if (!attendanceDate) continue;
      if (attendanceDate < startDate || attendanceDate > endDate) continue;

      const status = String(row[statusCol] || "").trim().toUpperCase();
      if (status === "C/IN") {
        attendance[studentId].add(toDateKey(attendanceDate));
      }
    }

    const dates = [];
    const cursor = new Date(startDate);
    while (cursor <= endDate) {
      dates.push(new Date(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }

    return { students, attendance, dates };
  }

  function normalizeStudentId(raw) {
    if (typeof raw === "number") {
      return Number.isInteger(raw) ? String(raw) : String(raw);
    }
    return String(raw).trim();
  }

  // Handles: native JS Date (from SheetJS cellDates), Excel serial numbers,
  // and the "M/D/YYYY h:mm:ss AM/PM" text format the original script expects.
  function parseAttendanceDate(raw) {
    if (raw instanceof Date && !isNaN(raw)) {
      return dateOnly(raw);
    }
    if (typeof raw === "number") {
      const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(raw) : null;
      if (parsed) {
        return dateOnly(new Date(parsed.y, parsed.m - 1, parsed.d));
      }
      return null;
    }
    const text = String(raw).trim();
    if (!text) return null;

    // Try "M/D/YYYY h:mm:ss AM/PM" (the format the source machine exports)
    const m = text.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})[, ]+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?$/i
    );
    if (m) {
      let [, mo, d, y] = m;
      return dateOnly(new Date(Number(y), Number(mo) - 1, Number(d)));
    }

    // Fall back to whatever the browser can parse
    const generic = new Date(text);
    if (!isNaN(generic)) return dateOnly(generic);

    return null;
  }

  function dateOnly(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function toDateKey(d) {
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  function pad(n) {
    return String(n).padStart(2, "0");
  }

  function toInputDate(d) {
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  function parseInputDate(value) {
    if (!value) return null;
    const [y, m, d] = value.split("-").map(Number);
    if (!y || !m || !d) return null;
    return new Date(y, m - 1, d);
  }

  function ordinalSuffix(n) {
    const v = n % 100;
    if (v >= 11 && v <= 13) return "th";
    switch (n % 10) {
      case 1: return "st";
      case 2: return "nd";
      case 3: return "rd";
      default: return "th";
    }
  }

  // Matches the template's day headers ("16th", "17th", ...). If the
  // range crosses a month boundary, the month is appended so the
  // columns stay unambiguous ("30th Aug", "1st Sep").
  function buildDayHeaders(dates) {
    const months = new Set(dates.map((d) => d.getMonth()));
    const spansMonths = months.size > 1;
    return dates.map((d) => {
      const label = `${d.getDate()}${ordinalSuffix(d.getDate())}`;
      return spansMonths ? `${label} ${MONTH_NAMES[d.getMonth()].slice(0, 3)}` : label;
    });
  }

  function formatNaira(n) {
    return "₦" + Math.round(n).toLocaleString("en-NG");
  }

  // Mirrors the naming convention of the college's own reports, e.g.
  // "HOSTEL_ABSENTEES_16TH_-_22ND_AUGUST_2026".
  function buildOutputBaseName(startDate, endDate) {
    const sameMonth = startDate.getMonth() === endDate.getMonth() && startDate.getFullYear() === endDate.getFullYear();
    const startLabel = `${startDate.getDate()}${ordinalSuffix(startDate.getDate())}`.toUpperCase();
    const endLabel = `${endDate.getDate()}${ordinalSuffix(endDate.getDate())}`.toUpperCase();
    if (sameMonth) {
      const month = MONTH_NAMES[startDate.getMonth()].toUpperCase();
      return `HOSTEL_ABSENTEES_${startLabel}_-_${endLabel}_${month}_${endDate.getFullYear()}`;
    }
    const startMonth = MONTH_NAMES[startDate.getMonth()].toUpperCase();
    const endMonth = MONTH_NAMES[endDate.getMonth()].toUpperCase();
    return `HOSTEL_ABSENTEES_${startLabel}_${startMonth}_-_${endLabel}_${endMonth}_${endDate.getFullYear()}`;
  }

  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  // ---------- Preview ----------

  function renderPreview(students, attendance, dates, dayHeaders, rate) {
    const ids = Object.keys(students);

    let totalMissedDays = 0;
    let totalFine = 0;

    const theadCells = ["Name/Date", "Hostel", ...dayHeaders, "Total Fine"];
    let thead = "<thead><tr>" + theadCells.map((c) => `<th>${escapeHtml(c)}</th>`).join("") + "</tr></thead>";

    let tbody = "<tbody>";
    for (const id of ids) {
      const s = students[id];
      let missed = 0;
      let cells = "";
      for (const d of dates) {
        const present = attendance[id].has(toDateKey(d));
        if (present) {
          cells += `<td></td>`;
        } else {
          cells += `<td class="absent">x</td>`;
          missed++;
        }
      }
      const fine = missed * rate;
      totalMissedDays += missed;
      totalFine += fine;
      tbody += `<tr><td>${escapeHtml(s.name)}</td><td>${escapeHtml(s.hostel)}</td>${cells}<td class="total">${fine.toLocaleString("en-NG")}</td></tr>`;
    }
    tbody += "</tbody>";

    previewTable.innerHTML = thead + tbody;

    summaryEl.innerHTML = [
      ["Students processed", ids.length],
      ["Date range", `${dayHeaders[0]} – ${dayHeaders[dayHeaders.length - 1]}`],
      ["Rate per day", formatNaira(rate)],
      ["Total absent days", totalMissedDays],
      ["Total boarding deduction", formatNaira(totalFine)],
    ].map(([label, value]) => `<div><dt>${label}</dt><dd>${value}</dd></div>`).join("");

    resultsEl.hidden = false;
    resultsEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  // ---------- Excel export (matches the college's plain print template) ----------

  async function downloadWorkbook(students, attendance, dates, dayHeaders, rate, outputName) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Eastern Oak College of Nursing Sciences";
    workbook.created = new Date();

    const ws = workbook.addWorksheet("Hostel Absentees", {
      views: [{ state: "frozen", xSplit: 2, ySplit: 1 }],
      pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    const header = ["Name/Date", "Hostel", ...dayHeaders, "Total Fine"];
    ws.addRow(header);
    ws.getRow(1).font = { bold: true };

    const ids = Object.keys(students);
    for (const id of ids) {
      const s = students[id];
      const row = [s.name, s.hostel];
      let missed = 0;
      for (const d of dates) {
        if (attendance[id].has(toDateKey(d))) {
          row.push("");
        } else {
          row.push("x");
          missed++;
        }
      }
      row.push(missed * rate);
      ws.addRow(row);
    }

    // Thin black borders on every used cell, plain alignment — matches
    // the plain printed-spreadsheet look of the template exactly.
    const totalColIndex = header.length;
    const thin = { style: "thin", color: { argb: "FF000000" } };
    for (let r = 1; r <= ws.rowCount; r++) {
      for (let c = 1; c <= totalColIndex; c++) {
        const cell = ws.getCell(r, c);
        cell.border = { top: thin, left: thin, bottom: thin, right: thin };
        if (c === 1 || c === 2) {
          cell.alignment = { horizontal: "left", vertical: "middle" };
        } else if (c === totalColIndex) {
          cell.alignment = { horizontal: "right", vertical: "middle" };
        } else {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }
    }

    ws.getColumn(1).width = 28;
    ws.getColumn(2).width = 14;
    for (let c = 3; c < totalColIndex; c++) {
      ws.getColumn(c).width = 7;
    }
    ws.getColumn(totalColIndex).width = 12;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    triggerDownload(blob, outputName);
  }

  // ---------- PDF export (matches the college's plain print template) ----------

  function downloadPdf(students, attendance, dates, dayHeaders, rate, outputName) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "letter" });

    const head = [["Name/Date", "Hostel", ...dayHeaders, "Total Fine"]];
    const body = [];
    for (const id of Object.keys(students)) {
      const s = students[id];
      let missed = 0;
      const marks = dates.map((d) => {
        const present = attendance[id].has(toDateKey(d));
        if (!present) missed++;
        return present ? "" : "x";
      });
      body.push([s.name, s.hostel, ...marks, (missed * rate).toLocaleString("en-NG")]);
    }

    const dayColCount = dayHeaders.length;
    const fontSize = dayColCount > 20 ? 6 : dayColCount > 12 ? 7 : 8;

    const columnStyles = { 0: { halign: "left" }, 1: { halign: "left" } };
    columnStyles[1 + dayColCount + 1] = { halign: "right" };

    doc.autoTable({
      head,
      body,
      theme: "grid",
      styles: {
        font: "helvetica",
        fontSize,
        cellPadding: 4,
        lineColor: [0, 0, 0],
        lineWidth: 0.75,
        textColor: [0, 0, 0],
        halign: "center",
        valign: "middle",
      },
      headStyles: {
        fillColor: false,
        textColor: [0, 0, 0],
        fontStyle: "bold",
        lineColor: [0, 0, 0],
        lineWidth: 0.75,
      },
      columnStyles,
      margin: { top: 30, left: 24, right: 24, bottom: 24 },
    });

    doc.save(outputName);
  }

  function triggerDownload(blob, outputName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = outputName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function setStatus(message, kind) {
    statusEl.textContent = message;
    statusEl.className = "status" + (kind ? " " + kind : "");
  }
})();
