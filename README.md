# Biometric-Report-Formatter

import { Router, type IRouter } from "express";

import multer from "multer";

import * as XLSX from "xlsx";
 
const router: IRouter = Router();

const upload = multer({ storage: multer.memoryStorage() });
 
function calcDuration(inTime: string, outTime: string): string {

  const [inH, inM] = inTime.split(":").map(Number);

  const [outH, outM] = outTime.split(":").map(Number);

  let diffMinutes = (outH * 60 + outM) - (inH * 60 + inM);

  if (diffMinutes < 0) diffMinutes += 24 * 60; // handles overnight shifts

  const hours = Math.floor(diffMinutes / 60);

  const mins = diffMinutes % 60;

  return `${String(hours).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;

}
 
function formatDate(d: Date): string {

  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  return `${String(d.getDate()).padStart(2, "0")}-${months[d.getMonth()]}-${d.getFullYear()}`;

}
 
router.post("/process-biometric", upload.single("file"), (req, res) => {

  if (!req.file) {

    res.status(400).json({ error: "No file uploaded" });

    return;

  }
 
  const workbook = XLSX.read(req.file.buffer, { type: "buffer" });

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const raw: (string | number | null)[][] = XLSX.utils.sheet_to_json(sheet, {

    header: 1,

    defval: null,

  });
 
  // Parse date range from row 2 e.g. "Mar 26 2026  To  Apr 25 2026"

  const dateRangeCell = String(raw[2]?.[1] ?? "");

  const dateMatches = dateRangeCell.match(/([A-Za-z]+ \d+ \d+)/g);

  if (!dateMatches || dateMatches.length < 2) {

    res.status(400).json({ error: "Could not parse date range from report" });

    return;

  }
 
  const startDate = new Date(dateMatches[0].trim());

  const endDate   = new Date(dateMatches[1].trim());
 
  // Build ordered list of all dates in range

  const orderedDates: Date[] = [];

  const cur = new Date(startDate);

  while (cur <= endDate) {

    orderedDates.push(new Date(cur));

    cur.setDate(cur.getDate() + 1);

  }
 
  // Map column index → formatted date string using row 6 (day headers)

  const datesRow = raw[6] ?? [];

  const colToDate: Record<number, string> = {};

  let dateIdx = 0;

  for (let col = 2; col < datesRow.length; col++) {

    const cell = datesRow[col];

    if (cell == null) continue;

    const dayNum = parseInt(String(cell).trim().split(/\s+/)[0], 10);

    if (!isNaN(dayNum) && dateIdx < orderedDates.length && orderedDates[dateIdx].getDate() === dayNum) {

      colToDate[col] = formatDate(orderedDates[dateIdx]);

      dateIdx++;

    }

  }
 
  // Process every employee block

  const finalRows = [];

  for (let i = 0; i < raw.length; i++) {

    const row = raw[i];

    if (!row) continue;

    if (String(row[0] ?? "").trim() !== "Emp. Code:") continue;
 
    const employeeName = String(raw[i]?.[13] ?? "").trim();

    const inTimeRow  = raw[i + 2] ?? [];

    const outTimeRow = raw[i + 3] ?? [];
 
    for (const [colStr, punchDate] of Object.entries(colToDate)) {

      const col = Number(colStr);

      const inVal  = inTimeRow[col];

      const outVal = outTimeRow[col];

      if (inVal == null && outVal == null) continue;
 
      const inTime  = inVal  != null ? String(inVal).trim()  : "";

      const outTime = outVal != null ? String(outVal).trim() : "";

      const duration = inTime && outTime ? calcDuration(inTime, outTime) : "";
 
      finalRows.push({

        "Employee Name": employeeName,

        "Punch Date":    punchDate,

        "In Time":       inTime,
 
