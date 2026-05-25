var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// server.ts
var import_express = __toESM(require("express"), 1);
var import_path = __toESM(require("path"), 1);
var import_fs = __toESM(require("fs"), 1);
var XLSX = __toESM(require("xlsx"), 1);
var import_vite = require("vite");
var app = (0, import_express.default)();
var PORT = 3e3;
app.use(import_express.default.json());
var DEFAULT_EXCEL_PATH = import_path.default.join(process.cwd(), "booking.xlsx");
function initDefaultExcel() {
  if (!import_fs.default.existsSync(DEFAULT_EXCEL_PATH)) {
    console.log("Initializing default booking.xlsx file on disk...");
    const sampleData = [
      {
        "Name": "Janadeepa",
        "Phone": "077 834 1657",
        "Facility": "Net Sessions",
        "Start Time": "06:03 PM",
        "End Time": "07:03 PM",
        "Status": "Confirmed"
      },
      {
        "Name": "Janadeepa",
        "Phone": "077 834 1657",
        "Facility": "Net Sessions",
        "Start Time": "02:26 PM",
        "End Time": "03:26 PM",
        "Status": "Confirmed"
      },
      {
        "Name": "Bawani",
        "Phone": "072 228 3147",
        "Facility": "Net Sessions",
        "Start Time": "01:30 PM",
        "End Time": "02:30 PM",
        "Status": "Confirmed"
      },
      {
        "Name": "Dilki",
        "Phone": "077 987 4564",
        "Facility": "Net Sessions",
        "Start Time": "01:40 PM",
        "End Time": "03:40 PM",
        "Status": "Confirmed"
      }
    ];
    try {
      const ws = XLSX.utils.json_to_sheet(sampleData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Today's Bookings");
      XLSX.writeFile(wb, DEFAULT_EXCEL_PATH);
      console.log("Default booking.xlsx successfully created!");
    } catch (err) {
      console.error("Failed to create default booking.xlsx:", err);
    }
  }
}
initDefaultExcel();
function getVal(row, keys) {
  const rowKeys = Object.keys(row);
  for (const k of keys) {
    const matchedKey = rowKeys.find((rk) => rk.toLowerCase().trim() === k.toLowerCase());
    if (matchedKey && row[matchedKey] !== void 0 && row[matchedKey] !== null) {
      return String(row[matchedKey]);
    }
  }
  return void 0;
}
function parseTimeValue(val) {
  if (val === void 0 || val === null) return "";
  const str = String(val).trim();
  const num = Number(str);
  if (!isNaN(num) && num > 0 && num < 1) {
    const totalMinutes = Math.round(num * 24 * 60);
    const hours24 = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    const period = hours24 >= 12 ? "PM" : "AM";
    const hours12 = hours24 % 12 === 0 ? 12 : hours24 % 12;
    const minStr = minutes < 10 ? `0${minutes}` : `${minutes}`;
    const hrStr = hours12 < 10 ? `0${hours12}` : `${hours12}`;
    return `${hrStr}:${minStr} ${period}`;
  }
  return str;
}
function parseExcelData(jsonData) {
  const nameKeys = ["name", "customer", "client", "player", "user", "naam", "lead", "booked by", "passenger"];
  const phoneKeys = ["phone no.", "phone no", "phone", "contact", "mobile", "telephone", "phone number", "number", "tel", "contact number"];
  const facilityKeys = ["type", "facility", "net", "lane", "booking type", "resource", "nets", "facility booked"];
  const timeKeys = ["time", "start time", "start", "time from", "starttime", "from", "start_time", "slot"];
  const statusKeys = ["status", "state", "booking status", "approved", "confirmed"];
  return jsonData.map((row, idx) => {
    const noKeys = ["no", "id", "serial"];
    const rowIdRaw = getVal(row, noKeys);
    const parsedId = rowIdRaw ? parseInt(rowIdRaw, 10) : idx + 1;
    const name = getVal(row, nameKeys) || `Guest ${idx + 1}`;
    const phone = getVal(row, phoneKeys) || "077 000 0000";
    const facility = getVal(row, facilityKeys) || "Net Sessions";
    const rawTime = getVal(row, timeKeys) || "09:00 AM - 10:00 AM";
    const statusVal = getVal(row, statusKeys) || "Confirmed";
    return {
      id: isNaN(parsedId) ? idx + 1 : parsedId,
      name,
      phone,
      facility,
      time: parseTimeValue(rawTime),
      status: String(statusVal).trim().toLowerCase().includes("pending") ? "Pending" : "Confirmed"
    };
  });
}
app.get("/api/bookings", async (req, res) => {
  const rawUrl = req.query.url || "https://www.dropbox.com/scl/fi/zsr8s25h7khrqiq3hegtx/Booking.xlsx?rlkey=x8r0yq1n4a61w148hz97o4tl3&st=mc1zyydf&dl=0";
  try {
    let workbook;
    if (rawUrl) {
      console.log("Fetching custom bookings file from URL:", rawUrl);
      let downloadUrl = rawUrl;
      if (downloadUrl.includes("dropbox.com")) {
        downloadUrl = downloadUrl.replace("www.dropbox.com", "dl.dropboxusercontent.com");
        downloadUrl = downloadUrl.replace("://dropbox.com", "://dl.dropboxusercontent.com");
        if (downloadUrl.includes("dl=0")) {
          downloadUrl = downloadUrl.replace("dl=0", "dl=1");
        } else if (!downloadUrl.includes("?")) {
          downloadUrl += "?dl=1";
        }
      }
      const response = await fetch(downloadUrl);
      if (!response.ok) {
        throw new Error(`Dropbox returned status ${response.status}`);
      }
      const buffer = await response.arrayBuffer();
      workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
    } else {
      console.log("Reading bookings from local booking.xlsx...");
      if (!import_fs.default.existsSync(DEFAULT_EXCEL_PATH)) {
        initDefaultExcel();
      }
      workbook = XLSX.readFile(DEFAULT_EXCEL_PATH);
    }
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawJson = XLSX.utils.sheet_to_json(worksheet);
    const formattedData = parseExcelData(rawJson);
    return res.json({ success: true, count: formattedData.length, data: formattedData });
  } catch (error) {
    console.error("Booking retrieval or parsing failure:", error.message);
    try {
      console.log("Error occurred fetching Dropbox link. Falling back to local file parsing...");
      const workbook = XLSX.readFile(DEFAULT_EXCEL_PATH);
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rawJson = XLSX.utils.sheet_to_json(worksheet);
      const formattedData = parseExcelData(rawJson);
      return res.json({
        success: true,
        count: formattedData.length,
        data: formattedData,
        warning: "Custom Dropbox sync link failed to load. Displaying local booking.xlsx contents."
      });
    } catch (fallbackError) {
      return res.status(500).json({ success: false, error: error.message });
    }
  }
});
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await (0, import_vite.createServer)({
      server: { middlewareMode: true },
      appType: "spa"
    });
    app.use(vite.middlewares);
  } else {
    const distPath = import_path.default.join(process.cwd(), "dist");
    app.use(import_express.default.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(import_path.default.join(distPath, "index.html"));
    });
  }
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Falcon Sport Complex Server booted successfully. Accessible at http://localhost:${PORT}`);
  });
}
startServer();
//# sourceMappingURL=server.cjs.map
