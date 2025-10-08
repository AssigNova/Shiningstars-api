// bulkRegister.js
const xlsx = require("xlsx");
const axios = require("axios");

// Path to your Excel file
const filePath = "./users2.xlsx";

// Load workbook
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

console.log("Reading Excel file from:", filePath);
console.log("Available sheet names:", workbook.SheetNames);

// API endpoint
const API_URL = "http://localhost:5000/api/auth/register"; // adjust to your server

// Array to hold results
const results = [];

// Delay helper
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Excel date → JS date string
function excelDateToISO(excelDate) {
  if (!excelDate) return null;

  // If it's already a string date (e.g., "1980-09-24")
  if (typeof excelDate === "string") {
    const parsed = new Date(excelDate);
    if (!isNaN(parsed)) {
      return new Date(parsed).toISOString();
    }
  }

  // If it's a number (Excel serial date)
  if (typeof excelDate === "number") {
    const epoch = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
    return new Date(epoch).toISOString();
  }

  return null;
}

async function registerUsers() {
  for (let i = 266; i < rows.length; i++) {
    const row = rows[i];
    const employeeId = row[0]; // column A
    const name = row[2]; // column C
    const gender = row[4]; // column E
    const dateOfBirth = row[5]; // column F
    const email = row[8]; // column I
    const contactNo = row[9]; // column J
    const department = row[11]; // column L

    // Format fields
    const formattedName = name?.toString().trim().toUpperCase() || "UNKNOWN";
    const formattedEmail = email?.toString().trim().toLowerCase() || "";
    const firstWord = formattedName.split(" ")[0];
    const password = `${firstWord}@123`;

    const formattedDOB = excelDateToISO(dateOfBirth);
    const createdAt = new Date().toISOString();

    const user = {
      name: formattedName,
      email: formattedEmail,
      password,
      avatar: "",
      employeeId: employeeId?.toString() || "",
      gender: gender || "",
      dateOfBirth: formattedDOB,
      department: department || "",
      contactNo: contactNo?.toString() || "",
      createdAt,
    };

    try {
      const res = await axios.post(API_URL, user);
      console.log(`✅ Registered: ${formattedName} (${formattedEmail})`);

      results.push({
        ...user,
        status: "Success",
        error: "",
      });
    } catch (error) {
      const errMsg = error.response?.data?.message || error.message;
      console.error(`❌ Failed: ${formattedName} (${formattedEmail}) -> ${errMsg}`);

      results.push({
        ...user,
        status: "Failed",
        error: errMsg,
      });
    }

    await sleep(500); // avoid server overload
  }

  // Save to Excel
  if (results.length > 0) {
    const newWB = xlsx.utils.book_new();
    const newWS = xlsx.utils.json_to_sheet(results);
    xlsx.utils.book_append_sheet(newWB, newWS, "Results");
    xlsx.writeFile(newWB, "submitted_users_with_errors.xlsx");
    console.log(`📂 Saved results to submitted_users_with_errors.xlsx`);
  } else {
    console.log("⚠️ No data to save.");
  }
}

registerUsers();
