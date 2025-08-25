// import express from "express";
// import nodemailer from "nodemailer";
// import xlsx from "xlsx";

// const app = express();
// app.use(express.json());

// // Nodemailer setup
// const transporter = nodemailer.createTransport({
//   host: "smtp-relay.brevo.com",
//   port: 587,
//   secure: false,
//   auth: {
//     user: "95859d001@smtp-brevo.com", 
//     pass: "5HxCqN2ZgLtv0RVd" 
//   }
// });

// // Function to read donor data from Excel
// function getDonorsFromExcel(filePath) {
//   const workbook = xlsx.readFile(filePath);
//   const sheetName = workbook.SheetNames[0];
//   const sheet = workbook.Sheets[sheetName];
//   const data = xlsx.utils.sheet_to_json(sheet);
//   return data; // array of objects with Name, Email, Last_Donation_Date, Next_Donation_Date
// }

// app.post("/send", async (req, res) => {
//   try {
//     // Load donors from Excel file
//     const donors = getDonorsFromExcel("./donors_with_dates.xlsx");

//     if (!donors.length) {
//       return res.status(400).json({ success: false, error: "No donor records found in Excel" });
//     }

//     // Loop over donors and send personalized email
//     for (let donor of donors) {
//       const { Name, Email, Last_Donation_Date, Next_Donation_Date } = donor;

//       const message = `
//         Dear ${Name},

//         Thank you for your generous support! 
//         We noticed your last donation was on ${Last_Donation_Date}. 
//         Your next donation date is scheduled for ${Next_Donation_Date}.

//         We truly appreciate your continued contributions to our NGO. 💙
        
//         Best regards,
//         Mayank
//       `;

//       await transporter.sendMail({
//         from: '"NGO Team" <mayankvatsa2@gmail.com>', 
//         to: Email,
//         subject: "Thank you for your support ❤️",
//         text: message
//       });

//       console.log(`Email sent to ${Name} (${Email})`);
//     }

//     res.json({ success: true, sent: donors.length });
//   } catch (err) {
//     console.error(err);
//     res.status(500).json({ success: false, error: err.message });
//   }
// });

// app.listen(3000, () => console.log("Server running on port 3000"));
// import express from "express";
// import nodemailer from "nodemailer";
// import xlsx from "xlsx";

// const app = express();
// app.use(express.json());

// // Nodemailer setup
// const transporter = nodemailer.createTransport({
//   host: "smtp-relay.brevo.com",
//   port: 587,
//   secure: false,
//   auth: {
//     user: "95859d001@smtp-brevo.com",
//     pass: "5HxCqN2ZgLtv0RVd"
//   }
// });

// // Function to read donor data from Excel
// function getDonorsFromExcel(filePath) {
//   const workbook = xlsx.readFile(filePath);
//   const sheetName = workbook.SheetNames[0];
//   const sheet = workbook.Sheets[sheetName];
//   return xlsx.utils.sheet_to_json(sheet);
// }

// // ✅ Improved date parser: handles Excel numbers, Date objects, and both string formats
// function parseDate(dateValue) {
//   if (!dateValue) return null;

//   // If Excel gave us a number (serial date)
//   if (typeof dateValue === "number") {
//     const excelEpoch = new Date(1900, 0, 1);
//     return new Date(excelEpoch.getTime() + (dateValue - 2) * 24 * 60 * 60 * 1000);
//   }

//   // If it's already a Date object
//   if (dateValue instanceof Date) {
//     return dateValue;
//   }

//   // Force to string
//   const str = String(dateValue).trim();

//   if (str.includes("-")) {
//     const parts = str.split("-");
//     if (parts[0].length === 4) {
//       // YYYY-MM-DD
//       return new Date(str);
//     } else {
//       // DD-MM-YYYY
//       return new Date(parts[2], parts[1] - 1, parts[0]);
//     }
//   }

//   return null;
// }

// // ✅ Format Date to DD/MM/YYYY
// function formatDate(dateValue) {
//   if (!dateValue) return "";
//   const date = parseDate(dateValue);
//   if (!date) return "";
//   return date.toLocaleDateString("en-GB"); // DD/MM/YYYY
// }

// app.post("/send", async (req, res) => {
//   try {
//     const donors = getDonorsFromExcel("./donors_with_dates.xlsx");

//     if (!donors.length) {
//       return res.status(400).json({ success: false, error: "No donor records found in Excel" });
//     }

//     const today = new Date();
//     today.setHours(0, 0, 0, 0); // normalize (ignore time)

//     let sentCount = 0;
//     let skipped = [];

//     for (let donor of donors) {
//       const { Name, Email, Last_Donation_Date, Next_Donation_Date } = donor;

//       const nextDate = parseDate(Next_Donation_Date);
//       if (!nextDate) {
//         skipped.push({ Name, Email, reason: "Invalid next donation date" });
//         continue;
//       }

//       nextDate.setHours(0, 0, 0, 0); // normalize

//       // ✅ Send only if Next Donation Date is today or in the past
//       if (nextDate <= today) {
//         const message = `
// Dear ${Name},

// Thank you for your generous support!
// We noticed your last donation was on ${formatDate(Last_Donation_Date)}.
// Your next donation date was scheduled for ${formatDate(Next_Donation_Date)}.

// You are now eligible to donate again 💙

// Best regards,
// Mayank
//         `;

//         await transporter.sendMail({
//           from: '"NGO Team" <mayankvatsa2@gmail.com>',
//           to: Email,
//           subject: "You're eligible to donate again ❤️",
//           text: message
//         });

//         console.log(`✅ Email sent to ${Name} (${Email})`);
//         sentCount++;
//       } else {
//         console.log(`⏭️ Skipped ${Name} (${Email}), next donation date is in the future.`);
//         skipped.push({ Name, Email, reason: "Next donation date in future" });
//       }
//     }

//     res.json({ success: true, sent: sentCount, skipped });
//   } catch (err) {
//     console.error(err);
//     res.status(500).json({ success: false, error: err.message });
//   }
// });

// app.listen(3000, () => console.log("🚀 Server running on port 3000"));



const nodemailer = require("nodemailer");
import xlsx =require("xlsx");

// ---------------- Nodemailer setup ----------------
const transporter = nodemailer.createTransport({
  host: "smtp-relay.brevo.com",
  port: 587,
  secure: false,
  auth: {
    user: process.env.SMTP_USER || "95859d001@smtp-brevo.com",
    pass: process.env.SMTP_PASS || "5HxCqN2ZgLtv0RVd"
  }
});

// ---------------- Read donor Excel ----------------
function getDonorsFromExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet);
}

// ---------------- Date parser ----------------
function parseDate(dateValue) {
  if (!dateValue) return null;

  if (typeof dateValue === "number") {
    const excelEpoch = new Date(1900, 0, 1);
    return new Date(excelEpoch.getTime() + (dateValue - 2) * 24 * 60 * 60 * 1000);
  }

  if (dateValue instanceof Date) return dateValue;

  const str = String(dateValue).trim();
  if (str.includes("-")) {
    const parts = str.split("-");
    if (parts[0].length === 4) {
      return new Date(str); // YYYY-MM-DD
    } else {
      return new Date(parts[2], parts[1] - 1, parts[0]); // DD-MM-YYYY
    }
  }

  return null;
}

// ---------------- Format date as DD/MM/YYYY ----------------
function formatDate(dateValue) {
  const date = parseDate(dateValue);
  if (!date) return "";
  return date.toLocaleDateString("en-GB"); // DD/MM/YYYY
}

// ---------------- Main function ----------------
async function sendEmails() {
  try {
    const donors = getDonorsFromExcel("./donors_with_dates.xlsx");

    if (!donors.length) {
      console.log("❌ No donor records found in Excel");
      return;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0); // normalize

    let sentCount = 0;
    let skipped = [];

    for (let donor of donors) {
      const { Name, Email, Last_Donation_Date, Next_Donation_Date } = donor;

      const nextDate = parseDate(Next_Donation_Date);
      if (!nextDate) {
        skipped.push({ Name, Email, reason: "Invalid next donation date" });
        continue;
      }

      nextDate.setHours(0, 0, 0, 0);

      if (nextDate <= today) {
        const message = `
Dear ${Name},

Thank you for your generous support!
We noticed your last donation was on ${formatDate(Last_Donation_Date)}.
Your next donation date was scheduled for ${formatDate(Next_Donation_Date)}.

You are now eligible to donate again 💙

Best regards,
Mayank
        `;

        await transporter.sendMail({
          from: '"NGO Team" <mayankvatsa2@gmail.com>',
          to: Email,
          subject: "You're eligible to donate again ❤️",
          text: message
        });

        console.log(`✅ Email sent to ${Name} (${Email})`);
        sentCount++;
      } else {
        console.log(`⏭️ Skipped ${Name} (${Email}), next donation date in future`);
        skipped.push({ Name, Email, reason: "Next donation date in future" });
      }
    }

    console.log(`\nSummary: Sent: ${sentCount}, Skipped: ${skipped.length}`);
    if (skipped.length > 0) console.log("Skipped details:", skipped);

  } catch (err) {
    console.error("Error sending emails:", err.message);
  }
}

// ---------------- Run script ----------------
sendEmails();
