import express from "express";
import cors from "cors";
import nodemailer from "nodemailer";
import XLSX from "xlsx";
import bodyParser from "body-parser";
import fs from "fs";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";
import { dirname } from "path";
import dotenv from "dotenv";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const PORT = 5000;

app.use(cors());
app.use(bodyParser.json());

const supervisorEmailMap = {
  "Jennifer Jurjens": "jjurjens@grayridge.com",
  "Rupesh Gautam": "rgautam@grayridge.com",
  "Darsan Sundesa": "dsundesa@grayridge.com",
  "Debie Cator": "dcator@grayridge.com",
};

app.post("/api/send", async (req, res) => {
  try {
    const { supervisorName, shift, checklist = {}, comments } = req.body;

    const checklistData = [
      { Field: "Supervisor Name", Value: supervisorName },
      { Field: "Shift", Value: shift },
      ...Object.entries(checklist).map(([key, value]) => ({
        Field: key
          .replace(/([A-Z])/g, " $1")
          .replace(/^./, (str) => str.toUpperCase()),
        Value: value ? "Yes" : "No",
      })),
      { Field: "Comments", Value: comments || "" },
    ];

    const worksheet = XLSX.utils.json_to_sheet(checklistData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Checklist");

    const tempDir = os.tmpdir();
    const filePath = path.join(tempDir, "checklist.xlsx");
    XLSX.writeFile(workbook, filePath);

    const supervisorEmail = supervisorEmailMap[supervisorName];
    if (!supervisorEmail) {
      throw new Error("Supervisor email not found.");
    }

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    });

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: supervisorEmail,
      cc: process.env.EMAIL_CC || "", // optional CC
      subject: "Power Equipment Safety Checklist",
      text: "Checklist submitted. See attached file.",
      attachments: [{ filename: "checklist.xlsx", path: filePath }],
    };

    await transporter.sendMail(mailOptions);
    fs.unlinkSync(filePath);

    res.status(200).send("Email sent successfully!");
  } catch (error) {
    console.error("Error:", error);
    res.status(500).send("Failed to send checklist.");
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
