import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import multer from "multer";
import * as XLSX from "xlsx";

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Logger
app.use((req, res, next) => {
  console.log(`[Server] ${req.method} ${req.url}`);
  next();
});

// File upload setup
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }
});

// API Router
const apiRouter = express.Router();

// Ping
apiRouter.get("/ping", (req, res) => {
  res.json({ message: "pong", time: new Date() });
});

// Health
apiRouter.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

// Upload students
apiRouter.post("/upload-students", upload.single("file"), (req: any, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    res.json({ students: data });
  } catch {
    res.status(500).json({ error: "Failed to process file" });
  }
});

// Upload timetable
apiRouter.post("/upload-timetable", upload.single("file"), (req: any, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    res.json({ timetable: data });
  } catch {
    res.status(500).json({ error: "Failed to process file" });
  }
});

// Upload subjects
apiRouter.post("/upload-subjects", upload.single("file"), (req: any, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    res.json({ subjects: data });
  } catch {
    res.status(500).json({ error: "Failed to process file" });
  }
});

// Mount API
app.use("/api", apiRouter);

// Start server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});