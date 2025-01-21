const express = require("express");
const sqlite3 = require("sqlite3").verbose();
const cors = require("cors");
const ExcelJS = require("exceljs");

const app = express();
const PORT = 5000;

// Middleware
app.use(cors());
app.use(express.json());

// SQLite Database Setup
const db = new sqlite3.Database("./activities.db", (err) => {
  if (err) console.error(err.message);
  else console.log("Connected to SQLite database.");
});

// Create Table
db.run(
  `CREATE TABLE IF NOT EXISTS activities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT,
    activity TEXT,
    status TEXT,
    completed_hours INTEGER DEFAULT 0
  )`
);

// Routes

// Get All Activities
app.get("/api/activities", (req, res) => {
  db.all("SELECT * FROM activities", [], (err, rows) => {
    if (err) return res.status(500).json(err.message);
    res.json(rows);
  });
});

// Add New Activity
app.post("/api/activities", (req, res) => {
  const { date, activity, status, completed_hours } = req.body;
  db.run(
    "INSERT INTO activities (date, activity, status, completed_hours) VALUES (?, ?, ?, ?)",
    [date, activity, status, completed_hours || 0],
    function (err) {
      if (err) return res.status(500).json(err.message);
      res.json({ id: this.lastID, date, activity, status, completed_hours });
    }
  );
});

// Update Activity
app.put("/api/activities/:id", (req, res) => {
  const { id } = req.params;
  const { date, activity, status, completed_hours } = req.body;
  db.run(
    "UPDATE activities SET date = ?, activity = ?, status = ?, completed_hours = ? WHERE id = ?",
    [date, activity, status, completed_hours, id],
    function (err) {
      if (err) return res.status(500).json(err.message);
      res.json({ id, date, activity, status, completed_hours });
    }
  );
});

// Delete Single Activity
app.delete("/api/activities/:id", (req, res) => {
  const { id } = req.params;
  db.run("DELETE FROM activities WHERE id = ?", [id], function (err) {
    if (err) return res.status(500).json(err.message);
    res.json({ message: "Activity deleted successfully." });
  });
});

// Delete All Activities
app.delete("/api/activities", (req, res) => {
  db.run("DELETE FROM activities", function (err) {
    if (err) return res.status(500).json(err.message);

    // Reset auto-increment ID
    db.run("UPDATE sqlite_sequence SET seq = 0 WHERE name = 'activities'", (seqErr) => {
      if (seqErr) return res.status(500).json(seqErr.message);
      res.json({ message: "All activities deleted, and IDs reset." });
    });
  });
});

// Download Excel Report
app.get("/api/export", async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Activities");

  worksheet.columns = [
    { header: "ID", key: "id", width: 10 },
    { header: "Date", key: "date", width: 20 },
    { header: "Activity", key: "activity", width: 30 },
    { header: "Status", key: "status", width: 20 },
    { header: "Completed Hours", key: "completed_hours", width: 20 },
  ];

  db.all("SELECT * FROM activities", [], async (err, rows) => {
    if (err) return res.status(500).json(err.message);

    rows.forEach((row) => worksheet.addRow(row));
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=Activity_Report.xlsx"
    );
    await workbook.xlsx.write(res);
    res.end();
  });
});

// Start Server
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
