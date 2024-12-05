const express = require('express');
const xlsx = require('xlsx');
const app = express();
const PORT = 3000;

// Middleware to parse JSON
app.use(express.json());

// Dummy data generator for income/expense report
const generateDummyReport = () => {
  const data = [
    { date: '2024-11-28', moneyIn: 500, moneyOut: 100, remarks: 'Groceries' },
    { date: '2024-11-29', moneyIn: 200, moneyOut: 50, remarks: 'Coffee' },
    { date: '2024-11-30', moneyIn: 300, moneyOut: 200, remarks: 'Transport' },
    { date: '2024-12-01', moneyIn: 0, moneyOut: 100, remarks: 'Entertainment' },
    { date: '2024-12-02', moneyIn: 1000, moneyOut: 500, remarks: 'Salary' },
    { date: '2024-12-03', moneyIn: 200, moneyOut: 50, remarks: 'Miscellaneous' },
    { date: '2024-12-04', moneyIn: 400, moneyOut: 200, remarks: 'Shopping' },
  ];
  return data;
};

// Endpoint to generate and download Excel report
app.get('/report', (req, res) => {
  // Generate dummy data
  const reportData = generateDummyReport();

  // Create a workbook and add a worksheet
  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(reportData);

  // Add worksheet to workbook
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Weekly Report');

  // Write workbook to buffer
  const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

  // Set response headers
  res.setHeader('Content-Disposition', 'attachment; filename=income_expense_report.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  // Send buffer as response
  res.send(buffer);
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
