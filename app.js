const express = require('express')
const multer = require('multer')
const xlsx = require('xlsx')
const fs = require('fs')
const cors = require('cors')

const app = express()
const PORT = 3003

app.use(express.json())
app.use(cors())

// Define the Excel file path
const filePath = './data.xlsx'

// Function to load or create an Excel file
const loadWorkbook = () => {
  if (fs.existsSync(filePath)) {
    return xlsx.readFile(filePath)
  } else {
    const wb = xlsx.utils.book_new()
    const ws = xlsx.utils.json_to_sheet([])
    xlsx.utils.book_append_sheet(wb, ws, 'Data')
    xlsx.writeFile(wb, filePath)
    return wb
  }
}

// Endpoint to add data to Excel
app.post('/add', (req, res) => {
  const { name, email, phoneNumber, wilaya, result } = req.body

  if (!name || !email || !phoneNumber || !wilaya) {
    return res.status(400).json({ error: 'All fields are required' })
  }

  const wb = loadWorkbook()
  const ws = wb.Sheets['Data']
  const data = xlsx.utils.sheet_to_json(ws) || []

  data.push({ Name: name, Email: email, Phone: phoneNumber, Wilaya: wilaya })

  const newWs = xlsx.utils.json_to_sheet(data)
  wb.Sheets['Data'] = newWs

  xlsx.writeFile(wb, filePath)

  res.json({ message: 'Data added successfully' })
})

// Endpoint to download Excel file
app.get('/download', (req, res) => {
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'No data available' })
  }

  res.download(filePath, 'data.xlsx')
})

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`)
})
