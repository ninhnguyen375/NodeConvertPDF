const express = require('express')
const app = express()
const path = require('path')
const { execSync } = require("child_process");
const fs = require('fs');
const execWordPath = "C:/\"Program Files\"/LibreOffice/program/swriter.exe"
const execExcelPath = "C:/\"Program Files\"/LibreOffice/program/scalc.exe"

app.use(express.json({ limit: '50mb' }))
app.use(express.urlencoded({ extended: false }))

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname + '/index.html'))
})

app.post('/api/convert-pdf', async (req, res) => {
  const base64ToConvert = req.body.base64
  const type = req.body.type

  if (base64ToConvert === undefined) {
    res.statusCode = 400
    res.send({ error: "No body data" })
    return
  }

  const buffer = Buffer.from(base64ToConvert, 'base64')
  const now = Date.now().toString()
  const filePath = path.join(__dirname, "tmp", now) 
  const convertedPath = path.join(__dirname, "converted")
  const convertedItem = path.join(convertedPath, now + ".pdf")
  fs.writeFileSync(filePath, buffer)

  let cmd = ""

  if (type === "word") {
    cmd = execWordPath + " --headless --convert-to pdf " + filePath + " --outdir " + convertedPath
  } else if (type === "excel") {
    cmd = execExcelPath + " --headless --convert-to pdf " + filePath + " --outdir " + convertedPath
  } else {
    res.send({ error: "missing type" })
    return
  }

  execSync(cmd);
  
  const bytes = fs.readFileSync(convertedItem)
  const base64 = bytes.toString("base64")

  // delete converted files
  execSync("del /f " + convertedItem)
  execSync("del /f " + filePath)

  res.send(base64)
})

app.listen(3000, () => { })
