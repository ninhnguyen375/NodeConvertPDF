const express = require('express')
const app = express()
const path = require('path')
const { execSync } = require("child_process");
const fs = require('fs');

const execWordPath = "C:/\"Program Files\"/LibreOffice/program/swriter.exe"
const execExcelPath = "C:/\"Program Files\"/LibreOffice/program/scalc.exe"
const tmpPath = path.join(__dirname, "tmp")
const convertedPath = path.join(__dirname, "converted")
let fileNumber = 0

app.use(express.json({ limit: '50mb' }))
app.use(express.urlencoded({ extended: false }))

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname + '/index.html'))
})

app.post('/api/convert-pdf', async (req, res) => {
  const base64ToConvert = req.body.base64
  const type = req.body.type
  const now = Date.now().toString()
  const filePath = path.join(__dirname, "tmp", now + fileNumber)

  if (!fs.existsSync(tmpPath)) {
    fs.mkdirSync(tmpPath);
  }
  if (!fs.existsSync(convertedPath)) {
    fs.mkdirSync(convertedPath);
  }

  if (base64ToConvert === undefined) {
    res.statusCode = 400
    console.log("no body data")
    res.send({ error: "No body data" })
    return
  }

  const buffer = Buffer.from(base64ToConvert, 'base64')
  const convertedItem = path.join(convertedPath, now + fileNumber + ".pdf")
  await fs.writeFileSync(filePath, buffer)

  let cmd = ""

  if (type === "word") {
    cmd = execWordPath + " --headless --convert-to pdf " + filePath + " --outdir " + convertedPath
  } else if (type === "excel") {
    cmd = execExcelPath + " --headless --convert-to pdf " + filePath + " --outdir " + convertedPath
  } else {
    console.log("missing type")
    res.send({ error: "missing type" })
    return
  }

  await execSync(cmd);

  let bytes
  let base64
  try {
    bytes = await fs.readFileSync(convertedItem)
    base64 = bytes.toString("base64")
  } catch (error) {
    await execSync("del /f " + filePath)
    res.send({ error })
    return
  }

  // delete converted files
  await execSync("del /f " + convertedItem)
  await execSync("del /f " + filePath)

  console.log("done")
  if (fileNumber > 100) {
    fileNumber = 0
  } else {
    fileNumber++
  }
  res.send({ base64 })
})

app.listen(3000, () => { })
