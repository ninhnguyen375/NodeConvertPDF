const express = require('express');
const app = express();
const libre = require('libreoffice-convert');
const path = require('path');

app.use(express.json({ limit: '50mb' }))
app.use(express.urlencoded({ extended: false }))

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname + '/index.html'))
})

app.post('/api/convert-pdf', async (req, res) => {
  const b64string = req.body.base64

  if(b64string === undefined) {
    res.statusCode = 400;
    res.send({ error: "No body data" });
    return;
  }

  const buffer = Buffer.from(b64string, 'base64')

  libre.convert(buffer, '.pdf', undefined, (err, done) => {
    if (err) {
      res.send({
        error: err
      })
    } else {
      console.log("done");
      res.send({
        base64: done.toString('base64')
      })
    }
  });
});

app.listen(3000, () => {
})
