const express = require('express');
const app = new express();
require('dotenv').config();
const PORT = process.env.PORT || 8001;

app.get('/token', (req, res) => {
  res.send({
    error: 0,
    data: 'ok',
  });
});

const excel = require('./servers/excel/index');
app.use(excel());

app.listen(PORT, () => {
  console.log('listening port: ' + PORT);
});
