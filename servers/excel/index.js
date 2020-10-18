const express = require('express');

module.exports = () => {
  const app = express.Router();
  app.post('/excel/template', require('./excel-template'));
  return app;
};
