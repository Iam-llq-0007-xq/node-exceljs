const { Workbook } = require('exceljs');

async function excel() {
  const wb = new Workbook();
  await wb.xlsx.readFile('./static/渠道风控上传模板.xlsx');
  const ws = wb.getWorksheet('Sheet1'); // '渠道风控上传模板-V3.0.0'
  ws.views = [{ state: 'frozen', ySplit: 2 }];
  const colChannel = ws.getColumn(7);
  colChannel.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    if (rowNumber > 5) {
      cell.dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"One,Two,Three,Four"'],
        showErrorMessage: true,
        showInputMessage: true,
        promptTitle: '所属渠道公司',
        prompt: '请选择下拉框的值',
      };
    }
  });
  await wb.xlsx.writeFile('./生成.xlsx');
}

excel();
