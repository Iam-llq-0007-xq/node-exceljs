const { Workbook } = require('exceljs');
const excelExample = async (req, res) => {
  console.log(' ~> /excel/template');
  const fileName = 'ok.xlsx';
  res.writeHead(200, {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Access_Control-Allow-Origin': '*',
    'Content-Disposition': `attachment; filename=${fileName}`,
  });

  try {
    const wb = await wbFactory();
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.log(e);
    res.end('error');
  }
};

async function wbFactory() {
  const templateDir = './servers/excel/assets/demo.xlsx';
  const wb = new Workbook();
  await wb.xlsx.readFile(templateDir);
  const ws = wb.getWorksheet('Sheet1'); // 'demo-V3.0.0'
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
        promptTitle: '选择下拉框值',
        prompt: '请选择下拉框的值',
      };
    }
  });
  
  return wb;
}

module.exports = excelExample;
