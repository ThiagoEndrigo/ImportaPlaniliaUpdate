const XLSX = require('xlsx');
const fs = require('fs');

const inputFile = 'BRIGADERIA BELGA_LUCRO REAL OU PRESUMIDO_BA_EMPRESA 1.xlsx';
const outputFile = 'UPDATE_PRODU.txt';
const setFields = [
  'ALIQUOTA', 'CST', 'REDUCAO_ICMS', 'NCM', 'CODNATOPERACAOSAI', 'CSTPISCOFINSSAI',
  'CODCST_IS', 'CODCST_CLASSTRIB_IS', 'ALIQ_IS', 'REDUCAO_ALIQ_IS',
  'CODCST_CLASSTRIB_IBSCBS', 'CODCST_IBSCBS', 'ALIQ_IBS_UF', 'REDUCAO_ALIQ_IBS_UF',
  'ALIQ_EFETIVA_IBS_UF', 'ALIQ_IBS_MUN', 'REDUCAO_ALIQ_IBS_MUN', 'ALIQ_EFETIVA_IBS_MUN',
  'ALIQ_CBS', 'REDUCAO_ALIQ_CBS', 'ALIQ_EFETIVA_CBS'
];
const whereFields = ['CODEMP', 'CAT_COD', 'PROD_COD'];

const workbook = XLSX.readFile(inputFile);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
const headers = data[0].map(value => String(value || '').trim());
const indexOf = field => headers.findIndex(header => header.toUpperCase() === field);
const valueOf = (row, field) => String(row[indexOf(field)] ?? '').trim();
const quote = value => `'${String(value).replace(/'/g, "''")}'`;
const formatCst = value => ({ '0': '000', '00': '000', '20': '020', '40': '040', '41': '041', '60': '060' }[value] || value);

const updates = data.slice(1)
  .filter(row => row.some(value => String(value).trim() !== ''))
  .map(row => {
    const set = setFields
      .filter(field => indexOf(field) !== -1 && valueOf(row, field) !== '')
      .map(field => `${field} = ${quote(field === 'CST' ? formatCst(valueOf(row, field)) : valueOf(row, field))}`);
    const where = whereFields
      .filter(field => indexOf(field) !== -1 && valueOf(row, field) !== '')
      .map(field => `${field} = ${quote(valueOf(row, field))}`);
    return set.length && where.length ? `UPDATE PRODU SET ${set.join(', ')} WHERE ${where.join(' AND ')};` : null;
  })
  .filter(Boolean);

fs.writeFileSync(outputFile, updates.join('\r\n\r\n'), 'utf8');
console.log(`${updates.length} UPDATEs gerados em ${outputFile}`);
