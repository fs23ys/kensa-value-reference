const XLSX = require('xlsx');
const { LAB_COLUMNS, DISEASE_COLUMNS, labRows, diseaseRows } = require('./initialData');

function rowsToAOA(columns, rows) {
  const aoa = [columns];
  rows.forEach(r => {
    aoa.push(columns.map(c => (r[c] !== undefined ? r[c] : '')));
  });
  return aoa;
}

const wb = XLSX.utils.book_new();

const labAOA = rowsToAOA(LAB_COLUMNS, labRows);
const labWS = XLSX.utils.aoa_to_sheet(labAOA);
labWS['!cols'] = LAB_COLUMNS.map(c => {
  if (c === '患者説明例') return { wch: 45 };
  if (c === '薬剤師補足') return { wch: 35 };
  if (c === '高値の時のアドバイス' || c === '低値の時のアドバイス') return { wch: 32 };
  if (c === '参考ページ') return { wch: 14 };
  return { wch: 18 };
});
XLSX.utils.book_append_sheet(wb, labWS, '検査値');

const diseaseAOA = rowsToAOA(DISEASE_COLUMNS, diseaseRows);
const diseaseWS = XLSX.utils.aoa_to_sheet(diseaseAOA);
diseaseWS['!cols'] = DISEASE_COLUMNS.map(c => {
  if (c === '聞くこと' || c === '声のかけ方' || c === '見落としやすい点') return { wch: 45 };
  if (c === '参考ページ') return { wch: 14 };
  return { wch: 18 };
});
XLSX.utils.book_append_sheet(wb, diseaseWS, '疾患');

const outPath = process.argv[2];
XLSX.writeFile(wb, outPath);
console.log('written:', outPath, 'labRows:', labRows.length, 'diseaseRows:', diseaseRows.length);
