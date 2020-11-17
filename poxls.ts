import * as Excel from 'exceljs';
import * as fs from 'fs-extra';
import Po from 'pofile';
import yargs from 'yargs';

function escapePoStr(str: string) {
  return str.replace(/(")/g, '\\$1')
}

function isPoFile(filePath: string) {
  return /\.po$/.test(filePath.toLocaleLowerCase().trim());
}

function isExcelFile(filePath: string) {
  return /\.xls(x|m|b|)$/.test(filePath.toLocaleLowerCase().trim());
}

class PoItem {
  constructor(protected msgId: string, protected msgStr: string) {}

  toString() {
    return `msgid "${escapePoStr(this.msgId)}"\nmsgstr "${escapePoStr(this.msgStr)}"`;
  }
}

async function xlsToPo(filePath: string, outputPath: string) {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filePath);

  const translations: PoItem[] = [];

  const sheet = workbook.worksheets[0];
  sheet.eachRow(row => {
    const source = row.getCell(1);
    const translation = row.getCell(2);

    translations.push(new PoItem(source.text, translation.text));
  });

  const output = translations.map(t => t.toString()).join('\n\n');

  return fs.writeFile(outputPath, output);
}

interface PoToXlsOptions {
  template?: boolean;
}

async function poToXls(filePath: string, outputPath: string, options: PoToXlsOptions = {}) {
  const poStr = await fs.readFile(filePath, 'utf8');
  const po = Po.parse(poStr);

  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet();

  sheet.addRows(po.items.map(item => options.template ? [item.msgid, ''] : [item.msgid, item.msgstr[0]]));

  let maxLen = 0;
  for (const column of sheet.columns) {
    column.eachCell(cell => maxLen = Math.max(cell.value?.toString().length ?? 0, maxLen));
  }

  for (const column of sheet.columns) {
    column.width = Math.max(maxLen, 10);
  }

  return workbook.xlsx.writeFile(outputPath);
}

async function main() {
  const [source, output] = process.argv.slice(2);

  if (!source || !output) {
    console.log('You need to specify two paths');
    return 1;
  }

  if (isPoFile(source) && isExcelFile(output)) {
    console.log(`Converting PO-file to XLSX...`);
    await poToXls(source, output);
    console.log('Done');
  } else if (isPoFile(output) && isExcelFile(source)) {
    console.log(`Converting XLSX to PO...`);
    await xlsToPo(source, output);
    console.log('Done');
  } else {
    console.log('unknown file extensions!');
    return 1;
  }
}

main();
