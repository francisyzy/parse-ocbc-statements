import PdfParse from 'pdf-parse';

import { readdirSync, readFileSync } from 'fs';
import excel, { Workbook } from 'excel4node';
import { addHours } from 'date-fns';

// const dir = './statements/frank';
const dir = './statements/consolidated';

const files = readdirSync(dir);

function analyse() {
  let dataBuffer = readFileSync(
    `${dir}/${'Consolidated Statement-Apr-21.pdf'}`,
  );
  PdfParse(dataBuffer).then((data) => {
    console.log(data.text);
  });
}
// analyse();

async function main() {
  let allStatement: {
    transactionDate: Date;
    valueDate: Date;
    description: string;
    withdrawal: number;
    deposit: number;
    balance: number;
  }[][] = [];
  let sheetName: string[] = [];
  for await (const file of files) {
    sheetName.push(file.substring(file.length - 10, file.length - 4));

    // }
    // files.forEach(async (file) => {
    let dataBuffer = readFileSync(`${dir}/${file}`);
    const statement = PdfParse(dataBuffer).then((data) => {
      let originalBalance = 0;
      let monthStatement = [];
      for (let i = 0; i < data.text.split('\n').length; i++) {
        const line = data.text.split('\n')[i];
        if (line.indexOf('BALANCE B/F') !== -1) {
          // console.log(file);
          originalBalance = extractNumbers(data.text.split('\n')[i - 1])[0];
        }
        // if (file.indexOf('Oct-21') !== -1) {
        //   // console.log(line);
        //   console.log();
        //   // console.log(i - 1 + ': ' + line.length);
        // }
        if (
          file.indexOf('Consolidated') !== -1 &&
          line.substring(0, 7).match(/\d\d \w\w\w\d/g) !== null
        ) {
          //prevent crash
          if (!numbersPresent(line) || !numbersPresent(line.substring(0, 1))) {
            continue;
          }
          const transactionDate = convertUTCDateToLocalDate(
            new Date(line.substring(0, 6) + ' 2021'),
          );
          let description = '';
          const remaining = line.substring(6, line.length);
          let deposit = 0;
          let withdrawal = 0;
          const balance = extractNumbers(remaining)[1];
          if (balance > originalBalance) {
            deposit = extractNumbers(remaining)[0];
          } else {
            withdrawal = extractNumbers(remaining)[0];
          }
          let j = i + 1;
          while (
            data.text
              .split('\n')
              [j].substring(0, 6)
              .match(/\d\d \w\w\w/g) === null
          ) {
            description += data.text.split('\n')[j] + '\n';
            j++;
          }
          const valueDate = convertUTCDateToLocalDate(
            new Date(data.text.split('\n')[j++] + ' 2021'),
          );
          originalBalance = balance;
          monthStatement.push({
            transactionDate,
            valueDate,
            description,
            withdrawal,
            deposit,
            balance,
          });
        }
        if (
          file.indexOf('FRANK') !== -1 &&
          line.substring(0, 7).match(/\d\d \w\w\w /g) !== null
        ) {
          //prevent crash
          if (!numbersPresent(line)) {
            continue;
          }
          const transactionDate = convertUTCDateToLocalDate(
            new Date(line.substring(0, 6) + ' 2021'),
          );
          // console.log(
          //   convertUTCDateToLocalDate(new Date(line.substring(0, 6) + ' 2021')),
          // );
          const valueDate = convertUTCDateToLocalDate(
            new Date(line.substring(line.length - 6) + ' 2021'),
          );
          let description = '';
          const remaining = line.substring(6, line.length - 6);
          let deposit = 0;
          let withdrawal = 0;
          const balance = extractNumbers(remaining)[1];
          if (balance > originalBalance) {
            deposit = extractNumbers(remaining)[0];
          } else {
            withdrawal = extractNumbers(remaining)[0];
          }
          description += data.text.split('\n')[i + 1] + '\n';
          description += data.text.split('\n')[i + 2] + '\n';
          description += data.text.split('\n')[i + 3] + '\n';
          description += data.text.split('\n')[i + 4] + '\n';
          description +=
            data.text.split('\n')[i + 5].length < 50
              ? data.text.split('\n')[i + 5]
              : '';
          originalBalance = balance;
          monthStatement.push({
            transactionDate,
            valueDate,
            description,
            withdrawal,
            deposit,
            balance,
          });
        }
        // console.log(i - 1 + ': ' + line.length);
      }
      // console.log(monthStatement);
      return monthStatement;
    });
    allStatement.push(await statement);
    // });
  }
  const wb = new excel.Workbook();
  let everything = [];
  for (let i = 0; i < allStatement.length; i++) {
    const statement = allStatement[i];
    everything.push(...statement);

    // generateSheetArray(wb, statement, i.toString());
  }
  generateSheetArray(wb, everything, 'frank');
  wb.write('statements.xlsx');
}
main();

function extractNumbers(text: string): number[] {
  let numbers;
  numbers = text.match(/(-\d+|\d+)(,\d+)*(\.\d+)*/g);
  numbers = numbers!.map((n) => Number(n.replace(/,/g, '')));
  return numbers;
}

function numbersPresent(text: string): boolean {
  let numbers = text.match(/(-\d+|\d+)(,\d+)*(\.\d+)*/g);
  return numbers !== null;
}

function generateSheetArray(wb: Workbook, object: any, sheetName: string) {
  const ws = wb.addWorksheet(sheetName);

  //Check if there's any value in the DB
  if (object[0] === undefined) {
    return wb;
  }

  let headingColumnNames: string[] = [];

  Object.keys(object[0]).forEach((columnName) => {
    headingColumnNames.push(columnName);
  });

  let headingColumnIndex = 1;
  headingColumnNames.forEach((heading) => {
    ws.cell(1, headingColumnIndex++).string(heading);
  });

  let rowIndex = 2;
  object.forEach((record: any) => {
    let columnIndex = 1;
    Object.keys(record).forEach((columnName) => {
      const cellContent = record[columnName];
      //Check if cell is null. If null skip cell
      if (typeof cellContent !== 'undefined') {
        //Check types to set cell to correct type
        if (typeof cellContent === 'string') {
          ws.cell(rowIndex, columnIndex++).string(cellContent.toString());
        } else if (typeof cellContent === 'number') {
          ws.cell(rowIndex, columnIndex++).number(cellContent);
        } else if (typeof cellContent === 'boolean') {
          ws.cell(rowIndex, columnIndex++).bool(cellContent);
        } else if (typeof cellContent === 'object') {
          if (cellContent instanceof Date) {
            ws.cell(rowIndex, columnIndex++).date(
              convertUTCDateToLocalDate(cellContent),
            );
          } else if (cellContent instanceof Array) {
            ws.cell(rowIndex, columnIndex++).string(cellContent.toString());
          } else {
            ws.cell(rowIndex, columnIndex++).string(
              JSON.stringify(cellContent),
            );
          }
        }
      } else {
        columnIndex++;
      }
    });
    rowIndex++;
  });
  return wb;
}

//https://stackoverflow.com/a/18330682
//Date.toString() returns accurate local time but is in string
//Set time from db to local time so excel can have local time
function convertUTCDateToLocalDate(date: Date) {
  var newDate = new Date(date.getTime() + date.getTimezoneOffset() * 60 * 1000);

  var offset = date.getTimezoneOffset() / 60;
  var hours = date.getHours();

  newDate.setHours(hours - offset);

  // return newDate;
  // return addDays(newDate, 1);
  return addHours(newDate, 10);
}
