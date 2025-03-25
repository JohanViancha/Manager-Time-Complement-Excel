/* global Excel console */
import dayjs from "dayjs";
import "dayjs/locale/es";
dayjs.locale("es");

export async function insertText(text) {
  // Write text to the top left cell.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function getDataFromBook() {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("C7:CQ30");
      range.load("values");
      await context.sync();
      return range.values;
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function getPaper(date) {
  try {
    return await Excel.run(async (context) => {
      const dateFormated = dayjs(date).format("MMMM YYYY");
      const namePaper = dateFormated.charAt(0).toUpperCase() + dateFormated.slice(1);
      const sheet = context.workbook.worksheets.getItem(namePaper);
      return sheet;
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function calculateForActivityandCompany(formData) {
  try {
    return await Excel.run(async (context) => {
      const sheet =  context.workbook.worksheets.getItem(formData.book);
      const range = sheet.getRange("C8:CQ30");
      range.load("values");
      await context.sync();
      let account = 0;
      range.values.forEach((row) => {
        const accountActivity = row.reduce((acc, current, index) => {
          console.log(formData.company === row[index - 1]);
          if (formData.activity === current && formData.project === row[index - 1]) {
            acc = acc + 1;
          }
          return acc;
        }, 0);

        account = account + accountActivity;
      });

      const rangeProject = sheet.getRange("F33");
      const rangeActivity = sheet.getRange("G33");
      const rangeTotal= sheet.getRange("H33");
      const total = (account * 30) / 60;
      rangeProject.values = [[formData.project]];
      rangeActivity.values = [[formData.activity]];
      rangeTotal.values = [[`${total} horas`]];  
      rangeProject.format.autofitColumns();
      rangeActivity.format.autofitColumns();
      rangeTotal.format.autofitColumns();

      return total; 
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function registerActivity(company, activity, date, timeStart, time) {
  try {
    return await Excel.run(async (context) => {
      const dateFormated = dayjs(date).format("MMMM YYYY");
      const namePaper = dateFormated.charAt(0).toUpperCase() + dateFormated.slice(1);
      const sheet = context.workbook.worksheets.getItem(namePaper);
      const range = sheet.getRange("C33");
      range.values = [["Texto de prueba"]];

      const cell = await findCellByText(namePaper, date, "7:00:00 a. m.");
      console.log("Celda", cell);
      range.format.autofitColumns();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function findCellByText(sheetName, rowText, colText) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, address");

    await context.sync();

    let rowIndex = -1;
    let colIndex = -1;

    // 🔹 Buscar el texto en la columna (por ejemplo, en la primera columna)
    for (let row = 0; row < usedRange.values.length; row++) {
      if (usedRange.values[row][0] === rowText) {
        rowIndex = row;
        break;
      }
    }

    // 🔹 Buscar el texto en la fila (por ejemplo, en la primera fila)
    for (let col = 0; col < usedRange.values[4].length; col++) {
      if (usedRange.values[0][col] === colText) {
        colIndex = col;
        break;
      }
    }

    // Validar si encontramos ambos índices
    if (rowIndex !== -1 && colIndex !== -1) {
      const cell = usedRange.getCell(rowIndex, colIndex);
      cell.load("address");
      await context.sync();
      return cell.address; // Retorna la dirección de la celda encontrada
    } else {
      return "No encontrado";
    }
  });
}

function convertDateToExcelNumber(value) {
  console.log(value);
  if (value instanceof Date) {
    const excelSerialNumber = Math.floor((value - new Date(1899, 11, 30)) / (1000 * 60 * 60 * 24));
    return excelSerialNumber;
  } else {
    return "No es una fecha válida.";
  }
  return undefined;
}
