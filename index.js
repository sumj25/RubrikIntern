const fs = require("fs");
const XLSX = require("xlsx");

// Create a new Excel workbook
const workbook = XLSX.utils.book_new();
// Specify the output Excel file name
const excelFileName = "output.xlsx";

// Function to create a hyperlink formula for Excel
function createHyperlinkFormula(sheetName) {
  return `=HYPERLINK("#'${sheetName}'!A1", "SHEET::${sheetName}")`;
}

// Function for processing arrays in JSON
function processArray(jsonData, sheetName) {
  let Data = [];
  let cellFormula = "";
  Data[0] = [];

  try {
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i] === null) {
        // If the array element is null, append a space
        cellFormula += " ";
      } else if (Array.isArray(jsonData[i])) {
        // If the array element is an array, create a nested sheet and process it
        const nestedSheetName = `${sheetName}.${i}`;
        cellFormula += createHyperlinkFormula(nestedSheetName);
        processArray(jsonData[i], nestedSheetName);
      } else if (typeof jsonData[i] === "object") {
        // If the array element is an object, create a nested sheet and process it
        const nestedSheetName = `${sheetName}.${i}`;
        cellFormula += createHyperlinkFormula(nestedSheetName);
        ProcessObjects(jsonData[i], nestedSheetName);
      } else {
        // If the array element is a primitive, append it to the cell formula
        cellFormula += jsonData[i];
      }

      cellFormula += ", ";
    }

    // Push the cell formula to the data array
    Data[0].push(cellFormula);
    // Append the data array as a new sheet in the workbook
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(Data),
      sheetName
    );

    // Write the workbook to the Excel file
    XLSX.writeFile(workbook, excelFileName);
  } catch (error) {
    console.error(
      `Error processing array in sheet '${sheetName}':`,
      error.message
    );
  }
}

// Function for processing objects in JSON
function ProcessObjects(jsonData, sheetName) {
  let Data = [];
  Data[0] = Object.keys(jsonData);

  try {
    if (typeof jsonData === "object") {
      Data[1] = [];
      for (const key in jsonData) {
        if (
          typeof jsonData[key] === "object" &&
          jsonData[key] !== null &&
          !Array.isArray(jsonData[key])
        ) {
          // If the object property is an object, create a nested sheet and process it
          const nestedSheetName = `${sheetName}.${key}`;
          ProcessObjects(jsonData[key], nestedSheetName);
          Data[1].push(createHyperlinkFormula(nestedSheetName));
        } else if (Array.isArray(jsonData[key])) {
          // If the object property is an array, create a nested sheet and process it
          let cellFormula = "";
          for (let i = 0; i < jsonData[key].length; i++) {
            if (jsonData[key][i] === null) {
              // If the array element is null, append a space
              cellFormula += " ";
            } else if (Array.isArray(jsonData[key][i])) {
              // If the array element is an array, create a nested sheet and process it
              const nestedSheetName = `${sheetName}.${key}.${i}`;
              cellFormula += createHyperlinkFormula(nestedSheetName);
              processArray(jsonData[key][i], nestedSheetName);
            } else if (typeof jsonData[key][i] === "object") {
              // If the array element is an object, create a nested sheet and process it
              const nestedSheetName = `${sheetName}.${key}.${i}`;
              cellFormula += createHyperlinkFormula(nestedSheetName);
              ProcessObjects(jsonData[key][i], nestedSheetName);
            } else {
              // If the array element is a primitive, append it to the cell formula
              cellFormula += jsonData[key][i];
            }

            cellFormula += ", ";
          }

          // Push the cell formula to the data array
          Data[1].push(cellFormula);
        } else {
          // If the object property is a primitive, append it to the data array
          Data[1].push(jsonData[key]);
        }
      }
    } else if (Array.isArray(jsonData)) {
      // If the JSON data is an array, process it
      processArray(jsonData, sheetName);
    } else {
      // If the JSON data is empty, log a message
      console.log("JsonData is empty");
      return;
    }

    // Append the data array as a new sheet in the workbook
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(Data),
      sheetName
    );

    // Write the workbook to the Excel file
    XLSX.writeFile(workbook, excelFileName);
  } catch (error) {
    console.error(
      `Error processing objects in sheet '${sheetName}':`,
      error.message
    );
  }
}

// Function to read and process JSON from a file
function readAndProcessJSON(jsonPath) {
  try {
    // Read the JSON data from the file synchronously
    const JsonData = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
    // Function to Process JsonObjects
    ProcessObjects(JsonData, "Sheet1");
    console.log("Conversion successful. Excel file created:", excelFileName);
  } catch (error) {
    console.error("Error processing JSON:", error.message);
  }
}

// Path of the File
const JsonPath = "input.json";

// Read and process the JSON data
readAndProcessJSON(JsonPath);
