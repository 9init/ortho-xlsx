const fs = require('fs');
const XLSX = require('xlsx');

async function processXlsx(filePath, imgsPath) {
  try {
    // Read the existing workbook
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Assuming one sheet
    const worksheet = workbook.Sheets[sheetName];

    // Loop through rows starting from the second row (skipping headers)
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Read with headers
    const filteredData = data.filter(row => {
      const imageName = row[0]; // Assuming image name is in first column (index 0)
      const imagePath = `${imgsPath}/${imageName}`;
      return fs.existsSync(imagePath); // Check if image exists
    });

    // Write the filtered data to a new workbook
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(filteredData); // Write with headers
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);

    // Save the modified workbook
    XLSX.writeFile(newWorkbook, newFilePath(filePath)); // Generate new file name
    console.log(`Processed file: ${filePath}`);
  } catch (error) {
    console.error(`Error processing file: ${filePath}`, error);
  }
}

function newFilePath(filePath) {
  // Generate a new file name to avoid overwriting the original
  const parts = filePath.split('.');
  parts[parts.length - 1] = 'processed.xlsx';
  return parts.join('.');
}

// Replace with your actual file paths
const xlsxFilePath = 'combined.xlsx';
const imgsPath = '/mnt/d/Graduation/dataset_before/combined/imgs';

processXlsx(xlsxFilePath, imgsPath);