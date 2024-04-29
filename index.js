const XLSX = require("xlsx");
const fs = require('fs');

// Array of XLSX file paths
const xlsxFiles = [
  "/mnt/d/Graduation/dataset_before/combined/abdo_win32_sorted.xlsx",
  "/mnt/d/Graduation/dataset_before/combined/ahmed_darwin_sorted.xlsx",
  "/mnt/d/Graduation/dataset_before/combined/ammar_win32_sorted.xlsx",
  "/mnt/d/Graduation/dataset_before/combined/aya_darwin_sorted.xlsx",
  "/mnt/d/Graduation/dataset_before/combined/fares_darwin_sorted.xlsx",
  "/mnt/d/Graduation/dataset_before/combined/hager_win32_sorted.xlsx",
];

// Array to store the combined data
let combinedData = [];

// Iterate over each XLSX file
xlsxFiles.forEach((filePath) => {
  // Read the workbook
  const workbook = XLSX.readFile(filePath);

  // Get the first sheet name
  const sheetName = workbook.SheetNames[0];

  // Get the worksheet
  const worksheet = workbook.Sheets[sheetName];

  // Convert the worksheet to JSON, skipping the first row (headers)
  const jsonData = XLSX.utils.sheet_to_json(worksheet, {
    header: "A",
    range: 1,
  });

  // Append the data to the combinedData array
  combinedData = combinedData.concat(jsonData);
});

// Loop through rows starting from the second row (skipping headers)
// to filter out rows with missing images
const filteredData = combinedData.filter((row) => {
    console.log(row);
  const imageName = row["A"]; // Assuming image name is in first column
  const imagePath = `/mnt/d/Graduation/dataset_before/combined/imgs/${imageName}`;
  return fs.existsSync(imagePath); // Check if image exists
});

// Create a new workbook
const newWorkbook = XLSX.utils.book_new();
// Create a new worksheet
const newWorksheet = XLSX.utils.json_to_sheet(filteredData);
// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Combined Sheet");
// Define the output file path
const outputFile = "./combined.xlsx";
// Write the workbook to the output file
XLSX.writeFile(newWorkbook, outputFile);
console.log("Combined XLSX file created:", outputFile);
