// const XLSX = require('xlsx');
// const path = require('path');

// // List of XLS files
// const xlsFiles = [
//     'AMMASANDRA_110_15_10_2023.xls',
//     'AHERI_220_14_11_2023.xls',
//     'CHADCHAN_110_14_11_2023.xls',
//     'DAVALESHWARA_110_14_11_2023.xls',
//     'HIREMURAL_110_14_11_2023.xls',
//     'SATTI_110_14_11_2023.xls'
// ];

// // Function to process a single file and return data
// const processFile = (xlsFilePath) => {
//     const workbook = XLSX.readFile(xlsFilePath);
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = workbook.Sheets[sheetName];
//     const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
//     const stationName = worksheet['A1'].v;

//     const initialDateCell = worksheet['B6'];
//     let lastDate = '';
//     if (initialDateCell && typeof initialDateCell.v === 'number') {
//         const dateObj = XLSX.SSF.parse_date_code(initialDateCell.v);
//         lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
//         console.log(`Initial date from B6 parsed as ${lastDate}`);
//     } else {
//         console.error('Initial date in cell B6 is missing or not a number.');
//     }

//     const limitedData = jsonData.slice(4000, 8001);
//     const data = [];
//     const foundFeeders = new Set();
//     const allFeeders = new Set();

//     for (let rowIndex = 0; rowIndex < limitedData.length; rowIndex++) {
//         const row = limitedData[rowIndex];
//         const dateValue = row[1];
//         if (typeof dateValue === 'number') {
//             const dateObj = XLSX.SSF.parse_date_code(dateValue);
//             lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
//             console.log(`Row ${rowIndex + 4000}: Date parsed as ${lastDate}`);
//         }

//         row.forEach((cell, cellIndex) => {
//             if (typeof cell === 'string' && cell.startsWith('F')) {
//                 allFeeders.add(cell.split('-')[0]);
//                 if (lastDate === '') {
//                     console.error(`Row ${rowIndex + 4000}: Found feeder name ${cell} but no date has been set yet.`);
//                 } else {
//                     console.log(`Row ${rowIndex + 4000}: Found feeder name ${cell}`);
//                     foundFeeders.add(cell.split('-')[0]);
//                     data.push({
//                         'Date': lastDate,
//                         'Name of Station': stationName,
//                         'Name of Feeder': cell
//                     });
//                 }
//             }
//         });
//     }

//     const feederArray = Array.from(allFeeders).sort();
//     const highestFeederNumber = parseInt(feederArray[feederArray.length - 1].substring(1));
//     const expectedFeeders = Array.from({ length: highestFeederNumber }, (_, i) => `F${i + 1}`);

//     expectedFeeders.forEach(feeder => {
//         if (!foundFeeders.has(feeder)) {
//             data.push({
//                 'Date': lastDate,
//                 'Name of Station': stationName,
//                 'Name of Feeder': `${feeder} - Feeder missing`
//             });
//         }
//     });

//     return data.filter(entry => entry.Date !== '');
// };

// // Main function to process all files and output consolidated data
// const processAllFiles = (files) => {
//     let consolidatedData = [];

//     files.forEach(file => {
//         const filePath = path.resolve(__dirname, file);
//         const fileData = processFile(filePath);
//         consolidatedData = consolidatedData.concat(fileData);
//     });

//     if (consolidatedData.length === 0) {
//         console.error('No valid data entries found. Please check the date extraction logic.');
//     } else {
//         console.log(`Total valid data entries found: ${consolidatedData.length}`);
//     }

//     // Output the consolidated data as a table
//     console.table(consolidatedData);
// };

// // Process all files
// processAllFiles(xlsFiles);




// CODE FOR BREAKING THE SHEET INTO THREE PARTS //
// const XLSX = require('xlsx');
// const path = require('path');

// // Input files
// const xlsFiles = [
//     'AMMASANDRA_110_15_10_2023.xls',
// ];

// // Helper function to process a specific range of rows
// const processRange = (worksheet, jsonData, stationName, startRow, endRow, initialDate) => {
//     const limitedData = jsonData.slice(startRow, endRow);
//     const data = [];
//     const foundFeeders = new Set();
//     const allFeeders = new Set();
//     let lastDate = initialDate;

//     for (let rowIndex = 0; rowIndex < limitedData.length; rowIndex++) {
//         const row = limitedData[rowIndex];
//         const dateValue = row[1];
//         if (typeof dateValue === 'number') {
//             const dateObj = XLSX.SSF.parse_date_code(dateValue);
//             lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
//         }

//         row.forEach((cell) => {
//             if (typeof cell === 'string' && cell.startsWith('F')) {
//                 const feederName = cell.split('-')[0];
//                 allFeeders.add(feederName);
//                 if (lastDate === '') {
//                     console.error(`Row ${rowIndex + startRow}: Found feeder name ${cell} but no date has been set yet.`);
//                 } else {
//                     foundFeeders.add(feederName);
//                     data.push({
//                         'Date': lastDate,
//                         'Name of Station': stationName,
//                         'Name of Feeder': cell
//                     });
//                 }
//             }
//         });
//     }

//     if (allFeeders.size > 0) {
//         const feederArray = Array.from(allFeeders).sort();
//         const highestFeederNumber = parseInt(feederArray[feederArray.length - 1].substring(1));
//         const expectedFeeders = Array.from({ length: highestFeederNumber }, (_, i) => `F${i + 1}`);

//         expectedFeeders.forEach(feeder => {
//             if (!foundFeeders.has(feeder)) {
//                 data.push({
//                     'Date': lastDate,
//                     'Name of Station': stationName,
//                     'Name of Feeder': `${feeder} - Feeder missing`
//                 });
//             }
//         });
//     }

//     return data.filter(entry => entry.Date !== '');
// };

// // Processing single file
// const processFile = (xlsFilePath) => {
//     const workbook = XLSX.readFile(xlsFilePath);
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = workbook.Sheets[sheetName];
//     const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
//     const stationName = path.basename(xlsFilePath, path.extname(xlsFilePath)).split('_')[0];

//     const initialDateCell = worksheet['B6'];
//     let initialDate = '';
//     if (initialDateCell && typeof initialDateCell.v === 'number') {
//         const dateObj = XLSX.SSF.parse_date_code(initialDateCell.v);
//         initialDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
//     } else {
//         console.error('Initial date in cell B6 is missing or not a number.');
//     }

//     const data0To2000 = processRange(worksheet, jsonData, stationName, 0, 2000, initialDate);
//     const data2000To4000 = processRange(worksheet, jsonData, stationName, 2000, 4000, initialDate);
//     const data4000To8000 = processRange(worksheet, jsonData, stationName, 4000, 8000, initialDate);

//     return {
//         data0To2000,
//         data2000To4000,
//         data4000To8000
//     };
// };

// // Processing all the files
// const processAllFiles = (files) => {
//     let consolidatedData0To2000 = [];
//     let consolidatedData2000To4000 = [];
//     let consolidatedData4000To8000 = [];

//     files.forEach(file => {
//         const filePath = path.resolve(__dirname, file);
//         const fileData = processFile(filePath);
//         consolidatedData0To2000 = consolidatedData0To2000.concat(fileData.data0To2000);
//         consolidatedData2000To4000 = consolidatedData2000To4000.concat(fileData.data2000To4000);
//         consolidatedData4000To8000 = consolidatedData4000To8000.concat(fileData.data4000To8000);
//     });

//     if (consolidatedData0To2000.length === 0) {
//         console.error('No valid data entries found in range 0-2000. Please check the date extraction logic.');
//     } else {
//         console.log(`Total valid data entries found in range 0-2000: ${consolidatedData0To2000.length}`);
//         console.table(consolidatedData0To2000);
//     }

//     if (consolidatedData2000To4000.length === 0) {
//         console.error('No valid data entries found in range 2000-4000. Please check the date extraction logic.');
//     } else {
//         console.log(`Total valid data entries found in range 2000-4000: ${consolidatedData2000To4000.length}`);
//         console.table(consolidatedData2000To4000);
//     }

//     if (consolidatedData4000To8000.length === 0) {
//         console.error('No valid data entries found in range 4000-8000. Please check the date extraction logic.');
//     } else {
//         console.log(`Total valid data entries found in range 4000-8000: ${consolidatedData4000To8000.length}`);
//         console.table(consolidatedData4000To8000);
//     }
// };

// processAllFiles(xlsFiles);








// const XLSX = require('xlsx');
// const path = require('path');
// const fs = require('fs');

// // Directory containing the files
// const directoryPath = path.resolve('C:/Users/KIIT/Desktop/Code/Project/readexcel/BGK_14112023');

// console.log(`Processing files in directory: ${directoryPath}`);

// // Validation functions
// const isValidDate = (dateStr) => {
//     const date = new Date(dateStr);
//     return !isNaN(date.getTime());
// };

// const isValidStationName = (name) => {
//     return typeof name === 'string' && name.length > 0;
// };

// const isValidItemName = (name, prefix) => {
//     const regex = new RegExp(`^${prefix}\\d+$`);
//     return regex.test(name);
// };

// // Helper function to process a specific range of rows
// const processRange = (worksheet, jsonData, stationName, startRow, endRow, initialDate, prefix, columnName) => {
//     console.log(`Processing range: ${startRow} to ${endRow}`);
//     const limitedData = jsonData.slice(startRow, endRow);
//     const data = [];
//     const foundItems = new Set();
//     const allItems = new Set();
//     let lastDate = initialDate;

//     if (!isValidStationName(stationName)) {
//         console.error(`Invalid station name: ${stationName}`);
//         return [];
//     }

//     for (let rowIndex = 0; rowIndex < limitedData.length; rowIndex++) {
//         const row = limitedData[rowIndex];
//         const dateValue = row[1];
//         if (typeof dateValue === 'number') {
//             const dateObj = XLSX.SSF.parse_date_code(dateValue);
//             lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
//         }

//         if (!isValidDate(lastDate)) {
//             console.error(`Invalid date at row ${rowIndex + startRow}: ${lastDate}`);
//             continue; // Skip this row
//         }

//         row.forEach((cell) => {
//             if (typeof cell === 'string' && cell.startsWith(prefix)) {
//                 const itemName = cell.split('-')[0];
//                 allItems.add(itemName);
//                 if (!isValidItemName(itemName, prefix)) {
//                     console.error(`Invalid item name at row ${rowIndex + startRow}: ${itemName}`);
//                     return;
//                 }
//                 if (lastDate === '') {
//                     console.error(`Row ${rowIndex + startRow}: Found ${prefix} name ${cell} but no date has been set yet.`);
//                 } else {
//                     foundItems.add(itemName);
//                     data.push({
//                         'Date': lastDate,
//                         'Name of Station': stationName,
//                         [columnName]: cell
//                     });
//                 }
//             }
//         });
//     }

//     if (allItems.size > 0) {
//         const itemArray = Array.from(allItems).sort();
//         const highestItemNumber = parseInt(itemArray[itemArray.length - 1].substring(1));
//         const expectedItems = Array.from({ length: highestItemNumber }, (_, i) => `${prefix}${i + 1}`);

//         expectedItems.forEach(item => {
//             if (!foundItems.has(item)) {
//                 data.push({
//                     'Date': lastDate,
//                     'Name of Station': stationName,
//                     [columnName]: `${item} - ${prefix} missing`
//                 });
//             }
//         });
//     }

//     return data.filter(entry => entry.Date !== '');
// };

// // Processing single file
// const processFile = (xlsFilePath) => {
//     console.log(`Processing file: ${xlsFilePath}`);
//     const workbook = XLSX.readFile(xlsFilePath);
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = workbook.Sheets[sheetName];
//     const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
//     const stationName = path.basename(xlsFilePath, path.extname(xlsFilePath)).split('_')[0];

//     const initialDateCell = worksheet['B6'];
//     let initialDate = '';
//     if (initialDateCell && typeof initialDateCell.v === 'number') {
//         const dateObj = XLSX.SSF.parse_date_code(initialDateCell.v);
//         initialDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
//     } else {
//         console.error('Initial date in cell B6 is missing or not a number.');
//     }

//     const data0To2000 = processRange(worksheet, jsonData, stationName, 0, 2000, initialDate, 'F', 'Name of Feeder');
//     const data2000To4000 = processRange(worksheet, jsonData, stationName, 2000, 4000, initialDate, 'T', 'Name of Transformer');
//     const data4000To8000 = processRange(worksheet, jsonData, stationName, 4000, 8000, initialDate, 'F', 'Name of Feeder');

//     return {
//         data0To2000,
//         data2000To4000,
//         data4000To8000
//     };
// };

// // Processing all the files in the directory
// const processAllFiles = (directory) => {
//     fs.readdir(directory, (err, files) => {
//         if (err) {
//             return console.error('Unable to scan directory: ' + err);
//         }

//         console.log(`Found ${files.length} files in the directory.`);

//         let consolidatedData0To2000 = [];
//         let consolidatedData2000To4000 = [];
//         let consolidatedData4000To8000 = [];

//         // Filter and limit to the first 100 .xls files
//         const xlsFiles = files.filter(file => path.extname(file) === '.xls').slice(0, 100);

//         console.log(`Processing the first ${xlsFiles.length} .xls files`);

//         xlsFiles.forEach(file => {
//             const filePath = path.join(directory, file);
//             const fileData = processFile(filePath);
//             consolidatedData0To2000 = consolidatedData0To2000.concat(fileData.data0To2000);
//             consolidatedData2000To4000 = consolidatedData2000To4000.concat(fileData.data2000To4000);
//             consolidatedData4000To8000 = consolidatedData4000To8000.concat(fileData.data4000To8000);
//         });

//         if (consolidatedData0To2000.length === 0) {
//             console.error('No valid data entries found in range 0-2000. Please check the date extraction logic.');
//         } else {
//             console.log(`Total valid data entries found in range 0-2000: ${consolidatedData0To2000.length}`);
//             console.table(consolidatedData0To2000);
//         }

//         if (consolidatedData2000To4000.length === 0) {
//             console.error('No valid data entries found in range 2000-4000. Please check the date extraction logic.');
//         } else {
//             console.log(`Total valid data entries found in range 2000-4000: ${consolidatedData2000To4000.length}`);
//             console.table(consolidatedData2000To4000);
//         }

//         if (consolidatedData4000To8000.length === 0) {
//             console.error('No valid data entries found in range 4000-8000. Please check the date extraction logic.');
//         } else {
//             console.log(`Total valid data entries found in range 4000-8000: ${consolidatedData4000To8000.length}`);
//             console.table(consolidatedData4000To8000);
//         }
//     });
// };

// // Start processing all files in the specified directory
// processAllFiles(directoryPath);












const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Directory containing the files
const directoryPath = path.resolve('C:/Users/KIIT/Desktop/Code/Project/readexcel/BGK_14112023');

console.log(`Processing files in directory: ${directoryPath}`);

// Validation functions
const isValidDate = (dateStr) => {
    const date = new Date(dateStr);
    return !isNaN(date.getTime());
};

const isValidStationName = (name) => {
    return typeof name === 'string' && name.length > 0;
};

const isValidItemName = (name, prefix) => {
    const regex = new RegExp(`^${prefix}\\d+$`);
    return regex.test(name);
};

// Helper function to process a specific range of rows
const processRange = (worksheet, jsonData, stationName, startRow, endRow, initialDate, prefix, columnName) => {
    console.log(`Processing range: ${startRow} to ${endRow}`);
    const limitedData = jsonData.slice(startRow, endRow);
    const data = [];
    const foundItems = new Set();
    const allItems = new Set();
    let lastDate = initialDate;

    if (!isValidStationName(stationName)) {
        console.error(`Invalid station name: ${stationName}`);
        return [];
    }

    for (let rowIndex = 0; rowIndex < limitedData.length; rowIndex++) {
        const row = limitedData[rowIndex];
        const dateValue = row[1];
        if (typeof dateValue === 'number') {
            const dateObj = XLSX.SSF.parse_date_code(dateValue);
            lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
        }

        if (!isValidDate(lastDate)) {
            console.error(`Invalid date at row ${rowIndex + startRow}: ${lastDate}`);
            continue; // Skip this row
        }

        row.forEach((cell) => {
            if (typeof cell === 'string' && cell.startsWith(prefix)) {
                const itemName = cell.split('-')[0];
                allItems.add(itemName);
                if (!isValidItemName(itemName, prefix)) {
                    console.error(`Invalid item name at row ${rowIndex + startRow}: ${itemName}`);
                    return;
                }
                if (lastDate === '') {
                    console.error(`Row ${rowIndex + startRow}: Found ${prefix} name ${cell} but no date has been set yet.`);
                } else {
                    foundItems.add(itemName);
                    data.push({
                        'Date': lastDate,
                        'Name of Station': stationName,
                        [columnName]: cell
                    });
                }
            }
        });
    }

    if (allItems.size > 0) {
        const itemArray = Array.from(allItems).sort();
        const highestItemNumber = parseInt(itemArray[itemArray.length - 1].substring(1));
        const expectedItems = Array.from({ length: highestItemNumber }, (_, i) => `${prefix}${i + 1}`);

        expectedItems.forEach(item => {
            if (!foundItems.has(item)) {
                data.push({
                    'Date': lastDate,
                    'Name of Station': stationName,
                    [columnName]: `${item} - ${prefix} missing`
                });
            }
        });
    }

    return data.filter(entry => entry.Date !== '');
};

// Processing single file
const processFile = (xlsFilePath) => {
    console.log(`Processing file: ${xlsFilePath}`);
    const workbook = XLSX.readFile(xlsFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
    const stationName = path.basename(xlsFilePath, path.extname(xlsFilePath)).split('_')[0];

    const initialDateCell = worksheet['B6'];
    let initialDate = '';
    if (initialDateCell && typeof initialDateCell.v === 'number') {
        const dateObj = XLSX.SSF.parse_date_code(initialDateCell.v);
        initialDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
    } else {
        console.error('Initial date in cell B6 is missing or not a number.');
    }

    const data0To2000 = processRange(worksheet, jsonData, stationName, 0, 2000, initialDate, 'F', 'Feeder Information (Range 0-2000) [HV]');
    const data2000To4000 = processRange(worksheet, jsonData, stationName, 2000, 4000, initialDate, 'T', 'Number of Transformer (Range 2000-4000)');
    const data4000To8000 = processRange(worksheet, jsonData, stationName, 4000, 8000, initialDate, 'F', 'Feeder Information (Range 4000-8000) [LV]');

    return {
        data0To2000,
        data2000To4000,
        data4000To8000
    };
};

// Processing all the files in the directory
const processAllFiles = (directory) => {
    fs.readdir(directory, (err, files) => {
        if (err) {
            return console.error('Unable to scan directory: ' + err);
        }

        console.log(`Found ${files.length} files in the directory.`);

        let consolidatedData0To2000 = [];
        let consolidatedData2000To4000 = [];
        let consolidatedData4000To8000 = [];

        // Filter and limit to the first 100 .xls files
        const xlsFiles = files.filter(file => path.extname(file) === '.xls').slice(0, 100);

        console.log(`Processing the first ${xlsFiles.length} .xls files`);

        xlsFiles.forEach(file => {
            const filePath = path.join(directory, file);
            const fileData = processFile(filePath);
            consolidatedData0To2000 = consolidatedData0To2000.concat(fileData.data0To2000);
            consolidatedData2000To4000 = consolidatedData2000To4000.concat(fileData.data2000To4000);
            consolidatedData4000To8000 = consolidatedData4000To8000.concat(fileData.data4000To8000);
        });

        if (consolidatedData0To2000.length === 0) {
            console.error('No valid data entries found in range 0-2000. Please check the date extraction logic.');
        } else {
            console.log(`Total valid data entries found in range 0-2000: ${consolidatedData0To2000.length}`);
        }

        if (consolidatedData2000To4000.length === 0) {
            console.error('No valid data entries found in range 2000-4000. Please check the date extraction logic.');
        } else {
            console.log(`Total valid data entries found in range 2000-4000: ${consolidatedData2000To4000.length}`);
        }

        if (consolidatedData4000To8000.length === 0) {
            console.error('No valid data entries found in range 4000-8000. Please check the date extraction logic.');
        } else {
            console.log(`Total valid data entries found in range 4000-8000: ${consolidatedData4000To8000.length}`);
        }

        // Write the consolidated data to a new Excel file in the required format
        const writeExcel = (data, outputFilePath) => {
            const newWorkbook = XLSX.utils.book_new();
            const newSheet = XLSX.utils.aoa_to_sheet([
                ['Zone', 'Name of the Stations', 'Date', 'Voltage level', 'Feeder Information', '', 'Number of Transformer', 'Feeder Information'],
                ['', '', '', '(KV)', '(Range 0-2000) [HV]', '', '(Range 2000-4000)', '(Range 4000-8000) [LV]']
            ]);
            
            // Add data rows
            data.forEach(row => {
                const feeder0To2000 = row['Feeder Information (Range 0-2000) [HV]'] || '';
                const transformer = row['Number of Transformer (Range 2000-4000)'] || '';
                const feeder4000To8000 = row['Feeder Information (Range 4000-8000) [LV]'] || '';
                XLSX.utils.sheet_add_aoa(newSheet, [[row.Zone, row['Name of Station'], row.Date, row['Voltage level'], feeder0To2000, '', transformer, feeder4000To8000]], { origin: -1 });
            });

            // Merge cells for the header
            newSheet['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // Merge "Zone"
                { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, // Merge "Name of the Stations"
                { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, // Merge "Date"
                { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, // Merge "Voltage level"
                { s: { r: 0, c: 4 }, e: { r: 0, c: 5 } }, // Merge "Feeder Information (Range 0-2000) [HV]"
                { s: { r: 0, c: 6 }, e: { r: 1, c: 6 } }, // Merge "Number of Transformer (Range 2000-4000)"
                { s: { r: 0, c: 7 }, e: { r: 1, c: 7 } }  // Merge "Feeder Information (Range 4000-8000) [LV]"
            ];

            XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');
            XLSX.writeFile(newWorkbook, outputFilePath);
            console.log(`File written to ${outputFilePath}`);  // Log to confirm file writing
        };

        // Combine data into the required format
        const combinedData = consolidatedData0To2000.map((entry, index) => {
            return {
                Zone: 'BGM',
                'Name of Station': entry['Name of Station'],
                Date: entry['Date'],
                'Voltage level': '110/11',
                'Feeder Information (Range 0-2000) [HV]': entry['Feeder Information (Range 0-2000) [HV]'],
                'Number of Transformer (Range 2000-4000)': consolidatedData2000To4000[index] ? consolidatedData2000To4000[index]['Number of Transformer (Range 2000-4000)'] : '',
                'Feeder Information (Range 4000-8000) [LV]': consolidatedData4000To8000[index] ? consolidatedData4000To8000[index]['Feeder Information (Range 4000-8000) [LV]'] : ''
            };
        });

        const outputFilePath = path.join(directory, 'consolidated_output.xlsx');
        if (combinedData.length > 0) {
            writeExcel(combinedData, outputFilePath);
        }
    });
};

// Start processing all files in the specified directory
processAllFiles(directoryPath);





// LV Feeders finder //
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Update this to the correct path where your Excel files are located
const directoryPath = path.join(__dirname, 'HSN_15102023');

try {
    // Check if the directory exists
    if (!fs.existsSync(directoryPath)) {
        console.error('Directory does not exist:', directoryPath);
        process.exit(1);
    }

    // Read all files in the directory
    const filesInDirectory = fs.readdirSync(directoryPath);

    // Filter the files to only include .xls or .xlsx files
    const excelFiles = filesInDirectory.filter(file => file.endsWith('.xls') || file.endsWith('.xlsx'));

    if (excelFiles.length === 0) {
        console.error('No Excel files found in the directory:', directoryPath);
        process.exit(1);
    }

    const allUniqueValues = new Set();
    const feederData = [];

    excelFiles.forEach((fileName) => {
        const filePath = path.join(directoryPath, fileName);

        if (fs.existsSync(filePath)) {
            // Load the workbook
            const workbook = XLSX.readFile(filePath);
            console.log('Workbook loaded successfully:', filePath);

            const sheetName = workbook.SheetNames[0]; // Assuming the data is in the first sheet
            const worksheet = workbook.Sheets[sheetName];
            console.log('Worksheet loaded:', sheetName);

            // Utility function to get the value of a cell
            const getCellValue = (cellAddress) => worksheet[cellAddress]?.v || null;

            // Get the station name from cell A4001
            const stationName = getCellValue('A4001');

            // Pattern to match values with prefix 'F'
            const feederPattern = /^F.*/;

            // Iterate over all cells in the sheet within the specified range
            for (const cellAddress in worksheet) {
                // Decode the cell address to get the row number
                const cell = XLSX.utils.decode_cell(cellAddress);
                if (cell.r >= 4000 && cell.r <= 6000) {
                    const cellValue = getCellValue(cellAddress);
                    if (cellValue && feederPattern.test(cellValue)) {
                        allUniqueValues.add(cellValue);
                        feederData.push({
                            'Name of Feeder LV': cellValue,
                            'Station Name': stationName
                        });
                    }
                }
            }
        } else {
            console.error('File not found:', filePath);
        }
    });

    // Convert set to array for logging and storage
    const uniqueValuesArray = Array.from(allUniqueValues);

    // Prepare the data for logging
    const transformedData = feederData.map(value => ({
        'Name of Feeder LV': value['Name of Feeder LV'],
        'Station Name': value['Station Name']
    }));

    // Log out the unique values matching the pattern with the columns "Name of Feeder LV" and "Station Name"
    console.log('Unique values matching the pattern:');
    console.table(transformedData);

    // Save the results to an Excel file named "Feeder_Output"
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(transformedData);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Feeder Output');
    XLSX.writeFile(newWorkbook, 'Feeder_Output.xlsx');
    console.log('Unique values saved to Feeder_Output.xlsx');
} catch (error) {
    console.error('An error occurred:', error);
}
