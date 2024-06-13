// const XLSX = require('xlsx');
// const path = require('path');

// // List of XLS files
// const xlsFiles = [
//     'HIREMURAL_110_14_11_2023.xls',
//     'SATTI_110_14_11_2023.xls'
// ];

// // Function to generate time intervals for a 24-hour period
// const generateTimeIntervals = () => {
//     const intervals = [];
//     for (let hour = 0; hour < 24; hour++) {
//         for (let minute = 0; minute < 60; minute++) {
//             const start = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
//             const endMinute = (minute + 1) % 60;
//             const endHour = hour + Math.floor((minute + 1) / 60);
//             const end = `${String(endHour % 24).padStart(2, '0')}:${String(endMinute).padStart(2, '0')}`;
//             intervals.push(`${start} - ${end}`);
//         }
//     }
//     return intervals;
// };

// const timeIntervals = generateTimeIntervals();

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
//                         'Name of Feeder': cell,
//                         'Time': '' // Placeholder for time intervals, will be filled later
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
//                 'Name of Feeder': `${feeder} - Feeder missing`,
//                 'Time': '' // Placeholder for time intervals, will be filled later
//             });
//         }
//     });

//     // Fill in the time intervals for each feeder
//     const expandedData = [];
//     data.forEach(entry => {
//         timeIntervals.forEach(time => {
//             expandedData.push({
//                 'Date': entry.Date,
//                 'Name of Station': entry['Name of Station'],
//                 'Name of Feeder': entry['Name of Feeder'],
//                 'Time': time
//             });
//         });
//     });

//     return expandedData.filter(entry => entry.Date !== '');
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

//     // Save the consolidated data to a new Excel file
//     const newWorkbook = XLSX.utils.book_new();
//     const newWorksheet = XLSX.utils.json_to_sheet(consolidatedData);
//     XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Consolidated Data');
//     XLSX.writeFile(newWorkbook, 'Consolidated_Data.xlsx');
//     console.log('Consolidated data saved to: Consolidated_Data.xlsx');
// };

// // Process all files
// processAllFiles(xlsFiles);







const XLSX = require('xlsx');
const path = require('path');

// List of XLS files
const xlsFiles = [
    'HIREMURAL_110_14_11_2023.xls',
    'SATTI_110_14_11_2023.xls'
];

// Function to generate time intervals for a 24-hour period
const generateTimeIntervals = () => {
    const intervals = [];
    for (let hour = 0; hour < 24; hour++) {
        for (let minute = 0; minute < 60; minute++) {
            const start = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
            const endMinute = (minute + 1) % 60;
            const endHour = hour + Math.floor((minute + 1) / 60);
            const end = `${String(endHour % 24).padStart(2, '0')}:${String(endMinute).padStart(2, '0')}`;
            intervals.push(`${start} - ${end}`);
        }
    }
    return intervals;
};

const timeIntervals = generateTimeIntervals();

// Function to process a single file and return data
const processFile = (xlsFilePath) => {
    const workbook = XLSX.readFile(xlsFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
    const stationName = worksheet['A1'].v;

    const initialDateCell = worksheet['B6'];
    let lastDate = '';
    if (initialDateCell && typeof initialDateCell.v === 'number') {
        const dateObj = XLSX.SSF.parse_date_code(initialDateCell.v);
        lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
        console.log(`Initial date from B6 parsed as ${lastDate}`);
    } else {
        console.error('Initial date in cell B6 is missing or not a number.');
    }

    const limitedData = jsonData.slice(4000, 8001);
    const data = [];
    const foundFeeders = new Set();
    const allFeeders = new Set();

    for (let rowIndex = 0; rowIndex < limitedData.length; rowIndex++) {
        const row = limitedData[rowIndex];
        const dateValue = row[1];
        if (typeof dateValue === 'number') {
            const dateObj = XLSX.SSF.parse_date_code(dateValue);
            lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
            console.log(`Row ${rowIndex + 4000}: Date parsed as ${lastDate}`);
        }

        row.forEach((cell, cellIndex) => {
            if (typeof cell === 'string' && cell.startsWith('F')) {
                allFeeders.add(cell.split('-')[0]);
                if (lastDate === '') {
                    console.error(`Row ${rowIndex + 4000}: Found feeder name ${cell} but no date has been set yet.`);
                } else {
                    console.log(`Row ${rowIndex + 4000}: Found feeder name ${cell}`);
                    foundFeeders.add(cell.split('-')[0]);
                    data.push({
                        'Date': lastDate,
                        'Name of Station': stationName,
                        'Name of Feeder': cell,
                        'Time': '', // Placeholder for time intervals, will be filled later
                        'Active Power': '' // Placeholder for Active Power, will be filled later
                    });
                }
            }
        });
    }

    const feederArray = Array.from(allFeeders).sort();
    const highestFeederNumber = parseInt(feederArray[feederArray.length - 1].substring(1));
    const expectedFeeders = Array.from({ length: highestFeederNumber }, (_, i) => `F${i + 1}`);

    expectedFeeders.forEach(feeder => {
        if (!foundFeeders.has(feeder)) {
            data.push({
                'Date': lastDate,
                'Name of Station': stationName,
                'Name of Feeder': `${feeder} - Feeder missing`,
                'Time': '', // Placeholder for time intervals, will be filled later
                'Active Power': '' // Placeholder for Active Power, will be filled later
            });
        }
    });

    // Extract active power values for each name of feeder
    const activePowerValues = {};
    feederArray.forEach((feeder, index) => {
        const columnIndex = 13 + index * 13; // Start column 'M' (13th), then every 13th column
        const values = [];
        for (let row = 4006; row < 5446; row++) { // Iterate over 1440 rows
            const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: columnIndex - 1 }); // Convert to 0-based index
            const cellValue = worksheet[cellAddress]?.v || null;
            values.push(cellValue);
        }
        activePowerValues[feeder] = values;
    });

    // Fill in the time intervals and active power values for each feeder
    const expandedData = [];
    data.forEach(entry => {
        const feederId = entry['Name of Feeder'].split('-')[0]; // Extract feeder ID (e.g., 'F1')
        const activePowerList = activePowerValues[feederId] || [];

        timeIntervals.forEach((time, index) => {
            expandedData.push({
                'Date': entry.Date,
                'Name of Station': entry['Name of Station'],
                'Name of Feeder': entry['Name of Feeder'],
                'Time': time,
                'Active Power': activePowerList[index % activePowerList.length] || null
            });
        });
    });

    return expandedData.filter(entry => entry.Date !== '');
};

// Main function to process all files and output consolidated data
const processAllFiles = (files) => {
    let consolidatedData = [];

    files.forEach(file => {
        const filePath = path.resolve(__dirname, file);
        const fileData = processFile(filePath);
        consolidatedData = consolidatedData.concat(fileData);
    });

    if (consolidatedData.length === 0) {
        console.error('No valid data entries found. Please check the date extraction logic.');
    } else {
        console.log(`Total valid data entries found: ${consolidatedData.length}`);
    }

    // Output the consolidated data as a table
    console.table(consolidatedData);

    // Save the consolidated data to a new Excel file
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(consolidatedData);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Consolidated Data');
    XLSX.writeFile(newWorkbook, 'Consolidated_Data.xlsx');
    console.log('Consolidated data saved to: Consolidated_Data.xlsx');
};

// Process all files
processAllFiles(xlsFiles);
