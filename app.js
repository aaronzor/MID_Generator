const fs = require('fs');
const XLSX = require('xlsx');

// Check if the required command line arguments are provided
if (process.argv.length < 4) {
  console.log('Usage: node app.js [midPrefix] [numRows]');
  process.exit(1);
}

// Extract the MID prefix and number of rows from the command line arguments
const midPrefix = process.argv[2];
const numRows = parseInt(process.argv[3], 10);

// Get a random channel for each MID
function getRandomChannel() {
  const channels = ['Online', 'Offline', 'Unrestricted'];
  return channels[Math.floor(Math.random() * channels.length)];
}

// Generate a random Narrative for each MID
// Will generate a number from 0001 to 9999 and append " - Test Narr"
function getRandomNarrative() {
  const randomNum = Math.floor(Math.random() * 10000);
  return `${randomNum} - Test Narr`;
}

// Generate MIDs
function generateMids(midPrefix, numRows) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new();

  // Define the headers for the worksheet
  const headers = [
    'MID', 'Acquirer', 'Channel', 'Show Maps', 'Address1', 'Address2', 'City', 'Country', 'Postcode', 'Reference', 'Status', 'Narrative',
  ];

  // Initialize the data array with the headers
  const data = [headers];

  // Populate the data array with the specified number of rows
  for (let i = 1; i <= numRows; i++) {
    const rowData = [
      `${midPrefix}${i}`, '', getRandomChannel(), '', `${i} Test Street`, '', 'Belfast', 'IRL', 'BT1 1AA', '', 'Live', getRandomNarrative(),
    ];
    data.push(rowData);
  }

  // Create a new worksheet with the generated data
  const worksheet = XLSX.utils.aoa_to_sheet(data);

  // Add the worksheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

  // Write the workbook to a file
  const filePath = `${numRows} - testMIDs.xlsx`;
  XLSX.writeFile(workbook, filePath);
  console.log(`Generated XLSX file: ${filePath} - ${numRows} MIDs Generated`);
}

generateMids(midPrefix, numRows);
