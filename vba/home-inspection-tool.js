// This code will generate an Excel file for a home inspection tool with multiple sheets
import * as XLSX from 'xlsx';

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Create the main inspection sheet
const mainSheetData = [
  ['HOME INSPECTION REPORT', '', '', '', ''],
  ['', '', '', '', ''],
  ['Property Address:', '', '', 'Inspection Date:', ''],
  ['City, State, Zip:', '', '', 'Inspector Name:', ''],
  ['Client Name:', '', '', 'Client Phone:', ''],
  ['Client Email:', '', '', 'Weather Conditions:', ''],
  ['Year Built:', '', '', 'Square Footage:', ''],
  ['', '', '', '', ''],
  ['RATING SYSTEM', '', '', '', ''],
  ['1 - Immediate Attention Required', '', '', '', ''],
  ['2 - Repair/Replace Soon', '', '', '', ''],
  ['3 - Monitor/Maintenance Item', '', '', '', ''],
  ['4 - Normal Wear and Tear', '', '', '', ''],
  ['5 - Good Condition', '', '', '', ''],
  ['N/A - Not Applicable', '', '', '', ''],
  ['', '', '', '', ''],
  ['INSPECTION SUMMARY', '', '', '', ''],
  ['Area', 'Rating', 'Deficiency', 'Recommendation', 'Photo Reference'],
];

// Add rows for each inspection area in the summary
const inspectionAreas = [
  'Roof', 'Exterior', 'Foundation', 'Basement', 'Crawlspace', 
  'Plumbing', 'Electrical', 'HVAC', 'Interior', 'Attic', 
  'Insulation', 'Ventilation', 'Kitchen', 'Bathrooms', 'Garage'
];

inspectionAreas.forEach(area => {
  mainSheetData.push([area, '', '', '', '']);
});

// Add final notes section
mainSheetData.push(
  ['', '', '', '', ''],
  ['ADDITIONAL NOTES', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', '']
);

// Create worksheet from data
const mainSheet = XLSX.utils.aoa_to_sheet(mainSheetData);

// Set column widths
const colWidths = [{ wch: 20 }, { wch: 10 }, { wch: 40 }, { wch: 40 }, { wch: 15 }];
mainSheet['!cols'] = colWidths;

// Add the main sheet to the workbook
XLSX.utils.book_append_sheet(workbook, mainSheet, 'Main Report');

// Function to create detailed inspection sheets
function createInspectionSheet(areaName, checkItems) {
  const sheetData = [
    [areaName.toUpperCase() + ' INSPECTION DETAILS', '', '', ''],
    ['', '', '', ''],
    ['Item', 'Condition (1-5)', 'Notes', 'Photo Reference'],
  ];
  
  checkItems.forEach(item => {
    sheetData.push([item, '', '', '']);
  });
  
  // Add final notes
  sheetData.push(
    ['', '', '', ''],
    ['ADDITIONAL NOTES:', '', '', ''],
    ['', '', '', '']
  );
  
  const sheet = XLSX.utils.aoa_to_sheet(sheetData);
  
  // Set column widths
  sheet['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 50 }, { wch: 15 }];
  
  return sheet;
}

// Create individual area sheets
// Roof inspection items
const roofItems = [
  'Roof Covering', 'Roof Flashing', 'Roof Drainage', 'Skylights', 
  'Chimneys', 'Roof Penetrations', 'Signs of Leaking', 'Roof Ventilation',
  'Roof Structure', 'Estimated Remaining Life'
];
const roofSheet = createInspectionSheet('Roof', roofItems);
XLSX.utils.book_append_sheet(workbook, roofSheet, 'Roof');

// Exterior inspection items
const exteriorItems = [
  'Siding/Cladding', 'Exterior Doors', 'Windows', 'Trim', 
  'Eaves/Soffits/Fascia', 'Exterior Lighting', 'Walkways', 'Driveway', 
  'Steps/Stoops', 'Porches/Decks', 'Railings', 'Grading/Drainage',
  'Vegetation', 'Retaining Walls', 'Fences/Gates'
];
const exteriorSheet = createInspectionSheet('Exterior', exteriorItems);
XLSX.utils.book_append_sheet(workbook, exteriorSheet, 'Exterior');

// Foundation inspection items
const foundationItems = [
  'Foundation Walls', 'Visible Structural Components', 'Signs of Water Penetration',
  'Cracks', 'Settlement', 'Foundation Type', 'Anchor Bolts', 'Floor Framing',
  'Wall Framing', 'Support Posts/Columns', 'Support Beams'
];
const foundationSheet = createInspectionSheet('Foundation', foundationItems);
XLSX.utils.book_append_sheet(workbook, foundationSheet, 'Foundation');

// Plumbing inspection items
const plumbingItems = [
  'Water Supply Lines', 'Drain/Waste/Vent Pipes', 'Main Water Shut-off',
  'Water Pressure', 'Water Heater', 'Toilets', 'Sinks', 'Tubs/Showers',
  'Faucets', 'Visible Leaks', 'Sump Pump', 'Sewage Ejector Pump',
  'Gas Lines', 'Main Gas Shut-off'
];
const plumbingSheet = createInspectionSheet('Plumbing', plumbingItems);
XLSX.utils.book_append_sheet(workbook, plumbingSheet, 'Plumbing');

// Electrical inspection items
const electricalItems = [
  'Service Entrance', 'Main Panel', 'Circuit Breakers/Fuses', 'Branch Wiring',
  'Grounding', 'GFCI Protection', 'AFCI Protection', 'Outlets', 'Switches',
  'Light Fixtures', 'Ceiling Fans', 'Smoke Detectors', 'Carbon Monoxide Detectors'
];
const electricalSheet = createInspectionSheet('Electrical', electricalItems);
XLSX.utils.book_append_sheet(workbook, electricalSheet, 'Electrical');

// HVAC inspection items
const hvacItems = [
  'Heating System Type', 'Heating System Age', 'Heating Operation', 'Cooling System Type',
  'Cooling System Age', 'Cooling Operation', 'Distribution System', 'Thermostat',
  'Filters', 'Humidifier', 'Dehumidifier', 'Ductwork', 'Ventilation'
];
const hvacSheet = createInspectionSheet('HVAC', hvacItems);
XLSX.utils.book_append_sheet(workbook, hvacSheet, 'HVAC');

// Interior inspection items
const interiorItems = [
  'Floors', 'Walls', 'Ceilings', 'Windows', 'Interior Doors', 'Stairs',
  'Railings', 'Countertops', 'Cabinets', 'Appliances', 'Evidence of Pests',
  'Evidence of Water Damage', 'Evidence of Mold'
];
const interiorSheet = createInspectionSheet('Interior', interiorItems);
XLSX.utils.book_append_sheet(workbook, interiorSheet, 'Interior');

// Attic inspection items
const atticItems = [
  'Access', 'Insulation Type', 'Insulation Depth', 'Ventilation', 
  'Visible Electrical', 'Visible Plumbing', 'Visible Framing', 
  'Signs of Leaking', 'Signs of Pests', 'Exhaust Venting'
];
const atticSheet = createInspectionSheet('Attic', atticItems);
XLSX.utils.book_append_sheet(workbook, atticSheet, 'Attic');

// Create a photo log sheet
const photoLogData = [
  ['PHOTO LOG', '', '', ''],
  ['', '', '', ''],
  ['Photo #', 'Location', 'Description', 'Date Taken'],
];

// Add 20 empty rows for photos
for (let i = 1; i <= 20; i++) {
  photoLogData.push([i.toString(), '', '', '']);
}

const photoLogSheet = XLSX.utils.aoa_to_sheet(photoLogData);
photoLogSheet['!cols'] = [{ wch: 10 }, { wch: 25 }, { wch: 50 }, { wch: 15 }];
XLSX.utils.book_append_sheet(workbook, photoLogSheet, 'Photo Log');

// Create a cost estimate sheet
const costEstimateData = [
  ['REPAIR COST ESTIMATES', '', '', '', ''],
  ['', '', '', '', ''],
  ['Item', 'Priority (1-5)', 'Estimated Cost (Low)', 'Estimated Cost (High)', 'Notes'],
];

// Add 15 empty rows for estimates
for (let i = 1; i <= 15; i++) {
  costEstimateData.push(['', '', '', '', '']);
}

// Add totals row
costEstimateData.push(
  ['TOTALS', '', '=SUM(C4:C18)', '=SUM(D4:D18)', '']
);

const costEstimateSheet = XLSX.utils.aoa_to_sheet(costEstimateData);
costEstimateSheet['!cols'] = [{ wch: 40 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 30 }];
XLSX.utils.book_append_sheet(workbook, costEstimateSheet, 'Cost Estimates');

// Create a maintenance checklist
const maintenanceData = [
  ['HOME MAINTENANCE CHECKLIST', '', '', ''],
  ['', '', '', ''],
  ['Task', 'Frequency', 'Last Completed', 'Next Due'],
  
  // Monthly tasks
  ['MONTHLY MAINTENANCE', '', '', ''],
  ['Test smoke/CO detectors', 'Monthly', '', ''],
  ['Check HVAC filters', 'Monthly', '', ''],
  ['Check water softener', 'Monthly', '', ''],
  ['Clean range hood filters', 'Monthly', '', ''],
  ['Check for plumbing leaks', 'Monthly', '', ''],
  
  // Quarterly tasks
  ['QUARTERLY MAINTENANCE', '', '', ''],
  ['Test GFCIs', 'Quarterly', '', ''],
  ['Run water in unused fixtures', 'Quarterly', '', ''],
  ['Check water heater', 'Quarterly', '', ''],
  ['Check garage door operation', 'Quarterly', '', ''],
  
  // Biannual tasks
  ['BIANNUAL MAINTENANCE', '', '', ''],
  ['Service HVAC systems', 'Biannual', '', ''],
  ['Check fire extinguishers', 'Biannual', '', ''],
  ['Clean gutters', 'Biannual', '', ''],
  ['Check exterior drainage', 'Biannual', '', ''],
  
  // Annual tasks
  ['ANNUAL MAINTENANCE', '', '', ''],
  ['Inspect roof', 'Annual', '', ''],
  ['Clean chimney', 'Annual', '', ''],
  ['Inspect attic', 'Annual', '', ''],
  ['Check exterior paint/siding', 'Annual', '', ''],
  ['Check foundation', 'Annual', '', ''],
  ['Flush water heater', 'Annual', '', ''],
  ['Check weatherstripping', 'Annual', '', ''],
  ['Check caulking around showers/tubs', 'Annual', '', '']
];

const maintenanceSheet = XLSX.utils.aoa_to_sheet(maintenanceData);
maintenanceSheet['!cols'] = [{ wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
XLSX.utils.book_append_sheet(workbook, maintenanceSheet, 'Maintenance');

// Client contact information sheet
const contactData = [
  ['CLIENT CONTACT INFORMATION', '', '', ''],
  ['', '', '', ''],
  ['Client Name:', '', 'Spouse/Partner Name:', ''],
  ['Phone (Primary):', '', 'Phone (Secondary):', ''],
  ['Email:', '', 'Alternate Email:', ''],
  ['Mailing Address:', '', '', ''],
  ['', '', '', ''],
  ['PROPERTY INFORMATION', '', '', ''],
  ['Property Address:', '', '', ''],
  ['City:', '', 'State:', ''],
  ['Zip Code:', '', 'County:', ''],
  ['Year Built:', '', 'Square Footage:', ''],
  ['Stories:', '', 'Bedrooms:', ''],
  ['Bathrooms:', '', 'Garage Spaces:', ''],
  ['Basement:', '', 'Crawlspace:', ''],
  ['', '', '', ''],
  ['IMPORTANT CONTACTS', '', '', ''],
  ['', '', '', ''],
  ['Contact Type', 'Name', 'Phone', 'Email'],
  ['Real Estate Agent', '', '', ''],
  ['Insurance Agent', '', '', ''],
  ['Mortgage Lender', '', '', ''],
  ['Electrician', '', '', ''],
  ['Plumber', '', '', ''],
  ['HVAC Contractor', '', '', ''],
  ['Roofer', '', '', ''],
  ['Landscaper', '', '', ''],
  ['Pest Control', '', '', ''],
  ['', '', '', ''],
  ['UTILITIES INFORMATION', '', '', ''],
  ['', '', '', ''],
  ['Utility', 'Provider', 'Account Number', 'Contact'],
  ['Electricity', '', '', ''],
  ['Gas', '', '', ''],
  ['Water/Sewer', '', '', ''],
  ['Garbage', '', '', ''],
  ['Internet', '', '', ''],
  ['Cable/Satellite', '', '', '']
];

const contactSheet = XLSX.utils.aoa_to_sheet(contactData);
contactSheet['!cols'] = [{ wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }];
XLSX.utils.book_append_sheet(workbook, contactSheet, 'Contact Info');

// Create an instructions sheet
const instructionsData = [
  ['HOME INSPECTION TOOL - INSTRUCTIONS', '', ''],
  ['', '', ''],
  ['Welcome to the Home Inspection Tool. This workbook contains multiple sheets to help you perform a comprehensive home inspection.', '', ''],
  ['', '', ''],
  ['SHEET DESCRIPTIONS', '', ''],
  ['', '', ''],
  ['Main Report', 'This is the primary sheet where you\'ll summarize your findings for each inspection area.', ''],
  ['Individual Area Sheets', 'There are separate sheets for detailed inspection of each area (Roof, Exterior, etc.)', ''],
  ['Photo Log', 'Use this sheet to catalog all photos taken during the inspection.', ''],
  ['Cost Estimates', 'Record estimated repair costs for identified issues.', ''],
  ['Maintenance', 'A checklist for regular home maintenance tasks.', ''],
  ['Contact Info', 'Record client information and important contacts.', ''],
  ['', '', ''],
  ['HOW TO USE THIS TOOL', '', ''],
  ['', '', ''],
  ['1. Begin by filling out the client and property information in the Main Report and Contact Info sheets.', '', ''],
  ['2. As you inspect each area of the home, complete the corresponding detailed sheet.', '', ''],
  ['3. Take photos of important findings and log them in the Photo Log sheet.', '', ''],
  ['4. After completing the detailed inspection, summarize your findings in the Main Report.', '', ''],
  ['5. For items needing repair, provide cost estimates in the Cost Estimates sheet.', '', ''],
  ['6. Review the Maintenance sheet with the client to establish a maintenance schedule.', '', ''],
  ['', '', ''],
  ['RATING SYSTEM', '', ''],
  ['', '', ''],
  ['1 - Immediate Attention Required: Items that pose a safety or significant damage risk and require immediate repair.', '', ''],
  ['2 - Repair/Replace Soon: Items that should be addressed in the near future (within 3-6 months).', '', ''],
  ['3 - Monitor/Maintenance Item: Issues to monitor or address through regular maintenance.', '', ''],
  ['4 - Normal Wear and Tear: Normal aging or wear consistent with the age of the home.', '', ''],
  ['5 - Good Condition: Items in good working order with no visible defects.', '', ''],
  ['N/A - Not Applicable: Items that don\'t apply to this property.', '', ''],
  ['', '', ''],
  ['TIPS FOR EFFECTIVE INSPECTIONS', '', ''],
  ['', '', ''],
  ['- Always take plenty of photos, especially of deficiencies.', '', ''],
  ['- Be thorough and methodical; inspect each area completely before moving to the next.', '', ''],
  ['- Use consistent terminology throughout your report.', '', ''],
  ['- Focus on facts rather than opinions.', '', ''],
  ['- Include both visible defects and potential concerns.', '', ''],
  ['- Note limited access areas or items that couldn\'t be fully inspected.', '', ''],
  ['- Use the Photo Log to create a clear reference system for your findings.', '', '']
];

const instructionsSheet = XLSX.utils.aoa_to_sheet(instructionsData);
instructionsSheet['!cols'] = [{ wch: 25 }, { wch: 70 }, { wch: 15 }];
XLSX.utils.book_append_sheet(workbook, instructionsSheet, 'Instructions');

// Create a sample filled report to use as reference
const sampleData = [
  ['SAMPLE INSPECTION ENTRY', '', '', ''],
  ['', '', '', ''],
  ['This sheet provides examples of how to fill out the inspection sheets properly.', '', '', ''],
  ['', '', '', ''],
  ['MAIN REPORT EXAMPLE:', '', '', ''],
  ['Area', 'Rating', 'Deficiency', 'Recommendation', 'Photo Reference'],
  ['Roof', '2', 'Missing shingles on south slope', 'Replace missing shingles to prevent water intrusion', 'Photos #3-5'],
  ['', '', '', '', ''],
  ['DETAILED SHEET EXAMPLE (ROOF):', '', '', ''],
  ['Item', 'Condition (1-5)', 'Notes', 'Photo Reference'],
  ['Roof Covering', '2', 'Asphalt shingles with approximately 5-7 years of remaining life. Missing shingles observed on south slope.', 'Photos #3-5'],
  ['Signs of Leaking', '4', 'No active leaks observed. Previous stain noted in attic - appears dry.', 'Photo #12'],
  ['', '', '', ''],
  ['PHOTO LOG EXAMPLE:', '', '', ''],
  ['Photo #', 'Location', 'Description', 'Date Taken'],
  ['3', 'Roof - South Slope', 'Missing shingles near chimney flashing', '4/15/2025'],
  ['4', 'Roof - South Slope', 'Close-up of damaged shingle area', '4/15/2025'],
  ['', '', '', ''],
  ['COST ESTIMATE EXAMPLE:', '', '', ''],
  ['Item', 'Priority (1-5)', 'Estimated Cost (Low)', 'Estimated Cost (High)', 'Notes'],
  ['Roof repair - replace missing shingles', '2', '$350', '$500', 'Recommend licensed roofer for proper installation'],
  ['', '', '', '', ''],
  ['GENERAL TIPS:', '', '', ''],
  ['1. Be specific about locations (e.g., "south slope" rather than just "roof").', '', '', ''],
  ['2. Include measurements where applicable.', '', '', ''],
  ['3. Note both the condition and the material/type of items being inspected.', '', '', ''],
  ['4. Cross-reference photos with findings for clarity.', '', '', ''],
  ['5. Provide specific recommendations - "repair" is too vague; "replace missing shingles" is better.', '', '', '']
];

const sampleSheet = XLSX.utils.aoa_to_sheet(sampleData);
sampleSheet['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 40 }, { wch: 40 }, { wch: 15 }];
XLSX.utils.book_append_sheet(workbook, sampleSheet, 'Sample Entries');

// Write to binary string
const binaryXLSX = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

// Convert to base64 for download
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xFF;
  }
  return buf;
}

// The base64 string can be used to download the file
// In a real application, you would trigger a download with this data
const base64 = btoa(String.fromCharCode.apply(null, new Uint8Array(s2ab(binaryXLSX))));

console.log('Excel file generated successfully!');
// In a real environment, you would use:
// const a = document.createElement('a');
// a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64;
// a.download = 'Home-Inspection-Tool.xlsx';
// a.click();
