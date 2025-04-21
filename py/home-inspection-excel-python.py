import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

def create_home_inspection_excel(filename="Home_Inspection_Tool.xlsx"):
    """
    Creates a comprehensive home inspection Excel workbook with multiple sheets
    for different areas of inspection, ratings, and data collection.
    """
    # Create a Pandas Excel writer using openpyxl as the engine
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    
    # Create main report sheet
    create_main_report_sheet(writer)
    
    # Create individual inspection sheets
    create_roof_sheet(writer)
    create_exterior_sheet(writer)
    create_foundation_sheet(writer)
    create_plumbing_sheet(writer)
    create_electrical_sheet(writer)
    create_hvac_sheet(writer)
    create_interior_sheet(writer)
    create_attic_sheet(writer)
    
    # Create support sheets
    create_photo_log_sheet(writer)
    create_cost_estimate_sheet(writer)
    create_maintenance_sheet(writer)
    create_contact_sheet(writer)
    create_instructions_sheet(writer)
    create_sample_sheet(writer)
    
    # Apply formatting and save workbook
    workbook = writer.book
    
    # Set Instructions sheet as the first sheet to open
    workbook.active = workbook['Instructions']
    
    # Save the workbook
    writer.close()
    
    print(f"Excel file '{filename}' created successfully!")

def create_main_report_sheet(writer):
    """Create the main inspection report summary sheet"""
    
    # Create data for main sheet
    main_data = [
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
    ]
    
    # Add rows for each inspection area
    inspection_areas = [
        'Roof', 'Exterior', 'Foundation', 'Basement', 'Crawlspace', 
        'Plumbing', 'Electrical', 'HVAC', 'Interior', 'Attic', 
        'Insulation', 'Ventilation', 'Kitchen', 'Bathrooms', 'Garage'
    ]
    
    for area in inspection_areas:
        main_data.append([area, '', '', '', ''])
    
    # Add final notes section
    main_data.extend([
        ['', '', '', '', ''],
        ['ADDITIONAL NOTES', '', '', '', ''],
        ['', '', '', '', ''],
        ['', '', '', '', '']
    ])
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(main_data)
    df.to_excel(writer, sheet_name='Main Report', index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets['Main Report']
    
    # Set column widths
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 10
    sheet.column_dimensions['C'].width = 40
    sheet.column_dimensions['D'].width = 40
    sheet.column_dimensions['E'].width = 15
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A17'].font = Font(bold=True)
    for col in ['A', 'B', 'C', 'D', 'E']:
        sheet[f'{col}18'].font = Font(bold=True)
    
    # Add data validation for ratings
    dv = DataValidation(type="list", formula1='"1,2,3,4,5,N/A"', allow_blank=True)
    sheet.add_data_validation(dv)
    
    # Apply validation to rating column
    for row in range(19, 19 + len(inspection_areas)):
        dv.add(f'B{row}')

def create_inspection_sheet(writer, area_name, check_items):
    """Helper function to create detailed inspection sheets"""
    
    # Create sheet data
    sheet_data = [
        [f'{area_name.upper()} INSPECTION DETAILS', '', '', ''],
        ['', '', '', ''],
        ['Item', 'Condition (1-5)', 'Notes', 'Photo Reference'],
    ]
    
    # Add rows for check items
    for item in check_items:
        sheet_data.append([item, '', '', ''])
    
    # Add notes section
    sheet_data.extend([
        ['', '', '', ''],
        ['ADDITIONAL NOTES:', '', '', ''],
        ['', '', '', '']
    ])
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(sheet_data)
    df.to_excel(writer, sheet_name=area_name, index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets[area_name]
    
    # Set column widths
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['C'].width = 50
    sheet.column_dimensions['D'].width = 15
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=14)
    for col in ['A', 'B', 'C', 'D']:
        sheet[f'{col}3'].font = Font(bold=True)
    
    # Add data validation for condition
    dv = DataValidation(type="list", formula1='"1,2,3,4,5,N/A"', allow_blank=True)
    sheet.add_data_validation(dv)
    
    # Apply validation to condition column
    for row in range(4, 4 + len(check_items)):
        dv.add(f'B{row}')

def create_roof_sheet(writer):
    """Create the roof inspection sheet"""
    roof_items = [
        'Roof Covering', 'Roof Flashing', 'Roof Drainage', 'Skylights', 
        'Chimneys', 'Roof Penetrations', 'Signs of Leaking', 'Roof Ventilation',
        'Roof Structure', 'Estimated Remaining Life'
    ]
    create_inspection_sheet(writer, 'Roof', roof_items)

def create_exterior_sheet(writer):
    """Create the exterior inspection sheet"""
    exterior_items = [
        'Siding/Cladding', 'Exterior Doors', 'Windows', 'Trim', 
        'Eaves/Soffits/Fascia', 'Exterior Lighting', 'Walkways', 'Driveway', 
        'Steps/Stoops', 'Porches/Decks', 'Railings', 'Grading/Drainage',
        'Vegetation', 'Retaining Walls', 'Fences/Gates'
    ]
    create_inspection_sheet(writer, 'Exterior', exterior_items)

def create_foundation_sheet(writer):
    """Create the foundation inspection sheet"""
    foundation_items = [
        'Foundation Walls', 'Visible Structural Components', 'Signs of Water Penetration',
        'Cracks', 'Settlement', 'Foundation Type', 'Anchor Bolts', 'Floor Framing',
        'Wall Framing', 'Support Posts/Columns', 'Support Beams'
    ]
    create_inspection_sheet(writer, 'Foundation', foundation_items)

def create_plumbing_sheet(writer):
    """Create the plumbing inspection sheet"""
    plumbing_items = [
        'Water Supply Lines', 'Drain/Waste/Vent Pipes', 'Main Water Shut-off',
        'Water Pressure', 'Water Heater', 'Toilets', 'Sinks', 'Tubs/Showers',
        'Faucets', 'Visible Leaks', 'Sump Pump', 'Sewage Ejector Pump',
        'Gas Lines', 'Main Gas Shut-off'
    ]
    create_inspection_sheet(writer, 'Plumbing', plumbing_items)

def create_electrical_sheet(writer):
    """Create the electrical inspection sheet"""
    electrical_items = [
        'Service Entrance', 'Main Panel', 'Circuit Breakers/Fuses', 'Branch Wiring',
        'Grounding', 'GFCI Protection', 'AFCI Protection', 'Outlets', 'Switches',
        'Light Fixtures', 'Ceiling Fans', 'Smoke Detectors', 'Carbon Monoxide Detectors'
    ]
    create_inspection_sheet(writer, 'Electrical', electrical_items)

def create_hvac_sheet(writer):
    """Create the HVAC inspection sheet"""
    hvac_items = [
        'Heating System Type', 'Heating System Age', 'Heating Operation', 'Cooling System Type',
        'Cooling System Age', 'Cooling Operation', 'Distribution System', 'Thermostat',
        'Filters', 'Humidifier', 'Dehumidifier', 'Ductwork', 'Ventilation'
    ]
    create_inspection_sheet(writer, 'HVAC', hvac_items)

def create_interior_sheet(writer):
    """Create the interior inspection sheet"""
    interior_items = [
        'Floors', 'Walls', 'Ceilings', 'Windows', 'Interior Doors', 'Stairs',
        'Railings', 'Countertops', 'Cabinets', 'Appliances', 'Evidence of Pests',
        'Evidence of Water Damage', 'Evidence of Mold'
    ]
    create_inspection_sheet(writer, 'Interior', interior_items)

def create_attic_sheet(writer):
    """Create the attic inspection sheet"""
    attic_items = [
        'Access', 'Insulation Type', 'Insulation Depth', 'Ventilation', 
        'Visible Electrical', 'Visible Plumbing', 'Visible Framing', 
        'Signs of Leaking', 'Signs of Pests', 'Exhaust Venting'
    ]
    create_inspection_sheet(writer, 'Attic', attic_items)

def create_photo_log_sheet(writer):
    """Create the photo log sheet"""
    # Create photo log data
    photo_log_data = [
        ['PHOTO LOG', '', '', ''],
        ['', '', '', ''],
        ['Photo #', 'Location', 'Description', 'Date Taken'],
    ]
    
    # Add 20 empty rows for photos
    for i in range(1, 21):
        photo_log_data.append([str(i), '', '', ''])
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(photo_log_data)
    df.to_excel(writer, sheet_name='Photo Log', index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets['Photo Log']
    
    # Set column widths
    sheet.column_dimensions['A'].width = 10
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 50
    sheet.column_dimensions['D'].width = 15
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=14)
    for col in ['A', 'B', 'C', 'D']:
        sheet[f'{col}3'].font = Font(bold=True)

def create_cost_estimate_sheet(writer):
    """Create the cost estimate sheet"""
    # Create cost estimate data
    cost_data = [
        ['REPAIR COST ESTIMATES', '', '', '', ''],
        ['', '', '', '', ''],
        ['Item', 'Priority (1-5)', 'Estimated Cost (Low)', 'Estimated Cost (High)', 'Notes'],
    ]
    
    # Add 15 empty rows
    for i in range(1, 16):
        cost_data.append(['', '', '', '', ''])
    
    # Add totals row
    cost_data.append(['TOTALS', '', '=SUM(C4:C18)', '=SUM(D4:D18)', ''])
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(cost_data)
    df.to_excel(writer, sheet_name='Cost Estimates', index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets['Cost Estimates']
    
    # Set column widths
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 30
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=14)
    for col in ['A', 'B', 'C', 'D', 'E']:
        sheet[f'{col}3'].font = Font(bold=True)
    
    # Format totals row
    sheet['A19'].font = Font(bold=True)
    
    # Add data validation for priority
    dv = DataValidation(type="list", formula1='"1,2,3,4,5"', allow_blank=True)
    sheet.add_data_validation(dv)
    
    # Apply validation to priority column
    for row in range(4, 19):
        dv.add(f'B{row}')

def create_maintenance_sheet(writer):
    """Create the home maintenance checklist"""
    # Create maintenance data
    maintenance_data = [
        ['HOME MAINTENANCE CHECKLIST', '', '', ''],
        ['', '', '', ''],
        ['Task', 'Frequency', 'Last Completed', 'Next Due'],
        
        # Monthly tasks
        ['MONTHLY MAINTENANCE', '', '', ''],
        ['Test smoke/CO detectors', 'Monthly', '', ''],
        ['Check HVAC filters', 'Monthly', '', ''],
        ['Check water softener', 'Monthly', '', ''],
        ['Clean range hood filters', 'Monthly', '', ''],
        ['Check for plumbing leaks', 'Monthly', '', ''],
        
        # Quarterly tasks
        ['QUARTERLY MAINTENANCE', '', '', ''],
        ['Test GFCIs', 'Quarterly', '', ''],
        ['Run water in unused fixtures', 'Quarterly', '', ''],
        ['Check water heater', 'Quarterly', '', ''],
        ['Check garage door operation', 'Quarterly', '', ''],
        
        # Biannual tasks
        ['BIANNUAL MAINTENANCE', '', '', ''],
        ['Service HVAC systems', 'Biannual', '', ''],
        ['Check fire extinguishers', 'Biannual', '', ''],
        ['Clean gutters', 'Biannual', '', ''],
        ['Check exterior drainage', 'Biannual', '', ''],
        
        # Annual tasks
        ['ANNUAL MAINTENANCE', '', '', ''],
        ['Inspect roof', 'Annual', '', ''],
        ['Clean chimney', 'Annual', '', ''],
        ['Inspect attic', 'Annual', '', ''],
        ['Check exterior paint/siding', 'Annual', '', ''],
        ['Check foundation', 'Annual', '', ''],
        ['Flush water heater', 'Annual', '', ''],
        ['Check weatherstripping', 'Annual', '', ''],
        ['Check caulking around showers/tubs', 'Annual', '', '']
    ]
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(maintenance_data)
    df.to_excel(writer, sheet_name='Maintenance', index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets['Maintenance']
    
    # Set column widths
    sheet.column_dimensions['A'].width = 35
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['C'].width = 15
    sheet.column_dimensions['D'].width = 15
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=14)
    for col in ['A', 'B', 'C', 'D']:
        sheet[f'{col}3'].font = Font(bold=True)
    
    # Format section headers
    for row in [4, 10, 15, 20]:
        sheet[f'A{row}'].font = Font(bold=True)

def create_contact_sheet(writer):
    """Create the client contact information sheet"""
    # Create contact data
    contact_data = [
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
    ]
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(contact_data)
    df.to_excel(writer, sheet_name='Contact Info', index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets['Contact Info']
    
    # Set column widths
    sheet.column_dimensions['A'].width = 25
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 25
    sheet.column_dimensions['D'].width = 25
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A8'].font = Font(bold=True, size=14)
    sheet['A17'].font = Font(bold=True, size=14)
    sheet['A31'].font = Font(bold=True, size=14)
    
    # Format table headers
    for col in ['A', 'B', 'C', 'D']:
        sheet[f'{col}19'].font = Font(bold=True)
        sheet[f'{col}33'].font = Font(bold=True)

def create_instructions_sheet(writer):
    """Create the instructions sheet"""
    # Create instructions data
    instructions_data = [
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
    ]
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(instructions_data)
    df.to_excel(writer, sheet_name='Instructions', index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets['Instructions']
    
    # Set column widths
    sheet.column_dimensions['A'].width = 25
    sheet.column_dimensions['B'].width = 70
    sheet.column_dimensions['C'].width = 15
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=16)
    sheet['A5'].font = Font(bold=True, size=12)
    sheet['A14'].font = Font(bold=True, size=12)
    sheet['A23'].font = Font(bold=True, size=12)
    sheet['A32'].font = Font(bold=True, size=12)
    
    # Format section items
    for row in range(7, 13):
        sheet[f'A{row}'].font = Font(bold=True)

def create_sample_sheet(writer):
    """Create a sample filled report to use as reference"""
    # Create sample data
    sample_data = [
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
    ]
    
    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(sample_data)
    df.to_excel(writer, sheet_name='Sample Entries', index=False, header=False)
    
    # Get the sheet and apply formatting
    sheet = writer.sheets['Sample Entries']
    
    # Set column widths
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['C'].width = 40
    sheet.column_dimensions['D'].width = 40
    sheet.column_dimensions['E'].width = 15
    
    # Format headers
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A5'].font = Font(bold=True)
    sheet['A9'].font = Font(bold=True)
    sheet['A14'].font = Font(bold=True)
    sheet['A19'].font = Font(bold=True)
    sheet['A23'].font = Font(bold=True)
    
    # Format subheaders
    for row in [6, 10, 15, 20]:
        for col in range(65, 70):  # A-E
            sheet[f'{chr(col)}{row}'].font = Font(bold=True)

# Execute the script
if __name__ == "__main__":
    create_home_inspection_excel()
