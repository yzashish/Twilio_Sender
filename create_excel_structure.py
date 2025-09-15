import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

def create_excel_file(filename="TwilioSender.xlsx"):
    """
    Creates an Excel workbook with the required structure for the Twilio Sender.
    """
    wb = openpyxl.Workbook()

    # --- Create Sheets ---
    # The default 'Sheet' is created, rename it to Dashboard
    ws_dashboard = wb.active
    ws_dashboard.title = "Dashboard"

    # Create other sheets
    ws_contacts = wb.create_sheet("Contacts")
    ws_templates = wb.create_sheet("Templates")
    ws_settings = wb.create_sheet("Settings")
    ws_log = wb.create_sheet("Log")

    # --- Move Dashboard to the front ---
    wb.move_sheet(ws_dashboard, -len(wb.sheetnames)+1)


    # --- Style Definitions ---
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    info_font = Font(italic=True, color="595959")
    cell_alignment = Alignment(vertical='top', wrap_text=True)
    title_font = Font(bold=True, size=18, color="4F81BD")

    # --- Populate Dashboard Sheet ---
    ws_dashboard.column_dimensions['A'].width = 25
    ws_dashboard.column_dimensions['B'].width = 50
    ws_dashboard.column_dimensions['C'].width = 15
    ws_dashboard.column_dimensions['D'].width = 50


    ws_dashboard['A1'] = "Twilio WhatsApp Messenger"
    ws_dashboard['A1'].font = title_font
    ws_dashboard.merge_cells('A1:D1')

    ws_dashboard['A3'] = "How to Use"
    ws_dashboard['A3'].font = header_font
    ws_dashboard['A3'].fill = header_fill
    ws_dashboard.merge_cells('A3:D3')

    instructions = """1. Go to the 'Settings' sheet and enter your Twilio Account SID, Auth Token, and 'From' number.
2. Go to the 'Contacts' sheet and add your contacts' names and numbers. Set 'Send?' to 'Yes' for those you want to message.
3. Go to the 'Templates' sheet to define your messages. Make sure each message has a unique 'Message UID'.
4. Come back to this Dashboard and enter the Message UID for your chosen message into the yellow box at C5.
5. Click the 'Send Messages' button (which will be added after embedding the macro) to send the messages."""

    ws_dashboard['A4'] = instructions
    ws_dashboard['A4'].alignment = Alignment(vertical='top', wrap_text=True)
    ws_dashboard['A4'].font = Font(size=12)
    ws_dashboard.merge_cells('A4:D4')
    ws_dashboard.row_dimensions[4].height = 100

    ws_dashboard['B5'] = "Active Message UID:"
    ws_dashboard['B5'].font = Font(bold=True)
    ws_dashboard['B5'].alignment = Alignment(horizontal='right')

    ws_dashboard['C5'] = "GOLD_HINDI_01" # Default value
    ws_dashboard['C5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # Yellow fill
    ws_dashboard['C5'].font = Font(bold=True)

    ws_dashboard['D5'] = "<- Enter the Message UID of the template you want to send."
    ws_dashboard['D5'].font = info_font

    ws_dashboard['B6'] = "Messages Per Second (MPS):"
    ws_dashboard['B6'].font = Font(bold=True)
    ws_dashboard['B6'].alignment = Alignment(horizontal='right')

    ws_dashboard['C6'] = 1 # Default value
    ws_dashboard['C6'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # Yellow fill
    ws_dashboard['C6'].font = Font(bold=True)

    ws_dashboard['D6'] = "<- Set your desired sending rate (e.g., 1)."
    ws_dashboard['D6'].font = info_font


    ws_dashboard['A8'] = "Button Placeholder"
    ws_dashboard['A7'].font = header_font
    ws_dashboard['A7'].fill = header_fill
    ws_dashboard.merge_cells('A7:D7')
    ws_dashboard['A8'] = "The 'Send Messages' button will be located in this area. You will need to add it manually after the macro is embedded. See User Guide for instructions."
    ws_dashboard['A8'].font = info_font
    ws_dashboard['A8'].alignment = Alignment(vertical='top', wrap_text=True)
    ws_dashboard.merge_cells('A8:D8')
    ws_dashboard.row_dimensions[8].height = 50


    # --- Populate Contacts Sheet ---
    # Define widths for first 4 columns, then 15 variable columns
    widths = [20, 20, 10, 30] + [15] * 15
    for i, width in enumerate(widths, 1):
        ws_contacts.column_dimensions[get_column_letter(i)].width = width

    headers = ["Name", "WhatsApp Number", "Send?", "Status"] + [f"Variable {{{{{i}}}}}" for i in range(1, 16)]
    for col_num, header in enumerate(headers, 1):
        cell = ws_contacts.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill

    # Sample data - expanded to have a few variable examples
    sample_contacts = [
        ("Test Contact 1", "919876543210", "Yes", "", "Value1A", "Value2A", "Value3A"),
        ("Test Contact 2", "14155238886", "No", "", "Value1B", "Value2B"),
        ("Test Contact 3", "447123456789", "Yes", "", "Value1C", "Value2C", "Value3C", "Value4C"),
    ]
    for row_data in sample_contacts:
        # Pad each row with empty strings to match the 15 variable columns
        padded_row = list(row_data) + [""] * (19 - len(row_data))
        ws_contacts.append(padded_row)

    # Add data validation for "Send?" column
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv.add('C2:C1048576')
    ws_contacts.add_data_validation(dv)


    # --- Populate Templates Sheet ---
    ws_templates.column_dimensions['A'].width = 25
    ws_templates.column_dimensions['B'].width = 80
    ws_templates.column_dimensions['C'].width = 40

    headers = ["Message UID", "Template Content", "Description"]
    for col_num, header in enumerate(headers, 1):
        cell = ws_templates.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill

    # Sample data
    sample_template_1_content = """नमस्ते! 🙏
आज का सोने का भाव:

22k: {{1}}
20k: {{2}}

{{3}}
{{4}}
सराफा बाज़ार, मेरठ शहर

शुभ दिन! ✨"""
    sample_template_1 = ("GOLD_HINDI_01", sample_template_1_content, "Gold Price Template (Hindi)")

    sample_template_2_content = "Hello {{1}}! This is a test message from the Twilio Excel Sender."
    sample_template_2 = ("GREETING_ENG_01", sample_template_2_content, "Simple English Greeting")

    for row_data in [sample_template_1, sample_template_2]:
        ws_templates.append(row_data)

    # Apply alignment to multi-line cells
    ws_templates['B2'].alignment = cell_alignment
    ws_templates.row_dimensions[2].height = 200
    ws_templates['B3'].alignment = cell_alignment
    ws_templates.row_dimensions[3].height = 30


    # --- Populate Log Sheet ---
    ws_log.column_dimensions['A'].width = 25
    ws_log.column_dimensions['B'].width = 25
    ws_log.column_dimensions['C'].width = 20
    ws_log.column_dimensions['D'].width = 20
    ws_log.column_dimensions['E'].width = 15
    ws_log.column_dimensions['F'].width = 50

    log_headers = ["Timestamp", "Contact Name", "Contact Number", "Message UID", "Status", "Twilio Response"]
    for col_num, header in enumerate(log_headers, 1):
        cell = ws_log.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill


    # --- Populate Settings Sheet ---
    ws_settings.column_dimensions['A'].width = 25
    ws_settings.column_dimensions['B'].width = 50

    headers = ["Setting", "Value"]
    for col_num, header in enumerate(headers, 1):
        cell = ws_settings.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill

    settings_data = [
        ("Twilio Account SID", "Enter your ACxxxxxxxx SID here"),
        ("Twilio Auth Token", "Enter your Auth Token here"),
        ("Twilio 'From' Number", "Enter your Twilio WhatsApp number here (e.g. +14155238886)"),
    ]
    for row_num, (setting, value) in enumerate(settings_data, 2):
        ws_settings.cell(row=row_num, column=1, value=setting).font = Font(bold=True)
        ws_settings.cell(row=row_num, column=2, value=value).font = info_font

    # Protect the sheet but leave value cells editable
    for cell_letter in ['B']:
        for i in range(2, 5):
            ws_settings[f'{cell_letter}{i}'].protection = Protection(locked=False)

    ws_settings.protection.sheet = True
    ws_settings.protection.password = 'twilio' # Set a password for protection

    # --- Save the workbook ---
    wb.save(filename)
    print(f"Workbook '{filename}' created successfully.")

if __name__ == "__main__":
    # This script is intended to be used by another script that will also embed the macro.
    # For standalone testing, you can run this.
    create_excel_file()
