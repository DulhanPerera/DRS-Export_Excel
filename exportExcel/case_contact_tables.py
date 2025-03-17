import logging
import sys
from .table_utils import create_table  # Import from table_utils
from .data_fetcher import get_arrears_band_value
from .excel_styles import format_with_thousand_separator

logger = logging.getLogger('excel_data_writer')

def create_case_details_table(worksheet, case_data, x_pointer, y_pointer, db, styles):
    """
    Create the Case Details table in the worksheet with a horizontal layout and a start_process header.
    """
    try:
        logger.info("Creating Case Details table...")
        
        # Define the start_process header
        main_header = "Case Details"
        
        # Write the start_process header
        worksheet.merge_cells(start_row=x_pointer, start_column=y_pointer, end_row=x_pointer, end_column=y_pointer + 1)
        main_header_cell = worksheet.cell(row=x_pointer, column=y_pointer, value=main_header)
        main_header_cell.font = styles["header_font"]
        main_header_cell.fill = styles["main_header_fill"]
        main_header_cell.border = styles["cell_border"]
        main_header_cell.alignment = styles["main_header_alignment"]
        
        # Move to the next row for the table
        x_pointer += 1
        
        # Define headers for the table
        headers = [
            "Case ID", "Incident ID", "Account No.", "Customer Ref", "Area",
            "BSS Arrears Amount", "Current Arrears Amount", "Action type", "Filtered reason",
            "Last Payment Date", "Last BSS Reading Date", "Commission", "Case Current Status",
            "Current Arrears band", "DRC Commission Rule", "Created dtm", "Implemented dtm",
            "RTOM", "Monitor months"
        ]
        
        # Map MongoDB data to headers
        data_mapping = {
            "Case ID": case_data.get("case_id"),
            "Incident ID": case_data.get("incident_id"),
            "Account No.": case_data.get("account_no"),
            "Customer Ref": case_data.get("customer_ref"),
            "Area": case_data.get("area"),
            "BSS Arrears Amount": format_with_thousand_separator(case_data.get("bss_arrears_amount")),
            "Current Arrears Amount": format_with_thousand_separator(case_data.get("current_arrears_amount")),
            "Action type": case_data.get("action_type"),
            "Filtered reason": case_data.get("filtered_reason"),
            "Last Payment Date": case_data.get("last_payment_date"),
            "Last BSS Reading Date": case_data.get("last_bss_reading_date"),
            "Commission": format_with_thousand_separator(case_data.get("commission")),
            "Case Current Status": case_data.get("case_current_status"),
            "Current Arrears band": case_data.get("current_arrears_band"),
            "DRC Commission Rule": case_data.get("drc_commision_rule"),
            "Created dtm": case_data.get("created_dtm"),
            "Implemented dtm": case_data.get("implemented_dtm"),
            "RTOM": case_data.get("rtom"),
            "Monitor months": case_data.get("monitor_months")
        }
        
        # Retrieve arrears band value
        current_arrears_band = case_data.get("current_arrears_band")
        if current_arrears_band:
            arrears_band_value = get_arrears_band_value(db, current_arrears_band)
            if arrears_band_value:
                data_mapping["Current Arrears band"] = arrears_band_value
            else:
                logger.warning(f"No value found for arrears band: {current_arrears_band}")
        
        # Write headers and data horizontally
        for index, header in enumerate(headers):
            # Write header in the first column
            worksheet.cell(row=x_pointer + index, column=y_pointer, value=header).font = styles["header_font"]
            worksheet.cell(row=x_pointer + index, column=y_pointer).fill = styles["sub_header_fill"]
            worksheet.cell(row=x_pointer + index, column=y_pointer).border = styles["cell_border"]
            worksheet.cell(row=x_pointer + index, column=y_pointer).alignment = styles["sub_header_alignment"]
            
            # Write corresponding data in the next column
            value = data_mapping.get(header)
            if isinstance(value, (list, dict)):
                value = str(value)
            worksheet.cell(row=x_pointer + index, column=y_pointer + 1, value=value).border = styles["cell_border"]
            if header in ["Case ID", "Incident ID"]:
                worksheet.cell(row=x_pointer + index, column=y_pointer + 1).font = styles["bold_font"]
        
        # Adjust column widths
        for col in range(y_pointer, y_pointer + 2):  # Only two columns (headers and data)
            max_length = 0
            column_letter = chr(64 + col)
            for row in range(x_pointer, x_pointer + len(headers)):
                cell_value = worksheet.cell(row=row, column=col).value
                if cell_value and len(str(cell_value)) > max_length:
                    max_length = len(str(cell_value))
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        logger.info("Case Details table created successfully.")
        return x_pointer + len(headers) + 1
    except Exception as failed_case_details_table_creation:
        logger.error(f"Failed to create Case Details table: {failed_case_details_table_creation}")
        sys.exit(1)

def create_contact_details_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Contact Details table in the worksheet.
    """
    try:
        logger.info("Creating Contact Details table...")
        contacts_headers = ["Mobile", "Email", "Home Phone", "Address"]
        
        # Prepare data for the table
        contacts = case_data.get("contact", [])
        data = [[contact.get("mob"), contact.get("email"), contact.get("lan"), contact.get("address")] for contact in contacts]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Contact Info", contacts_headers, data, styles)
        
        logger.info("Contact Details table created successfully.")
        return next_row
    except Exception as failed_contact_details_table_creation:
        logger.error(f"Failed to create Contact Details table: {failed_contact_details_table_creation}")
        sys.exit(1)