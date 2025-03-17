from datetime import datetime  # Module for handling date and time operations
import logging  # Module for logging errors and debugging information
import os  # Module for interacting with the operating system
import sys  # Module for system-specific parameters and functions
from openpyxl import Workbook  # Library for working with Excel files
from .data_fetcher import get_settlement_data, get_settlement_plan_data 
from .case_contact_tables import create_case_details_table, create_contact_details_table
from .settlements_remarks import create_remarks_table, create_settlement_table, create_settlement_plan_table
from .approve_table import create_approve_table, create_case_status_table, create_abnormal_stop_table
from .drc_ro_tables import create_drc_table, create_ro_table
from .payments_table import create_payments_table
from .ro_tables import create_ro_negotiations_table, create_ro_requests_table

logger = logging.getLogger('excel_data_writer')

def create_all_tables(workBook, case_data, db, styles):
    """
    Create all tables in a structured format.
    
    Args:
        workBook (Workbook): An openpyxl Workbook object where the sheet will be created.
        case_data (dict): Dictionary containing case-related data.
        db: Database connection object for fetching additional data.
        styles (dict): Predefined styles for formatting.
    
    Returns:
        worksheet: The worksheet object containing the generated tables.
    
    Outputs:
        - Logs success or failure messages while creating the sheet.
    """
    try:
        logger.info("Creating Case Details sheet...")
        worksheet = workBook.active
        worksheet.title = "Case Details"
        
        # Define starting row and column for the first table
        x_pointer, y_pointer = 1, 1
        
        # Create the Case Details table
        next_row = create_case_details_table(worksheet, case_data, x_pointer, y_pointer, db, styles)
        
        # Add a two-row gap between the tables
        gap_row = next_row + 1
        
        # Create the Contact Details table
        next_row = create_contact_details_table(worksheet, case_data, gap_row, y_pointer, styles)
        
        # Add a two-row gap between the Contact Info and Remarks tables
        gap_row = next_row + 1
        
        # Create the Remarks table
        next_row = create_remarks_table(worksheet, case_data, gap_row, y_pointer, styles)
        
        # Add a two-row gap between the Remarks and Settlement tables
        gap_row = next_row + 1
        
        # Retrieve settlement data for the case
        case_id = case_data.get("case_id")
        settlements = get_settlement_data(db, case_id)
        
        # Create the Settlement table if settlement data exists
        if settlements:
            next_row = create_settlement_table(worksheet, settlements, gap_row, y_pointer, styles)
            gap_row = next_row + 1  # Add a gap after the Settlement table
        
        # Retrieve settlement plan data for the case
        settlement_plans = get_settlement_plan_data(db, case_id)
        
        # Create the Settlement Plan table if settlement plan data exists
        if settlement_plans:
            next_row = create_settlement_plan_table(worksheet, settlement_plans, gap_row, y_pointer, styles)
            gap_row = next_row + 1  # Add a gap after the Settlement Plan table
        
        # Create the Approve table
        next_row = create_approve_table(worksheet, case_data, gap_row, y_pointer, styles)
        gap_row = next_row + 1  # Add a gap after the Approve table
        
        # Create the Case Status table
        next_row = create_case_status_table(worksheet, case_data, gap_row, y_pointer, styles)
        gap_row = next_row + 1  # Add a gap after the Case Status table
        
        # Create the Abnormal Stop table
        next_row = create_abnormal_stop_table(worksheet, case_data, gap_row, y_pointer, styles)
        gap_row = next_row + 1  # Add a gap after the Abnormal Stop table
        
        # Create the DRC table
        next_row = create_drc_table(worksheet, case_data, gap_row, y_pointer, styles)
        gap_row = next_row + 1  # Add a gap after the DRC table
        
        # Create the Recovery Officer (RO) table
        next_row = create_ro_table(worksheet, case_data, gap_row, y_pointer, styles)
        gap_row = next_row + 1  # Add a gap after the RO table
        
        # Create the Payments table
        next_row = create_payments_table(worksheet, db, case_id, gap_row, y_pointer, styles)
        gap_row = next_row + 1  # Add a gap after the Payments table
        
        # Create the Recovery Officer Negotiations table
        next_row = create_ro_negotiations_table(worksheet, case_data, gap_row, y_pointer, styles)
        gap_row = next_row + 1  # Add a gap after the RO Negotiations table
        
        # Create the Recovery Officer Requests table
        next_row = create_ro_requests_table(worksheet, case_data, gap_row, y_pointer, styles)
        
        logger.info("Case Details sheet created successfully.")
        return worksheet
    except Exception as create_all_sheet_failed:
        logger.error(f"Failed to create all tables in sheet: {create_all_sheet_failed}")
        sys.exit(1)

def export_all_tables(db, incident_id, output_path, collection_name, styles):
    """
    Export case details from MongoDB to an Excel file.
    
    - Fetches case data based on `incident_id`.
    - Generates tables for case details, contacts, remarks, settlements, and settlement plans.
    - Saves the Excel file with a unique name to avoid overwriting.
    
    Args:
        db: Database connection object.
        incident_id (str): The incident ID to fetch case data.
        output_path (str): The directory to save the Excel file.
        collection_name (str): The MongoDB collection name.
        styles (dict): Predefined styles for formatting.

    Returns:
        None (Writes the Excel file to disk)
    """
    try:
        case_collection = db[collection_name]
        case_data = case_collection.find_one({"incident_id": incident_id})
        
        if not case_data:
            logger.error(f"No case details found for Incident ID: {incident_id}")
            sys.exit(1)
        
        # logger.info(f"Case data found: {case_data}")
        logger.info(f"Case data found!, Exporting case details for Incident ID: {incident_id}")
        
        workBook = Workbook()
        worksheet = create_all_tables(workBook, case_data, db, styles)
        
        # Get the current date and time
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        
        # Create the base output file name with incident_id and current date/time
        file_name = f"Case_Details_{incident_id}_{current_time}.xlsx"
        
        # Ensure the output path ends with the correct file name
        if not output_path.endswith('.xlsx'):
            output_path = os.path.join(output_path, file_name)
        else:
            output_path = os.path.join(output_path, file_name)
        
        # Check if the file exists and append a number to make the filename unique
        if os.path.exists(output_path):
            base, extension = os.path.splitext(output_path)
            counter = 1
            while os.path.exists(f"{base}_{counter}{extension}"):
                counter += 1
            output_path = f"{base}_{counter}{extension}"
        
        # Create directories if they don't exist
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Save the workbook
        try:
            workBook.save(output_path)
            logger.info(f"Case details exported to {output_path}")
        except Exception as failed_export:
            logger.error(f"Failed to save Excel file: {failed_export}")
            sys.exit(1)
    except Exception as failed_tables_all_export:
        logger.error(f"Failed to export all tables: {failed_tables_all_export}")
        sys.exit(1)