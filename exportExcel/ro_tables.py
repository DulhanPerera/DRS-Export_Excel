import logging
import sys
from .table_utils import create_table

logger = logging.getLogger('excel_data_writer')

def create_ro_negotiations_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Recovery Officer Negotiations table in the worksheet.
    """
    try:
        logger.info("Creating Recovery Officer Negotiations table...")
        headers = [
            "DRC ID", "RO ID", "Created DTM", "Field Reason ID", "Field Reason", "Remark"
        ]
        
        # Prepare data for the table
        ro_negotiations_data = case_data.get("ro_negotiation", [])
        data = [
            [
                negotiation.get("drc_id"),
                negotiation.get("ro_id"),
                negotiation.get("created_dtm"),
                negotiation.get("field_reason_id"),
                negotiation.get("field_reason"),
                negotiation.get("remark")
            ]
            for negotiation in ro_negotiations_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Recovery Officer Negotiations", headers, data, styles)
        
        logger.info("Recovery Officer Negotiations table created successfully.")
        return next_row
    except Exception as failed_ro_negotiations_table_creation:
        logger.error(f"Failed to create Recovery Officer Negotiations table: {failed_ro_negotiations_table_creation}")
        sys.exit(1)

def create_ro_requests_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Recovery Officer Requests table in the worksheet.
    """
    try:
        logger.info("Creating Recovery Officer Requests table...")
        headers = [
            "DRC ID", "RO ID", "Created DTM", "RO Request ID", "RO Request", "ToDo On", "Completed On"
        ]
        
        # Prepare data for the table
        ro_requests_data = case_data.get("ro_requests", [])
        data = [
            [
                request.get("drc_id"),
                request.get("ro_id"),
                request.get("created_dtm"),
                request.get("ro_request_id"),
                request.get("ro_request"),
                request.get("todo_on"),
                request.get("completed_on")
            ]
            for request in ro_requests_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Recovery Officer Requests", headers, data, styles)
        
        logger.info("Recovery Officer Requests table created successfully.")
        return next_row
    except Exception as failed_ro_requests_table_creation:
        logger.error(f"Failed to create Recovery Officer Requests table: {failed_ro_requests_table_creation}")
        sys.exit(1)