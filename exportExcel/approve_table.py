import logging
import sys
from .table_utils import create_table

logger = logging.getLogger('excel_data_writer')

def create_approve_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Approve table in the worksheet.
    """
    try:
        logger.info("Creating Approve table...")
        headers = ["Approved Process", "Approved By", "Approved On", "Remark"]
        
        # Prepare data for the table
        approve_data = case_data.get("approve", [])
        data = [
            [
                approve.get("approved_process"),
                approve.get("approved_by"),
                approve.get("approved_on"),
                approve.get("remark")
            ]
            for approve in approve_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Approve Details", headers, data, styles)
        
        logger.info("Approve table created successfully.")
        return next_row
    except Exception as failed_approve_table_creation:
        logger.error(f"Failed to create Approve table: {failed_approve_table_creation}")
        sys.exit(1)

def create_case_status_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Case Status table in the worksheet.
    """
    try:
        logger.info("Creating Case Status table...")
        headers = ["Case Status", "Status Reason", "Created DTM", "Created By", "Notified DTM", "Expire DTM"]
        
        # Prepare data for the table
        case_status_data = case_data.get("case_status", [])
        data = [
            [
                status.get("case_status"),
                status.get("status_reason"),
                status.get("created_dtm"),
                status.get("created_by"),
                status.get("notified_dtm"),
                status.get("expire_dtm")
            ]
            for status in case_status_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Case Status", headers, data, styles)
        
        logger.info("Case Status table created successfully.")
        return next_row
    except Exception as failed_case_status_table_creation:
        logger.error(f"Failed to create Case Status table: {failed_case_status_table_creation}")
        sys.exit(1)

def create_abnormal_stop_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Abnormal Stop table in the worksheet.
    """
    try:
        logger.info("Creating Abnormal Stop table...")
        headers = ["Remark", "Done By", "Done On", "Action"]
        
        # Prepare data for the table
        abnormal_stop_data = case_data.get("abnormal_stop", [])
        data = [
            [
                stop.get("remark"),
                stop.get("done_by"),
                stop.get("done_on"),
                stop.get("action")
            ]
            for stop in abnormal_stop_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Abnormal Stop", headers, data, styles)
        
        logger.info("Abnormal Stop table created successfully.")
        return next_row
    except Exception as failed_abnormal_stop_table_creation:
        logger.error(f"Failed to create Abnormal Stop table: {failed_abnormal_stop_table_creation}")
        sys.exit(1)