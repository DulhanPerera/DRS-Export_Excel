import logging
import sys
from .table_utils import create_table

logger = logging.getLogger('excel_data_writer')

def create_drc_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Debt Recovery Company (DRC) table in the worksheet.
    """
    try:
        logger.info("Creating DRC table...")
        headers = [
            "Order ID", "DRC ID", "DRC Name", "Created DTM", "DRC Status", "Status DTM",
            "Expire DTM", "Case Removal Remark", "Removed By", "Removed DTM",
            "DRC Selection Logic", "Case Distribution Batch ID"
        ]
        
        # Prepare data for the table
        drc_data = case_data.get("drc", [])
        data = [
            [
                drc.get("order_id"),
                drc.get("drc_id"),
                drc.get("drc_name"),
                drc.get("created_dtm"),
                drc.get("drc_status"),
                drc.get("status_dtm"),
                drc.get("expire_dtm"),
                drc.get("case_removal_remark"),
                drc.get("removed_by"),
                drc.get("removed_dtm"),
                drc.get("drc_selection_logic"),
                drc.get("case_distribution_batch_id")
            ]
            for drc in drc_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Debt Recovery Company (DRC)", headers, data, styles)
        
        logger.info("DRC table created successfully.")
        return next_row
    except Exception as failed_drc_table_creation:
        logger.error(f"Failed to create DRC table: {failed_drc_table_creation}")
        sys.exit(1)

def create_ro_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Recovery Officer (RO) table in the worksheet.
    """
    try:
        logger.info("Creating Recovery Officer (RO) table...")
        headers = [
            "RO ID", "Assigned DTM", "Assigned By", "Removed DTM", "Case Removal Remark", "DRC ID", "DRC Name"
        ]
        
        # Prepare data for the table
        drc_data = case_data.get("drc", [])
        data = []
        
        # Iterate through each DRC object to extract RO data
        for drc in drc_data:
            ro_data = drc.get("recovery_officers", [])
            for ro in ro_data:
                data.append([
                    ro.get("ro_id"),
                    ro.get("assigned_dtm"),
                    ro.get("assigned_by"),
                    ro.get("removed_dtm"),
                    ro.get("case_removal_remark"),
                    drc.get("drc_id"),  # DRC ID from the parent DRC object
                    drc.get("drc_name")  # DRC Name from the parent DRC object
                ])
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Recovery Officer (RO)", headers, data, styles)
        
        logger.info("Recovery Officer (RO) table created successfully.")
        return next_row
    except Exception as failed_ro_table_creation:
        logger.error(f"Failed to create RO table: {failed_ro_table_creation}")
        sys.exit(1)