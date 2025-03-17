import logging
import sys
from openpyxl import Workbook
from exportExcel.table_utils import create_table
from .data_fetcher import get_settlement_data, get_settlement_plan_data
from .excel_styles import format_with_thousand_separator

logger = logging.getLogger('excel_data_writer')

def create_remarks_table(worksheet, case_data, x_pointer, y_pointer, styles):
    """
    Create the Remarks table in the worksheet.
    """
    try:
        logger.info("Creating Remarks table...")
        headers = ["Remark", "Remark Added by", "Remark Added Date"]
        
        # Prepare data for the table
        remarks = case_data.get("remark", [])
        data = [[remark.get("remark"), remark.get("remark_added_by"), remark.get("remark_added_date")] for remark in remarks]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Remarks", headers, data, styles)
        
        logger.info("Remarks table created successfully.")
        return next_row
    except Exception as failed_remarks_table_creation:
        logger.error(f"Failed to create Remarks table: {failed_remarks_table_creation}")
        sys.exit(1)

def create_settlement_table(worksheet, settlements, x_pointer, y_pointer, styles):
    """
    Create the Settlement table in the worksheet.
    """
    try:
        logger.info("Creating Settlement table...")
        headers = [
            "Settlement ID", "Case ID", "DRC Name", "RO Name", "Status", "Status reason",
            "Status DTM", "Settlement Type", "Settlement Amount", "Settlement Phase",
            "Settlement Created by", "Settlement Created DTM", "Last Monitoring DTM", "Remark"
        ]
        
        # Prepare data for the table
        data = [
            [
                settlement.get("settlement_id"), settlement.get("case_id"), settlement.get("drc_id"),
                settlement.get("ro_id"), settlement.get("settlement_status"), settlement.get("status_reason"),
                settlement.get("status_dtm"), settlement.get("settlement_type"), settlement.get("settlement_amount"),
                settlement.get("settlement_phase"), settlement.get("created_by"), settlement.get("created_on"),
                settlement.get("last_monitoring_dtm"), settlement.get("remark")
            ]
            for settlement in settlements
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Settlement Details", headers, data, styles)
        
        logger.info("Settlement table created successfully.")
        return next_row
    except Exception as failed_settlement_table_creation:
        logger.error(f"Failed to create Settlement table: {failed_settlement_table_creation}")
        sys.exit(1)

def create_settlement_plan_table(worksheet, settlement_plans, x_pointer, y_pointer, styles):
    """
    Create the Settlement Plan table in the worksheet.
    """
    try:
        logger.info("Creating Settlement Plan table...")
        headers = [
            "Settlement ID", "Installment Sequence", "Installment Settle Amount",
            "Accumulated Amount", "Plan Date and Time"
        ]
        
        # Prepare data for the table
        data = [
            [
                plan.get("settlement_id"), plan.get("installment_seq"), plan.get("installment_settle_amount"),
                plan.get("accumulated_amount"), plan.get("plan_date")
            ]
            for plan in settlement_plans
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Settlement Plan", headers, data, styles)
        
        logger.info("Settlement Plan table created successfully.")
        return next_row
    except Exception as failed_settlement_plan_table_creation:
        logger.error(f"Failed to create Settlement Plan table: {failed_settlement_plan_table_creation}")
        sys.exit(1)