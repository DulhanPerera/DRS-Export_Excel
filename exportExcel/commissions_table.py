import logging
import sys
from .table_utils import create_table

logger = logging.getLogger('excel_data_writer')

def create_commissions_table(worksheet, db, case_id, x_pointer, y_pointer, styles):
    """
    Create the Commissions table in the worksheet.
    """
    try:
        logger.info("Creating Commissions table...")
        headers = [
            "Money Transaction ID", "Transaction Type", "Paid DTM", "Arrears", 
            "Transaction", "Running Credit", "Running Debt", 
            "Cumulative Settled Balance", "Commissioned Amount"
        ]
        
        # Fetch commissions data from the Commissions collection
        commissions_collection = db["Commissions"]
        commissions_data = list(commissions_collection.find({"case_id": case_id}))
        
        # Prepare data for the table
        data = [
            [
                transaction.get("money_transaction_id"),
                transaction.get("transaction_type"),
                transaction.get("paid_dtm"),
                transaction.get("arrears"),
                transaction.get("transaction"),
                transaction.get("running_credit"),
                transaction.get("running_debt"),
                transaction.get("cummulative_settled_balance"),
                transaction.get("commissioned_amount")
            ]
            for transaction in commissions_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Commissions", headers, data, styles)
        
        logger.info("Commissions table created successfully.")
        return next_row
    except Exception as failed_commissions_table_creation:
        logger.error(f"Failed to create Commissions table: {failed_commissions_table_creation}")
        sys.exit(1)