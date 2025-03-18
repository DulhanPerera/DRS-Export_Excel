import logging
import sys
from .table_utils import create_table

logger = logging.getLogger('excel_data_writer')

def create_commissions_table(worksheet, db, case_id, x_pointer, y_pointer, styles):
    """
    Create the Commissions table in the worksheet.
    Fetches money_transaction_id from Case_payments collection based on case_id,
    then fetches commission data from Commissions collection using those transaction IDs.
    """
    try:
        logger.info("Creating Commissions table...")
        headers = [
            "Money Transaction ID", "Transaction Type", "Paid DTM", "Arrears", 
            "Transaction", "Running Credit", "Running Debt", 
            "Cumulative Settled Balance", "Commissioned Amount"
        ]
        
        # Fetch money_transaction_id(s) from Case_payments collection related to the case_id
        payments_collection = db["Case_payments"]
        money_transaction_ids = payments_collection.distinct("money_transaction_id", {"case_id": case_id})
        
        # Fetch commission data for each money_transaction_id from Commissions collection
        commissions_collection = db["Commissions"]
        commissions_data = []
        for money_transaction_id in money_transaction_ids:
            transactions = list(commissions_collection.find({
                "money_transaction_id": money_transaction_id
            }))
            commissions_data.extend(transactions)
        
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
        
        # Create the table only if data exists
        if data:
            next_row = create_table(worksheet, x_pointer, y_pointer, "Commissions", headers, data, styles)
            logger.info("Commissions table created successfully.")
            return next_row
        else:
            logger.warning("No commission data found for the given case_id.")
            return x_pointer  # Return the same row pointer if no data is found
    except Exception as failed_commissions_table_creation:
        logger.error(f"Failed to create Commissions table: {failed_commissions_table_creation}")
        sys.exit(1)