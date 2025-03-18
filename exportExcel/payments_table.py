import logging
import sys
from .table_utils import create_table

logger = logging.getLogger('excel_data_writer')

def create_payments_table(worksheet, db, case_id, x_pointer, y_pointer, styles):
    """
    Create the Payments table in the worksheet.
    """
    try:
        logger.info("Creating Payments table...")
        headers = [
            "Payment ID", "Settlement ID", "Installment Sequence", "Bill Payment Sequence", "Bill Paid Amount", "Bill Paid Date",
            "Bill Payment Status", "Bill Payment Type", "Settled Balance", "Cumulative Settled Balance", "Created Date and Time",
            "Account No", "Money Transaction Reference Type", "Money Transaction ID"
        ]
        
        # Fetch payments data from the Case_payments collection
        payments_collection = db["Case_payments"]
        payments_data = list(payments_collection.find({"case_id": case_id}))
        
        # Prepare data for the table
        data = [
            [
                payment.get("payment_id"),
                payment.get("settlement_id"),
                payment.get("installment_seq"),
                payment.get("bill_payment_seq"),
                payment.get("bill_paid_amount"),
                payment.get("bill_paid_date"),
                payment.get("bill_payment_status"),
                payment.get("bill_payment_type"),
                payment.get("settled_balance"),
                payment.get("cumulative_settled_balance"),
                payment.get("created_dtm"),
                payment.get("account_no"),
                payment.get("money_transaction_Reference_type"),
                payment.get("money_transaction_id")
            ]
            for payment in payments_data
        ]
        
        # Create the table
        next_row = create_table(worksheet, x_pointer, y_pointer, "Payments", headers, data, styles)
        
        logger.info("Payments table created successfully.")
        return next_row
    except Exception as failed_payments_table_creation:
        logger.error(f"Failed to create Payments table: {failed_payments_table_creation}")
        sys.exit(1)