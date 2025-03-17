import logging  # Module for logging errors and debug information

# Initialize logger for this module
logger = logging.getLogger('excel_data_writer')


def get_arrears_band_value(db, current_arrears_band):
    """
    Retrieve the value for the given arrears band from the 'Arrears_bands' collection.

    Args:
        db (pymongo.database.Database): The MongoDB database instance.
        current_arrears_band (str): The key representing the specific arrears band.

    Returns:
        any: The value corresponding to the given arrears band, or None if not found.

    Outputs:
        - Logs a warning if no arrears bands document is found.
        - Logs an error if an exception occurs.

    Exceptions:
        - Returns None if an error occurs while retrieving the arrears band value.
    """
    try:
        # Access the 'Arrears_bands' collection in the database
        arrears_bands_collection = db["Arrears_bands"]

        # Retrieve a single document from the collection
        arrears_bands_doc = arrears_bands_collection.find_one({})

        # Return the requested arrears band value if the document exists
        if arrears_bands_doc:
            return arrears_bands_doc.get(current_arrears_band)
        else:
            logger.warning("No arrears bands document found in the collection.")
            return None
    except Exception as failed_arrears_band_retrieval:
        logger.error(f"Failed to retrieve arrears band value: {failed_arrears_band_retrieval}")
        return None


def get_settlement_data(db, case_id):
    """
    Retrieve settlement data for the given case_id from the 'Case_settlements' collection.

    Args:
        db (pymongo.database.Database): The MongoDB database instance.
        case_id (int or str): The unique identifier of the case.

    Returns:
        list: A list of settlement records for the given case_id, or an empty list if none are found.

    Outputs:
        - Logs the number of settlements found.
        - Logs a warning if no settlements are found.
        - Logs an error if an exception occurs.

    Exceptions:
        - Returns an empty list if an error occurs while retrieving settlement data.
    """
    try:
        # Access the 'Case_settlements' collection in the database
        settlements_collection = db["Case_settlements"]

        # Retrieve all settlements matching the given case_id
        settlements = list(settlements_collection.find({"case_id": case_id}))

        # Log and return results
        if settlements:
            logger.info(f"Found {len(settlements)} settlement records for case_id: {case_id}")
        else:
            logger.warning(f"No settlement records found for case_id: {case_id}")
        
        return settlements
    except Exception as failed_settlement_retrieval:
        logger.error(f"Failed to retrieve settlement data: {failed_settlement_retrieval}")
        return []


def get_settlement_plan_data(db, case_id):
    """
    Retrieve settlement plan data for the given case_id from the 'Case_settlements' collection.

    Args:
        db (pymongo.database.Database): The MongoDB database instance.
        case_id (int or str): The unique identifier of the case.

    Returns:
        list: A list of settlement plan records, each including a settlement_id, or an empty list if none are found.

    Outputs:
        - Logs the number of settlement plans found.
        - Logs a warning if no settlement plans are found.
        - Logs an error if an exception occurs.

    Exceptions:
        - Returns an empty list if an error occurs while retrieving settlement plan data.
    """
    try:
        # Access the 'Case_settlements' collection in the database
        settlements_collection = db["Case_settlements"]

        # Retrieve all settlements matching the given case_id
        settlements = list(settlements_collection.find({"case_id": case_id}))

        # Initialize a list to store extracted settlement plans
        settlement_plans = []

        # Iterate through each settlement record
        for settlement in settlements:
            if "settlement_plan" in settlement:
                for plan in settlement["settlement_plan"]:
                    # Add settlement_id to each plan for reference
                    plan["settlement_id"] = settlement.get("settlement_id")
                    settlement_plans.append(plan)

        # Log and return results
        if settlement_plans:
            logger.info(f"Found {len(settlement_plans)} settlement plan records for case_id: {case_id}")
        else:
            logger.warning(f"No settlement plan records found for case_id: {case_id}")
        
        return settlement_plans
    except Exception as failed_settlement_plan_retrieval:
        logger.error(f"Failed to retrieve settlement plan data: {failed_settlement_plan_retrieval}")
        return []
