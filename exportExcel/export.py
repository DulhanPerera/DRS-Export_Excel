import sys # Module for system-specific parameters and functions
import logging # Module for logging errors and debugging information
# Import custom modules for configuration loading, styling, and Excel writing
from .config_loader import load_config
from .excel_styles import load_styles
from .excel_writer import export_all_tables
import logging.config # Module for loading logging configurations
from pymongo import MongoClient # MongoDB client for database interactions


# Initialize logger for this module
logger = logging.getLogger('excel_data_writer')

def connect_db(mongo_uri, db_name):
    """
    Connect to the MongoDB database using the provided URI and database name.

    Args:
        mongo_uri (str): The MongoDB connection string.
        db_name (str): The name of the database to connect to.

    Returns:
        pymongo.database.Database: A MongoDB database object.

    Outputs:
        - Logs a success message if the connection is successful.
        - Logs an error and exits if the connection fails.

    Exceptions:
        - Exits the program if unable to establish a database connection.
    """
    try:
        # Establish a connection to MongoDB
        client = MongoClient(mongo_uri)
        db = client[db_name]
        # Log successful connection
        logger.info("Successfully connected to the database.")
        return db
    except Exception as failed_db_connection:
        # Log the error and exit the program
        logger.error(f"Failed to connect to the database: {failed_db_connection}")
        sys.exit(1)

def setup_logger(logger_config_path):
    """
    Set up logging configuration using the specified logger config file.

    Args:
        logger_config_path (str): The path to the logger configuration file.

    Returns:
        logging.Logger: Configured logger instance.

    Outputs:
        - Logs an error and exits if the logger configuration file cannot be loaded.

    Exceptions:
        - Exits the program if logging setup fails.
    """
    try:
        # Load logging configuration from the file
        logging.config.fileConfig(logger_config_path)

        # Get and return the configured logger
        logger = logging.getLogger('excel_data_writer')
        return logger

    except Exception as failed_logger_setup:
        # Print the error message and exit
        print(f"Failed to set up logger: {failed_logger_setup}")
        sys.exit(1)

def start_process():
    """
    Main function to execute the case details export process.

    Args:
        None

    Returns:
        None

    Outputs:
        - Logs the progress of the export process.
        - Logs an error and exits if any step in the process fails.

    Exceptions:
        - Exits the program if an unexpected error occurs.
    """
    try:
        # Set up the logger from the configuration file
        logger = setup_logger('Config/logger/loggers.ini')
        logger.info("Starting case details export process...")

        # Load configuration settings
        config = load_config('Config/Config.ini')

        # Connect to MongoDB database
        db = connect_db(config['DATABASE']['MONGO_URI'], config['DATABASE']['DB_NAME'])

        # Load styling configurations for the Excel export
        styles = load_styles('Config/styles.ini')

        # Define necessary parameters
        incident_id = 2025  # Replace with the actual incident ID
        export_path = config['EXCEL_EXPORT_FOLDER']['WIN_DB']
        collection_name = config['COLLECTIONS']['CASE_DETAIL_COLLECTION']

        # Call function to export case details into an Excel file
        export_all_tables(db, incident_id, export_path, collection_name, styles)

        # Log successful completion of the process
        logger.info("Case details export process completed.")

    except Exception as failed_process:
        # Log any unexpected errors and exit the program
        logger.error(f"An unexpected error occurred: {failed_process}")
        sys.exit(1)