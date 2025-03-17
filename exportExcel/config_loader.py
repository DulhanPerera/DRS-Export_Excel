import configparser # Module for reading configuration files
import sys # Module for system-specific parameters and functions
import logging # Module for logging errors and debugging information

# Set up a logger for this module
logger = logging.getLogger('excel_data_writer')


def load_config(config_file_path):
    """
    Loads and returns the configuration from the specified config file.

    Args:
        config_file_path (str): The path to the configuration file.

    Returns:
        configparser.ConfigParser: A ConfigParser object containing the loaded configuration.

    Outputs:
        - Logs an error and exits if the configuration file is empty or not found.
        - Logs an error and exits if any exception occurs while reading the file.

    Exceptions:
        - Exits the program if the configuration file is empty or cannot be read.
    """
    try:
        # Create a ConfigParser object
        config = configparser.ConfigParser()
        config.read(config_file_path) # Read the configuration file
        # Check if the file contains any sections
        if not config.sections():
            logger.error("Configuration file is empty or not found.")
            sys.exit(1) # Terminate the program with an error code
        return config # Return the loaded configuration object
    except Exception as failed_config_load:
        # Log the exception and terminate the program
        logger.error(f"Failed to load configuration: {failed_config_load}")
        sys.exit(1)