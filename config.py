from datetime import datetime

START_YEAR = 2015
END_YEAR = 2025
REQUEST_SHEET = "REQUEST_TABLE"

COUNTRY_CODE_INPUT = "country_code.xlsx"

ENTITY_INPUT_FOLDER = "data-split-by-entity"
ENTITY_OUTPUT_FOLDER = "data-split-by-variable"
ENTITY_LOG_FILE = f"entity_integrate_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

VARIABLE_OUTPUT_FOLDER = "data"
VARIABLE_LOG_FILE = f"variable_integrate_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

YEAR_OUTPUT_FOLDER = f"data-{START_YEAR}-{END_YEAR}"
YEAR_LOG_FILE = f"year_integrate_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
