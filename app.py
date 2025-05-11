import os
import logging # Import logging
from flask import Flask, render_template, jsonify, request, session # Import session
from google.oauth2 import service_account # Use service_account credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv # Import load_dotenv
import json # Import json for parsing errors
from waitress import serve # Import serve

load_dotenv() # Load environment variables from .env file

# --- Basic Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# --- End Logging Setup ---

# --- Configuration ---
# Removed SPREADSHEET_ID and RANGE_NAME constants as they come from request now
# API_KEY = os.getenv('GOOGLE_API_KEY') # No longer using API Key
SERVICE_ACCOUNT_FILE = 'credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets' # Need read AND write scope now
]

# Define default columns and max rows for the range
DEFAULT_START_COLUMN_LETTER = 'A'
DEFAULT_END_COLUMN_LETTER = 'YZ' # Fetch up to column YZ (650th column)
DEFAULT_MAX_ROWS = 500 # Fetch up to 500 rows initially
# --- End Configuration ---

app = Flask(__name__)

# --- IMPORTANT: Set a Secret Key for Sessions ---
# In production, use a strong, random key set via environment variable
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'dev-secret-key-replace-me-in-prod')
if app.secret_key == 'dev-secret-key-replace-me-in-prod':
    logging.warning("Using default FLASK_SECRET_KEY. Set a proper secret key in your environment for production!")
# --- End Secret Key ---

# --- Utility to get column letter ---
def get_col_letter(col_index_zero_based):
    """Converts 0-based column index to A1 notation letter (A, B, ..., Z, AA, AB)."""
    letter = ''
    while col_index_zero_based >= 0:
        letter = chr(65 + (col_index_zero_based % 26)) + letter
        col_index_zero_based = col_index_zero_based // 26 - 1
    return letter
# --- End Utility ---

# --- Authentication Helper ---
def get_credentials():
    """Loads service account credentials from ENV var or file."""
    creds = None
    google_credentials_json_str = os.getenv('GOOGLE_CREDENTIALS_JSON')

    if google_credentials_json_str:
        logging.info("Loading credentials from GOOGLE_CREDENTIALS_JSON env var.")
        try:
            # Parse the JSON string from the environment variable
            credentials_info = json.loads(google_credentials_json_str)
            creds = service_account.Credentials.from_service_account_info(
                credentials_info, scopes=SCOPES)
            logging.info("Credentials loaded successfully from env var.")
        except json.JSONDecodeError:
            logging.error("CRITICAL: Failed to parse GOOGLE_CREDENTIALS_JSON env var.")
            return None
        except Exception as e:
            logging.error(f"CRITICAL: Error loading credentials from env var info: {e}", exc_info=True)
            return None
    else:
        logging.info(f"GOOGLE_CREDENTIALS_JSON not set. Attempting to load from file: {SERVICE_ACCOUNT_FILE}")
        try:
            creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
            logging.info("Credentials loaded successfully from file.")
        except FileNotFoundError:
            logging.error(f"CRITICAL: Service account file not found at {SERVICE_ACCOUNT_FILE}. Set GOOGLE_CREDENTIALS_JSON env var for Vercel.")
            return None
        except Exception as e:
            logging.error(f"CRITICAL: Error loading credentials from file: {e}", exc_info=True)
            return None
    return creds
# --- End Auth Helper ---

def get_sheet_data(spreadsheet_id, range_name):
    """Fetches data from the Google Sheet using an API Key.

    Args:
        spreadsheet_id: The ID of the Google Spreadsheet.
        range_name: The A1 notation of the range to retrieve (e.g., 'Sheet1!A1:S150').
    """
    logging.info(f"Attempting to get sheet data for ID: {spreadsheet_id}, Range: {range_name}")

    if not SERVICE_ACCOUNT_FILE:
        logging.error("Service account file not found in environment variables.")
        # Return a specific error message that can be passed to the frontend
        return None, "Server configuration error: Missing Service Account File."

    try:
        logging.info(f"Building Google Sheets API service (v4) with Service Account.")
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        logging.info(f"Executing sheet.values().get() for spreadsheetId={spreadsheet_id}, range={range_name}")
        result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                    range=range_name).execute()
        values = result.get('values', [])
        logging.info(f"Successfully fetched {len(values)} rows from {spreadsheet_id}/{range_name}")
        return values, None # Return data and no error
    except HttpError as err:
        logging.error(f"Google Sheets API error occurred: {err}", exc_info=True) # Log stack trace
        error_content = err.resp.get('content', b'{}')
        error_details_msg = f"API Error: {err.resp.status} {err.reason}."
        try:
            # Decode safely
            decoded_content = error_content.decode('utf-8', errors='ignore')
            logging.info(f"Raw error content from API: {decoded_content}")
            # Use json.loads if possible, fall back to eval cautiously or just use the string
            try:
                error_details = json.loads(decoded_content)
                logging.info(f"Parsed JSON error details: {error_details}")
                if isinstance(error_details, dict) and 'error' in error_details:
                    error_details_msg += f" Details: {error_details['error'].get('message', '{}')}"
            except json.JSONDecodeError:
                 logging.warning(f"Could not parse API error content as JSON: {decoded_content}")
                 # Attempt eval as a fallback ONLY if necessary and understood risks
                 # For now, just append raw decoded content might be safer
                 error_details_msg += f" Raw details: {decoded_content}"
        except Exception as parse_err:
            logging.error(f"Error processing API error details: {parse_err}")
            error_details_msg += f" Raw undecoded content: {error_content}"

        # Check for common specific errors based on status and potentially content
        # Ensure range_name is defined here before using it in f-string
        current_range_name = range_name if 'range_name' in locals() else '[range not available]'
        if err.resp.status == 400 and "Unable to parse range" in str(error_content):
             error_details_msg = f"Error: Invalid Sheet Name or Range ('{current_range_name}'). Please check the sheet name exists and is spelled correctly, and the range format is valid."
        elif err.resp.status == 403:
            error_details_msg = "Error: Permission denied. Make sure the Google Sheet is shared as 'Anyone with the link can view'."
        elif err.resp.status == 404:
            error_details_msg = f"Error: Spreadsheet not found. Check the URL and ensure the ID ('{spreadsheet_id}') is correct."

        return None, error_details_msg # Return no data and the error message

    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}", exc_info=True) # Log stack trace
        return None, f"An unexpected server error occurred: {e}"

@app.route('/')
def index():
    # Just render the basic page, data will be loaded via JS
    return render_template('index.html')

@app.route('/load-data')
def load_data():
    spreadsheet_id = request.args.get('id')
    sheet_name = request.args.get('sheet')
    # Get start/end row numbers, default to None if not provided or invalid
    try:
        start_row_num = request.args.get('start_row', type=int)
        if start_row_num is not None and start_row_num < 1:
            logging.warning(f"Invalid start_row parameter ({start_row_num}), ignoring.")
            start_row_num = None # Treat invalid as None
    except ValueError:
        logging.warning(f"Non-integer start_row parameter received, ignoring.")
        start_row_num = None
    try:
        end_row_num = request.args.get('end_row', type=int)
        if end_row_num is not None and end_row_num < 1:
             logging.warning(f"Invalid end_row parameter ({end_row_num}), ignoring.")
             end_row_num = None # Treat invalid as None
    except ValueError:
        logging.warning(f"Non-integer end_row parameter received, ignoring.")
        end_row_num = None

    # Validate start <= end if both are provided
    if start_row_num is not None and end_row_num is not None and start_row_num > end_row_num:
         logging.warning(f"start_row ({start_row_num}) > end_row ({end_row_num}), ignoring row filtering.")
         start_row_num = None
         end_row_num = None

    logging.info(f"Request: /load-data | ID: '{spreadsheet_id}', Sheet: '{sheet_name}', StartRow: {start_row_num}, EndRow: {end_row_num}")

    if not spreadsheet_id or not sheet_name:
        # Error logging handled inside the check
        if not spreadsheet_id: logging.warning("Request missing ID.")
        if not sheet_name: logging.warning("Request missing Sheet Name.")
        return jsonify({"error": "Missing required parameters (id, sheet)."}), 400

    # Fetch data including notes using the new function
    header, all_rows, notes, data_row_sheet_indices, error_msg = get_sheet_data_with_notes(spreadsheet_id, sheet_name)

    if error_msg:
        logging.error(f"Error from get_sheet_data_with_notes: {error_msg}")
        return jsonify({"error": error_msg}), 500

    # --- Apply Row Filtering (if applicable) ---
    filtered_rows = [] # Initialize filtered_rows
    filtered_row_indices = []
    if all_rows and (start_row_num is not None or end_row_num is not None):
        # Adjust to 0-based index for slicing (start_row_num=1 maps to index 0)
        slice_start = (start_row_num - 1) if start_row_num is not None else 0
        slice_end = end_row_num if end_row_num is not None else len(all_rows)

        # Ensure indices are within bounds of the actual data_rows
        slice_start = max(0, slice_start)
        slice_end = min(len(all_rows), slice_end)

        if slice_start < slice_end:
            filtered_rows = all_rows[slice_start:slice_end]
            filtered_row_indices = data_row_sheet_indices[slice_start:slice_end]
            logging.info(f"Applied filter: start_row={start_row_num}, end_row={end_row_num}. Resulting rows: {len(filtered_rows)} (Indices {slice_start}-{slice_end-1})")
        else:
            filtered_rows = []
            filtered_row_indices = []
            logging.info(f"Applied filter: start_row={start_row_num}, end_row={end_row_num}. No rows selected after bounds check.")
    else:
        # No filtering requested or no data_rows to filter
        filtered_rows = all_rows
        filtered_row_indices = data_row_sheet_indices
        if all_rows:
            logging.info(f"No row filtering applied. Using all {len(filtered_rows)} data rows.")
    # --- End Filtering ---

    # Return PADDED header and the (potentially filtered) rows
    logging.info(f"Returning JSON. Padded header length: {len(header)}, Filtered rows count: {len(filtered_rows)}, Notes count: {len(notes)}")
    return jsonify({"header": header, "rows": filtered_rows, "notes": notes, "data_row_sheet_indices": filtered_row_indices})

@app.route('/get-sheet-names')
def get_sheet_names():
    spreadsheet_id = request.args.get('id')
    logging.info(f"Request: /get-sheet-names | ID: '{spreadsheet_id}'")
    if not spreadsheet_id: return jsonify({"error": "Missing ID."}, 400)

    creds = get_credentials()
    if not creds:
        return jsonify({"error": "Server authentication error."}), 500

    sheet_names = []
    error_msg = None
    try:
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet_metadata.get('sheets', [])
        for sheet in sheets:
            title = sheet.get('properties', {}).get('title')
            if title: sheet_names.append(title)
        logging.info(f"Fetched {len(sheet_names)} sheet names for ID: {spreadsheet_id}")
    except HttpError as err:
        logging.error(f"API error getting sheet names: {err}", exc_info=True)
        # Basic error reporting for brevity, could parse details like before
        error_msg = f"API Error ({err.resp.status}): Could not get sheet names. Check permissions and Sheet ID."
    except Exception as e:
        logging.error(f"Unexpected error getting sheet names: {e}", exc_info=True)
        error_msg = "Unexpected server error getting sheet names."

    if error_msg:
        return jsonify({"error": error_msg}), 500
    else:
        return jsonify({"sheet_names": sheet_names})

# --- Modified Data Fetching ---
def get_sheet_data_with_notes(spreadsheet_id, sheet_name):
    """Fetches sheet data including values and notes using spreadsheets.get."""
    logging.info(f"Fetching data and notes for ID: {spreadsheet_id}, Sheet: {sheet_name}")
    creds = get_credentials()
    if not creds: return None, None, None, None, "Server authentication error."

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Construct the range string for A1:YZ500 within the specified sheet
        range_a1 = f"{sheet_name}!{DEFAULT_START_COLUMN_LETTER}1:{DEFAULT_END_COLUMN_LETTER}{DEFAULT_MAX_ROWS}"
        logging.info(f"Calling spreadsheets.get with range '{range_a1}' and includeGridData=true")

        # Use spreadsheets.get to fetch grid data including notes
        # Specify fields to potentially limit response size, include necessary grid data
        fields = 'sheets(data(rowData(values(formattedValue,note))))'
        # Note: Fetching grid data can be slower/larger than values().get()
        result = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[range_a1],
            fields=fields,
            includeGridData=True
        ).execute()

        # --- Parse the GridData response --- 
        sheets_data = result.get('sheets', [])
        if not sheets_data or not sheets_data[0].get('data'):
            logging.warning(f"No sheet data found in response for range {range_a1}")
            return [], {}, {}, [], None # Return empty header/notes, no error

        grid_data = sheets_data[0]['data'][0] # Assuming single range fetch
        row_data = grid_data.get('rowData', [])

        header = []
        all_rows_values = []
        notes = {} # Store notes as { "A1_notation": "note text" }
        max_cols = 0
        data_row_sheet_indices = []  # List of actual sheet row numbers for each data row

        # Find max columns based *only* on formattedValue presence first
        # This avoids notes in empty columns creating extra padding later
        for r_idx, row in enumerate(row_data):
             values = row.get('values', [])
             max_cols = max(max_cols, len(values))

        logging.info(f"Max columns found based on cell values: {max_cols}")

        sheet_row_offset = 1 # Sheet rows are 1-based

        for r_idx, row in enumerate(row_data):
            current_row_values = []
            values_in_row = row.get('values', [])
            sheet_row_num = r_idx + sheet_row_offset

            for c_idx in range(max_cols):
                cell_data = values_in_row[c_idx] if c_idx < len(values_in_row) else {}
                formatted_value = cell_data.get('formattedValue', '')
                note = cell_data.get('note')
                current_row_values.append(formatted_value if formatted_value is not None else '')

                if note:
                    col_letter = get_col_letter(c_idx)
                    a1_notation = f"{sheet_name}!{col_letter}{sheet_row_num}"
                    notes[a1_notation] = note
                    # logging.debug(f"Found note at {a1_notation}: {note}") # Optional debug log

            # Extract header from the 2nd row (index 1)
            if r_idx == 1:
                header = current_row_values[:max_cols]
            # Add subsequent rows to data (starting from index 2)
            if r_idx >= 2:
                all_rows_values.append(current_row_values[:max_cols])
                data_row_sheet_indices.append(sheet_row_num)

        logging.info(f"Parsed {len(header)} header columns and {len(all_rows_values)} data rows. Found {len(notes)} notes.")
        return header, all_rows_values, notes, data_row_sheet_indices, None # Return header, rows, notes, row indices, no error

    except HttpError as err:
        logging.error(f"API error getting sheet data/notes: {err}", exc_info=True)
        error_msg = f"API Error ({err.resp.status}): Could not get sheet data. Check permissions, Sheet ID, and Sheet Name."
        # Add more specific error checks if needed
        return None, None, None, None, error_msg
    except Exception as e:
        logging.error(f"Unexpected error getting sheet data/notes: {e}", exc_info=True)
        return None, None, None, None, "Unexpected server error getting sheet data."
# --- End Data Fetching ---

# --- New Endpoint to Save Note ---
@app.route('/save-note', methods=['POST'])
def save_note_route():
    data = request.json
    spreadsheet_id = data.get('id')
    a1_notation = data.get('a1') # e.g., "Sheet1!C5"
    note_text = data.get('note')

    logging.info(f"Request: /save-note | ID: '{spreadsheet_id}', A1: '{a1_notation}', Note: '{note_text[:50]}...'")

    if not spreadsheet_id or not a1_notation:
        return jsonify({"error": "Missing required parameters (id, a1)."}), 400
    if note_text is None: # Allow empty string to clear note
        note_text = ''

    creds = get_credentials()
    if not creds:
        return jsonify({"error": "Server authentication error."}), 500

    error_msg = None
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Parse Sheet Name, Row, Col from A1 notation
        # Basic parsing, assumes standard format like SheetName!LetterNumber
        parts = a1_notation.split('!')
        if len(parts) != 2:
             raise ValueError("Invalid A1 notation format")
        sheet_name = parts[0]
        cell_ref = parts[1]
        col_letter = ''
        row_num_str = ''
        for char in cell_ref:
            if char.isalpha():
                col_letter += char
            elif char.isdigit():
                row_num_str += char
        if not col_letter or not row_num_str:
             raise ValueError("Could not parse column/row from A1 notation")
        row_index = int(row_num_str) - 1 # Convert to 0-based index
        col_index = 0
        for char in col_letter:
            col_index = col_index * 26 + (ord(char.upper()) - 65 + 1)
        col_index -= 1

        # --- Fetch sheetId BEFORE building the request body ---
        logging.info(f"Fetching sheetId for sheet name: '{sheet_name}'")
        spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id, fields='sheets/properties(sheetId,title)').execute()
        sheet_id = None
        for sheet in spreadsheet_metadata.get('sheets', []):
            if sheet.get('properties', {}).get('title') == sheet_name:
                sheet_id = sheet.get('properties', {}).get('sheetId')
                break
        if sheet_id is None:
            raise ValueError(f"Could not find sheet ID for sheet name '{sheet_name}'")
        logging.info(f"Found sheetId: {sheet_id} for sheet name: '{sheet_name}'")
        # --- End Fetch sheetId ---

        # Use batchUpdate to set the note
        requests_body = {
            'requests': [
                {
                    'updateCells': {
                        'rows': [
                            {
                                'values': [
                                    {
                                        'note': note_text if note_text else None
                                    }
                                ]
                            }
                        ],
                        'fields': 'note',
                        'range': {
                            # Now use the fetched sheet_id and indices
                            'sheetId': sheet_id,
                            'startRowIndex': row_index,
                            'endRowIndex': row_index + 1,
                            'startColumnIndex': col_index,
                            'endColumnIndex': col_index + 1
                        }
                    }
                }
            ]
        }

        logging.info(f"Executing batchUpdate to set note for range: sheetId={sheet_id}, row={row_index}, col={col_index}")
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=requests_body
        ).execute()

        logging.info(f"Successfully updated note for {a1_notation}. Response: {response}")

    except HttpError as err:
        logging.error(f"API error saving note: {err}", exc_info=True)
        error_msg = f"API Error ({err.resp.status}): Could not save note. Check permissions and cell reference."
    except ValueError as e:
        logging.error(f"Value error processing save note request: {e}")
        error_msg = f"Invalid request format: {e}"
    except Exception as e:
        logging.error(f"Unexpected error saving note: {e}", exc_info=True)
        error_msg = "Unexpected server error saving note."

    if error_msg:
        return jsonify({"error": error_msg}), 500
    else:
        return jsonify({"success": True, "a1": a1_notation, "note": note_text})
# --- End Save Note Endpoint ---

# --- New Endpoint to Ban Cell (Set Value to 0) ---
@app.route('/ban-cell', methods=['POST'])
def ban_cell_route():
    data = request.json
    spreadsheet_id = data.get('id')
    a1_notation = data.get('a1') # e.g., "Sheet1!D5"

    logging.info(f"Request: /ban-cell | ID: '{spreadsheet_id}', A1: '{a1_notation}'")

    if not spreadsheet_id or not a1_notation:
        return jsonify({"error": "Missing required parameters (id, a1)."}), 400

    creds = get_credentials()
    if not creds:
        return jsonify({"error": "Server authentication error."}), 500

    error_msg = None
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Define the request body for values.update
        # We set the value to the string "0"
        body = {
            'values': [['0']]
        }

        logging.info(f"Executing values.update to set cell {a1_notation} to '0'")
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=a1_notation, # values.update works directly with A1 range
            valueInputOption='USER_ENTERED', # Interpret '0' as a number if possible by Sheets
            body=body
        ).execute()

        logging.info(f"Successfully set cell {a1_notation} to '0'. Response: {result}")

    except HttpError as err:
        logging.error(f"API error banning cell: {err}", exc_info=True)
        error_msg = f"API Error ({err.resp.status}): Could not update cell value. Check permissions and cell reference."
        # Add more specific checks if needed (e.g., 400 for invalid range)
        if err.resp.status == 400:
            error_msg += " (Possible invalid cell reference)"

    except Exception as e:
        logging.error(f"Unexpected error banning cell: {e}", exc_info=True)
        error_msg = "Unexpected server error updating cell value."

    if error_msg:
        return jsonify({"error": error_msg}), 500
    else:
        # Return success along with which cell was updated
        return jsonify({"success": True, "a1": a1_notation, "newValue": "0"})
# --- End Ban Cell Endpoint ---

if __name__ == '__main__':
    # No longer need the templates check here, focus on serving
    # if not os.path.exists('templates'):
    #     os.makedirs('templates')
    port = 3000
    print(f"Starting Waitress server on http://0.0.0.0:{port}")
    # Use waitress.serve instead of app.run()
    serve(app, host='0.0.0.0', port=port)
    # app.run(debug=True) # Remove flask development server 