from flask import Flask, render_template, request, redirect, url_for, flash
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

app = Flask(__name__)
app.secret_key = "your_secret_key"  # Needed for flashing messages

# Google Sheets Setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("/etc/secrets/deposittracking-c75aa9c780f4.json", scope)

client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1BWROL9J53G4wSou5KpwI1rgXymNe31F-zDm_iXjBcok/edit#gid=0").sheet1

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        date_for = request.form["date_for"]  # Date the deposit was for
        actual_deposit = float(request.form["actual_deposit"])
        reference_number = request.form["reference_number"]

        # Find the row with the matching date in Column A (Date)
        cell = sheet.find(date_for)
        if cell:
            row = cell.row
            expected_deposit = float(sheet.cell(row, 2).value)  # Column B is "Expected Deposit"
            difference = actual_deposit - expected_deposit
            status = "Match" if difference == 0 else "Not Match"
            
            # Update Google Sheets
            sheet.update_cell(row, 3, actual_deposit)  # Column C: Actual Deposit
            sheet.update_cell(row, 4, status)  # Column D: Status
            sheet.update_cell(row, 5, difference)  # Column E: Difference
            sheet.update_cell(row, 6, reference_number)  # Column F: Reference #

            # Apply conditional formatting to the "Status" column
            apply_conditional_formatting()

            # Redirect to confirmation page with data
            return redirect(url_for("confirmation", 
                                    status=status, 
                                    date_for=date_for, 
                                    actual_deposit=actual_deposit,
                                    expected_deposit=expected_deposit,
                                    difference=difference))
        else:
            flash("Date not found in the Google Sheet. Please double-check the date entered.", "danger")

    return render_template("index.html")

@app.route("/confirmation")
def confirmation():
    status = request.args.get("status")
    date_for = request.args.get("date_for")
    actual_deposit = request.args.get("actual_deposit")
    expected_deposit = request.args.get("expected_deposit")
    difference = request.args.get("difference")
    
    return render_template("confirmation.html", 
                           status=status, 
                           date_for=date_for,
                           actual_deposit=actual_deposit,
                           expected_deposit=expected_deposit,
                           difference=difference)

# Google Sheets Conditional Formatting
def apply_conditional_formatting():
    credentials = Credentials.from_service_account_file("/etc/secrets/deposittracking.json")

    service = build('sheets', 'v4', credentials=credentials)

    spreadsheet_id = "1BWROL9J53G4wSou5KpwI1rgXymNe31F-zDm_iXjBcok"  # Your correct spreadsheet ID

    requests = [
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": 0, "startColumnIndex": 3, "endColumnIndex": 4}],  # Column D (Status)
                    "booleanRule": {
                        "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Match"}]},
                        "format": {"backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.83}}
                    }
                },
                "index": 0
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": 0, "startColumnIndex": 3, "endColumnIndex": 4}],  # Column D (Status)
                    "booleanRule": {
                        "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Not Match"}]},
                        "format": {"backgroundColor": {"red": 1.0, "green": 0.8, "blue": 0.8}}
                    }
                },
                "index": 1
            }
        }
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

if __name__ == "__main__":
    app.run(debug=True)