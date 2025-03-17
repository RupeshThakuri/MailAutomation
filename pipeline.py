import pickle
import re
import openpyxl
import os
from datetime import datetime
from google_apis import create_service
from gmail_api import get_email_messages, get_email_message_details

def run_pipeline(client_secret_file, selected_date):
    # Step 1: Initialize Gmail API service
    API_SERVICE_NAME = 'gmail'
    API_VERSION = 'v1'
    SCOPES = ['https://mail.google.com/']
    service = create_service(client_secret_file, API_SERVICE_NAME, API_VERSION, SCOPES)
    print(f"Gmail API service created successfully for {client_secret_file}")

    # Step 2: Fetch all messages
    messages = get_email_messages(service)
    print(f"Total messages fetched: {len(messages)}")

    # Step 3: Fetch details for each message
    message_details = []
    for msg in messages:
        details = get_email_message_details(service, msg['id'])
        if details:
            message_details.append(details)

    # Step 4: Filter messages by selected date
    filtered_messages = []
    for details in message_details:
        email_date_str = details["date"].replace(" GMT", " +0000")
        try:
            email_date_obj = datetime.strptime(email_date_str, "%a, %d %b %Y %H:%M:%S %z")
            email_date = email_date_obj.strftime("%Y-%m-%d")
            if email_date == selected_date:
                filtered_messages.append(details)
        except ValueError as e:
            print(f"Error parsing date: {email_date_str}")
            print(f"Exception: {e}")

    # Step 5: Extract transactions
    transactions = []
    send_pattern = r"You sent \$(\d+\.?\d*) to ([A-Za-z ]+)"
    receive_pattern = r"([A-Za-z ]+) sent you \$(\d+\.?\d*)"

    total_in = 0.0
    total_out = 0.0

    for details in filtered_messages:
        subject = details['subject']
        send_match = re.match(send_pattern, subject)
        receive_match = re.match(receive_pattern, subject)

        if send_match:
            amount = float(send_match.group(1))
            transactions.append(("You", send_match.group(2), amount, "Out"))
            total_out += amount
        elif receive_match:
            amount = float(receive_match.group(2))
            transactions.append((receive_match.group(1), "You", amount, "In"))
            total_in += amount

    # Calculate profit/loss
    profit_loss = total_in - total_out

    # Step 6: Save transactions to a unique Excel file inside the 'static' folder
    static_folder = os.path.join(os.path.dirname(__file__), "static")
    os.makedirs(static_folder, exist_ok=True)  # Ensure 'static' folder exists

    client_name = os.path.splitext(os.path.basename(client_secret_file))[0]  # Extract file name without extension
    file_name = os.path.join(static_folder, f"transactions_{client_name}.xlsx")  # Save inside 'static' folder

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"
    ws.append(["Sender", "Receiver", "Amount", "In/Out"])

    for transaction in transactions:
        ws.append(transaction)

    # Add total in/out and profit/loss to the sheet
    ws.append([])
    ws.append(["Total In", total_in])
    ws.append(["Total Out", total_out])
    ws.append(["Profit/Loss", profit_loss])

    wb.save(file_name)  # Save in 'static' folder
    print(f"Transactions saved to {file_name}")

    # Step 7: Save filtered messages as a pickle file inside 'static' folder
    pickle_file = os.path.join(static_folder, f"filtered_messages_{client_name}.pkl")
    with open(pickle_file, "wb") as f:
        pickle.dump(filtered_messages, f)

    return transactions, total_in, total_out, profit_loss
