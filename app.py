from flask import Flask, render_template, request
import pickle
from pipeline import run_pipeline

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    transactions = []
    selected_date = None
    if request.method == "POST":
        client_secret_files = ["client_secret_1.json", "client_secret_2.json"]
        selected_date = request.form.get('transaction_date')
        
        # Loop through both client secret files and run the pipeline for each
        for client_secret_file in client_secret_files:
            transactions_for_file, _, _, _ = run_pipeline(client_secret_file , selected_date)
            print(selected_date)
            transactions.extend(transactions_for_file)
        
        # Save all transactions to a single pickle file
        with open("transactions.pkl", "wb") as f:
            pickle.dump(transactions, f)
    
    return render_template("index.html", transactions=transactions)

if __name__ == "__main__":
    app.run(debug=False)
