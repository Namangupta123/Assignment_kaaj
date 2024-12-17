# Importing necessary libraries
import os
import json
import pandas as pd
import matplotlib.pyplot as plt
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
import re
from dotenv import load_dotenv
load_dotenv()

# Azure endpoint and key
endpoint = os.getenv('AZURE_ENDPOINT')
key = os.getenv('AZURE_API_KEY')

# Path to folder containing bank statements (PDFs) and output folder
input_folder = "Bank_statements"
output_folder = "output_folder"

# Initialize Document Analysis Client
document_analysis_client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(key))

def safe_float(value):
    """
    Attempt to convert a value to float. If it fails, return 0.0.
    """
    try:
        return float(value.replace(",", "").replace("$", ""))
    except ValueError:
        return 0.0

def analyze_bank_statement(file_path):
    """
    Analyze a single bank statement and extract relevant information.
    """
    with open(file_path, "rb") as document:
        print(f"Analyzing document: {file_path}")
        try:
            poller = document_analysis_client.begin_analyze_document(
                model_id="prebuilt-document", document=document
            )
            result = poller.result()

            # Initialize variables
            transactions = []
            starting_balance = None
            ending_balance = None

            # First pass: Look for balances in key-value pairs
            for kv_pair in result.key_value_pairs:
                key = kv_pair.key.content.lower() if kv_pair.key else ""
                value = kv_pair.value.content if kv_pair.value else ""

                # Starting balance detection
                if re.search(r'\b(starting|beginning|previous|opening|prior|initial)\b.*\b(balance|bal)\b', key):
                    starting_balance = safe_float(value)
                
                # Ending balance detection
                if re.search(r'\b(ending|current|closing|new|final)\b.*\b(balance|bal)\b', key):
                    ending_balance = safe_float(value)

            # Second pass: Look for balances in tables if not found
            if starting_balance is None or ending_balance is None:
                for table in result.tables:
                    for cell in table.cells:
                        cell_text = cell.content.lower()
                        
                        # Look for starting balance
                        if starting_balance is None and re.search(r'\b(starting|beginning|previous|opening|prior|initial)\b.*\b(balance|bal)\b', cell_text):
                            for adjacent_cell in table.cells:
                                if (adjacent_cell.row_index == cell.row_index and 
                                    adjacent_cell.column_index in [cell.column_index + 1, cell.column_index - 1]):
                                    potential_value = safe_float(adjacent_cell.content)
                                    if potential_value != 0.0:
                                        starting_balance = potential_value
                                        break

                        # Look for ending balance
                        if ending_balance is None and re.search(r'\b(ending|current|closing|new|final)\b.*\b(balance|bal)\b', cell_text):
                            for adjacent_cell in table.cells:
                                if (adjacent_cell.row_index == cell.row_index and 
                                    adjacent_cell.column_index in [cell.column_index + 1, cell.column_index - 1]):
                                    potential_value = safe_float(adjacent_cell.content)
                                    if potential_value != 0.0:
                                        ending_balance = potential_value
                                        break

            # Third pass: If still no starting balance, try to find it in the first row of transaction tables
            if starting_balance is None:
                for table in result.tables:
                    if len(table.cells) > 0:
                        first_row_cells = [cell for cell in table.cells if cell.row_index == 0]
                        for cell in first_row_cells:
                            potential_value = safe_float(cell.content)
                            if potential_value != 0.0:
                                starting_balance = potential_value
                                break

            # If still no ending balance, try to find it in the last row
            if ending_balance is None:
                for table in result.tables:
                    if len(table.cells) > 0:
                        max_row = max(cell.row_index for cell in table.cells)
                        last_row_cells = [cell for cell in table.cells if cell.row_index == max_row]
                        for cell in last_row_cells:
                            potential_value = safe_float(cell.content)
                            if potential_value != 0.0:
                                ending_balance = potential_value
                                break

            # Set default values if still not found
            if starting_balance is None:
                starting_balance = 0.0
                print(f"Warning: Could not find starting balance in {file_path}. Using 0.0 as default.")
            
            if ending_balance is None:
                ending_balance = 0.0
                print(f"Warning: Could not find ending balance in {file_path}. Using 0.0 as default.")

            # Extract transactions
            transactions = []
            for table in result.tables:
                for cell in table.cells:
                    if cell.row_index > 0:
                        row = [c.content for c in table.cells if c.row_index == cell.row_index]
                        if len(row) >= 3:
                            transactions.append({
                                "Date": row[0],
                                "Description": row[1],
                                "Amount": safe_float(row[2]),
                                "Direction": "Credit" if safe_float(row[2]) > 0 else "Debit"
                            })

            return {
                "starting_balance": starting_balance,
                "ending_balance": ending_balance,
                "transactions": transactions
            }

        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")
            return {
                "starting_balance": 0.0,
                "ending_balance": 0.0,
                "transactions": []
            }

def plot_variance_graph(variances, output_folder):
    """
    Create and save a variance graph showing balance discrepancies across all statements.
    """
    plt.figure(figsize=(14, 8))  # Increased size for better visibility
    
    # Extract data and sort by variance amount for better visualization
    files = [v['file'].replace('.pdf', '') for v in variances]  # Remove .pdf extension for cleaner labels
    variance_values = [v['variance'] for v in variances]
    
    # Create bars with different colors based on positive/negative values
    bars = plt.bar(files, variance_values, width=0.6)  # Adjusted bar width
    
    # Color the bars based on variance (red for negative, green for positive)
    for bar, variance in zip(bars, variance_values):
        if variance < 0:
            bar.set_color('#ff6b6b')  # Red for negative variance
        else:
            bar.set_color('#4ecdc4')  # Green for positive variance
    
    # Improve graph styling
    plt.title('Balance Variances Across Bank Statements', pad=20, fontsize=16, fontweight='bold')
    plt.xlabel('Statement Files', labelpad=10, fontsize=14)
    plt.ylabel('Variance Amount ($)', labelpad=10, fontsize=14)
    
    # Rotate labels for better readability
    plt.xticks(rotation=45, ha='right', fontsize=12)
    
    # Add horizontal line at y=0 to clearly show positive/negative split
    plt.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
    
    # Add grid for better readability
    plt.grid(True, axis='y', linestyle='--', alpha=0.3)
    
    # Add value labels on top of each bar
    for i, v in enumerate(variance_values):
        label_position = v + (0.1 if v >= 0 else -0.1)
        plt.text(i, label_position, f'${v:,.2f}', 
                ha='center', 
                va='bottom' if v >= 0 else 'top',
                fontsize=10)
    
    # Adjust layout to prevent label cutoff
    plt.tight_layout()
    
    # Save with high DPI for better quality
    plt.savefig(os.path.join(output_folder, "variance_graph.png"), dpi=300, bbox_inches='tight')
    plt.close()

def process_folder(input_folder, output_folder):
    """
    Process all PDFs in a folder, save results to CSV, generate output.json,
    create discrepancy report and variance graph.
    """
    all_data = []
    output_json_data = {}
    discrepancies = []
    variances = []

    for file_name in os.listdir(input_folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(input_folder, file_name)
            print(f"Processing file: {file_name}")

            data = analyze_bank_statement(file_path)
            if data:
                # Calculate total transactions amount
                total_transactions = sum(
                    t["Amount"] if t["Direction"] == "Credit" else -t["Amount"] 
                    for t in data["transactions"]
                )
                
                # Calculate expected ending balance
                expected_ending = data["starting_balance"] + total_transactions
                actual_ending = data["ending_balance"]
                variance = actual_ending - expected_ending

                # Record variance
                variances.append({
                    "file": file_name,
                    "variance": variance
                })

                # Check for discrepancy
                if abs(variance) > 0.01:
                    discrepancies.append({
                        "file": file_name,
                        "starting_balance": data["starting_balance"],
                        "total_transactions": total_transactions,
                        "expected_ending": expected_ending,
                        "actual_ending": actual_ending,
                        "variance": variance
                    })

                all_data.append({
                    "File": file_name,
                    "Starting Balance": data["starting_balance"],
                    "Ending Balance": data["ending_balance"],
                    "Transactions": len(data["transactions"])
                })

                # Add to output JSON data
                output_json_data[file_name] = {
                    "Starting Balance": data["starting_balance"],
                    "Ending Balance": data["ending_balance"],
                    "Transactions": data["transactions"]
                }

    # Save data to CSV
    df = pd.DataFrame(all_data)
    df.to_csv(os.path.join(output_folder, "extracted_transactions.csv"), index=False)

    # Save JSON output
    with open(os.path.join(output_folder, "output.json"), "w") as json_file:
        json.dump(output_json_data, json_file, indent=4)

    # Create discrepancy report
    with open(os.path.join(output_folder, "discrepancy_report.txt"), "w") as report:
        report.write("Bank Statement Discrepancy Report\n")
        report.write("==============================\n\n")
        
        if discrepancies:
            for d in discrepancies:
                report.write(f"File: {d['file']}\n")
                report.write(f"Starting Balance: ${d['starting_balance']:,.2f}\n")
                report.write(f"Total Transactions: ${d['total_transactions']:,.2f}\n")
                report.write(f"Expected Ending Balance: ${d['expected_ending']:,.2f}\n")
                report.write(f"Actual Ending Balance: ${d['actual_ending']:,.2f}\n")
                report.write(f"Variance: ${d['variance']:,.2f}\n")
                report.write("------------------------------\n\n")
        else:
            report.write("No discrepancies found in any statements.\n")

    # Create variance graph
    plot_variance_graph(variances, output_folder)

    print("Results saved to output.json, extracted_transactions.csv, discrepancy_report.txt, and variance_graph.png")

# Run the processing
process_folder(input_folder, output_folder)