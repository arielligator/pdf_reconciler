import os
import re
import csv
from os.path import join
import pdfplumber  # Library to extract text and tables from PDFs
from rapidfuzz import fuzz  # Library for fuzzy string matching
from itertools import combinations

# File paths (update these with your actual file locations)
pdf_path = "your file path"          # Path to the PDF file to process
csv_directory = "your file path"     # Directory containing CSV files for matching
output_csv = "your file path"        # Path where the output CSV will be saved

# ---------------------------- PDF Extraction ----------------------------

def extract_pdf(pdf_path):
    """
    Extracts table data from a PDF and returns a list of dictionaries,
    where each dictionary represents a row of data with cleaned fields.
    """
    pdf_data = []  # List to store extracted PDF records
    # Open the PDF file using pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        # Iterate over each page in the PDF
        for page in pdf.pages:
            # Attempt to extract a table from the current page
            table = page.extract_table()
            # Only process pages that contain table data
            if table:
                # Skip the header row (assumes first row is header)
                data_rows = table[1:]
                # Process each row in the table
                for row in data_rows:
                    # Only process rows that have at least 11 columns
                    if len(row) >= 11:
                        # Clean and extract the description field (assumes index 4 holds the description)
                        original_description = row[4].strip() if row[4] else ""
                        # Remove any unwanted prefix using a regular expression
                        cleaned_description = re.sub(
                            r"^\d{8}\s+Subscription\s+#\d{7}:\s+",
                            "",
                            original_description,
                        )

                        # Extract and clean the customer field (assumes index 1 holds customer info)
                        original_customer = row[1].strip() if row[1] else ""
                        # Remove a 10-digit prefix and any newline characters
                        cleaned_customer = re.sub(
                            r"^\d{10}\s*|[\n\r]+", " ", original_customer
                        )

                        # Extract and clean the quantity field (assumes index 8 holds quantity)
                        original_qty = row[8].strip() if row[8] else ""
                        # Remove any non-numeric characters (except decimal point)
                        cleaned_qty = re.sub(r"[^\d.]", "", original_qty)
                        # Convert quantity to an integer (default to 0 if empty)
                        cleaned_qty = int(float(cleaned_qty)) if cleaned_qty else 0
                        cleaned_qty = cleaned_qty if cleaned_qty else "0"  # default to "0" if empty

                        # Extract and clean the net unit price (assumes index 9 holds the price)
                        original_price = row[9].strip() if row[9] and row[9].strip() else "0.00"
                        # Remove commas and convert to a float
                        cleaned_price = float(original_price.replace(",", ""))

                        # Extract and clean the total amount (assumes index 10 holds the total)
                        total_amount = row[10].strip() if row[10] and row[10].strip() else "0.00"
                        cleaned_total_amount = float(total_amount.replace(",", ""))

                        # Append the cleaned record as a dictionary to the pdf_data list
                        pdf_data.append(
                            {
                                "Number": row[0],  # Record identifier
                                "End-Customer": cleaned_customer,
                                "Description": cleaned_description,
                                "Net Unit Price": cleaned_price,
                                "Qty": int(float(cleaned_qty)),
                                "Total Amount": cleaned_total_amount,
                                "SO/PO Number": row[5].strip() if row[5] else "",  # Sales/PO number
                            }
                        )
    return pdf_data

# ---------------------------- String Normalization ----------------------------

def normalize_string(s):
    """
    Normalize a string by converting it to lowercase and removing any
    non-alphanumeric characters (except spaces) to standardize it for matching.
    """
    return re.sub(r'[^a-zA-Z0-9\s]', '', s).strip().lower()

# Note: There were two definitions of normalize_string. The second one overwrites the first,
# so we keep only one consistent function for normalization.

# ---------------------------- CSV Matching ----------------------------

def find_matching_csv_files(company_name, csv_directory):
    """
    Finds CSV files in the specified directory that match the given company name.
    Uses both exact matching and fuzzy matching (with a threshold of 80).
    Returns a list containing at most one matched CSV file path.
    """
    matching_files = []  # List to collect potential matching CSV file paths
    # Normalize the company name for comparison
    normalized_company_name = normalize_string(company_name)

    print(f"Looking for a match for: {company_name} (normalized: {normalized_company_name})")

    # Iterate over all files in the CSV directory
    for filename in os.listdir(csv_directory):
        # Process only CSV files
        if filename.endswith(".csv"):
            # Remove the file extension and normalize the filename
            base_name = os.path.splitext(filename)[0]
            normalized_filename = normalize_string(base_name)

            print(f"Checking file: {filename} (normalized: {normalized_filename})")

            # First, try to find an exact match between the normalized company name and filename
            if normalized_company_name == normalized_filename:
                print(f"Exact match found: {filename}")
                return [os.path.join(csv_directory, filename)]
            
            # Remove trailing numbers from the filename for improved fuzzy matching
            base_without_number = re.sub(r"\s*\d+$", "", normalized_filename)

            # Use fuzzy matching with a threshold of 80 for flexibility
            if fuzz.ratio(normalized_company_name, normalized_filename) > 80 or \
               fuzz.ratio(normalized_company_name, base_without_number) > 80:
                print(f"Fuzzy match found: {filename}")
                matching_files.append(os.path.join(csv_directory, filename))
    
    # If multiple matching files are found, warn and return the first match
    if len(matching_files) > 1:
        print(f"Warning: Multiple CSV files found for {company_name}. Returning the first match.")

    # If no matches were found, log that information
    if not matching_files:
        print(f"No matching files found for {company_name}")

    # Return only the first matching file (or an empty list if none were found)
    return matching_files[:1]

    # Note: The code below this return statement is unreachable and can be removed.

def load_csv_data(csv_file_paths):
    """
    Loads CSV data from a list of CSV file paths.
    Returns a list of dictionaries, where each dictionary represents a CSV row.
    """
    csv_data = []
    # Iterate over each file path in the provided list
    for file_path in csv_file_paths:
        with open(file_path, newline='', encoding='utf-8') as csvfile:
            # Use DictReader to read CSV rows as dictionaries
            reader = csv.DictReader(csvfile)
            # Strip any leading or trailing whitespace from the CSV header names
            reader.fieldnames = [name.strip() for name in reader.fieldnames]
            # Append each row dictionary to the csv_data list
            for row in reader:
                csv_data.append(row)
    return csv_data

def extract_matching_csv_data(company_name, csv_directory):
    """
    Extracts CSV data for a given company name by finding the matching CSV file.
    If exactly one CSV file matches, its data is loaded and returned.
    If there is no match or multiple matches, an empty list is returned.
    """
    matching_files = find_matching_csv_files(company_name, csv_directory)
    
    if len(matching_files) == 1:
        # Load and return CSV data from the single matched file
        return load_csv_data(matching_files)
    else:
        # Log a warning if there isn't exactly one matching CSV file
        print(f"Warning: Expected one CSV file for {company_name}, found {len(matching_files)}.")
        return []

def find_matching_additions(pdf_record, csv_data):
    """
    Compares a single PDF record with CSV data rows to find matching records.
    It uses the 'Net Unit Price' and 'Qty' for an exact match and
    fuzzy string matching for the description.
    Returns a list of matches sorted by the relevance score (highest first).
    """
    matches = []
    # Normalize the PDF description for fuzzy matching
    pdf_description = normalize_string(pdf_record["Description"])
    pdf_cost = pdf_record["Net Unit Price"]
    pdf_qty = pdf_record["Qty"]

    # Define which CSV fields to keep when a match is found
    selected_fields = ["Description", "Unit Cost", "Total Quantity", "Agreement Name"]
    
    # Iterate over each CSV row to find potential matches
    for row in csv_data:
        # Ensure keys in the CSV row are trimmed of whitespace
        row = {key.strip(): value for key, value in row.items()}
        # Normalize the CSV description (default text if missing)
        csv_description = normalize_string(row.get("Description", "No Description Found"))
        
        # Dynamically determine which column contains cost and quantity info
        cost_column = next((col for col in row.keys() if "cost" in col.lower()), None)
        qty_column = next((col for col in row.keys() if "quantity" in col.lower()), None)
        agreement_name = row.get("Agreement Name", "No Agreement Found")
       
        # Retrieve cost and quantity values from CSV, removing commas if needed
        csv_cost = row.get(cost_column, "0").replace(",", "") if cost_column else "0"
        csv_qty = row.get(qty_column, "0").replace(",", "") if qty_column else "0"
        # Convert CSV quantity to an integer safely
        csv_qty = int(float(csv_qty))
        
        # Clean the PDF cost to ensure correct float comparison
        pdf_cost_cleaned = float(str(pdf_cost).replace(",", ""))
        
        # Check if both cost and quantity match exactly between PDF and CSV record
        if float(csv_cost) == float(pdf_cost) and int(csv_qty) == int(pdf_qty):
            # Compute the fuzzy matching score between PDF and CSV descriptions
            relevance = fuzz.ratio(pdf_description, csv_description)
            # Filter CSV data to only include selected fields for the output
            filtered_data = {key: value for key, value in row.items() if key in selected_fields}
            # Append the match along with its relevance score
            matches.append({"data": filtered_data, "relevance": relevance})
    
    # Return matches sorted by descending relevance (highest score first)
    return sorted(matches, key=lambda x: x["relevance"], reverse=True)

def get_pdf_matches(pdf_path, csv_directory):
    """
    Processes the PDF file to extract records, then for each record, finds the corresponding
    CSV data based on the 'End-Customer' field. It then finds matching CSV rows based on cost,
    quantity, and fuzzy matching of descriptions.
    Returns a list of dictionaries with both PDF data and its corresponding matches.
    """
    pdf_data = extract_pdf(pdf_path)
    matches = []
    # Iterate through each record extracted from the PDF
    for pdf_record in pdf_data:
        # Use the 'End-Customer' field to determine which CSV file to search
        company_name = pdf_record["End-Customer"]
        csv_data = extract_matching_csv_data(company_name, csv_directory)
        # Find matching additions between the PDF record and CSV data
        matching_records = [
            {"score": m["relevance"], "data": m["data"]}
            for m in find_matching_additions(pdf_record, csv_data)
        ]
        # Append the PDF record and its matches to the results list
        matches.append({ "pdf": pdf_record, "matches": matching_records })
    return matches

def find_unmatched_pdf_records(pdf_path, csv_directory):
    """
    Identifies and returns PDF records that do not have any matching CSV record.
    Useful for troubleshooting missing or incomplete matches.
    """
    pdf_data = extract_pdf(pdf_path)
    unmatched_records = []
    # Iterate through each PDF record
    for record in pdf_data:
        company_name = record["End-Customer"]
        csv_data = extract_matching_csv_data(company_name, csv_directory)
        matching_records = find_matching_additions(record, csv_data)
        # If no matching CSV records are found, add the PDF record to the unmatched list
        if not matching_records:
            unmatched_records.append(record)
    return unmatched_records

# ---------------------------- Writing Results to CSV ----------------------------

# Get the list of matches by comparing PDF data to CSV data
matches = get_pdf_matches(pdf_path, csv_directory)

# Define CSV headers for the output file
headers = [
    "Number", "End-Customer", "Description", "Qty", "Net Unit Price",
    "Total Amount", "SO/PO Number", "Match Description", "Match Cost",
    "Match Quantity", "Agreement Name", "Match Score"
]

# Set to track seen record numbers to avoid writing duplicate rows
seen_numbers = set()

# Open the output CSV file for writing the results
with open(output_csv, mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(headers)  # Write the header row
    
    # Iterate over each PDF record and its matches
    for entry in matches:
        pdf = entry['pdf']
        matches_for_record = entry['matches']
        number = pdf['Number']
        
        # Skip records that have already been processed (avoid duplicates)
        if number in seen_numbers:
            continue
        seen_numbers.add(number)
        
        # If there are matching CSV records, write each match as a separate row
        if matches_for_record:
            for match in matches_for_record:
                row = [
                    pdf['Number'],
                    pdf['End-Customer'],
                    pdf['Description'],
                    pdf['Qty'],
                    pdf['Net Unit Price'],
                    pdf['Total Amount'],
                    pdf['SO/PO Number'],
                    match['data']['Description'],
                    match['data']['Unit Cost'],
                    match['data']['Total Quantity'],
                    match['data']['Agreement Name'],
                    match['score']
                ]
                writer.writerow(row)
        else:
            # If no match was found, write the PDF record with blank fields for CSV data
            row = [
                pdf['Number'],
                pdf['End-Customer'],
                pdf['Description'],
                pdf['Qty'],
                pdf['Net Unit Price'],
                pdf['Total Amount'],
                pdf['SO/PO Number'],
                "", "", "", ""
            ]
            writer.writerow(row)
