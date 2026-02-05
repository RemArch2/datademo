import pandas as pd
import os
import sys

def analyze_je():
    file_path = 'je_samples.xlsx'

    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        sys.exit(1)

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Basic Statistics
    row_count = len(df)
    unique_je_numbers = df['JENumber'].nunique()

    # Date Ranges
    effective_date_min = df['EffectiveDate'].min()
    effective_date_max = df['EffectiveDate'].max()
    entry_date_min = df['EntryDate'].min()
    entry_date_max = df['EntryDate'].max()

    # Amount Statistics
    # Ensure numeric columns are actually numeric, handling non-numeric gracefully if needed
    # Based on previous inspection, Debit is object, Credit is float, Amount is float.
    # I'll convert Debit to numeric, coercing errors to NaN just in case.

    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')

    amount_stats = df['Amount'].describe()
    debit_sum = df['Debit'].sum()
    credit_sum = df['Credit'].sum()
    amount_sum = df['Amount'].sum()

    # Prepare Output
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)

    output_file = os.path.join(output_dir, 'analysis_report.txt')

    with open(output_file, 'w') as f:
        f.write("JE Samples Analysis Report\n")
        f.write("==========================\n\n")

        f.write(f"Total Row Count: {row_count}\n")
        f.write(f"Unique JE Numbers: {unique_je_numbers}\n\n")

        f.write("Date Ranges:\n")
        f.write(f"  EffectiveDate: {effective_date_min} to {effective_date_max}\n")
        f.write(f"  EntryDate:     {entry_date_min} to {entry_date_max}\n\n")

        f.write("Financial Statistics:\n")
        f.write(f"  Total Debit:   {debit_sum:,.2f}\n")
        f.write(f"  Total Credit:  {credit_sum:,.2f}\n")
        f.write(f"  Total Amount:  {amount_sum:,.2f}\n\n")

        f.write("Amount Column Statistics:\n")
        f.write(f"  Count: {amount_stats['count']}\n")
        f.write(f"  Mean:  {amount_stats['mean']:,.2f}\n")
        f.write(f"  Std:   {amount_stats['std']:,.2f}\n")
        f.write(f"  Min:   {amount_stats['min']:,.2f}\n")
        f.write(f"  25%:   {amount_stats['25%']:,.2f}\n")
        f.write(f"  50%:   {amount_stats['50%']:,.2f}\n")
        f.write(f"  75%:   {amount_stats['75%']:,.2f}\n")
        f.write(f"  Max:   {amount_stats['max']:,.2f}\n")

    print(f"Analysis complete. Report saved to {output_file}")

if __name__ == "__main__":
    analyze_je()
