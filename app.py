# File: maillage_niveaux.py

import os
import pandas as pd
from urllib.parse import urljoin

def detect_subcategories(url, urls):
    """
    Detects subcategories for a given URL by finding URLs that extend it with a single path.
    """
    base_path = url.rstrip("/")
    subcategories = [
        candidate for candidate in urls
        if candidate.startswith(base_path + "/") and candidate[len(base_path + "/"):].count("/") == 0
    ]
    return subcategories

def generate_links(data):
    """
    Generate links for N+1 categories based on the provided URLs.
    """
    urls = data["URL"].tolist()
    results = []

    for url in urls:
        subcategories = detect_subcategories(url, urls)
        for subcat in subcategories:
            anchor = subcat.rsplit("/", 1)[-1]
            results.append({"Original URL": url, "Generated URL": subcat, "Anchor Text": anchor})

    return pd.DataFrame(results)

def main(input_file, output_file):
    """
    Main function to read the input file, process the URLs, and export the results.
    """
    try:
        # Read the Excel file
        data = pd.read_excel(input_file)
        if "URL" not in data.columns:
            raise ValueError("The uploaded file must contain a column named 'URL'.")

        # Generate N+1 links
        result_data = generate_links(data)

        # Export results to Excel
        result_data.to_excel(output_file, index=False)
        print(f"Generated links saved to {output_file}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    # Input and output file paths
    input_file = "Test_Maillage.xlsx"  # Replace with your input file name
    output_file = "Generated_Links.xlsx"

    # Run the script
    main(input_file, output_file)
