import os
import json
import fitz  # PyMuPDF
import pandas as pd
from tqdm import tqdm
from openai import OpenAI
from dotenv import load_dotenv



load_dotenv()

# Define the folder containing the PDF files
folder_path = 'FundFactSheet'

# Initialize an empty list to store product data
product_data = []

# Set OpenAI client
openai_client = OpenAI(api_key=os.getenv('BATI_OPENAI_API_KEY'))

# Function to extract information using OpenAI API
def extract_info_with_openai(text):
    prompt=f"""
    Extract the following information from the text as JSON: {text}\n\n
    
    response format:
    ```json
    {{
        "Product Name": "",
        "Fund Category": "",
        "Effective Date": "",
        "Currency": "",
        "Minimum Initial Subscription": "",
        "Valuation Period": "",
        "Subscription Fee": "",
        "Redemption Fee": "",
        "Switching Fee": "",
        "Management Fee": "",
        "Custodian Bank": "",
        "Custodian Fee": "",
        "ISIN Code": "",
        "Bloomberg Ticker": "",
        "Benchmark": "",
        "Risk Factor": "",
        "Risk Level": "",
        "Top Holdings": "",
        "Investment Policy": "",
        "Asset Allocation as of Reporting Date": "",
        "1 Month Return": "",
        "3 Month Return": "",
        "6 Month Return": "",
        "YTD": "",
        "1 Year Return": "",
        "3 Year Return": "",
        "5 Year Return": "",
        "Since Inception": ""
    }}
    ```
    Note: 
    - All output should be in English.
    - Risk Level should be one of the following: Low, Medium, High.
    - Risk Factor: list all.
    - Top Holdings: list all.
    - Fund Category should be one of the following: Balanced, Index, Money Market, Fixed income, Equity.
    - Date format example: 16 Feb 2027
    - Currency should be either USD or IDR
    - Use 'None' for empty data
    - Use uniform terms: 'per annum', 'per transaction'
    """
    # Call the model
    completion = openai_client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user", 
                "content": [
                    {"type": "text", "text": prompt},
                    ],
            }
        ],
        temperature=0
    )  

    # Initialize a variable to store JSON block lines
    json_blocks = []
    current_json_block = None

    # Simulate stream of response with milliseconds delay
    accumulated_response = completion.choices[0].message.content
    for line in accumulated_response.split('\n'):
        # Check if the line starts a JSON block (```json indicates start of JSON data)
        if line.strip().startswith('```json'):
            current_json_block = []
            continue
        elif line.strip().startswith('```') and current_json_block is not None:
            json_blocks.append('\n'.join(current_json_block))
            current_json_block = None
            continue
        
        # Collect lines if they are part of a JSON block
        if current_json_block is not None:
            current_json_block.append(line)
            continue

    return json_blocks

# Iterate over each PDF file in the folder
for filename in tqdm(os.listdir(folder_path)):

    if filename.endswith('.pdf'):
        file_path = os.path.join(folder_path, filename)
        
        # Open the PDF
        doc = fitz.open(file_path)

        # Extract text from each page
        pdf_text = ''
        for page in doc:
            pdf_text += page.get_text("text")

        # # Save the extracted text to a .txt file
        # txt_filename = 'test.txt'
        # with open(txt_filename, 'w', encoding='utf-8') as txt_file:
        #     txt_file.write(pdf_text)

        # Use OpenAI API to extract information
        json_blocks = extract_info_with_openai(pdf_text)
        
        # Convert the extracted information string to a JSON object
        extracted_info = json.loads(json_blocks[0])
        
        # Process the extracted information into a dictionary
        product_info = {
            'Product Name': extracted_info.get('Product Name', ''),
            'Fund Category': extracted_info.get('Fund Category', ''),
            'Effective Date': extracted_info.get('Effective Date', ''),
            'Currency': extracted_info.get('Currency', ''),
            'Minimum Initial Subscription': extracted_info.get('Minimum Initial Subscription', ''),
            'Valuation Period': extracted_info.get('Valuation Period', ''),
            'Subscription Fee': extracted_info.get('Subscription Fee', ''),
            'Redemption Fee': extracted_info.get('Redemption Fee', ''),
            'Switching Fee': extracted_info.get('Switching Fee', ''),
            'Management Fee': extracted_info.get('Management Fee', ''),
            'Custodian Bank': extracted_info.get('Custodian Bank', ''),
            'Custodian Fee': extracted_info.get('Custodian Fee', ''),
            'ISIN Code': extracted_info.get('ISIN Code', ''),
            'Bloomberg Ticker': extracted_info.get('Bloomberg Ticker', ''),
            'Benchmark': extracted_info.get('Benchmark', ''),
            'Risk Factor': extracted_info.get('Risk Factor', ''),
            'Risk Level': extracted_info.get('Risk Level', ''),
            'Top Holdings': extracted_info.get('Top Holdings', ''),
            'Investment Policy': extracted_info.get('Investment Policy', ''),
            'Asset Allocation as of Reporting Date': extracted_info.get('Asset Allocation as of Reporting Date', ''),
            '1 Month Return': extracted_info.get('1 Month Return', ''),
            '3 Month Return': extracted_info.get('3 Month Return', ''),
            '6 Month Return': extracted_info.get('6 Month Return', ''),
            'YTD': extracted_info.get('YTD', ''),
            '1 Year Return': extracted_info.get('1 Year Return', ''),
            '3 Year Return': extracted_info.get('3 Year Return', ''),
            '5 Year Return': extracted_info.get('5 Year Return', ''),
            'Since Inception': extracted_info.get('Since Inception', '')
        }
        
        # Append the extracted information to the list
        product_data.append(product_info)

# Create a DataFrame from the list of product data
df = pd.DataFrame(product_data)

# Save the DataFrame to an Excel file
df.to_excel('product_data.xlsx', index=False)


