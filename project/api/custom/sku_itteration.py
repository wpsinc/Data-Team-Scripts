import requests
import json
import concurrent.futures
from dotenv import load_dotenv
import os

skus = [''] # Add your SKUs here. Ex: '<number>',

load_dotenv()

BEARER_TOKEN = os.getenv('BEARER_TOKEN')

url = "https://api.wps-inc.com/products?include=items&filter[sku]="

headers = {
    'Authorization': BEARER_TOKEN
    # Add BEARER_TOKEN to .env file for additional security to protect the private token
}

# Function to send GET request
def send_request(sku):
    response = requests.get(url + sku, headers=headers)
    response_json = response.json()
    response_size = len(json.dumps(response_json).encode('utf-8')) / (1024 * 1024)  # Convert bytes to megabytes
    return response_json, response_size

# Initialize an empty list to store all responses
all_responses = []

# Initialize a variable to store the total response size
total_response_size = 0

# Use ThreadPoolExecutor to send multiple requests at a time
with concurrent.futures.ThreadPoolExecutor() as executor:
    future_to_sku = {executor.submit(send_request, sku): sku for sku in skus}
    for future in concurrent.futures.as_completed(future_to_sku):
        response, response_size = future.result()
        all_responses.append(response)
        total_response_size += response_size

# Print the total response size
print(f'Total response size: {total_response_size} MB')

# Replace the line where you open the file with these lines
script_dir = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(script_dir, 'response.json')

with open(output_path, 'w') as f:
    json.dump(all_responses, f)