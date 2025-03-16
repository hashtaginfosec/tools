import requests
import json
import os
import argparse

def get_api_key():
    """Retrieve API key from environment variable or configuration file."""
    api_key = os.getenv("GREYNOISE_API_KEY")
    if not api_key:
        try:
            with open(".conf", "r") as conf_file:
                api_key = conf_file.read().strip()
        except FileNotFoundError:
            raise ValueError("API key not found. Set GREYNOISE_API_KEY or add it to .conf")
    
    if not api_key:
        raise ValueError("API key not found in environment variable or .conf file")
    
    return api_key

def fetch_cve_data(cve_id, api_key):
    """Fetch CVE data from the Greynoise API."""
    url = f"https://api.greynoise.io/v1/cve/{cve_id}"
    headers = {
        "accept": "application/json",
        "key": api_key
    }
    response = requests.get(url, headers=headers)
    return response.json()

def format_output(json_data):
    """Format the API response in a readable format."""
    output = []
    output.append(f"CVE ID: {json_data.get('id', 'N/A')}")
    
    details = json_data.get("details", {})
    output.append("\nDetails:")
    output.append(f"  Name: {details.get('vulnerability_name', 'N/A')}")
    output.append(f"  Description: {details.get('vulnerability_description', 'N/A')}")
    output.append(f"  CVSS Score: {details.get('cve_cvss_score', 'N/A')}")
    output.append(f"  Product: {details.get('product', 'N/A')}")
    output.append(f"  Vendor: {details.get('vendor', 'N/A')}")
    output.append(f"  Published to NIST NVD: {details.get('published_to_nist_nvd', 'N/A')}")
    
    timeline = json_data.get("timeline", {})
    output.append("\nTimeline:")
    output.append(f"  CVE Published Date: {timeline.get('cve_published_date', 'N/A')}")
    output.append(f"  Last Updated Date: {timeline.get('cve_last_updated_date', 'N/A')}")
    output.append(f"  First Known Published Date: {timeline.get('first_known_published_date', 'N/A')}")
    output.append(f"  CISA KEV Date Added: {timeline.get('cisa_kev_date_added', 'N/A')}")
    
    exploitation = json_data.get("exploitation_details", {})
    output.append("\nExploitation Details:")
    output.append(f"  Attack Vector: {exploitation.get('attack_vector', 'N/A')}")
    output.append(f"  Exploit Found: {exploitation.get('exploit_found', 'N/A')}")
    output.append(f"  Exploitation Registered in KEV: {exploitation.get('exploitation_registered_in_kev', 'N/A')}")
    output.append(f"  EPSS Score: {exploitation.get('epss_score', 'N/A')}")
    
    return "\n".join(output)

def main():
    parser = argparse.ArgumentParser(description="Fetch CVE details from Greynoise API")
    parser.add_argument("cve_id", help="The CVE ID to fetch details for (e.g., CVE-2019-19781)")
    args = parser.parse_args()
    
    try:
        api_key = get_api_key()
        json_data = fetch_cve_data(args.cve_id, api_key)
        
        formatted_text = format_output(json_data)
        print(formatted_text)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
