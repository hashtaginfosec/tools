##########################
# Collect vuln data from different sources.
# Add your greynoise and vulnerability.circl.lu API keys to config.ini file. 
# config.ini file content:
# [API_KEYS]
# GREYNOISE_API_KEY = XXXXXXXXXXXXXXXXXXXXX
# CIRCL_API_KEY = XXXXXXXXXXXXXXXXXXXXX
##########################

import requests
import json
import os
import argparse
from datetime import datetime, timedelta
import configparser

def get_api_keys():
    """Retrieve API keys from config.ini file."""
    config = configparser.ConfigParser()
    config.read("config.ini")
    
    greynoise_key = config.get("API_KEYS", "GREYNOISE_API_KEY", fallback=None)
    circl_key = config.get("API_KEYS", "CIRCL_API_KEY", fallback=None)

    if not greynoise_key:
        raise ValueError("GREYNOISE_API_KEY not found in config.ini")
    if not circl_key:
        raise ValueError("CIRCL_API_KEY not found in config.ini")

    return greynoise_key, circl_key

def fetch_cve_data(cve_id, api_key):
    """Fetch CVE data from the Greynoise API."""
    url = f"https://api.greynoise.io/v1/cve/{cve_id}"
    headers = {
        "accept": "application/json",
        "key": api_key
    }
    response = requests.get(url, headers=headers)
    return response.json()

def fetch_shodan_cve_data(cve_id):
    """Fetch CVE data from the Shodan CVEDB API."""
    url = f"https://cvedb.shodan.io/cve/{cve_id}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    return None

def fetch_circl_recent_vulnerabilities(circl_key):
    """Fetch recent vulnerabilities from CIRCL."""
    url = "https://vulnerability.circl.lu/api/vulnerability/recent"
    headers = {
        "accept": "application/json",
        "X-API-KEY": circl_key
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        try:
            data = response.json()
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []

def format_output(cve_id, json_data, shodan_data):
    """Format the API response in a readable format."""
    output = []
    output.append(f"CVE ID: {cve_id}")

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

    if shodan_data:
        ransomware_status = shodan_data.get("ransomware_campaign", "None detected")
        output.append(f"  Shodan Known Ransomware Campaigns: {ransomware_status}")
        references = shodan_data.get("references", [])
        if references:
            output.append("\nReferences:")
            for ref in references:
                output.append(f"  - {ref}")
    else:
        output.append("\nShodan CVEDB Information: No data available")

    return "\n".join(output)

def main():
    try:
        greynoise_key, circl_key = get_api_keys()
        circl_data = fetch_circl_recent_vulnerabilities(circl_key)

        if not circl_data:
            print("No recent vulnerabilities from CIRCL.")
            return

        for vuln in circl_data:
            cve_id = None

            # CVE ID might be under different keys depending on the structure
            if isinstance(vuln, dict):
                if "cveMetadata" in vuln and isinstance(vuln["cveMetadata"], dict):
                    cve_id = vuln["cveMetadata"].get("cveId")
                elif "vulnerabilities" in vuln and isinstance(vuln["vulnerabilities"], list):
                    for v in vuln["vulnerabilities"]:
                        if "cve" in v:
                            cve_id = v["cve"]
                            break

            if not cve_id:
                continue

            try:
                greynoise_data = fetch_cve_data(cve_id, greynoise_key)
            except Exception as e:
                greynoise_data = {"details": {}, "timeline": {}, "exploitation_details": {}}
                print(f"Failed to fetch Greynoise data for {cve_id}: {e}")

            try:
                shodan_data = fetch_shodan_cve_data(cve_id)
            except Exception as e:
                shodan_data = None
                print(f"Failed to fetch Shodan data for {cve_id}: {e}")

            formatted = format_output(cve_id, greynoise_data, shodan_data)
            print("\n" + formatted + "\n" + ("=" * 80) + "\n")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
