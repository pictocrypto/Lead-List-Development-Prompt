import pandas as pd
from bs4 import BeautifulSoup
import requests
from tqdm import tqdm
import time

# Load the input Excel file
input_file = "ycom600.xlsx"
output_file = "output600.xlsx"

# Read the Excel file and extract the profile URLs
df = pd.read_excel(input_file)
profile_urls = df["Profile URL"]

# Initialize the data for the output file
output_data = []

# Function to extract company information
def extract_company_info(soup):
    # Extract company information from the provided HTML structure
    # Modify this function to extract the required company information
    company_info_card = soup.find("div", class_="ycdc-card")

    if company_info_card:
        company_name = company_info_card.find("div", class_="text-lg font-bold").text
        company_info = company_info_card.find_all("div", class_="flex flex-row justify-between")

        founded = None
        team_size = None
        location = None
        linkedin = None
        twitter = None
        crunchbase = None
        facebook = None
        github = None

        for info in company_info:
            label = info.find("span", recursive=False).text
            value = info.find("span", recursive=False).find_next("span").text

            if label == "Founded:":
                founded = value
            elif label == "Team Size:":
                team_size = value
            elif label == "Location:":
                location = value

        # Extract LinkedIn, Twitter, Crunchbase, Facebook, and GitHub profiles
        links = company_info_card.find_all("a")
        for link in links:
            href = link.get("href")
            if href:
                if "linkedin.com" in href:
                    linkedin = href
                elif "twitter.com" in href:
                    twitter = href
                elif "crunchbase.com" in href:
                    crunchbase = href
                elif "facebook.com" in href:
                    facebook = href
                elif "github.com" in href:
                    github = href

        # Add the extracted company information to a dictionary
        company_info_dict = {
            "Company Name": company_name,
            "Founded": founded,
            "Team Size": team_size,
            "Location": location,
            "LinkedIn Profile": linkedin,
            "Twitter Profile": twitter,
            "Crunchbase Profile": crunchbase,
            "Facebook Profile": facebook,
            "GitHub Profile": github
        }
        return company_info_dict
    else:
        return None

# Function to extract founder information
def extract_founders(soup):
    # Extract founder information from the provided HTML structure
    # Modify this function to extract the required founder information
    founder_cards = soup.find_all("div", class_="ycdc-card shrink-0 space-y-1.5 sm:w-[300px]")

    founders_info = []
    for i, card in enumerate(founder_cards):
        name = card.find("div", class_="font-bold").text
        title = card.find("div", class_="leading-snug").text

        # Remove the name from the title
        title = title.replace(name, '').strip()

        # Assume there's only one link (Twitter or LinkedIn) per founder
        links = card.find_all("a")
        twitter = next((link["href"] for link in links if "twitter.com" in link["href"]), None)
        linkedin = next((link["href"] for link in links if "linkedin.com" in link["href"]), None)

        # Add the extracted founder information to a dictionary
        founder_info = {
            f"Name {i + 1}": name,
            f"Title {i + 1}": title,
            f"Twitter {i + 1}": twitter,
            f"LinkedIn {i + 1}": linkedin,
            f"Email {i + 1}": None  # Update this based on extraction logic
        }

        founders_info.append(founder_info)

    return founders_info



# Iterate through each profile URL and extract information
for url in tqdm(profile_urls, desc="Processing profiles"):
    # Make a GET request to the profile URL and parse the HTML content
    response = requests.get(url)
    if response.status_code != 200:
        print(f"Failed to fetch data for {url}. Skipping...")
        continue

    soup = BeautifulSoup(response.content, "html.parser")

    # Extract company website URL
    company_website = soup.find("a", class_="whitespace-nowrap")
    company_website = company_website["href"] if company_website else None

    # Extract company information
    company_info = extract_company_info(soup)

    # Extract founder information
    founders_info = extract_founders(soup)

    # Combine all extracted information into a dictionary
    output_row = {
        "Website": company_website,
        **company_info,
        **{key: value for founder_info in founders_info for key, value in founder_info.items()}
    }

    # Append the extracted information to the output data
    output_data.append(output_row)

    # Delay for 1 seconds between each profile URL
    time.sleep(1)

# Create a DataFrame from the extracted data and write to the output file
output_df = pd.DataFrame(output_data)
df_with_output = pd.concat([df, output_df], axis=1)
df_with_output.to_excel(output_file, index=False)
print("Extraction completed. Data saved to", output_file)
