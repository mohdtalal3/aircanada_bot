from seleniumbase import SB
import csv
import time
import os
from datetime import datetime
import re

def read_csv(file_path):
    """Read CSV file and return list of dictionaries"""
    data_list = []
    with open(file_path, 'r', encoding='utf-8-sig') as file:  # utf-8-sig handles BOM
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            # Skip empty rows
            if row.get('Email') and row.get('Email').strip():
                data_list.append(row)
    return data_list


def get_processed_emails(processed_file):
    """Get list of already processed emails"""
    if not os.path.exists(processed_file):
        return set()
    
    processed_emails = set()
    with open(processed_file, 'r', encoding='utf-8-sig') as file:  # utf-8-sig handles BOM
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            if row.get('Email'):
                processed_emails.add(row.get('Email').strip())
    return processed_emails


def save_processed_entry(processed_file, row_data, aeroplan_number=None):
    """Save processed entry to CSV file"""
    file_exists = os.path.exists(processed_file)
    
    with open(processed_file, 'a', encoding='utf-8', newline='') as file:
        # Include Aeroplan number in fieldnames if not already there
        fieldnames = list(row_data.keys())
        if 'Processed_Date' not in fieldnames:
            fieldnames.append('Processed_Date')
        if 'Aeroplan number' not in fieldnames:
            fieldnames.append('Aeroplan number')
            
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        
        # Write header if file doesn't exist
        if not file_exists:
            writer.writeheader()
        
        # Add timestamp and aeroplan number
        row_data['Processed_Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if aeroplan_number:
            row_data['Aeroplan number'] = aeroplan_number
        else:
            row_data['Aeroplan number'] = 'Failed to extract'
            
        writer.writerow(row_data)


# Fixed rotating proxy - no need to read from CSV
ROTATING_PROXY = "e955120e5c61b9eb64e9__cr.us:27e87286655e11fe@gw.dataimpulse.com:10000"
#ROTATING_PROXY = "LV23106126-IqS4nHKett-33:3a2bcxb537@174.138.175.186:7383"

def fill_form(sb, data):
    """Fill the form with data from CSV row"""
    
    # Open the registration page
    sb.open("https://www.aircanada.com/aeroplan/member/enrolment")
    input("Press Enter to continue after the page loads...")
    time.sleep(3)
    
    # Fill email
    sb.js_click("#emailFocus", scroll=True,timeout=15)
    sb.type("#emailFocus", data.get("Email"))

    # Fill password
    sb.js_click("#pwd", scroll=True,timeout=15)
    sb.type("#pwd", data.get("Password"))

    # Click checkbox
    sb.js_click('input[id="checkBox-input"]', timeout=15)

    # Click continue button
    sb.js_click("button[data-analytics-val*='continue']", timeout=15)
    time.sleep(5)

    # Fill first name
    sb.js_click("input[name='firstName']", scroll=True, timeout=15)
    sb.type("input[name='firstName']", data.get("First name"))

    # Fill last name
    sb.js_click("input[name='lastName']", scroll=True, timeout=15)
    sb.type("input[name='lastName']", data.get("Last name"))

    # Select gender
    sb.js_click('mat-select[formcontrolname="gender"]', scroll=True, timeout=15)
    if data.get("Gender") == "Male":
        sb.js_click("//mat-option//span[text()=' Male ']", by="xpath", timeout=15)
    else:
        sb.js_click("//mat-option//span[text()=' Female ']", by="xpath", timeout=15)
    time.sleep(2)

    # Select day
    sb.js_click("mat-select[formcontrolname='d']", scroll=True, timeout=15)
    element = sb.find_element(f'//mat-option//span[normalize-space(text())="{data.get("Day")}"]', by="xpath", timeout=15)
    sb.execute_script("arguments[0].click();", element)

    # Select month
    sb.js_click("mat-select[formcontrolname='m']", scroll=True, timeout=15)
    sb.js_click(f"//mat-option//span[text()=' {data.get("Month")} ']", by="xpath", timeout=15)

    # Select year
    sb.js_click("mat-select[formcontrolname='y']", scroll=True, timeout=15)
    sb.js_click(f"//mat-option//span[text()=' {data.get("Year")} ']", by="xpath", timeout=15)

    time.sleep(2)


        # Click next button
    sb.js_click("button[data-analytics-val*='continue']", timeout=5)
    time.sleep(5)


    # Fill address
    sb.js_click('input[formcontrolname="addressLine1"]', scroll=True, timeout=15)
    sb.type('input[formcontrolname="addressLine1"]', data.get("Address"), timeout=15)

    # Fill city
    sb.js_click('input[formcontrolname="city"]', scroll=True, timeout=15)
    sb.type('input[formcontrolname="city"]', data.get("City"), timeout=15)



    # Select country
    sb.js_click('mat-select[formcontrolname="country"]', scroll=True, timeout=15)
    sb.js_click(f"//mat-option//span[text()=' {data.get("Country")} ']", by="xpath", timeout=15)
    time.sleep(2)

    # Select province/state
    sb.js_click('mat-select[formcontrolname="state"]', scroll=True, timeout=15)
    sb.js_click(f"//mat-option//span[text()=' {data.get("Province")} ']", by="xpath", timeout=15)

    # Fill postal code
    sb.js_click('input[formcontrolname="zip"]', scroll=True, timeout=15)
    sb.type('input[formcontrolname="zip"]', data.get("Postal Code"), timeout=15)
    # Fill phone number
    sb.js_click('input[formcontrolname="phoneNumber"]', scroll=True, timeout=15)
    sb.type('input[formcontrolname="phoneNumber"]', data.get("Phone number"), timeout=15)

    # Click privacy policy checkbox
    sb.js_click('input[id="privacyPolicycheckBox"]', scroll=True, timeout=15)

    # Wait for user to solve captcha
    print("\n⚠️  Please solve the CAPTCHA manually...")
    input("Press Enter after solving the CAPTCHA to submit the form...")

    # Submit form
    sb.js_click('button[data-analytics-val*="create my account"]', scroll=True,timeout=15)

    time.sleep(5)
    try:
        number_text = sb.get_text("span.aeroplan-number")
        aeroplan_number = re.sub(r"\D", "", number_text)
        print(f"\n🎉 Aeroplan Number Extracted: {aeroplan_number}")
        return aeroplan_number
    except Exception as e:
        print(f"❌ Could not extract Aeroplan Number: {str(e)}")
        return None


def main():
    """Main function to process CSV entries"""
    
    # File paths
    input_csv = "Test data.csv"
    processed_csv = "processed_data.csv"
    
    # Read CSV data
    print("📖 Reading CSV file...")
    data_list = read_csv(input_csv)
    print(f"✅ Found {len(data_list)} entries in CSV")
    
    # Get already processed emails
    processed_emails = get_processed_emails(processed_csv)
    print(f"✅ Found {len(processed_emails)} already processed entries")
    
    # Filter out already processed entries
    remaining_data = [row for row in data_list if row.get('Email') not in processed_emails]
    print(f"📋 {len(remaining_data)} entries remaining to process\n")
    
    if not remaining_data:
        print("✨ All entries have been processed!")
        return
    
    # Process each row
    for index, data in enumerate(remaining_data, 1):
        email = data.get('Email')
        print(f"\n{'='*60}")
        print(f"🔄 Processing entry {index}/{len(remaining_data)}")
        print(f"📧 Email: {email}")
        print(f"👤 Name: {data.get('First name')} {data.get('Last name')}")
        print(f"{'='*60}\n")
        
        # Use the fixed rotating proxy
        print(f"🌐 Using rotating proxy: {ROTATING_PROXY}")
        
        try:
            # Initialize browser with rotating proxy
            aeroplan_number = None
            with SB(uc=True, proxy=ROTATING_PROXY) as sb:
                aeroplan_number = fill_form(sb, data)
            
            # Save to processed CSV with Aeroplan number
            save_processed_entry(processed_csv, data, aeroplan_number)
            
            if aeroplan_number:
                print(f"✅ Successfully processed and saved: {email}")
                print(f"✈️  Aeroplan Number: {aeroplan_number}\n")
            else:
                print(f"⚠️  Processed but could not extract Aeroplan number: {email}\n")
            
        except Exception as e:
            print(f"❌ Error processing {email}: {str(e)}")
            print("⏭️  Skipping to next entry...\n")
            continue
    
    print("\n" + "="*60)
    print("🎉 All entries have been processed!")
    print("="*60)


if __name__ == "__main__":
    main()