import pandas as pd
import numpy as np
from bs4 import BeautifulSoup

# Define car type arrays
full_23m = ['91', '94', '97', '98', '99', '100', '101', '102', '117', '118', '120', '121', '122', '123',
            '124', '125', '126', '127', '128', '129', '131', '132', '133', '134', '135', '136', '137',
            '138', '139', '140', '141', '142', '143', '144', '145', '146', '147', '148', '149', '150', 'VL-02']
full_25m = ['151', '152', '153', '154', '155', '156',
            '157', '158', '159', '160', '161', '162', '163', '164']
flatbed = ['F01', 'F02', 'F03', 'F04', 'F05',
           'F06', 'F07', 'F08', 'F09', 'F10', 'F11', 'F12']
type_s = ['S6', 'S7']
type_sb = ['SB2']

# Driver KM bonus data from the provided table
driver_km_bonuses = {
    'นาย วะสัน ยินเสียง': 1500,
    'นาย วิชิต เรืองชาญ': 1500,
    'นาย อำนาจ เขียวงาม': 300,
    'นาย ณัฐวุฒิ ดีสลาม': 3000,
    'นาย ประยูร ดวงใจ': 1500,
    'นาย ประเสริฐ แชขำ': 1500,
    'นาย วันชัย วิหกรัตน์': 1500,
    'นาย สรายุทธ กลีบรัง': 1500,
}

# Define the list of values in 'ต้นทาง' that should cause zero calculations
zero_calculation_sources = [
    'รถฝึก', 'รถเข้าศูนย์', 'รถซ่อม', 'รถเสีย', 'AAT TOY', 'FTM TOY']


def normalize_driver_name(name):
    """
    Normalize driver names by ensuring only one space between words
    Example: "นาย  วิชิต  เรืองชาญ" becomes "นาย วิชิต เรืองชาญ"
    """
    if not name or not isinstance(name, str):
        return name

    # Split by spaces and filter out empty strings
    parts = [part for part in name.split() if part]
    # Rejoin with single spaces
    return " ".join(parts)


# Normalize driver names in the driver_km_bonuses dictionary
normalized_driver_km_bonuses = {}
for driver, bonus in driver_km_bonuses.items():
    normalized_driver_km_bonuses[normalize_driver_name(driver)] = bonus
driver_km_bonuses = normalized_driver_km_bonuses


def get_car_type(car_number):
    if car_number in full_23m:
        return "full_23m"
    elif car_number in full_25m:
        return "full_25m"
    elif car_number in flatbed:
        return "flatbed"
    elif car_number in type_s:
        return "type_s"
    elif car_number in type_sb:
        return "type_sb"
    else:
        return "unknown"


def get_km_bonus(driver_name, total_km):
    """Get the kilometer bonus for a driver if they exceed 5000km"""
    if driver_name in driver_km_bonuses and total_km > 5000:
        print(driver_name, total_km)
        return driver_km_bonuses[driver_name]
    return 0


def calculate_bonus(row, rate_table):
    # Get vehicle type and distance category
    car_type = row["ประเภทรถ"]
    distance_cat = row["ประเภทระยะทาง"]

    # Convert kilometers and rate to float
    try:
        kilometers = float(row["กิโลเมตร"])
        actual_rate = float(row["เรท"]) if pd.notna(
            row["เรท"]) and row["เรท"] != "" else 0
    except (ValueError, TypeError):
        # Handle cases where conversion fails
        return 0

    # Skip calculation if kilometers is 0
    if kilometers == 0:
        return 0

    # Find applicable rate from rate table
    old_rate = None  # เรทเดิม
    new_rate = None  # เรทใหม่

    # Find the right row in rate table based on distance category
    for i, rate_row in enumerate(rate_table):
        # Use exact match instead of substring match
        if distance_cat == rate_row[1]:
            # Find the right column based on car type
            if car_type == "full_23m":
                old_rate = float(rate_row[3])
                new_rate = float(rate_row[4])
            elif car_type == "full_25m":
                old_rate = float(rate_row[5])
                new_rate = float(rate_row[6])
            elif car_type == "flatbed":
                old_rate = float(rate_row[7])
                new_rate = float(rate_row[8])
            elif car_type == "type_s":
                old_rate = float(rate_row[9])
                new_rate = float(rate_row[10])
            elif car_type == "type_sb":
                old_rate = float(rate_row[11])
                new_rate = float(rate_row[12])
            break

    if old_rate is None or new_rate is None:
        return 0  # Cannot calculate if rates not found

    # Calculate expected fuel usage
    expected_old = kilometers / old_rate  # Bottom ceiling (more fuel)
    expected_new = kilometers / new_rate  # Top ceiling (less fuel)
    actual_used = kilometers / actual_rate if actual_rate > 0 else 0

    # Calculate bonus based on the rules
    if actual_rate > new_rate:  # Better than new rate
        saved_from_old_to_new = expected_old - expected_new
        saved_beyond_new = expected_new - actual_used
        bonus = (saved_from_old_to_new * 0.75) + (saved_beyond_new * 0.5)
    elif actual_rate > old_rate:  # Between old and new rate
        saved = expected_old - actual_used
        bonus = (saved * 0.75) - (actual_used - expected_new)
    else:  # Worse than old rate
        extra_used = actual_used - expected_new
        bonus = -extra_used  # Driver pays back 100% below ceiling
        print(distance_cat, actual_used, actual_rate, expected_new,
              new_rate, extra_used, bonus)

    return bonus


def process_driver_data(df, all_job_counts, oil_price):
    # Define rate table from the provided information
    rate_table = [
        [1, "น้อยกว่า 350 กม. (ต่อเนื่อง)", "", 3.80, 4.00, 3.80, 3.90, 4.50,
            5.00, 5.50, 6.00, 6.00, 7.50, "ใกล้(ต่อเนื่อง)"],
        [2, "น้อยกว่า 350 กม.", "", 3.80, 4.20, 3.80, 4.10,
            4.50, 5.50, 5.50, 6.00, 6.00, 7.50, "ใกล้"],
        [3, "มากกว่า 350 และน้อยว่า 800 กม.", "", 3.80, 4.40, 3.80,
            4.30, 4.50, 6.00, 5.50, 6.50, 6.00, 7.50, "กลาง"],
        [4, "มากกว่า 800 กม.", "", 3.80, 4.60, 3.80,
            4.50, 4.50, 6.40, 5.50, 7.00, 6.00, "", "ไกล"]
    ]

    # Reset index to make sure we can iterate reliably
    df = df.reset_index(drop=True)

    # Rename columns to match expected names
    column_mapping = {
        'ชื่อ พขร.': 'ชื่อ-นามสกุล',  # Map the actual driver name column to expected name
        # Fix the spacing issue in column name
        'น้ำม                      มัน(ลิตร)': 'น้ำมัน(ลิตร)'
    }

    df = df.rename(columns=column_mapping)

    # Normalize driver names if the driver name column exists
    if 'ชื่อ พขร.' in df.columns:
        df['ชื่อ พขร.'] = df['ชื่อ พขร.'].apply(normalize_driver_name)
    elif 'ชื่อ-นามสกุล' in df.columns:
        df['ชื่อ-นามสกุล'] = df['ชื่อ-นามสกุล'].apply(normalize_driver_name)

    # Flag rows that match the zero calculation sources but don't set them to zero yet
    special_case_mask = None
    if 'ต้นทาง' in df.columns:
        # Create a mask for rows that match any of the special terms
        special_case_mask = df['ต้นทาง'].str.contains('|'.join(zero_calculation_sources),
                                                      case=False,
                                                      na=False)
        print(f"Found {special_case_mask.sum()} rows with special source values that will have zero calculations after categorization")

    # Ensure numeric columns are properly converted with safe handling
    try:
        df["กิโลเมตร"] = pd.to_numeric(df["กิโลเมตร"].astype(str).str.replace(
            ',', '').replace('', '0'), errors='coerce').fillna(0)
        df["เรท"] = pd.to_numeric(df["เรท"].astype(str).str.replace(
            ',', '').replace('', '0'), errors='coerce').fillna(0)
        if "น้ำมัน(ลิตร)" in df.columns:
            df["น้ำมัน(ลิตร)"] = pd.to_numeric(df["น้ำมัน(ลิตร)"].astype(
                str).str.replace(',', '').replace('', '0'), errors='coerce').fillna(0)
    except KeyError as e:
        print(
            f"Warning: Column not found - {e}. Continuing with available columns.")

    # Add car type column if เบอร์รถ exists
    if "เบอร์รถ" in df.columns:
        df.insert(df.columns.get_loc("เบอร์รถ") + 1,
                  "ประเภทรถ", df["เบอร์รถ"].apply(get_car_type))

    # Add drivers count column next to 'เลข Job' if it exists
    if "เลข Job" in df.columns:
        df.insert(df.columns.get_loc("เลข Job") + 1,
                  "จำนวน พขร.", df["เลข Job"].map(all_job_counts))

    # Create a new column for distance category if กิโลเมตร exists
    if "กิโลเมตร" in df.columns:
        # Calculate total kilometers per driver for KM bonus
        if "ชื่อ-นามสกุล" in df.columns:
            driver_total_km = df.groupby(
                "ชื่อ-นามสกุล")["กิโลเมตร"].sum().to_dict()
            # Add a new column for total kilometers for each driver
            df["รวมกิโลเมตร"] = df["ชื่อ-นามสกุล"].map(driver_total_km)
            # Add km bonus column based on total kilometers
            df["โบนัสกิโลเมตร"] = df.apply(
                lambda row: get_km_bonus(
                    row["ชื่อ-นามสกุล"], row["รวมกิโลเมตร"])
                if pd.notna(row["ชื่อ-นามสกุล"]) else 0,
                axis=1
            )
            # Since we only want to count the bonus once per driver, we'll only
            # apply it to the first occurrence of each driver
            driver_first_occurrence = {}
            for i, row in df.iterrows():
                driver_name = row["ชื่อ-นามสกุล"]
                if pd.notna(driver_name):
                    if driver_name not in driver_first_occurrence:
                        driver_first_occurrence[driver_name] = i
                    elif i != driver_first_occurrence[driver_name]:
                        df.at[i, "โบนัสกิโลเมตร"] = 0

        distance_categories = []

        for i in range(len(df)):
            km = df.loc[i, "กิโลเมตร"]

            # Skip the zero calculation check for trip categorization
            # This way, zero calculation sources won't trigger continuous trip logic

            if km == 0:
                # Check if this is part of a continuous trip
                if i < len(df) - 1:  # Make sure there's a next row
                    next_km = df.loc[i + 1, "กิโลเมตร"]

                    # Mark both current and next row
                    if next_km > 800:
                        distance_categories.append("มากกว่า 800 กม.")
                    elif next_km > 350:
                        distance_categories.append(
                            "มากกว่า 350 และน้อยว่า 800 กม.")
                    else:
                        distance_categories.append(
                            "น้อยกว่า 350 กม. (ต่อเนื่อง)")
                else:
                    # Last row with 0 km
                    distance_categories.append("น้อยกว่า 350 กม. (ต่อเนื่อง)")
            else:
                # Check if previous row had 0 km (part of continuous trip)
                if i > 0 and df.loc[i - 1, "กิโลเมตร"] == 0:
                    if km > 800:
                        distance_categories.append("มากกว่า 800 กม.")
                    elif km > 350:
                        distance_categories.append(
                            "มากกว่า 350 และน้อยว่า 800 กม.")
                    else:
                        distance_categories.append(
                            "น้อยกว่า 350 กม. (ต่อเนื่อง)")
                else:
                    # Regular categorization
                    if km > 800:
                        distance_categories.append("มากกว่า 800 กม.")
                    elif km > 350:
                        distance_categories.append(
                            "มากกว่า 350 และน้อยว่า 800 กม.")
                    else:
                        distance_categories.append("น้อยกว่า 350 กม.")

        # Insert the distance category column after กิโลเมตร
        df.insert(df.columns.get_loc("กิโลเมตร") + 1,
                  "ประเภทระยะทาง", distance_categories)

        # NOW apply the zero calculation sources mask AFTER the categorization is done
        if special_case_mask is not None and special_case_mask.any():
            if 'กิโลเมตร' in df.columns:
                df.loc[special_case_mask, 'กิโลเมตร'] = 0
            if 'น้ำมัน(ลิตร)' in df.columns:
                df.loc[special_case_mask, 'น้ำมัน(ลิตร)'] = 0
            if 'เรท' in df.columns:
                df.loc[special_case_mask, 'เรท'] = 0

        # Calculate bonus for each row
        df["เบี้ยคำนวณ"] = df.apply(lambda row: calculate_bonus(
            row, rate_table) if "เรท" in row and pd.notna(row["เรท"]) and row["เรท"] != "" else 0, axis=1)

        # Divide bonus by 2 if จำนวน พขร. is 2
        if "จำนวน พขร." in df.columns:
            df["เบี้ยคำนวณ"] = df.apply(lambda row: row["เบี้ยคำนวณ"] / 2
                                        if pd.notna(row["จำนวน พขร."]) and row["จำนวน พขร."] == 2
                                        else row["เบี้ยคำนวณ"], axis=1)

        # Add column with bonus multiplied by oil price
        df["เบี้ยคำนวณ x ราคาน้ำมัน"] = df["เบี้ยคำนวณ"] * oil_price

        # Add total bonus column (fuel efficiency bonus + km bonus)
        if "โบนัสกิโลเมตร" in df.columns:
            df["รวมโบนัสทั้งหมด"] = df["เบี้ยคำนวณ x ราคาน้ำมัน"] + \
                df["โบนัสกิโลเมตร"]

    return df


def html_to_excel():
    try:
        # Get oil price from user
        try:
            oil_price = float(input("กรุณาใส่ราคาน้ำมัน: "))
            print(f"Using oil price: {oil_price} บาท")
        except ValueError:
            print("ใส่ราคาน้ำมันไม่ถูกต้อง กำหนดเป็น 30 บาท")
            oil_price = 30.0

        # Get input file from user
        input_file = input(
            "กรุณาใส่ชื่อไฟล์ที่ต้องการประมวลผล (เช่น vis_job_driver_2025-04-21_14_11_56.xls): ")
        if not input_file:
            input_file = 'vis_job_driver_2025-04-21_14_11_56.xls'
            print(f"ไม่ได้ระบุชื่อไฟล์ ใช้ไฟล์เริ่มต้น: {input_file}")

        # Read the HTML file
        try:
            with open(input_file, 'r', encoding='utf-8') as file:
                html_content = file.read()
                print(f"HTML file '{input_file}' loaded successfully")
        except Exception as e:
            print(f"Error reading HTML file: {e}")
            return

        # Parse HTML with BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')

        # Find all tables in the HTML
        tables = soup.find_all('table')
        print(f"Found {len(tables)} tables in the HTML file")

        if len(tables) == 0:
            print("No tables found in the HTML file")
            return

        # First, collect all job numbers across all tables to count occurrences
        all_jobs = []
        for table in tables:
            # Find all data rows
            rows = table.find_all('tr')
            if len(rows) < 3:  # Need at least header + data row
                continue

            for tr in rows[2:]:  # Skip the header rows
                # Check if this is a data row or a summary row
                if tr.text and 'รวม' not in tr.text:  # Skip summary row
                    cells = tr.find_all(['td'])
                    if cells and len(cells) > 2:  # Make sure we have cells
                        # Assuming job number is in the 3rd column
                        job_number = cells[2].text.strip() if len(
                            cells) > 2 else ""
                        if job_number:
                            all_jobs.append(job_number)

        # Count occurrences of each job number
        job_counts = {}
        for job in all_jobs:
            if job in job_counts:
                job_counts[job] = 2  # If it appears more than once, set to 2
            else:
                job_counts[job] = 1

        print(f"Collected {len(job_counts)} unique job numbers")

        # Create a writer to save the processed data
        output_path = "processed_data_with_km_bonus.xlsx"
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Dictionary to collect driver summaries
            driver_summaries = {}
            processed_tables = 0

            # List to store all processed data for the combined sheet
            all_processed_data = []

            for table_index, table in enumerate(tables):
                # Extract headers
                header_rows = table.find_all('tr')
                if len(header_rows) < 2:
                    print(
                        f"Table {table_index} has insufficient rows, skipping")
                    continue

                header_cells = header_rows[1].find_all('th')
                if not header_cells:
                    print(f"Table {table_index} has no header cells, skipping")
                    continue

                headers = [th.text.strip() for th in header_cells]

                if not headers:
                    print(f"Table {table_index} has empty headers, skipping")
                    continue

                print(f"Table {table_index} headers: {headers}")

                # Find all data rows
                rows = []
                for tr in table.find_all('tr')[2:]:  # Skip the header rows
                    # Check if this is a data row or a summary row
                    if 'รวม' not in tr.text:  # Skip summary row
                        cells = tr.find_all(['td'])
                        if cells:  # Make sure we have cells
                            row_data = [cell.text.strip() for cell in cells]
                            if len(row_data) == len(headers):  # Ensure row matches headers
                                rows.append(row_data)
                            else:
                                print(
                                    f"Row length {len(row_data)} doesn't match headers length {len(headers)}, adjusting")
                                # Adjust row to match headers length
                                if len(row_data) < len(headers):
                                    row_data.extend(
                                        [''] * (len(headers) - len(row_data)))
                                else:
                                    row_data = row_data[:len(headers)]
                                rows.append(row_data)

                if not rows:
                    print(f"Table {table_index} has no data rows, skipping")
                    continue

                print(f"Table {table_index} has {len(rows)} data rows")

                # Create DataFrame
                df = pd.DataFrame(rows, columns=headers)

                # Check if required columns exist for processing (using more flexible column matching)
                required_base_columns = ['เบอร์รถ', 'กิโลเมตร', 'เลข Job']

                # Check for driver name column with alternative names
                driver_name_columns = ['ชื่อ-นามสกุล', 'ชื่อ พขร.', 'พขร.']
                has_driver_name = any(
                    col in df.columns for col in driver_name_columns)

                missing_base_columns = [
                    col for col in required_base_columns if col not in df.columns]

                if missing_base_columns or not has_driver_name:
                    print(
                        f"Table {table_index} missing required columns: {missing_base_columns}")
                    if not has_driver_name:
                        print("Missing driver name column")
                    print("Skipping this table")
                    continue

                # Process the data
                processed_df = process_driver_data(df, job_counts, oil_price)
                processed_tables += 1

                # Calculate totals for each driver if the required columns exist
                if 'ชื่อ-นามสกุล' in processed_df.columns:
                    # Create a DataFrame to store driver totals
                    driver_totals_df = pd.DataFrame(columns=[
                                                    'ชื่อ-นามสกุล', 'เบี้ยคำนวณ x ราคาน้ำมัน', 'โบนัสกิโลเมตร', 'รวมโบนัสทั้งหมด'])

                    # Get unique drivers
                    unique_drivers = processed_df['ชื่อ-นามสกุล'].unique()

                    for driver in unique_drivers:
                        if pd.notna(driver) and driver != '':
                            driver_data = processed_df[processed_df['ชื่อ-นามสกุล'] == driver]

                            # Calculate totals
                            fuel_bonus = driver_data['เบี้ยคำนวณ x ราคาน้ำมัน'].sum(
                            ) if 'เบี้ยคำนวณ x ราคาน้ำมัน' in driver_data.columns else 0
                            km_bonus = driver_data['โบนัสกิโลเมตร'].max(
                            ) if 'โบนัสกิโลเมตร' in driver_data.columns else 0
                            total_bonus = fuel_bonus + km_bonus

                            # Add to driver summaries
                            if driver in driver_summaries:
                                driver_summaries[driver]['fuel_bonus'] += fuel_bonus
                                # km bonus is fixed per driver
                                driver_summaries[driver]['km_bonus'] = km_bonus
                                driver_summaries[driver]['total_bonus'] = driver_summaries[driver]['fuel_bonus'] + \
                                    driver_summaries[driver]['km_bonus']
                            else:
                                driver_summaries[driver] = {
                                    'fuel_bonus': fuel_bonus,
                                    'km_bonus': km_bonus,
                                    'total_bonus': total_bonus
                                }

                # Add a total row to the dataframe
                if any(col in processed_df.columns for col in ['เบี้ยคำนวณ', 'เบี้ยคำนวณ x ราคาน้ำมัน', 'โบนัสกิโลเมตร', 'รวมโบนัสทั้งหมด']):
                    total_row = {'ชื่อ-นามสกุล': ['รวม']}

                    # Add totals for each relevant column
                    if 'เบี้ยคำนวณ' in processed_df.columns:
                        total_row['เบี้ยคำนวณ'] = [
                            processed_df['เบี้ยคำนวณ'].sum()]

                    if 'เบี้ยคำนวณ x ราคาน้ำมัน' in processed_df.columns:
                        total_row['เบี้ยคำนวณ x ราคาน้ำมัน'] = [
                            processed_df['เบี้ยคำนวณ x ราคาน้ำมัน'].sum()]

                    if 'โบนัสกิโลเมตร' in processed_df.columns:
                        total_row['โบนัสกิโลเมตร'] = [
                            processed_df['โบนัสกิโลเมตร'].sum()]

                    if 'รวมโบนัสทั้งหมด' in processed_df.columns:
                        total_row['รวมโบนัสทั้งหมด'] = [
                            processed_df['รวมโบนัสทั้งหมด'].sum()]

                    total_row_df = pd.DataFrame(total_row)

                    # Add columns that are in processed_df but not in total_row_df
                    for col in processed_df.columns:
                        if col not in total_row_df.columns:
                            total_row_df[col] = ['']

                    # Make sure columns are in the same order
                    total_row_df = total_row_df[processed_df.columns]

                    # Concatenate processed_df and total_row_df
                    processed_df = pd.concat(
                        [processed_df, total_row_df], ignore_index=True)

                # Add a separator row for visual clarity in the combined sheet
                separator_row = pd.DataFrame(
                    [['']*len(processed_df.columns)], columns=processed_df.columns)

                # Add the processed data to our list for the combined sheet
                all_processed_data.append(processed_df)
                # Add separator after each table
                all_processed_data.append(separator_row)

            if processed_tables == 0:
                print("No tables were successfully processed")
                return

            print(f"Successfully processed {processed_tables} tables")

            # Create and save the combined sheet with all driver data
            if all_processed_data:
                # Remove the last separator as it's not needed
                all_processed_data.pop()

                # Combine all data into a single dataframe
                combined_df = pd.concat(all_processed_data, ignore_index=True)

                # Add a grand total row at the end
                grand_total_columns = [
                    'เบี้ยคำนวณ', 'เบี้ยคำนวณ x ราคาน้ำมัน', 'โบนัสกิโลเมตร', 'รวมโบนัสทั้งหมด']
                grand_total_row = {'ชื่อ-นามสกุล': ['รวมทั้งหมด']}

                for col in grand_total_columns:
                    if col in combined_df.columns:
                        # Calculate sums excluding the separator rows (which have empty values)
                        col_sum = combined_df[col].replace(
                            '', np.nan).dropna().sum()
                        grand_total_row[col] = [col_sum]

                grand_total_df = pd.DataFrame(grand_total_row)

                # Add columns that are in combined_df but not in grand_total
                # Add columns that are in combined_df but not in grand_total_df
                for col in combined_df.columns:
                    if col not in grand_total_df.columns:
                        grand_total_df[col] = ['']

                # Make sure columns are in the same order
                grand_total_df = grand_total_df[combined_df.columns]

                # Concatenate with the combined_df
                combined_df = pd.concat(
                    [combined_df, grand_total_df], ignore_index=True)

                # Save the combined data to a single sheet
                combined_df.to_excel(
                    writer, sheet_name="All_Drivers", index=False)
                print(f"Saved all driver data to a single sheet 'All_Drivers'")

            # Create driver summary sheet
            if driver_summaries:
                summary_data = {
                    'ชื่อ-นามสกุล': [],
                    'เบี้ยประหยัดน้ำมัน': [],
                    'โบนัสกิโลเมตร': [],
                    'รวมโบนัสทั้งหมด': []
                }

                for driver, bonuses in driver_summaries.items():
                    summary_data['ชื่อ-นามสกุล'].append(driver)
                    summary_data['เบี้ยประหยัดน้ำมัน'].append(
                        bonuses['fuel_bonus'])
                    summary_data['โบนัสกิโลเมตร'].append(bonuses['km_bonus'])
                    summary_data['รวมโบนัสทั้งหมด'].append(
                        bonuses['total_bonus'])

                    summary_df = pd.DataFrame(summary_data)

                # Calculate totals
                summary_totals = {
                    'ชื่อ-นามสกุล': 'รวมทั้งหมด',
                    'เบี้ยประหยัดน้ำมัน': summary_df['เบี้ยประหยัดน้ำมัน'].sum(),
                    'โบนัสกิโลเมตร': summary_df['โบนัสกิโลเมตร'].sum(),
                    'รวมโบนัสทั้งหมด': summary_df['รวมโบนัสทั้งหมด'].sum()
                }

                # Add total row to summary
                summary_df = pd.concat(
                    [summary_df, pd.DataFrame([summary_totals])], ignore_index=True)

                # Save summary to Excel
                summary_df.to_excel(
                    writer, sheet_name="Driver_Summary", index=False)
                print("Saved Driver_Summary to Excel")
            else:
                print("No driver summaries to report")

        print(f"Processed data saved to {output_path}")

    except Exception as e:
        import traceback
        print(f"Error: {e}")
        print(traceback.format_exc())


if __name__ == "__main__":
    html_to_excel()
