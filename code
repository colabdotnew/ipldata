import zipfile
import json
import os
import pandas as pd

zip_file_path = '/content/ipl_json.zip'
extracted_dir = '/content/extracted_ipl_data'
all_results = pd.DataFrame()  # Initialize an empty DataFrame to store all results
excel_output_path = '/content/all_ipl_match_results_sorted_and_formatted.xlsx'

try:
    # 1. Extract ZIP (if not already done)
    if not os.path.exists(extracted_dir):
        os.makedirs(extracted_dir, exist_ok=True)
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_dir)
        print(f"Successfully extracted contents of '{zip_file_path}' to '{extracted_dir}'.")
    else:
        print(f"Directory '{extracted_dir}' already exists. Skipping extraction.")

    # 2. Iterate through JSON files and collect data
    match_data_list = []
    for filename in os.listdir(extracted_dir):
        if filename.endswith('.json'):
            file_path = os.path.join(extracted_dir, filename)
            try:
                with open(file_path, 'r') as f:
                    match_data = json.load(f)

                date_str = match_data['info']['dates'][0] if 'dates' in match_data['info'] and match_data['info']['dates'] else None
                stadium = match_data['info']['venue'] if 'venue' in match_data['info'] else None
                result = ""
                if 'outcome' in match_data['info']:
                    outcome = match_data['info']['outcome']
                    if 'winner' in outcome:
                        result = f"{outcome['winner']} won"
                        if 'by' in outcome:
                            by_type = list(outcome['by'].keys())[0]
                            by_value = outcome['by'][by_type]
                            result += f" by {by_value} {by_type}"
                    else:
                        result = "Match Tied" if 'tie' in outcome else outcome.get('result', 'No Result')
                else:
                    result = "Result Not Available"

                innings_data = []
                for inning in match_data.get('innings', []):
                    team_name = inning.get('team')
                    total_runs = 0
                    wickets_fallen = 0
                    for over in inning.get('overs', []):
                        for delivery in over.get('deliveries', []):
                            total_runs += delivery.get('runs', {}).get('total', 0)
                            if 'wickets' in delivery:
                                wickets_fallen += len(delivery['wickets'])
                    if team_name:
                        innings_data.append({'team': team_name, 'score': f"{total_runs}/{wickets_fallen}"})

                first_innings_team = innings_data[0]['team'] if len(innings_data) > 0 else None
                first_innings_score = innings_data[0]['score'] if len(innings_data) > 0 else None
                second_innings_team = innings_data[1]['team'] if len(innings_data) > 1 else None
                second_innings_score = innings_data[1]['score'] if len(innings_data) > 1 else None

                match_result = {'DATE_STR': date_str,  # Store original date string for sorting
                                'STADIUM': stadium,
                                '1st INNING TEAM': first_innings_team,
                                '1st INNING SCORE': first_innings_score,
                                '2nd INNING TEAM': second_innings_team,
                                '2nd INNING SCORE': second_innings_score,
                                'RESULT': result}

                match_data_list.append(match_result)

            except Exception as e:
                print(f"Error processing file '{filename}': {e}")

    # 3. Create DataFrame from the list of dictionaries
    all_results = pd.DataFrame(match_data_list)

    # 4. Sort DataFrame by date (using the original date string)
    all_results['DATE'] = pd.to_datetime(all_results['DATE_STR'], errors='coerce')
    all_results = all_results.sort_values(by='DATE', ascending=True)
    all_results.drop(columns=['DATE'], inplace=True)  # Remove the temporary datetime column

    # 5. Format the date to DDMMYYYY
    all_results['DATE_STR'] = all_results['DATE_STR'].apply(
        lambda x: pd.to_datetime(x, errors='coerce').strftime('%d%m%Y') if pd.notna(x) else None
    )
    all_results.rename(columns={'DATE_STR': 'DATE'}, inplace=True)  # Rename for clarity

    # 6. Save sorted and formatted results to Excel
    all_results.to_excel(excel_output_path, index=False)
    print(f"\nAll match results extracted, sorted by date (ascending), with date format DDMMYYYY, and saved to '{excel_output_path}'")

except FileNotFoundError:
    print(f"Error: The file '{zip_file_path}' or directory '{extracted_dir}' was not found.")
except zipfile.BadZipFile:
    print(f"Error: The file '{zip_file_path}' is not a valid ZIP file.")
except Exception as e:
    print(f"An error occurred during the process: {e}")
