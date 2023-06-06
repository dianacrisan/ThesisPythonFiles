import os
import pandas as pd

# Set the path to the folder containing the excel files
folder_path = '/Users/dianacrisan/Desktop/Dizertatie/ref-miner/DesigniteJava/CommitsResults'

# Create a list of all the file names in the folder in chronological order
excel_files = sorted([f for f in os.listdir(folder_path) if f.endswith('.xlsx')], key=lambda f: os.stat(os.path.join(folder_path, f)).st_mtime)

# Initialize a dictionary to store the statistics
statistics = {}

# Dedine the sheet names that contain smells
smell_sheets = ['ArchitectureSmells', 'ImplementationSmells', 'DesignSmells', 'TestabilitySmells']
sheet_analyzed = 'ArchitectureSmells'

# Loop through each Excel file
for file in excel_files:
    print('file: ' + file)
    # Load the smells sheet from the current file
    df = pd.read_excel(os.path.join(folder_path, file), sheet_name=sheet_analyzed)

    # Get the commit state from the file name
    commit_state = file.split('/')[-1].split('.')[0]

    # Loop through each row in the sheet
    for _, row in df.iterrows():
        smell = tuple(row)
        if smell not in statistics:
            # If the smell is not in the dictionary, add it with a list of commits
            statistics[smell] = [commit_state]
        else:
            # If the smell is already in the dictionary, add the current commit state to the list
            statistics[smell].append(commit_state)

# Create a new DataFrame to store the statistics
stats_df = pd.DataFrame(columns=['Smell', 'Count', 'First Commit', 'Last Commit', 'Survival Length'])

# Loop through each smell in the dictionary
for smell, commits in statistics.items():
    # Calculate the count and first/last commit for the smell
    count = len(commits)
    first_commit = min(commits)
    last_commit = max(commits)
    survival_length = len(commits)

    # Check if the smell was removed in a later commit
    for file in excel_files:
        if file.split('/')[-1].split('.')[0] > last_commit:
            print('file with commit: ' + file.split('/')[-1].split('.')[0])
            df = pd.read_excel(os.path.join(folder_path, file), sheet_name=sheet_analyzed)
            for _, row in df.iterrows():
                if tuple(row) == smell:
                    # If the smell is found in a later commit, update the last commit and survival length
                    last_commit = file.split('/')[-1].split('.')[0]
                    survival_length = len(commits) + 1

    # Add the smell and its statistics to the DataFrame
    stats_df = stats_df._append({
        'Smell': smell,
        'Count': count,
        'First Commit': first_commit,
        'Last Commit': last_commit,
        'Survival Length': survival_length
    }, ignore_index=True)

# Save the DataFrame to a new Excel file
output_path = '/Users/dianacrisan/Desktop/Dizertatie/ref-miner/GeneratedExcelFiles/stats' + sheet_analyzed + '.xlsx'
stats_df.to_excel(output_path, sheet_name=sheet_analyzed + ' Stats', index=False)
