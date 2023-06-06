import pandas as pd
import os
import subprocess
import xlsxwriter

# Read commits from Refactoring Miner excel file
commitsDataFrame = pd.read_excel('/Users/dianacrisan/Desktop/Dizertatie/ref-miner/RefactoringMiner/ref-miner Refactoring Miner findings.xlsx', sheet_name=1)
commitsDataFrame.drop(index=commitsDataFrame.index[0], axis=0, inplace=True) # drop first row
commitsDataFrame.drop_duplicates() # drop duplicate commitIds

# Save commits in list
commitsList = commitsDataFrame['commitId'].values.tolist() # transform from pandas dataframe to list
commitsList.reverse() # commits now from old to new

# Clone the JUnit4 repository
repo_url = "https://github.com/tsantalis/RefactoringMiner"
local_path = "/Users/dianacrisan/Desktop/Dizertatie/ref-miner/Repo"
subprocess.run(["git", "clone", repo_url, local_path], check=True)

# Loop over the commit IDs and run DesigniteJava on each commit
results = []
for commit in commitsList:
    # Checkout the repository at the current commit
    subprocess.run(["git", "checkout", commit], cwd=local_path, check=True)

    # Run DesigniteJava on the cloned repository
    designite_path = "/Users/dianacrisan/.designite/DesigniteJava.jar"
    input_path = "/Users/dianacrisan/Desktop/Dizertatie/ref-miner/Repo"
    output_path = "/Users/dianacrisan/Desktop/Dizertatie/ref-miner/DesigniteJava/CommitsResults/ref-miner-" + commit

    command = f"java -jar {designite_path} -i {input_path} -o {output_path}"
    proc = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Wait for DesigniteJava to finish and capture its output
    stdout, stderr = proc.communicate()
    if proc.returncode != 0:
        # DesigniteJava failed, print the error message and exit
        print("DesigniteJava failed with error:")
        print(stderr.decode())
        exit(1)

    # Combine the output CSV files into a single Excel file
    output_folder = "/Users/dianacrisan/Desktop/Dizertatie/ref-miner/DesigniteJava/CommitsResults"

    # Create an Excel workbook object
    output_file = os.path.join(output_folder, 'output-' + commit + '.xlsx')
    workbook = xlsxwriter.Workbook(output_file)

    # Loop through each csv file in the directory
    for file_name in os.listdir(output_path):
        if file_name.endswith(".csv"):
            sheet_name = os.path.splitext(file_name)[0]  # Get sheet name from filename
            worksheet = workbook.add_worksheet(sheet_name)  # Add new worksheet to workbook
            
            # Open the csv file and read its contents
            with open(os.path.join(output_path, file_name), "r") as csv_file:
                lines = csv_file.readlines()
                
                # Write the data to the worksheet
                for row_num, line in enumerate(lines):
                    row_data = line.strip().split(",")
                    if row_num == 0:
                        # Write header row
                        worksheet.write_row(0, 0, row_data)
                    else:
                        # Write data rows
                        worksheet.write_row(row_num, 0, row_data)
                        
    # Close the workbook
    workbook.close()
