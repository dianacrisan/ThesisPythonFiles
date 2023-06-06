import pandas as pd
import os

# Initialize smells file names and path to folder
designSmellsFile = 'statsDesignSmells.xlsx'
architectureSmellsFile = 'statsArchitectureSmells.xlsx'
implementationSmellsFile = 'statsImplementationSmells.xlsx'
testabilitySmellsFile = 'statsTestabilitySmells.xlsx'
folder_path = '/Users/dianacrisan/Desktop/Dizertatie/ref-miner/GeneratedExcelFiles'

# Assign smell type file to be evaluated
evaluatedSmellsTypeFile = architectureSmellsFile
print('Evaluating ' + evaluatedSmellsTypeFile)

# Read the Excel file into a DataFrame
df = pd.read_excel(os.path.join(folder_path, evaluatedSmellsTypeFile))

# Extract the columnNumber string from each list in the first column
columnNumber = 2 # for architectureSmells
# columnNumber = 3 # for designSmells
# columnNumber = 4 # for implementationSmells
df['Smell Type'] = df.iloc[:, 0].str.split(',').str[columnNumber].str.strip(" '")

# Remove the first two columns
df.drop(df.columns[[0, 1]], axis=1, inplace=True)

# Reorder the columns to have the smell in the first column
columns = ['Smell Type'] + list(df.columns[:-1])
df = df[columns]

# Specify path to output folder and the new simple Excel file name 
output_path = '/Users/dianacrisan/Desktop/Dizertatie/ref-miner/Statistics/' + evaluatedSmellsTypeFile[5:]

# Save the updated DataFrame back to the Excel file
df.to_excel(output_path, index=False)
