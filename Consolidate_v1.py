import os
import pandas as pd

# Define the folder where your Excel files are stored
Main_Folder = 'D:\Stocks'
data_directory = r'D:\Stocks\Execute1000'
folder_path = data_directory  # Replace with the path to your folder containing Excel files
output_file = 'D:\Stocks\consolidated_Execute1000_4_2_10.csv'  # Name of the output CSV file

# List to hold data from each Excel file
all_data = []
df = []

i = 0
# Loop through all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Construct full file path
        file_path = os.path.join(folder_path, filename)
        # Read the Excel file
        #df = pd.read_excel(file_path, engine='openpyxl')
        df = pd.read_excel(file_path)
        #print(df)
        # Append the data to the list
        all_data.append(df)
        i = i + 1
        print(str(i) + ":" + filename + ":" + str(len(all_data)))

print("Concate in progress")
# Concatenate all the data into a single DataFrame
combined_df = pd.concat(all_data, ignore_index=True)

print("Writing to csv")
# Export the consolidated DataFrame to a CSV file
combined_df.to_csv(os.path.join(folder_path, output_file), index=False)
print("Consolidation completed")
shape = combined_df.shape
# Print the shape
print(f"Rows: {shape[0]}, Columns: {shape[1]}")

# Filter out rows where FieldName is blank or null
filtered_df = combined_df[combined_df['signal'].notna() & (combined_df['signal'] != '')]
shape = filtered_df.shape
# Print the shape
print(f"Rows: {shape[0]}, Columns: {shape[1]}")
#print(filtered_df)

pivot_df = filtered_df.pivot_table(index='Stock', columns='signal', aggfunc='size', fill_value=0)
# Rename the columns to match your example
pivot_df.columns = [f'Signal {col}' for col in pivot_df.columns]
print(pivot_df)

pivot_df1 = filtered_df.pivot_table(index='signal', columns='Year', aggfunc='size', fill_value=0)
# Rename the columns to match your example
pivot_df1.columns = [f'{col}' for col in pivot_df1.columns]
print(pivot_df1)

Main_Folder = 'D:\\Stocks'  # Make sure to use double backslashes for Windows paths
excel_path = os.path.join(Main_Folder, "Final_Results_v4_Execute1000_4_2_10.xlsx")
pivot_df.to_excel(excel_path, index=True, engine='openpyxl')  # Use 'excel_path' here instead of 'Main_Folder'