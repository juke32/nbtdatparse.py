import nbtlib
import os
import gzip
from openpyxl import Workbook

# Set the directory path
directory_path = 'C:/Users/juke32/AppData/Roaming/.minecraft'
#directory_path = 'D:/dump'

# Create a new Excel workbook
wb = Workbook()
ws_data = wb.active
ws_data.title = "Data"
ws_errors = wb.create_sheet(title="Errors")
ws_log = wb.create_sheet(title="Log Results")

# Set header row for Data tab
ws_data['A1'] = 'File Name'
ws_data['B1'] = 'Random Seed'
ws_data['C1'] = 'Total Time'
ws_data['D1'] = 'Generator Name'
ws_data['E1'] = 'Level Name'
ws_data['F1'] = 'Game Mode'
ws_data['G1'] = 'Spawn Location'
ws_data['H1'] = 'Path'

# Set header row for Errors tab
ws_errors['A1'] = 'File Name'
ws_errors['B1'] = 'Error Message'
ws_errors['C1'] = 'Path'

# Set header row for Log Results tab
ws_log['A1'] = 'File Name'
ws_log['B1'] = 'Path'
ws_log['C1'] = 'Log Line'

# Adjust column widths
ws_data.column_dimensions['A'].width = 20
ws_data.column_dimensions['B'].width = 20
ws_data.column_dimensions['C'].width = 15
ws_data.column_dimensions['D'].width = 20
ws_data.column_dimensions['E'].width = 20
ws_data.column_dimensions['F'].width = 15
ws_data.column_dimensions['G'].width = 20
ws_data.column_dimensions['H'].width = 50

ws_errors.column_dimensions['A'].width = 20
ws_errors.column_dimensions['B'].width = 50
ws_errors.column_dimensions['C'].width = 50

ws_log.column_dimensions['A'].width = 20
ws_log.column_dimensions['B'].width = 50
ws_log.column_dimensions['C'].width = 100

# Initialize row counters
row_data = 2
row_errors = 2
row_log = 2

# Initialize counters for processed and saved files
processed_files = 0
saved_entries = 0
errors_encountered = 0

# Iterate through all .dat files in the directory and its subdirectories
for root, dirs, files in os.walk(directory_path):
    for filename in files:
        if filename.endswith(".dat"):
            file_path = os.path.join(root, filename)
            
            processed_files += 1
            
            try:
                # Load the NBT file
                var = nbtlib.load(file_path)
                
                # Extract and write data to Excel
                ws_data[f'A{row_data}'] = filename
                ws_data[f'H{row_data}'] = os.path.dirname(file_path)
                
                # Try to extract each piece of data individually
                try:
                    ws_data[f'B{row_data}'] = str(var.root['Data']['RandomSeed'])
                except Exception as e:
                    ws_data[f'B{row_data}'] = f"Error: {e}"
                    print(f"Error extracting RandomSeed from {filename}: {e}")
                    
                try:
                    ws_data[f'C{row_data}'] = var.root['Data']['Time']
                except Exception as e:
                    ws_data[f'C{row_data}'] = f"Error: {e}"
                    print(f"Error extracting Time from {filename}: {e}")
                    
                try:
                    ws_data[f'D{row_data}'] = var.root['Data']['generatorName']
                except Exception as e:
                    ws_data[f'D{row_data}'] = f"Error: {e}"
                    print(f"Error extracting generatorName from {filename}: {e}")
                    
                try:
                    ws_data[f'E{row_data}'] = var.root['Data']['LevelName']
                except Exception as e:
                    ws_data[f'E{row_data}'] = f"Error: {e}"
                    print(f"Error extracting LevelName from {filename}: {e}")
                    
                try:
                    if var.root['Data']['GameType'] == 0:
                        ws_data[f'F{row_data}'] = 'Survival'
                    elif var.root['Data']['GameType'] == 1:
                        ws_data[f'F{row_data}'] = 'Creative'
                    elif var.root['Data']['GameType'] == 2:
                        ws_data[f'F{row_data}'] = 'Adventure'
                    elif var.root['Data']['GameType'] == 3:
                        ws_data[f'F{row_data}'] = 'Spectator'
                except Exception as e:
                    ws_data[f'F{row_data}'] = f"Error: {e}"
                    print(f"Error extracting GameType from {filename}: {e}")
                    
                try:
                    spawn_location = f"X={var.root['Data']['SpawnX']}, Y={var.root['Data']['SpawnY']}, Z={var.root['Data']['SpawnZ']}"
                    ws_data[f'G{row_data}'] = spawn_location
                except Exception as e:
                    ws_data[f'G{row_data}'] = f"Error: {e}"
                    print(f"Error extracting Spawn Location from {filename}: {e}")
                
                saved_entries += 1
                row_data += 1
            except ValueError as e:
                ws_errors[f'A{row_errors}'] = filename
                ws_errors[f'B{row_errors}'] = f"ValueError: {e}"
                ws_errors[f'C{row_errors}'] = os.path.dirname(file_path)
                print(f"ValueError in {filename}: {e}")
                errors_encountered += 1
                row_errors += 1
            except TypeError as e:
                ws_errors[f'A{row_errors}'] = filename
                ws_errors[f'B{row_errors}'] = f"TypeError: {e}"
                ws_errors[f'C{row_errors}'] = os.path.dirname(file_path)
                print(f"TypeError in {filename}: {e}")
                errors_encountered += 1
                row_errors += 1
            except Exception as e:
                ws_errors[f'A{row_errors}'] = filename
                ws_errors[f'B{row_errors}'] = f"Error: {e}"
                ws_errors[f'C{row_errors}'] = os.path.dirname(file_path)
                print(f"Error processing {filename}: {e}")
                errors_encountered += 1
                row_errors += 1

# Iterate through all .gz files in the directory and its subdirectories
for root, dirs, files in os.walk(directory_path):
    for filename in files:
        if filename.endswith(".gz"):
            file_path = os.path.join(root, filename)
            
            try:
                with gzip.open(file_path, 'rt') as f:
                    for line in f:
                        if 'seed' in line.lower() and 'level.dat' not in file_path:
                            ws_log[f'A{row_log}'] = filename
                            ws_log[f'B{row_log}'] = os.path.dirname(file_path)
                            ws_log[f'C{row_log}'] = line.strip()
                            row_log += 1
            except Exception as e:
                print(f"Error reading {filename}: {e}")

# Save the workbook to location
output_dir = directory_path  
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
wb.save(os.path.join(output_dir, "minecraft_worlds.xlsx"))
print("Workbook saved successfully.")

# Print summary
print(f"Total .dat files processed: {processed_files}")
print(f"Entries successfully saved to Excel: {saved_entries}")
print(f"Errors encountered during processing: {errors_encountered}")

# Pause before closing
input("Press Enter to continue...")
