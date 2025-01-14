import os
import pandas as pd

def extract_mean_row_values(folder_path):
    """
    Extract just the mean values from each Excel file and format them simply
    
    Parameters:
    folder_path (str): Path to the folder containing Excel files
    
    Returns:
    list: List of mean value strings in the requested format
    """
    mean_values = []
    
    try:
        # List all Excel files in the folder, excluding temporary files
        excel_files = [f for f in os.listdir(folder_path) 
                      if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
        
        if not excel_files:
            print(f"No Excel files found in {folder_path}")
            return mean_values
            
        # Process each Excel file
        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            try:
                # Read Excel file
                df = pd.read_excel(file_path)
                
                # Find the row containing "Mean" in the first column
                mean_row_idx = df[df.iloc[:, 0].astype(str).str.contains('Mean', case=False)].index
                
                if not mean_row_idx.empty:
                    # Get the last row containing "Mean"
                    mean_row = df.iloc[mean_row_idx[-1]]
                    
                    # Format the values as required
                    formatted_mean = f"Mean {mean_row[6]:.4f} {mean_row[7]:.4f} {mean_row[11]:.4f} {mean_row[12]:.4f} {mean_row[13]:.4f} {mean_row[14]:.4f}"
                    mean_values.append(formatted_mean)
                    print(f"Successfully extracted mean values from {file}")
                else:
                    print(f"No Mean row found in {file}")
                    
            except Exception as e:
                print(f"Error processing file {file}: {str(e)}")
                
    except Exception as e:
        print(f"Error accessing folder: {str(e)}")
        
    return mean_values

# Example usage
if __name__ == "__main__":
    # Replace this with your folder path
    folder_path = r"C:\Users\adity\Downloads\COCO Dataset.v8-yolov8m.coco\Data_percobaan\02012025\Data_pagi"
    
    # Extract values
    results = extract_mean_row_values(folder_path)
    
    if results:
        # Save to text file
        output_file = "MEAN_DATA_02012025.txt"
        with open(output_file, 'w') as f:
            f.write(' '.join(results))
        
        # Display results
        print("\nExtracted mean values:")
        print(' '.join(results))
        print(f"\nResults have been saved to {output_file}")
    else:
        print("\nNo mean values were found to save.")