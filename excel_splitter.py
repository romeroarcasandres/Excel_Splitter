import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox, simpledialog

def get_user_inputs():
    """Get user inputs for directory, chunk size, and header preference"""
    root = Tk()
    root.withdraw()  # Hide the root window
    
    # Select directory
    directory = filedialog.askdirectory(title="Select directory containing Excel files")
    if not directory:
        print("No directory selected.")
        return None, None, None
    
    # Get chunk size
    try:
        chunk_size = simpledialog.askinteger(
            "Chunk Size", 
            "How many rows per split file?",
            minvalue=1,
            initialvalue=100
        )
        if chunk_size is None:
            print("No chunk size provided.")
            return None, None, None
    except:
        print("Invalid chunk size.")
        return None, None, None
    
    # Ask about header
    keep_header = messagebox.askyesno(
        "Header Option", 
        "Do you want to keep the header row in all split files?"
    )
    
    root.destroy()
    return directory, chunk_size, keep_header

def split_excel_file(file_path, chunk_size, keep_header):
    """Split a single Excel file into chunks"""
    try:
        # Load the Excel file
        df = pd.read_excel(file_path, dtype=str)
        
        if len(df) == 0:
            print(f"Skipping empty file: {file_path}")
            return
        
        # Get file directory and base name
        file_dir = os.path.dirname(file_path)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        
        # Create splits folder
        splits_dir = os.path.join(file_dir, f"{file_name}_splits")
        os.makedirs(splits_dir, exist_ok=True)
        
        # Handle header
        if keep_header and len(df) > 0:
            header_row = df.iloc[0:1]  # First row as header
            data_rows = df.iloc[1:]    # Rest of the data
        else:
            header_row = None
            data_rows = df
        
        # Split and save chunks
        total_chunks = 0
        for i, chunk_start in enumerate(range(0, len(data_rows), chunk_size), start=1):
            chunk_end = min(chunk_start + chunk_size, len(data_rows))
            df_chunk = data_rows.iloc[chunk_start:chunk_end]
            
            # Add header if requested
            if keep_header and header_row is not None:
                df_chunk = pd.concat([header_row, df_chunk], ignore_index=True)
            
            output_path = os.path.join(splits_dir, f"{file_name}_part_{i}.xlsx")
            df_chunk.to_excel(output_path, index=False)
            print(f"Saved: {output_path}")
            total_chunks += 1
        
        print(f"Split {file_name} into {total_chunks} files")
        
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")

def main():
    """Main function to process all Excel files in directory"""
    # Get user inputs
    directory, chunk_size, keep_header = get_user_inputs()
    
    if directory is None:
        return
    
    # Find all Excel files in directory
    excel_files = []
    for file in os.listdir(directory):
        if file.lower().endswith(('.xlsx', '.xls')):
            excel_files.append(os.path.join(directory, file))
    
    if not excel_files:
        print(f"No Excel files found in {directory}")
        return
    
    print(f"Found {len(excel_files)} Excel file(s)")
    print(f"Chunk size: {chunk_size} rows")
    print(f"Keep header: {'Yes' if keep_header else 'No'}")
    print("-" * 50)
    
    # Process each Excel file
    for file_path in excel_files:
        print(f"\nProcessing: {os.path.basename(file_path)}")
        split_excel_file(file_path, chunk_size, keep_header)
    
    print("\n" + "=" * 50)
    print("All files processed successfully!")

if __name__ == "__main__":
    main()