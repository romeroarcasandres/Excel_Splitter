# Excel File Splitter Script
Advanced Excel file splitting tool with batch processing and customizable header management

## Overview:
This comprehensive script splits large Excel files into smaller, manageable chunks with full user control over processing parameters. It provides an interactive interface for selecting directories containing Excel files, allows customization of chunk sizes, and offers flexible header row management across all split files. The script intelligently processes multiple Excel files in batch mode while organizing output files in structured directories for easy navigation and management.

## Requirements:
- Python 3.6+
- tkinter library (for interactive dialog interface)
- pandas library (for Excel file processing and manipulation)
- openpyxl library (for Excel file reading and writing)
- os library (for file and directory operations)

## Files
excel_splitter.py

## Installation
Before running the script, install the required dependencies:
```bash
pip install pandas openpyxl
```

## Usage
1. Run the script: `python excel_splitter.py`
2. A directory dialog will prompt you to select a folder containing Excel files
3. Enter the desired number of rows per split file when prompted
4. Choose whether to preserve header rows in all split files (Yes/No dialog)
5. The script processes all Excel files in the directory and generates:
   - Individual split files organized in dedicated subfolders
   - Progress feedback showing processing status for each file
   - Summary report of total files processed and chunks created

## Key Features
- **Batch Directory Processing**: Processes all Excel files (.xlsx and .xls) in a selected directory
- **Interactive Parameter Setup**: User-friendly dialogs for chunk size and header preferences
- **Flexible Chunk Sizing**: Customizable row count per split file (minimum 1 row)
- **Smart Header Management**: Option to preserve original headers across all split files
- **Organized Output Structure**: Creates dedicated subfolders for each original file's splits
- **Robust Error Handling**: Gracefully handles empty files, invalid formats, and processing errors
- **Progress Monitoring**: Real-time feedback on file processing status
- **File Type Support**: Compatible with both .xlsx and .xls Excel formats
- **Unicode Compatibility**: Full UTF-8 support for international character sets
- **Memory Efficient**: Processes large files without excessive memory usage

## Supported File Formats
- **Excel 2007+**: .xlsx files (OpenXML format)
- **Excel 97-2003**: .xls files (Binary format)
- **Data Types**: All Excel data types preserved including text, numbers, dates, formulas
- **Encoding**: Full Unicode (UTF-8) support for international characters
- **Cell Formatting**: Basic formatting preserved during split operations

## File Organization Structure
The script creates an organized directory structure:
```
selected_directory/
├── original_file1.xlsx
├── original_file1_splits/
│   ├── original_file1_part_1.xlsx
│   ├── original_file1_part_2.xlsx
│   ├── original_file1_part_3.xlsx
│   └── ...
├── original_file2.xlsx
├── original_file2_splits/
│   ├── original_file2_part_1.xlsx
│   ├── original_file2_part_2.xlsx
│   └── ...
└── ...
```

## Header Management Options
- **Keep Headers (Yes)**: First row from original file is added to every split file
- **No Headers (No)**: Split files contain only data rows without header preservation
- **Smart Detection**: Automatically handles files with or without header rows

## Example Usage Scenarios

### Scenario 1: Large Dataset Distribution
For a 10,000-row Excel file that needs to be distributed in 500-row chunks:
1. Select directory containing the large file
2. Enter "500" for chunk size
3. Choose "Yes" to keep headers
4. Result: 20 split files, each with headers + 500 data rows

### Scenario 2: Batch Processing Multiple Files
For a directory with multiple Excel files requiring 100-row splits:
```
input_directory/
├── sales_data_2023.xlsx (5,000 rows)
├── customer_list.xlsx (2,300 rows)
└── inventory_report.xlsx (800 rows)
```
Result: 3 split folders with 50, 23, and 8 split files respectively

### Scenario 3: Data Analysis Preparation
For preparing large datasets for analysis tools with memory limitations:
1. Select directory with analysis files
2. Set chunk size based on tool requirements
3. Preserve headers for consistent data structure
4. Import split files individually into analysis software

## Interactive Dialog Sequence
```
Step 1: Directory Selection
[Folder Browser Dialog]
"Select directory containing Excel files"

Step 2: Chunk Size Input
[Number Input Dialog]
"How many rows per split file?"
Default: 100, Minimum: 1

Step 3: Header Preference
[Yes/No Dialog]
"Do you want to keep the header row in all split files?"
```

## Output and Progress Feedback
```
Found 3 Excel file(s)
Chunk size: 100 rows
Keep header: Yes
--------------------------------------------------

Processing: sales_data.xlsx
Saved: sales_data_splits/sales_data_part_1.xlsx
Saved: sales_data_splits/sales_data_part_2.xlsx
Saved: sales_data_splits/sales_data_part_3.xlsx
Split sales_data into 3 files

Processing: customer_data.xlsx
Saved: customer_data_splits/customer_data_part_1.xlsx
Saved: customer_data_splits/customer_data_part_2.xlsx
Split customer_data into 2 files

==================================================
All files processed successfully!
```

## Performance Considerations
- **Memory Usage**: Processes files in chunks to minimize memory footprint
- **Processing Speed**: Optimized for large files with efficient pandas operations
- **Storage Efficiency**: Maintains original file compression where possible
- **Scalability**: Handles directories with multiple large files efficiently

## Error Handling and Recovery
- **Empty Files**: Automatically skips empty Excel files with notification
- **Invalid Formats**: Detects and reports corrupted or invalid Excel files
- **Permission Errors**: Handles file access restrictions with clear error messages
- **Disk Space**: Monitors available storage during processing
- **User Cancellation**: Graceful handling of dialog cancellations
- **Processing Interruption**: Safe termination without partial file corruption

## Advanced Features
- **Dynamic Memory Management**: Adjusts processing based on available system memory
- **Concurrent Processing**: Efficient handling of multiple files in sequence
- **Progress Tracking**: Real-time status updates for long-running operations
- **File Validation**: Pre-processing validation to identify potential issues
- **Cleanup Operations**: Automatic handling of temporary files and resources
- **Format Preservation**: Maintains original Excel formatting and data types

## Common Use Cases
- **Database Import Preparation**: Split large exports for database batch imports
- **Email Distribution**: Create manageable file sizes for email attachments
- **Analysis Tool Integration**: Prepare data for tools with row limitations
- **Team Collaboration**: Distribute sections of large datasets to team members
- **System Migration**: Break down large files for system transfer processes
- **Quality Assurance**: Create smaller files for detailed review and validation

## Troubleshooting
- **"No directory selected"**: Ensure a valid directory is chosen in the dialog
- **"No Excel files found"**: Verify directory contains .xlsx or .xls files
- **"Invalid chunk size"**: Enter a positive integer value for rows per file
- **"Error processing file"**: Check file permissions and Excel file integrity
- **"No file selected"**: Complete the directory selection process
- **Memory errors**: Reduce chunk size for very large files

## File Size Guidelines
- **Small Files** (< 1MB): Any chunk size suitable
- **Medium Files** (1-10MB): Chunk sizes of 100-1000 rows recommended
- **Large Files** (10-100MB): Chunk sizes of 500-2000 rows for optimal performance
- **Very Large Files** (> 100MB): Chunk sizes of 1000+ rows, monitor memory usage

## Important Notes
- Ensure sufficient disk space for split files (approximately same size as original)
- Original files remain unchanged during the splitting process
- Split files maintain original data types and basic formatting
- Header preservation works best with consistently formatted source files
- The script handles various Excel cell types including formulas and dates
- Processing time scales with file size and number of files in directory
- All text encoding is preserved using UTF-8 standards

## License
This project is governed by the CC BY-NC 4.0 license. For comprehensive details, kindly refer to the LICENSE file included with this project.
