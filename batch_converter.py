import os
import sys
from excel_to_pdf import excel_to_pdf
import win32com.client

class ExcelToPDFBatchConverter:
    """
    A class to batch convert Excel files to PDF with custom print areas
    """
    
    def __init__(self, folder_path, output_folder=None, fit_to_one_page=True, print_area=None):
        """
        Initialize the batch converter
        
        Args:
            folder_path (str): Path to folder containing Excel files
            print_area (str): Print area in Excel format (e.g., "A1:Z50")
            output_folder (str): Optional folder for PDF output. If None, uses same folder as Excel files
            fit_to_one_page (bool): If True, fits content to one page
        """
        self.folder_path = folder_path
        self.print_area = print_area
        self.output_folder = output_folder
        self.fit_to_one_page = fit_to_one_page
        self.converted_files = []
        self.failed_files = []
        
    def get_excel_files(self):
        """Get list of Excel files in the folder"""
        excel_files = []
        if not os.path.exists(self.folder_path):
            print(f"Error: Folder '{self.folder_path}' not found.")
            return excel_files
            
        for file in os.listdir(self.folder_path):
            if file.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')):
                excel_files.append(file)
        
        return excel_files
    
    def convert_single_file(self, excel_file):
        """Convert a single Excel file to PDF"""
        excel_path = os.path.join(self.folder_path, excel_file)
        auto_detect = False
        
        # Generate output path
        if self.output_folder:
            os.makedirs(self.output_folder, exist_ok=True)
            pdf_name = os.path.splitext(excel_file)[0] + '.pdf'
            pdf_path = os.path.join(self.output_folder, pdf_name)
        else:
            pdf_path = None  # Use default naming

        if self.print_area is None:
            print("Auto-detecting data range for each file...")
            self.print_area = self.detect_data_range(excel_path)
            auto_detect = True

        print(f"Converting: {excel_file}")
        result = excel_to_pdf(excel_path, self.print_area, pdf_path, self.fit_to_one_page)
        
        if result:
            self.converted_files.append((excel_file, result))
            print(f"✓ Success: {excel_file} -> {result}")
        else:
            self.failed_files.append(excel_file)
            print(f"✗ Failed: {excel_file}")

        if auto_detect:
            self.print_area = None
        
        return result
    
    def convert_all(self):
        """Convert all Excel files in the folder"""
        excel_files = self.get_excel_files()
        
        if not excel_files:
            print("No Excel files found in the folder.")
            return
        
        print(f"Found {len(excel_files)} Excel files to convert...")
        print(f"Print area: {self.print_area if self.print_area else 'Default'}")
        print(f"Fit to one page: {self.fit_to_one_page}")
        print("-" * 50)
        
        for excel_file in excel_files:
            self.convert_single_file(excel_file)
        
        self.print_summary()
    
    def print_summary(self):
        """Print conversion summary"""
        print("\n" + "=" * 50)
        print("CONVERSION SUMMARY")
        print("=" * 50)
        print(f"Successfully converted: {len(self.converted_files)}")
        print(f"Failed conversions: {len(self.failed_files)}")
        
        if self.converted_files:
            print("\nSuccessfully converted files:")
            for excel_file, pdf_file in self.converted_files:
                print(f"  {excel_file} -> {pdf_file}")
        
        if self.failed_files:
            print("\nFailed files:")
            for excel_file in self.failed_files:
                print(f"  {excel_file}")

    def detect_data_range(self, excel_path):
        """
        Automatically detect the data range using End(xlUp) method
        
        Args:
            excel_path (str): Path to Excel file
            
        Returns:
            str: Print area in Excel format (e.g., "A1:Z50")
        """
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            try:
                workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
                worksheet = workbook.ActiveSheet
                
                # Constants for Excel
                xlUp = -4162  # Excel constant for End(xlUp)
                
                # Find the last column with data by checking from bottom of each column
                max_row = 0
                max_col = 0
                
                # Start with column A and go right (columns 1 to 256 = A to IV)
                for col in range(1, 257):  # 1-based indexing for Excel columns
                    # Start from the bottom of the column (row 65536 for Excel 2003 format)
                    # and simulate Ctrl+Up arrow to find the last cell with content
                    bottom_cell = worksheet.Cells(65536, col)
                    last_cell_in_col = bottom_cell.End(xlUp)
                    
                    # Check if we found a cell with content (not in row 1)
                    if last_cell_in_col.Row > 1:
                        # Check if this cell actually has content
                        if last_cell_in_col.Value is not None and str(last_cell_in_col.Value).strip() != "":
                            # Found content in this column
                            current_row = last_cell_in_col.Row
                            current_col = col
                            if current_row > max_row:
                                max_row = current_row
                            if current_col > max_col:
                                max_col = current_col
                            
                            print(f"Found data in column {self.number_to_column_letter(col)} at row {current_row}")
                
                if max_row == 0:
                    print(f"No data found in {excel_path}")
                    workbook.Close(False)
                    return None
                
                # The first cell will always be A1 (top-left)
                first_cell = worksheet.Range("A1")
                
                if first_cell is None:
                    print(f"No data found in {excel_path}")
                    workbook.Close(False)
                    return None
                
                first_row = first_cell.Row
                first_col = first_cell.Column
                
                # Add 1 to each for better spacing
                last_row = max_row + 1
                last_col = max_col + 1
                
                # Convert column numbers to letters
                first_col_letter = self.number_to_column_letter(first_col)
                last_col_letter = self.number_to_column_letter(last_col)
                
                # Create print area
                print_area = f"{first_col_letter}{first_row}:{last_col_letter}{last_row}"
                
                print(f"Detected range: {print_area}")
                print(f"First cell: {first_col_letter}{first_row}, Last cell: {last_col_letter}{max_row}")
                
                workbook.Close(False)
                return print_area
                
            except Exception as e:
                print(f"Error detecting range for {excel_path}: {str(e)}")
                workbook.Close(False)
                return None
            finally:
                excel.Quit()
                
        except Exception as e:
            print(f"Error opening Excel for range detection: {str(e)}")
            return None

    def number_to_column_letter(self, n):
        """
        Convert a column number to a letter (e.g., 1 -> A, 2 -> B, etc.)
        
        Args:
            n (int): Column number
        
        Returns:
            str: Column letter
        """
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

def main():
    """Main function for command line usage"""
    if len(sys.argv) < 2:
        print("Usage: python batch_converter.py <folder_path> [fit_to_one_page] [output_folder] [print_area]")
        print('Example: python batch_converter.py ./excel_files ./pdf_output true "A1:Z50"')
        print("Example: python batch_converter.py ./excel_files ./pdf_output")
        print("Example: python batch_converter.py ./excel_files")
        return
    
    folder_path = sys.argv[1]
    output_folder = sys.argv[2] if len(sys.argv) > 2 else None
    fit_one_page = sys.argv[3].lower() == 'true' if len(sys.argv) > 3 else True
    print_area = sys.argv[4] if len(sys.argv) > 4 else None


    # Create converter and run
    converter = ExcelToPDFBatchConverter(
        folder_path=folder_path,
        print_area=print_area,
        output_folder=output_folder,
        fit_to_one_page=fit_one_page
    )
    
    converter.convert_all()

if __name__ == "__main__":
    main()
