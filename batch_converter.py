import os
import sys
from excel_to_pdf import excel_to_pdf

class ExcelToPDFBatchConverter:
    """
    A class to batch convert Excel files to PDF with custom print areas
    """
    
    def __init__(self, folder_path, print_area=None, output_folder=None, fit_to_one_page=True):
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
        
        # Generate output path
        if self.output_folder:
            os.makedirs(self.output_folder, exist_ok=True)
            pdf_name = os.path.splitext(excel_file)[0] + '.pdf'
            pdf_path = os.path.join(self.output_folder, pdf_name)
        else:
            pdf_path = None  # Use default naming
        
        print(f"Converting: {excel_file}")
        result = excel_to_pdf(excel_path, self.print_area, pdf_path, self.fit_to_one_page)
        
        if result:
            self.converted_files.append((excel_file, result))
            print(f"✓ Success: {excel_file} -> {result}")
        else:
            self.failed_files.append(excel_file)
            print(f"✗ Failed: {excel_file}")
        
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

def main():
    """Main function for command line usage"""
    if len(sys.argv) < 2:
        print("Usage: python batch_converter.py <folder_path> [print_area] [output_folder] [fit_to_one_page]")
        print("Example: python batch_converter.py ./excel_files 'A1:Z50' ./pdf_output true")
        print("Example: python batch_converter.py ./excel_files 'A1:Z50'")
        print("Example: python batch_converter.py ./excel_files")
        return
    
    folder_path = sys.argv[1]
    print_area = sys.argv[2] if len(sys.argv) > 2 else None
    output_folder = sys.argv[3] if len(sys.argv) > 3 else None
    fit_one_page = sys.argv[4].lower() == 'true' if len(sys.argv) > 4 else True
    
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
