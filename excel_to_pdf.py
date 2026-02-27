import os
import sys
import win32com.client

def excel_to_pdf(excel_file_path, print_area=None, pdf_output_path=None, fit_to_one_page=True):
    """
    Convert Excel file to PDF format using Win32 COM (Windows only)
    
    Args:
        excel_file_path (str): Path to the Excel file
        print_area (str): Optional print area in Excel format (e.g., "A1:Z50"). If None, uses current print area
        pdf_output_path (str): Optional path for PDF output. If None, uses same name as Excel file
        fit_to_one_page (bool): If True, fits content to one page wide and tall
    """
    try:
        # Generate PDF output path if not provided
        if pdf_output_path is None:
            base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
            pdf_output_path = f"{base_name}.pdf"
        
        print(f"Converting {excel_file_path} to PDF: {pdf_output_path}")
        
        # Initialize Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            # Open workbook
            workbook = excel.Workbooks.Open(os.path.abspath(excel_file_path))
            
            # Get active worksheet
            worksheet = workbook.ActiveSheet
            
            # Set print area if specified
            if print_area:
                worksheet.PageSetup.PrintArea = print_area
                print(f"Set print area to: {print_area}")
            
            # Fit to one page if requested
            if fit_to_one_page:
                worksheet.PageSetup.Zoom = False
                worksheet.PageSetup.FitToPagesWide = 1
                worksheet.PageSetup.FitToPagesTall = 1
                print("Fitting content to one page")
            
            # Export to PDF
            worksheet.ExportAsFixedFormat(0, os.path.abspath(pdf_output_path))
            
            # Close workbook
            workbook.Close(False)
            
            print(f"Successfully converted to PDF: {pdf_output_path}")
            return pdf_output_path
            
        except Exception as e:
            print(f"Error during Excel operations: {str(e)}")
            return None
        finally:
            # Quit Excel application
            excel.Quit()
            
    except Exception as e:
        print(f"Error converting Excel to PDF: {str(e)}")
        print("Make sure Microsoft Excel is installed and pywin32 is available.")
        return None

def main():
    """Main function to handle command line arguments"""
    if len(sys.argv) < 2:
        print("Usage: python excel_to_pdf.py <excel_file_path> [print_area] [pdf_output_path] [fit_to_one_page]")
        print("Example: python excel_to_pdf.py data.xlsx 'A1:Z50' output.pdf true")
        print("Example: python excel_to_pdf.py data.xlsx 'A1:Z50' output.pdf false")
        print("Example: python excel_to_pdf.py data.xlsx")
        return
    
    excel_file = sys.argv[1]
    print_area = sys.argv[2] if len(sys.argv) > 2 else None
    pdf_file = sys.argv[3] if len(sys.argv) > 3 else None
    fit_one_page = sys.argv[4].lower() == 'true' if len(sys.argv) > 4 else True
    
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found.")
        return
    
    result = excel_to_pdf(excel_file, print_area, pdf_file, fit_one_page)
    
    if result:
        print(f"Conversion completed successfully!")
    else:
        print("Conversion failed!")

if __name__ == "__main__":
    main()
