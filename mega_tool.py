import os
import shutil
import win32com.client as win32
import logging

class MegaTool:
    def __init__(self):
        self.logger = logging.getLogger('MegaTool')
        logging.basicConfig(level=logging.INFO)
        
    def copy_file(self, src, dest):
        """Copies a file from the source to the destination directory."""
        try:
            shutil.copy(src, dest)
            self.logger.info(f"File {src} copied to {dest}")
        except Exception as e:
            self.logger.error(f"Failed to copy file: {e}")

    def move_file(self, src, dest):
        """Moves a file from the source to the destination directory."""
        try:
            shutil.move(src, dest)
            self.logger.info(f"File {src} moved to {dest}")
        except Exception as e:
            self.logger.error(f"Failed to move file: {e}")

    def delete_file(self, path):
        """Deletes a file at the given path."""
        try:
            os.remove(path)
            self.logger.info(f"File {path} deleted")
        except Exception as e:
            self.logger.error(f"Failed to delete file: {e}")

    def excel_automation(self, file_path):
        """Demonstrates basic Excel automation."""
        try:
            excel = win32.Dispatch('Excel.Application')
            workbook = excel.Workbooks.Open(file_path)
            sheet = workbook.Sheets(1)
            cell_value = sheet.Cells(1, 1).Value
            self.logger.info(f"Value in cell A1: {cell_value}")
            workbook.Close(SaveChanges=0)
            excel.Quit()
        except Exception as e:
            self.logger.error(f"Excel automation failed: {e}")

if __name__ == "__main__":
    tool = MegaTool()

    # Example usage
    tool.copy_file('path/to/source/file.txt', 'path/to/destination/')
    tool.move_file('path/to/source/file.txt', 'path/to/destination/')
    tool.delete_file('path/to/destination/file.txt')
    tool.excel_automation('path/to/excel/file.xlsx')