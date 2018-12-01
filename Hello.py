"""Script to gather IMDB keywords from 2013's top grossing movies."""
import sys
import openpyxl

def main():
    """Main entry point for the script."""
    roster_file = openpyxl.load_workbook("C:\\Users\\User\\Desktop\\Roster.xlsx")
    print (roster_file)
    print (roster_file.sheetnames)

if __name__ == '__main__':
    sys.exit(main())
