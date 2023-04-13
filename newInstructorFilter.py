import xlsxFinder
import findDifferences

def main():
    findDifferences.findDifferences(*xlsxFinder.findXlsxFiles())    

if __name__ == '__main__':
    main()