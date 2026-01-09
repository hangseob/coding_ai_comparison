import xlwings as xw
import sys

def find_cell_value():
    try:
        # Set output to utf-8
        sys.stdout.reconfigure(encoding='utf-8')
        
        found = False
        for b in xw.books:
            for s in b.sheets:
                if '가나다' in s.name:
                    val = s.range('A1').value
                    print(f"FOUND|{b.name}|{s.name}|{val}")
                    found = True
        
        if not found:
            # Print current state if not found
            print("NOT_FOUND")
            print("Current Open Books and Sheets:")
            for b in xw.books:
                print(f" - Book: {b.name}, Sheets: {[s.name for s in b.sheets]}")
                
    except Exception as e:
        print(f"ERROR: {str(e)}")

if __name__ == "__main__":
    find_cell_value()

