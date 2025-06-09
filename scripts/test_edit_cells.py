import xlwings as xw
import sys

print("=== Testing Edit Cells ===")

try:
    # Excel接続テスト
    print("1. Checking Excel connection...")
    app = xw.apps.active
    print(f"   ✓ Connected to Excel app: {app}")
    
    # アクティブブックの確認
    print("\n2. Checking active workbook...")
    wb = app.books.active
    if wb:
        print(f"   ✓ Active workbook: {wb.name}")
    else:
        print("   ✗ No active workbook found")
        sys.exit(1)
    
    # アクティブシートの確認
    print("\n3. Checking active sheet...")
    sheet = wb.sheets.active
    print(f"   ✓ Active sheet: {sheet.name}")
    
    # セル編集テスト
    print("\n4. Testing cell edit...")
    test_value = "Test from Python"
    sheet.range("A1").value = test_value
    print(f"   ✓ Set A1 to: {test_value}")
    
    # 値の読み取り確認
    print("\n5. Verifying value...")
    read_value = sheet.range("A1").value
    print(f"   ✓ Read A1 value: {read_value}")
    
    if read_value == test_value:
        print("\n✅ Test passed! Edit cells is working correctly.")
    else:
        print("\n❌ Test failed! Values don't match.")
        
except Exception as e:
    print(f"\n❌ Error: {type(e).__name__}: {str(e)}")
    print("\n📋 Troubleshooting tips:")
    print("   1. Make sure Excel is running")
    print("   2. Make sure you have an Excel file open")
    print("   3. Make sure the file is not read-only")
    print("   4. Check if you have xlwings installed: pip install xlwings")