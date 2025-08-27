import os
import openpyxl

base_file = "INS(xxx).xlsx"
ref_file = "PaintCalculateFinal-01.xlsx"

if not os.path.exists(base_file):
    print(f"❌ فایل پایه پیدا نشد: {base_file}")
    raise SystemExit

if not os.path.exists(ref_file):
    print(f"❌ فایل مرجع پیدا نشد: {ref_file}")
    raise SystemExit

num_out = input("شماره ins: ").strip()
output_file = f"INS({num_out}).xlsx"

wb = openpyxl.load_workbook(base_file)
ws = wb.active

ref_wb = openpyxl.load_workbook(ref_file, data_only=True)
ref_ws = ref_wb.active

ws.title = f"نصب {num_out}"
ws["E2"].value = f"شماره صورتمجلس\n INS-{num_out}"

max_elements = 40
row_start = 6
elements_entered = 0

for i in range(1, max_elements + 1):
    row = row_start + (i - 1)

    if i == 1 or input(f"آیا می‌خواهی المان {i} را وارد کنی؟ (y/n): ").strip().lower() == "y":
        desc = input(f"توضیحات المان {i}: ")

        sharh = ""
        tol = ""
        tedad = ""
        vazn = ""

        for r in range(2, ref_ws.max_row + 1):
            if str(ref_ws[f"D{r}"].value).strip() == desc.strip():
                sharh = ref_ws[f"F{r}"].value
                tol_val = ref_ws[f"H{r}"].value
                tol = float(tol_val) / 1000 if tol_val else ""
                tedad = ref_ws[f"M{r}"].value
                vazn = ref_ws[f"L{r}"].value
                print(f"✅ '{desc}' پیدا شد.")
                break
        else:
            print(f"⚠️ '{desc}'  پیدا نشد → سلول‌های G, I, J, K خالی می‌مانند.")

        ws[f"D{row}"].value = desc
        ws[f"G{row}"].value = sharh
        ws[f"I{row}"].value = tol
        ws[f"J{row}"].value = tedad
        ws[f"K{row}"].value = vazn

        elements_entered += 1
    else:
        break

last_row = row_start + elements_entered
if elements_entered < max_elements:
    ws.delete_rows(last_row, max_elements - elements_entered)

print(f"✅ {elements_entered} المان وارد شد. ردیف‌های اضافی حذف شدند.")

wb.save(output_file)
print("✅ فایل خروجی ساخته شد:", output_file)
