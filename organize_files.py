import os
import glob
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास'
OUTPUT_DIR = os.path.join(BASE_DIR, "Final_PYQ_Output")

# We need to find the folders again because they have Hindi names
history_src = glob.glob(os.path.join(BASE_DIR, "*इतिहास*"))[0]
culture_src = glob.glob(os.path.join(BASE_DIR, "*कला*"))[0]

hist_target = os.path.join(OUTPUT_DIR, "A - राजस्थान का इतिहास")
cult_target = os.path.join(OUTPUT_DIR, "B - कला एवं संस्कृति")

os.makedirs(hist_target, exist_ok=True)
os.makedirs(cult_target, exist_ok=True)

# Original filenames mapping
hist_names = [os.path.basename(f).replace(".xlsx", "_PYQ_Format.xlsx") for f in glob.glob(os.path.join(history_src, "*.xlsx"))]
cult_names = [os.path.basename(f).replace(".xlsx", "_PYQ_Format.xlsx") for f in glob.glob(os.path.join(culture_src, "*.xlsx"))]

for f in glob.glob(os.path.join(OUTPUT_DIR, "*_PYQ_Format.xlsx")):
    fname = os.path.basename(f)
    if fname in hist_names:
        shutil.move(f, os.path.join(hist_target, fname))
        print(f"Moved {fname} to History")
    elif fname in cult_names:
        shutil.move(f, os.path.join(cult_target, fname))
        print(f"Moved {fname} to Culture")

print("Organization Complete!")
