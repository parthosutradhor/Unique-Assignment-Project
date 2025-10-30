import openpyxl
import hashlib
import re
import subprocess
import os
import shutil
import glob
import time
from datetime import datetime


# ==========================================
# CONFIGURATION
# ==========================================
ASSESSMENT_TYPE = "Assignment - 01"
SEMESTER_NAME = "FALL 2025"
COURSE_CODE = "MAT215"
COURSE_NAME = r"Complex Variables \& Laplace Transform"
SECTION = "12"
TOTAL_POINTS="100"

TEMPLATE_PATH = "question_template.txt"
WORKBOOK_PATH = "course-attendee.xlsx"
SHEET_NAME = "Worksheet"
LOGO_FILE = "Brac_University_Logo.png"

# Number of rows to skip before starting data (adjust as needed)
START_ROW = 39


# ==========================================
# UTILITY FUNCTIONS
# ==========================================
def generate_integers_range(ID, variation, n, a, b):
    """
    Generate deterministic pseudo-random integers from ID and variation.
    Ensures reproducibility based on hash seed.
    """
    seed = f"{ID}{variation}"
    hash_value = hashlib.md5(seed.encode()).hexdigest()
    return [(int(hash_value[i:i + 2], 32) % (b - a + 1)) + a for i in range(0, n * 2, 2)]


def replace_template_placeholders(template_text, **kwargs):
    """Replace placeholders like @Name@, @ID@, etc., in the LaTeX template."""
    for key, value in kwargs.items():
        template_text = template_text.replace(f"@{key}@", str(value))
    return template_text


def compile_latex_to_pdf(tex_path):
    """
    Compile a .tex file into a PDF using pdflatex.
    Runs twice for stable references.
    Shows log snippet if failed.
    """
    tex_dir = os.path.dirname(tex_path) or "."
    tex_file = os.path.basename(tex_path)

    cmd = ["pdflatex", "-interaction=nonstopmode", tex_file]

    for attempt in range(2):
        result = subprocess.run(
            cmd,
            cwd=tex_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        if result.returncode != 0 and attempt == 1:
            print(f"⚠️  pdflatex returned non-zero exit code on {tex_file}")
            print("\n".join(result.stdout.splitlines()[-15:]))
            break

    pdf_path = os.path.splitext(tex_path)[0] + ".pdf"

    # Wait up to 1 second for PDF write
    for _ in range(10):
        if os.path.exists(pdf_path):
            print(f"✅ PDF generated successfully at: {pdf_path}")
            return True
        time.sleep(0.1)

    print(f"❌ PDF generation failed for {tex_file}")
    return False


def safe_filename(name):
    """Convert name to filesystem-safe format."""
    return re.sub(r"[^\w\s-]", "", str(name)).strip().replace(" ", "_")


def delete_files_by_extension(directory, extensions):
    """Delete all files in directory (recursive) matching given extensions."""
    for ext in extensions:
        for path in glob.glob(os.path.join(directory, f"**/*{ext}"), recursive=True):
            try:
                os.remove(path)
            except Exception as e:
                print(f"⚠️ Could not delete {path}: {e}")


def delete_empty_dirs(directory):
    """Delete all empty subdirectories inside the directory."""
    for root, dirs, _ in os.walk(directory, topdown=False):
        for d in dirs:
            dir_path = os.path.join(root, d)
            try:
                shutil.rmtree(dir_path)
                print(f"🗑️ Deleted inner folder: {dir_path}")
            except Exception as e:
                print(f"⚠️ Could not remove folder {dir_path}: {e}")


# ==========================================
# PREPARATION
# ==========================================
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
print(f"📄 LaTeX PDF Batch Generator | Started at {timestamp}\n")

output_dir = ASSESSMENT_TYPE
os.makedirs(output_dir, exist_ok=True)
print(f"📁 Output directory ready: {output_dir}")

# Copy logo if available
if os.path.exists(LOGO_FILE):
    shutil.copy(LOGO_FILE, output_dir)
    print(f"🖼️ Copied logo → {output_dir}/")
else:
    print(f"⚠️ Logo not found: '{LOGO_FILE}' (skipped)")

# Load workbook and template
workbook = openpyxl.load_workbook(WORKBOOK_PATH)
sheet = workbook[SHEET_NAME]

with open(TEMPLATE_PATH, "r", encoding="utf-8") as f_template:
    template = f_template.read().replace("\r\n", "\n")

print("\n🚀 Starting PDF generation...\n")


# ==========================================
# MAIN LOOP
# ==========================================
count_success, count_fail = 0, 0
row = START_ROW

while True:
    ID = sheet.cell(row=row, column=1).value
    Name = sheet.cell(row=row, column=2).value

    if not ID:
        print(f"\n✅ Completed processing. {count_success} succeeded, {count_fail} failed.")
        break

    Name = Name or "Unknown"
    print(f"👤 Processing: {Name}")

    tex_content = replace_template_placeholders(
        template,
        Name=Name,
        ID=ID,
        Section=SECTION,
        Course_Name=COURSE_NAME,
        Course_Code=COURSE_CODE,
        Semester_Name=SEMESTER_NAME,
        Assesment_Type=ASSESSMENT_TYPE,
        Total_Points=TOTAL_POINTS
    )

    safe_name = safe_filename(Name)
    tex_filename = f"{ID}_{safe_name}.tex"
    tex_path = os.path.join(output_dir, tex_filename)

    # Write LaTeX file
    with open(tex_path, "w", encoding="utf-8") as f_out:
        f_out.write(tex_content + "\n")

    if compile_latex_to_pdf(tex_path):
        count_success += 1
    else:
        count_fail += 1

    row += 1


# ==========================================
# CLEANUP
# ==========================================
print("\n🧹 Performing cleanup...")

# Move all PDFs up one level (if nested)
for root, _, files in os.walk(output_dir):
    for file in files:
        if file.lower().endswith(".pdf"):
            src = os.path.join(root, file)
            dest = os.path.join(output_dir, file)
            if src != dest:
                shutil.move(src, dest)

# Remove auxiliary LaTeX files and logo
delete_files_by_extension(output_dir, [".tex", ".aux", ".log", ".out"])
if os.path.exists(os.path.join(output_dir, LOGO_FILE)):
    os.remove(os.path.join(output_dir, LOGO_FILE))

# Remove subfolders
delete_empty_dirs(output_dir)

print(f"\n✅ All done! {count_success} PDFs generated successfully.")
print(f"📂 Clean folder ready at: {output_dir}")
print(f"🕓 Finished at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
