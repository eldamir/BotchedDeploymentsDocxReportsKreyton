import re
import os
import zipfile
import shutil
from bs4 import BeautifulSoup
from openpyxl.reader.excel import load_workbook

# The values to put in the PieChart
list_of_labels = ["foo", "bar", "baz"]
list_of_values = [25, 42, 30]

row_count = len(list_of_labels)

template_path = "temp.docx"
temp_dir = "/tmp/workdir"

# ------------------------------
# Unzip the docx
# ------------------------------
os.makedirs(temp_dir, exist_ok=True)
with zipfile.ZipFile(template_path, "r") as zip_ref:
    zip_ref.extractall(temp_dir)

# ------------------------------
# Load and fix the docx xlsx
# ------------------------------
xlsx_path = os.path.join(
    temp_dir, "word", "embeddings", "Microsoft_Excel_Worksheet.xlsx"
)
workbook = load_workbook(xlsx_path)
sheet = workbook.active
for i, label in enumerate(list_of_labels):
    sheet[f"A{i+2}"] = label
for i, value in enumerate(list_of_values):
    sheet[f"B{i+2}"] = value

# We can delete any extra rows of data left behind
sheet.delete_rows(row_count + 2, 100)

workbook.save(xlsx_path)
workbook.close()

# ------------------------------
# Load and fix the docx xml
# ------------------------------
chart_xml_path = os.path.join(temp_dir, "word", "charts", "chart1.xml")
with open(chart_xml_path) as xml_file:
    contents = xml_file.read()

soup = BeautifulSoup(contents, "xml")
plot_area = soup.find("c:plotArea")

# Fix categories/labels of the pie chart
# First the labels in the cache
cat = plot_area.find("c:ser").find("c:cat")
cache = cat.find("c:strCache")

cache.clear()
ptCount = soup.new_tag("c:ptCount", val=str(len(list_of_labels)))
cache.append(ptCount)
for i, key in enumerate(list_of_labels):
    pt = soup.new_tag("c:pt", idx=str(i))
    v = soup.new_tag("c:v")
    v.string = key
    pt.append(v)
    cache.append(pt)
# Then adjust the range to fit the amount of rows used
sheetRange = cat.find("c:strRef").find("c:f")
sheetRange.string = re.sub(r"(^.*!\$\w+\$\d+:\$\w+\$)(\d+)$", f"\\g<1>{str(row_count + 1)}", sheetRange.string)

# Fix values of the chart
# First the values in the cache
val = plot_area.find("c:ser").find("c:val")
cache = val.find("c:numCache")

cache.clear()
ptCount = soup.new_tag("c:ptCount", val=str(len(list_of_values)))
cache.append(ptCount)
for i, key in enumerate(list_of_values):
    pt = soup.new_tag("c:pt", idx=str(i))
    v = soup.new_tag("c:v")
    v.string = str(key)
    pt.append(v)
    cache.append(pt)
# Then adjust the range to fit the amount of rows used
sheetRange = val.find("c:numRef").find("c:f")
sheetRange.string = re.sub(r"(^.*!\$\w+\$\d+:\$\w+\$)(\d+)$", f"\\g<1>{str(row_count + 1)}", sheetRange.string)

with open(chart_xml_path, "w") as xml_file:
    xml_file.write(str(soup))

# ------------------------------
# Recompress and remove tmp folder
# ------------------------------
destination_file = os.path.join(
    os.path.dirname(__file__),
    "docx_templates",
    f"my_finished_report.docx",
)
with zipfile.ZipFile(destination_file, "w") as new_zip:
    for foldername, subfolders, filenames in os.walk(temp_dir):
        for filename in filenames:
            file_path = os.path.join(foldername, filename)
            arcname = os.path.relpath(file_path, temp_dir)
            new_zip.write(file_path, arcname)

shutil.rmtree(temp_dir)