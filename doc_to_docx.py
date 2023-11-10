import os
import subprocess

input_path = "/Users/harvin/code/docx_merge/data/07"

for root, folders, files in os.walk(input_path):
    for file in files:
        if file.endswith(".doc"):
            file_path = os.path.abspath(os.path.join(root, file))
            path = os.path.dirname(file_path)
            subprocess.run(["/Applications/LibreOffice.app/Contents/MacOS/soffice", "--headless", "--convert-to", "docx", file_path, "--outdir", path])
            os.remove(file_path)

print('Success!')
