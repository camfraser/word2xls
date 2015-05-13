# Copy .docx files into the directory
Copy all docx files you'd like to process and convert into a single XLS file to the
current directory. Each docx file will be represented in a Sheet in the Excel workbook.
Urban legend suggests the maximum number of sheets an Excel document can handle is 255,
but others report successfully including over 2000 sheets. YMMV.

# Create your virtual environment
    virtualenv -q --prompt="(word2xls)" env
    . env/bin/activate
    pip install -r requirements.txt

# Run the python script
    python word2xls.py

You will receive a notice similar to the following:

    5 docx file(s) converted to docs_2015-05-13T02-25-53.xls

Open the Excel file and reap the rewards that only scripting can offer.