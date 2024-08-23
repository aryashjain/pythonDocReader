import os
from flask import Flask, redirect, request, render_template, send_file, send_from_directory
from scrip import extract_rows_by_first_entity, write_to_new_docx, extract_rows_by_t5

app = Flask(__name__)

# Define the upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx', 'doc'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def insert_newlines( text,case):
    arr = [
    "Serial number (typical):", "Manufacturer Part Number :", "THIS PART IS USED FOR:",
    "Basic Data Text:", "Long Description", "Inspection Text:", "Parent Equipment:",
    "Classification Text:", "OEM Serial Number :", "Manufacturer number :", 
    "OEM Model Number :", "Drawing number:", "Sub Assesmbly:", "Position Number :", 
    "Assembly Number :", "Item Description:", "Manufacturer:", "PGCODE:", "Crossref:",
    "Old Number:", "Characteristics :", "MPN Text:", "Tag Number :", "Part Number :", 
    "Model(*):", "MODEL:", "P/N", "Class :", "Certificate:", "Description(*):"
    ]

    
    # Step 3 & 4: Iterate over the phrases and insert newlines
    for phrase in arr:
        if(case):
            text = text.replace(phrase, "\n" +"\n" + phrase)
        else:
            text = text.replace(phrase,"<br/> <br/>"+phrase)
    
    return text

# Function to check if a file has an allowed extension
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/process', methods=['POST'])
def process():
    # Process the form data
    # For demonstration purposes, we'll just redirect to the home page
    return redirect('/')

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')
        
        file = request.files['file']
        
        # Check if file is selected
        if file.filename == '':
            return render_template('index.html', error='No file selected')
        
        if not allowed_file(file.filename):
            return render_template('index.html', error='Invalid file extension')

        # Save the uploaded file
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)

        # Process the uploaded file
        t1 = ["Overview", "Owner", "Event Type","Currency"]
        t2 = ["Timing Rules", "Publish time", "Due Date"]
        t3 = [ "Currency Rules", "Allow Participants to select bidding currency"]
        t4 = ["Information No 1","Inco Term", "Inco Term Location"]
        t5 = ["Item No", "Item Description", "Quantity","UOM" ,"Requested Delivery Date","Price", "InternalNote" ]
       

        rows_dict1 = extract_rows_by_first_entity(filename, t1)
        rows_dict2 = extract_rows_by_first_entity(filename, t2)
        rows_dict3 = extract_rows_by_first_entity(filename, t3)
        rows_dict4 = extract_rows_by_first_entity(filename, t4)
        rows_2D =extract_rows_by_t5(filename, t5)
    
        d1 = {
        "Owner":rows_dict1["Owner"][0][1] if len(rows_dict1["Owner"]) > 0 else '-',
        "Event Type":rows_dict1["Event Type"][0][1] if len(rows_dict1["Event Type"]) > 0 else '-',
        "Currency":rows_dict1["Currency"][0][1] if len(rows_dict1["Currency"]) > 0 else '-',
        }
        d2 = {
        "Publish time":rows_dict2["Publish time"][0][1] if len(rows_dict2["Publish time"]) > 0 else '-',
        "Due Date":rows_dict2["Due Date"][0][1] if len(rows_dict2["Due Date"]) > 0 else '-',
        }
        d3 = {
        "Allow Participants to select bidding currency":rows_dict3["Allow Participants to select bidding currency"][0][1] if len(rows_dict3["Allow Participants to select bidding currency"]) > 0 else '-', 
        }
       
        d4 = {
        "Inco Term":rows_dict4["Inco Term"][0][2] if len(rows_dict4["Inco Term"]) > 0 else '-',
        "Inco Term Location":rows_dict4["Inco Term Location"][0][2] if len(rows_dict4["Inco Term Location"]) > 0 else '-',
        }

        rowsHTML=rows_2D

  
        for i in range(len(rows_2D)):   
            text = rows_2D[i][6]
            mst= insert_newlines(text,True)
            rows_2D[i][6] =mst

  

        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'output_document.docx')

        write_to_new_docx(rows_dict1,rows_dict2,rows_dict3,rows_dict4,rows_2D, output_file)
        output_file_link = f"/download/{output_file.split('/')[-1]}"
        for i in range(len(rowsHTML)):   
            t1 = rows_2D[i][6]
            modified_str = insert_newlines(t1,False)
            rowsHTML[i][6] = modified_str
        return render_template('index.html', d1 = d1, d2 = d2, d3=d3, d4=d4, data=rowsHTML, output_file=output_file_link)

    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)
