import os
from flask import Flask, request, render_template
from scrip import extract_rows_by_first_entity, write_to_new_docx

app = Flask(__name__)

# Define the upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Function to check if a file has an allowed extension
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Main route to upload and process the document
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
        entity_names = ["Overview", "Owner", "Event Type", "Currency", "Timing Rules", "Publish time", "Due Date", "Currency Rules", "Allow Participants to select bidding currency","Inco Term", "Inco Term Location", "Requested Delivery Date", "InternalNote" ]

        rows_dict = extract_rows_by_first_entity(filename, entity_names)
      
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'output_document.docx')
        write_to_new_docx(rows_dict, output_file)

        output_file_link = f"/download/{output_file.split('/')[-1]}"

        return render_template('index.html', output_file=output_file_link)

    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)
