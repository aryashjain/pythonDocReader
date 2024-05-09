from docx import Document
import os

def extract_rows_by_first_entity(docx_file, entity_names):
    document = Document(docx_file)
    rows_dict = {entity_name: [] for entity_name in entity_names}
    st =set()

    for table in document.tables:
        for row in table.rows:
            # Extract the value of the first cell in the row
            first_cell_value = row.cells[0].text.strip()

            # Check if the value matches any of the specified entity names
            for entity_name in entity_names:
                if first_cell_value.lower() == entity_name.lower() and entity_name not in st:
                    rows_dict[entity_name].append([cell.text.strip() for cell in row.cells])
                    if(entity_name!="InternalNote"):
                        st.add(entity_name)
                    break

    return rows_dict

def write_to_new_docx(rows_dict, output_file):
    document = Document()
    document.add_heading('Selected Rows -:', level=1)
    for entity_name, rows in rows_dict.items():
        if rows:
            table = document.add_table(rows=1, cols=len(rows[0]))
            table.style = 'Table Grid'

            for i, cell_value in enumerate(rows[0]):
                table.cell(0, i).text = cell_value

            for row in rows[1:]:
                row_cells = table.add_row().cells
                for i, cell_value in enumerate(row):
                    row_cells[i].text = cell_value
    if os.path.exists(output_file):
        os.remove(output_file)

    document.save(output_file)


def main():
    # Provide the path to your Word document
    docx_file = 'test1.docx'


    entity_names = ["Overview", "Owner", "Event Type", "Currency", "Timing Rules", "Publish time", "Due Date", "Currency Rules", "Allow Participants to select bidding currency","Inco Term", "Inco Term Location", "Requested Delivery Date", "InternalNote" ]

    output_file = 'output.docx'  
    rows_dict = extract_rows_by_first_entity(docx_file, entity_names)
    write_to_new_docx(rows_dict, output_file)
    print("Rows extracted and written to new document successfully.")

if __name__ == "__main__":
    main()
