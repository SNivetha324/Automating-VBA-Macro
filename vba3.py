import os
import win32com.client
import pythoncom
import pydotplus
from docx import Document
import matplotlib.pyplot as plt
import seaborn as sns
from flask import Flask, render_template, request, send_file

# Initialize COM
pythoncom.CoInitialize()

# Flask app setup
app = Flask(__name__)

# Function to extract VBA code from Excel
def extract_vba_code(excel_file):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(excel_file)
        vba_project = workbook.VBProject

        vba_code = {}
        for component in vba_project.VBComponents:
            if component.Type in [1, 2, 3]:  # Standard, Class, and Module
                vba_code[component.Name] = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
        
        workbook.Close(False)
        excel.Quit()
        return vba_code
    except Exception as e:
        print(f"Error extracting VBA code: {e}")
        return None

# Function to parse VBA code structure
def parse_vba_code(vba_code):
    parsed_code = {}
    for module_name, code in vba_code.items():
        # Implement your parsing logic here
        # Example: Use regular expressions or custom logic to parse the VBA code
        # For demonstration, we'll just store the code as-is
        parsed_code[module_name] = code
    return parsed_code

# Function to generate process flow visualization
def generate_process_flow(parsed_code, output_dir):
    os.makedirs(output_dir, exist_ok=True)

    for module_name, code in parsed_code.items():
        # Create a new graph for each module
        graph = pydotplus.Dot(graph_type='digraph', rankdir='LR')
        
        # Nodes represent functions/subs and edges represent flow between them
        lines = code.splitlines()
        current_node = None
        nodes = {}

        for line in lines:
            line = line.strip()
            if line.startswith('Sub ') or line.startswith('Function '):
                func_name = line.split('(')[0].split()[-1]
                current_node = pydotplus.Node(func_name, shape='box')
                graph.add_node(current_node)
                nodes[func_name] = current_node
            elif line.startswith('End Sub') or line.startswith('End Function'):
                current_node = None
            elif current_node and line.startswith(('Call ', func_name)):
                called_func = line.split('(')[0].split()[-1]
                if called_func in nodes:
                    graph.add_edge(pydotplus.Edge(nodes[func_name], nodes[called_func]))

        graph.write_png(os.path.join(output_dir, f'{module_name}_process_flow.png'))

# Main function to run the entire process
def main(excel_file, output_dir):
    vba_code = extract_vba_code(excel_file)
    parsed_code = parse_vba_code(vba_code)
    generate_process_flow(parsed_code, output_dir)

# Flask routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        excel_file = os.path.join('uploads', file.filename)
        file.save(excel_file)
        output_dir = os.path.join('outputs', os.path.splitext(file.filename)[0])
        main(excel_file, output_dir)
        return send_file(output_dir, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
