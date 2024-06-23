import win32com.client
import os 

def get_word_documents_in_current_path():
    # Get the current working directory
    current_path = os.getcwd()
    # Find all subdirectories in the current directory
    subdirectories = [os.path.join(current_path, subdir) for subdir in os.listdir(current_path) if os.path.isdir(os.path.join(current_path, subdir))]
    word_files = []
    for directory in subdirectories:
        # Use glob to find all .docx files in the current directory and subdirectories
        word_files.extend(glob.glob(os.path.join(directory, '**', '*.docx'), recursive=True))
    return word_files

def count_pages_in_word_doc(file_path):
    # Open Word application
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    # Open the document
    doc = word_app.Documents.Open(file_path)
    # Count the number of pages
    num_pages = doc.ComputeStatistics(2)
    # Close the document and the Word application
    doc.Close(False)
    word_app.Quit()
    return num_pages

word_documents = get_word_documents_in_current_path()

for file in word_documents:
    count_pages_in_word_doc(file)