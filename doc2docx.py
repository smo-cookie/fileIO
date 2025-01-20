import glob
import re
import os
import win32com.client

def convert_doc_to_docx(directory='.'):
    """
    Converts all .doc files in the specified directory to .docx format.

    Parameters:
    directory (str): The directory to search for .doc files. Defaults to the current directory.

    Returns:
    None
    """
    word = win32com.client.Dispatch('Word.Application')
    
    try:
        doc_files = glob.glob(os.path.join(directory, '*.doc'))
        if not doc_files:
            print('No .doc files found.')
            return

        for file in doc_files:
            full_path = os.path.abspath(file)
            print(f'Converting {file} to docx...')
            
            doc = word.Documents.Open(full_path)
            new_path = re.sub(r'\.\w+$', '.docx', full_path)
            doc.SaveAs(new_path, 16)  # 16 is the format code for .docx
            doc.Close()
            
        print('Conversion completed.')
    finally:
        word.Quit()

if __name__ == "__main__":
    # Replace '.' with the desired directory if needed
    convert_doc_to_docx('.')

