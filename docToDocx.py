# Convert Microsoft Word 'doc' files to 'docx'
import os.path
import win32com.client
#import sys

def convertir(path):
    #path = 'C:/Users/julian.lastra/Documents/Testings/DocToDocx/documento.doc'#sys.argv[0]
    word = win32com.client.Dispatch("Word.application")

    #file_path = os.path.join(dir_path, file_name)
    docx_file = '{0}{1}'.format(path, 'x')
    if not os.path.isfile(docx_file): # Skip conversion where docx file already exists
        print('Converting: {0}'.format(path))
        try:
            wordDoc = word.Documents.Open(path, False, False, False)
            wordDoc.SaveAs2(docx_file, FileFormat = 16)
            wordDoc.Close()
            print("File converted succesfully") 
        except Exception: 
            print('Failed to Convert: {0}'.format(path))
    word.Quit()
