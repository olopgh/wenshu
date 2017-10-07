import docx
import os
from xml.etree.cElementTree import XML
from xml.etree.ElementTree import XML
import zipfile

base_dir = ur'C:\\Users\\yuyi4\\Desktop\\document\\fazhi\\'
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'

###################################################################################
'''
Owner: Yu Yi
function Module that extract text from MS XML Word document (.docx).
inspired from: http://etienned.github.io/posts/extract-text-from-word-docx-simply/
'''
###################################################################################
def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)

###################################################################################
'''
Owner: Yu Yi
function read all the file names under a certain folder and return in a list of file names
'''
###################################################################################
def file_name(file_dir): 
    ls_file = []  
    for root, dirs, files in os.walk(file_dir):  
         break
    return files

if __name__ == "__main__":
    ls_files = file_name(base_dir + 'sample_verdit')
    print ls_files
    file_path = base_dir + 'sample_verdit\\' + ls_files[4]

    print get_docx_text(file_path)
