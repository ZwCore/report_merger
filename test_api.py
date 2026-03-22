
import requests
import os
from docx import Document

def create_docx(filename, text):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(filename)

def test_api():
    # Setup dummy files
    create_docx('test_template.docx', 'Content before {{Test}} content after.')
    create_docx('Test.docx', 'This is the merged content.')

    url = 'http://127.0.0.1:5000/api/process'
    
    files = [
        ('template', ('test_template.docx', open('test_template.docx', 'rb'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')),
        ('reports', ('Test.docx', open('Test.docx', 'rb'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))
    ]
    
    try:
        response = requests.post(url, files=files)
        print(f"Status Code: {response.status_code}")
        print(f"Response: {response.text}")
        
        if response.status_code == 200:
            print("Test Passed!")
        else:
            print("Test Failed!")
            
    except Exception as e:
        print(f"Exception: {e}")
    finally:
        # Cleanup
        try:
            # os.remove('test_template.docx')
            # os.remove('Test.docx')
            pass
        except:
            pass

if __name__ == '__main__':
    test_api()
