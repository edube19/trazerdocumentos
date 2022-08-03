from email.mime import image
from importlib.resources import path
from PIL import Image 
from pytesseract import pytesseract 
cliente='RENZO ZEBALLOS DEZA'
path_to_tesseract = r'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe'
#path_to_tesseract = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
#image_path = r"sampletext.png"
image_path = r'C:/Users/DELL/Desktop/pruebaword/minuta/Minuta de Compra Venta_3956235_2 (2)_pages-to-jpg-0001.jpg'
img = Image.open(image_path) 
  
pytesseract.tesseract_cmd = path_to_tesseract 
  
text = pytesseract.image_to_string(img) 

r=text.find(cliente) 
if r!=-1:
    print('Se encontro')
else:
    print('No encontro')
#print(text[:-1])