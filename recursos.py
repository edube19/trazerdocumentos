from array import array
from cgi import print_arguments
from re import T
from docx import Document
from docx.shared import Cm
import aspose.words as aw
from fitz import *
from email.mime import image
from importlib.resources import path
from PIL import Image 
from pytesseract import pytesseract 
from pdf2image import convert_from_path
from recurso_prueba import*

#from PyPDF2 import PDFFileReader

def evitar_sobreescritura(cadena,string):
    cond=True
    a=-1
    r=cadena.find(string)#string = 'por ciento'
    if r!=a:
        cond=False
    return cond

def principal(expediente,documento,nombre,string):
    doc=f'C:/Users/DELL/Desktop/pruebaword/{expediente}/{documento}/{nombre}.docx'
    #docu=str(doc)
    print(doc)
    print(type(doc))
    try:
        buscar_palabra(string,doc)
        cambiarcomas(doc)
        cambiar_porcentaje(doc)
        camrbiarmillon(doc)
    except Exception as e:
        print(e)

    doc.save(f'C:/Users/DELL/Desktop/pruebaword/{expediente}/{documento}/{nombre}modificado.docx')

def principal2(doc,string):
    buscar_palabra(string,doc)
    cambiarcomas(doc)
    cambiar_porcentaje(doc)
    camrbiarmillon(doc)
    doc.save('C:/Users/DELL/Desktop/pruebaword/modificado.docx')

def cantidad_lineas(doc):
    l=len(doc.paragraphs)
    return l

def cambiar_porcentaje(doc):
    l=cantidad_lineas(doc)
    a=-1
    for i in range(l):
        cadena=doc.paragraphs[i].text
        r=cadena.find('%')#posicion en que se encuentra la coma
        if r!=a:
            #print('Se encontro en la linea '+str(i+1))
            #print(doc.paragraphs[i].text)
            num1=cadena[r-1]
            cond1=num1.isdigit()
            numero=num1
            if (cond1):
                num2=cadena[r-2]
                cond2=num2.isdigit()
                if(cond2):
                    numero=num2+num1
                    num3=cadena[r-3]
                    cond3=num3.isdigit()
                    if (cond3):
                        numero=num3+num2+num1
            #print(numero)     
            try:      
                numeroconv=int(numero)
                #intnum=int(numero)
                porcentaje=numero_to_letras(numeroconv)
                porcentajeminus=porcentaje.lower()
                nuevostring='%('+porcentajeminus+' por ciento)'
                nueva_cadena=cadena.replace('%',nuevostring)
                doc.paragraphs[i].text=nueva_cadena
            except Exception as e:
                print(e)
    porcentaje_tablas(doc)
    doc.save('C:/Users/DELL/Desktop/pruebaword/porcentajedocprueba.docx')

def buscar_palabra(string,doc):
    #str_doc=str(doc)
    l=cantidad_lineas(doc)
    a=-1
    c=0
    for i in range(l):
        cadena=doc.paragraphs[i].text
        r=cadena.find(string)
        if r!=a:
            #print('Se encontro en la linea '+str(i+1))
            #print(doc.paragraphs[i].text)
            c=c+1
    if (c==0):
        buscar_en_tabla(string,doc)
    else:
        print('Se encontro '+str(c)+' veces')
    return c

def cambiarcomas(doc):
    l=cantidad_lineas(doc)
    a=-1
    caracteres=['S/','US$']
    for i in range(l):
        cadena=doc.paragraphs[i].text
        for caracter in caracteres:
            r=cadena.find(caracter)#posicion en que se encuentra la coma
            if r!=a:
                #print('Se encontro en la linea '+str(i+1))
                #print(doc.paragraphs[i].text)
                nueva_cadena=cadena.replace(',','.')
                #print(nueva_cadena)
                doc.paragraphs[i].text=nueva_cadena
    comas_en_tabla(doc)
    doc.save('C:/Users/DELL/Desktop/pruebaword/porcentajedocprueba.docx')

def camrbiarmillon(doc):
    l=cantidad_lineas(doc)
    a=-1
    caracteres=['S/','US$']
    for i in range(l):
        cadena=doc.paragraphs[i].text
        for caracter in caracteres:
            r=cadena.find(caracter)#posicion en que se encuentra la coma
            if r!=a:
                #print('Se encontro en la linea '+str(i+1))
                #print(doc.paragraphs[i].text)
                nueva_cadena=cadena.replace('´','.')
                #print(nueva_cadena)
                doc.paragraphs[i].text=nueva_cadena
    comas_en_tabla(doc)
    doc.save('C:/Users/DELL/Desktop/pruebaword/porcentajedocprueba.docx')

def buscar_en_tabla(string,doc):
    leertabla=doc.tables
    resultado=0
    a=-1
    for x in range(len(leertabla)):
        tabla=leertabla[x]#obtener la primera tabla
        for i in range(0,len(tabla.rows)):#filas
            for j in range(0,len(tabla.columns)):#columnas
                #print(tabla.cell(i,j).text)
                cadena=tabla.cell(i,j).text
                #nc=cadena[20:30]
                #print(nc)
                r=cadena.find(string)
                if (r!=a):
                    print('Se encontro')
                    resultado=resultado+1
                    break
    if (resultado==0):
        print('No se encontro')

def comas_en_tabla(doc):
    leertabla=doc.tables
    a=-1
    caracteres=['S/','US$']
    for x in range(len(leertabla)):
        tabla=leertabla[x]#obtener la primera tabla
        for i in range(0,len(tabla.rows)):#filas
            for j in range(0,len(tabla.columns)):#columnas
                #print(tabla.cell(i,j).text)
                cadena=tabla.cell(i,j).text
                #nc=cadena[20:30]
                #print(nc)
                for caracter in caracteres:
                    r=cadena.find(caracter)
                    if (r!=a):
                        nueva_cadena=cadena.replace(',','.')
                        #print(nueva_cadena)
                        tabla.cell(i,j).text=nueva_cadena

def porcentaje_tablas(doc):
    leertabla=doc.tables
    a=-1
    for x in range(len(leertabla)):
        tabla=leertabla[x]#obtener la primera tabla
        for i in range(0,len(tabla.rows)):#filas
            for j in range(0,len(tabla.columns)):#columnas
                #print(tabla.cell(i,j).text)
                cadena=tabla.cell(i,j).text
                #nc=cadena[20:30]
                #print(nc)
                if (evitar_sobreescritura(cadena,'por ciento')):
                    r=cadena.find('%')
                    if (r!=a):
                #print('Se encontro en la linea '+str(i+1))
                #print(doc.paragraphs[i].text)
                        num1=cadena[r-1]
                        cond1=num1.isdigit()
                        numero=num1
                        if (cond1):
                            num2=cadena[r-2]
                            cond2=num2.isdigit()
                            if(cond2):
                                numero=num2+num1
                                num3=cadena[r-3]
                                cond3=num3.isdigit()
                                if (cond3):
                                    numero=num3+num2+num1
                        #print(numero)     
                        try:      
                            numeroconv=int(numero)
                            #intnum=int(numero)
                            porcentaje=numero_to_letras(numeroconv)
                            porcentajeminus=porcentaje.lower()
                            nuevostring='%('+porcentajeminus+' por ciento)'
                            nueva_cadena=cadena.replace('%',nuevostring)
                            tabla.cell(i,j).text=nueva_cadena
                        except Exception as e:
                            print(e)

def agregar_parrafo(string,doc):
    p=doc.add_paragraph(string)
    return p

def pdf_a_word(pdf):
    pdfn=pdf.replace('.pdf','')
    doc = aw.Document(pdf)
    doc.save(f'{pdfn}.docx')

def leerpdf(pdf,palabra):
    pdf_documento = pdf
    documento=fitz.open(pdf_documento)
    p=documento.page_count
    a=-1
    c=0
    for i in range(p):
        pagina=documento.load_page(i)
        text=pagina.get_text('text')
        r=text.find(palabra)
        if r!=a:
            #print('Se encontro en la linea '+str(i+1))
            #print('Se encontro')
            c=c+1
    if (c==0):
        print('No se encontro')
    else:
        print('Se encontro '+str(c)+ ' veces')

def extraerTexto_imagen(palabra,imagen):
    image_path =imagen
    path_to_tesseract = r'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe'
    img = Image.open(image_path) 
    pytesseract.tesseract_cmd = path_to_tesseract 
    text = pytesseract.image_to_string(img) 
    r=text.find(palabra) 
    if r!=-1:
        print('Se encontro')
    else:
        print('No encontro')

def pdf_imagen(pdf):
    # import module
    pages = convert_from_path(pdf)
    for i in range(len(pages)):
        pages[i].save('page'+ str(i) +'.jpg', 'JPEG')

def extraerTexto_imagenV2(nombre,cantidad,palabra,tipodoc,expediente):#modificar la ruta de la imagen
    for i in range(cantidad):
        imagen='C:/Users/DELL/Desktop/pruebaword/'+expediente+'/'+tipodoc+'/'+nombre+f'{i+1}.jpg'
        print('Pagina '+str(i+1)) 
        extraerTexto_imagen(palabra,imagen)

