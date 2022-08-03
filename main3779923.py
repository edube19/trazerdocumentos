from pydoc import doc
import sys
from xmlrpc import client
from docx import Document
from docx.shared import Cm
import docx
from recursos import*
import sys, fitz



if __name__ == "__main__":
    while True:
        
        cliente='GIANFRANCO DELSO PAREDES'
        clienteInmueble='GIANCARLO'
        empleador='VULCO PERU'
        ruta1='C:/Users/DELL/Desktop/pruebaword/exp3779923/contrato/SOLCREDITOHIPOTECARIO_3779923_22.docx'

        #ruta1 = sys.argv[1]  # get document filename
        try:
            doc = fitz.open(ruta1)  # open document
            out = open(ruta1, "wb")  # open text output
            for page in doc:  # iterate the document pages
                text = page.get_text().encode("utf8")  # get plain text (is in UTF-8)
                out.write(text)  # write text of page
                out.write(bytes((12,)))  # write page delimiter (form feed 0x0C)
            out.close()
        except Exception as e:
            print(e)
        #CONTRATO
        #doc1=Document(ruta1)
        #buscar_palabra(cliente,doc1)
        #INMUEBLE
        #ruta2='C:/Users/DELL/Desktop/pruebaword/exp3779923/inmueble/Solicituddesegurodeincendiodeusoviviendaousomixto_3779923_1.docx'
        #ruta2='C:/Users/DELL/Desktop/pruebaword/exp3779923/inmueble/SOLCREDITOHIPOTECARIO_3779923_21.docx'
        
        #doc2=Document(ruta2)
        #buscar_palabra(cliente,doc2)
        #MINUTA


        #TITULO


        sys.exit()  