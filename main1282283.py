from pydoc import doc
import sys
from xmlrpc import client
from docx import Document
from docx.shared import Cm
import docx
from recursos import*

if __name__ == "__main__":
    while True:
        cliente='CHRISTIAN JOHN HANN VILLA'
        clientev2='Hann Villa Christian John'
        clienteInmueble='CHRISTIAN'
        empleador='REPSOL COMERCIAL'

        try:
        #CLAUSULA
        #Requiere buscar con clientev2
            ruta1='C:/Users/DELL/Desktop/pruebaword/exp1282283/clausula/SOLCREDITOHIPOTECARIO_1282283_18.docx'
            doc1=docx.Document(ruta1)
            buscar_palabra(clientev2,doc1)
        except Exception as e:
            print("Error en CLAUSULA :"+str(e))

        try:   
        #EVIDENCIA
        #extraerTexto_imagenV2('SOLCREDITOHIPOTECARIO_1282283_14_page-',1,clienteInmueble,'evidencia','exp1282283')
            pass
        except Exception as e:
            print('Error en EVIDENCIA :'+str(e))

        sys.exit()    