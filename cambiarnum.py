from email.utils import decode_rfc2231
from pydoc import doc
import sys
from xmlrpc import client
from docx import Document
from docx.shared import Cm
import docx
from recursos import*
from recurso_prueba import*


if __name__ == "__main__":
    while True:

        cliente = 'JULIO CESAR LABAN RAMIREZ'
        ruta_guardar='C:/Users/DELL/Desktop/pruebaword/porcentajedocprueba.docx'
        #principal('exp3974679','prestamofirma','SOLCREDITOHIPOTECARIO_3974679_9',cliente)

        ruta1='C:/Users/DELL/Desktop/pruebaword/exp3892562/prestamo/SOLCREDITOHIPOTECARIO_3892562_15.docx'
        principalv3(ruta1,cliente,ruta_guardar)

        """doc1=docx.Document(ruta1)
        #leerdoc(doc1)
        buscar_palabra(cliente,doc1)
        editar_linea(doc1,'SEÑOR NOTARIO:','‎',ruta_guardar)#U+200E es un caracter vacio
        porcentaje_decimalv2(doc1,ruta_guardar)
        cambiarcomas(doc1,ruta_guardar)
        camrbiarmillon(doc1,ruta_guardar)

        cambiarcomas(doc1,ruta_guardar)
        cambiar_porcentaje(doc1,ruta_guardar)
        buscar_palabra(cliente,doc1)
        camrbiarmillon(doc1,ruta_guardar)
        editar_linea(doc1,string,'',ruta_guardar)"""
        sys.exit()