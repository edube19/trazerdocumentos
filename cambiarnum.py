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
        cliente = 'LOURDES HIDALGO CASTILLO'
        ruta_guardar='C:/Users/DELL/Desktop/pruebaword/porcentajedocprueba.docx'
        ruta1='C:/Users/DELL/Desktop/pruebaword/exp3892562/prestamo/SOLCREDITOHIPOTECARIO_3892562_15.docx'
        #ruta2='C:/Users/DELL/Desktop/pruebaword/pruebaformato.docx'
        principalv3(ruta1,cliente,ruta_guardar)
        #modificartamano(ruta2,'20100047218',ruta2)
        sys.exit()