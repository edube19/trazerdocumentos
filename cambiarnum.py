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
        cliente = 'GIANFRANCO DELSO PAREDES'#funciona incluso en minusculas
        palabra= 'SEÑOR NOTARIO'
        ruta_guardar='C:/Users/DELL/Desktop/pruebaword/porcentajedocprueba.docx'
        ruta1='C:/Users/DELL/Desktop/pruebaword/exp3779923/prestamo/SOLCREDITOHIPOTECARIO_3779923_22.docx'
        string_prueba='Tasa Fija: 8.37 % (ocho punto treinta y  siete por ciento) Tasa Efectiva Anual (TEA), la misma que no será modificada unilateralmente por el BANCO.'
        #ruta2='C:/Users/DELL/Desktop/pruebaword/pruebaformato.docx'
        #eliminar_linea(ruta1,palabra,ruta_guardar)
        principalv3(ruta1,cliente,ruta_guardar)
        #modificartamano(ruta2,'20100047218',ruta2)
        sys.exit()