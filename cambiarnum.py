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

        cliente='RENZO ZEBALLOS DEZA'
        clienteInmueble='RENZO'
        empleador='CORPORACION SECURITY TECH'
        
        #principal('exp3974679','prestamofirma','SOLCREDITOHIPOTECARIO_3974679_9',cliente)

        """print('formato 1')
        ruta1='C:/Users/DELL/Desktop/pruebaword/exp3974679/prestamofirma/SOLCREDITOHIPOTECARIO_3974679_9.docx'
        doc1=docx.Document(ruta1)
        cambiarcomas(doc1)
        cambiar_porcentaje(doc1)"""

        ruta1='C:/Users/DELL/Desktop/pruebaword/exp3956235/prestamo/SOLCREDITOHIPOTECARIO_3956235_13.docx'
        doc1=docx.Document(ruta1)
        cambiarcomas(doc1)
        cambiar_porcentaje(doc1)
        buscar_palabra(cliente,doc1)
        camrbiarmillon(doc1)

        """print('formato 2')
        ruta2='C:/Users/DELL/Desktop/pruebaword/exp3956235/titulo/SOLCREDITOHIPOTECARIO_3956235_14.docx'
        doc2=docx.Document(ruta2)
        cambiarcomas(doc2)
        camrbiarmillon(doc2)"""

        """ruta3='C:/Users/DELL/Desktop/pruebaword/exp3892562/titulo/SOLCREDITOHIPOTECARIO_3892562_11.docx'
        doc3=docx.Document(ruta3)
        cambiarcomas(doc3)"""

        #pdf_a_word('C:/Users/DELL/Desktop/pruebaword/exp3974679/minuta/Minuta de Compra Venta_3993816_3_TASACIONANTICIPADA_17062022.pdf')

        #doc='C:/Users/DELL/Desktop/pruebaword/exp3974679/prestamofirma/SOLCREDITOHIPOTECARIO_3974679_9.docx'
        #principal('exp3974679','prestamofirma','SOLCREDITOHIPOTECARIO_3974679_9',cliente)
        #principal2(doc,cliente)

        sys.exit()