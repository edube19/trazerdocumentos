from pydoc import doc
import sys
from xmlrpc import client
from docx import Document
from docx.shared import Cm
import docx
from recursos import*

if __name__ == "__main__":
    while True:
        
        cliente='NORMA ANGULO PANDURO'
        clienteInmueble='NORMA'
        empleador='NORMA ANGULO PANDURO'
        
        """#TITULO
        try:
            print('TITULO')
            ruta1='C:/Users/DELL/Desktop/pruebaword/exp312294/titulo/SOLCREDITOHIPOTECARIO_312294_8.docx'
            doc1 = docx.Document(ruta1)
            buscar_palabra(cliente,doc1)
            
        except Exception as e:
            print("Error en TITULO: " + str(e))

        try:    
        #MINUTA , se encontraron en la pag 13
            print('MINUTA')
            extraerTexto_imagenV2('SOLCREDITOHIPOTECARIO_312294_3_page-',22,cliente)
            extraerTexto_imagenV2('TESTIMONIO Y GRAVAMEN_312294_1_page-',53,cliente)
            
        except Exception as e:
            print("Error en MINUTA: " + str(e))

        try:
        #PRESTAMO 
            print('PRESTAMO')
            ruta2='C:/Users/DELL/Desktop/pruebaword/exp312294/firma/SOLCREDITOHIPOTECARIO_312294_12.docx'
            doc2 = docx.Document(ruta2)
            buscar_palabra(cliente,doc2)
            #con firmas completas
            extraerTexto_imagenV2('SOLCREDITOHIPOTECARIO_312294_14_page-',26,cliente,'firma','exp312294')#pag 22 y pag 24

            #extraerTexto_imagenV2('SOLCREDITOHIPOTECARIO_312294_15_page-',3,cliente,'firma','exp312294')#pag1

        except Exception as e:
            print("Error en PRESTAMO: " + str(e))

        try:
        #CLAUSULA
            print('CLAUSULA')
            extraerTexto_imagenV2('SOLCREDITOHIPOTECARIO_312294_4_page-',3,cliente,'clausula','exp312294')#pag1
            
        except Exception as e:
            print("Error en CLAUSULA: " + str(e))"""
        
        try:
        #PAGO
            print('PAGO')
            extraerTexto_imagenV2('IMPDECEM_312294_3_page-',1,clienteInmueble,'pago','exp312294')
            
        except Exception as e:
            print("Error en PAGO: " + str(e))

        sys.exit()