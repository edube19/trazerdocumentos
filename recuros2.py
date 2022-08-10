MONEDA_SINGULAR = 'NUEVO SOL'
MONEDA_PLURAL = 'NUEVOS SOLES'

CENTIMOS_SINGULAR = 'CENTIMO'
CENTIMOS_PLURAL = 'CENTIMOS'

MAX_NUMERO = 999999999999

UNIDADES = (
    'CERO',
    'UNO',
    'DOS',
    'TRES',
    'CUATRO',
    'CINCO',
    'SEIS',
    'SIETE',
    'OCHO',
    'NUEVE'
)

DECENAS = (
    'DIEZ',
    'ONCE',
    'DOCE',
    'TRECE',
    'CATORCE',
    'QUINCE',
    'DIECISEIS',
    'DIECISIETE',
    'DIECIOCHO',
    'DIECINUEVE'
)

DIEZ_DIEZ = (
    'CERO',
    'DIEZ',
    'VEINTE',
    'TREINTA',
    'CUARENTA',
    'CINCUENTA',
    'SESENTA',
    'SETENTA',
    'OCHENTA',
    'NOVENTA'
)

CIENTOS = (
    '_',
    'CIENTO',
    'DOSCIENTOS',
    'TRESCIENTOS',
    'CUATROSCIENTOS',
    'QUINIENTOS',
    'SEISCIENTOS',
    'SETECIENTOS',
    'OCHOCIENTOS',
    'NOVECIENTOS'
)


def numero_a_letras(numero):
    numero_entero = int(numero)
    if numero_entero > MAX_NUMERO:
        raise OverflowError('Número demasiado alto')
    if numero_entero < 0:
        negativo_letras = numero_a_letras(abs(numero))
        return f"MENOS {negativo_letras}"
    letras_decimal = ''
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    if parte_decimal > 9:
        letras_decimal = f"PUNTO {numero_a_letras(parte_decimal)}"
    elif parte_decimal > 0:
        letras_decimal = f"PUNTO CERO {numero_a_letras(parte_decimal)}"
    if numero_entero <= 99:
        resultado = leer_decenas(numero_entero)
    elif numero_entero <= 999:
        resultado = leer_centenas(numero_entero)
    elif numero_entero <= 999999:
        resultado = leer_miles(numero_entero)
    elif numero_entero <= 999999999:
        resultado = leer_millones(numero_entero)
    else:
        resultado = leer_millardos(numero_entero)
    resultado = resultado.replace('UNO MIL', 'UN MIL')
    resultado = resultado.strip()
    resultado = resultado.replace(' _ ', ' ')
    resultado = resultado.replace('  ', ' ')
    if parte_decimal > 0:
        resultado = f"{resultado} {letras_decimal}"
    return resultado


def numero_a_moneda(numero):
    numero_entero = int(numero)
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    centimos = CENTIMOS_SINGULAR if parte_decimal == 1 else CENTIMOS_PLURAL
    moneda = MONEDA_SINGULAR if numero_entero == 1 else MONEDA_PLURAL
    letras = numero_a_letras(numero_entero)
    letras = letras.replace('UNO', 'UN')
    aux_decimal = numero_a_letras(parte_decimal).replace('UNO', 'UN')
    letras_decimal = f"con {aux_decimal} {centimos}"
    letras = f"{letras} {letras_decimal} {moneda}"
    return letras


def leer_decenas(numero):
    if numero < 10:
        return UNIDADES[numero]
    decena, unidad = divmod(numero, 10)
    if numero <= 19:
        resultado = DECENAS[unidad]
    elif numero <= 29:
        resultado = f"VEINTI{UNIDADES[unidad]}"
    else:
        resultado = DIEZ_DIEZ[decena]
        if unidad > 0:
            resultado = f"{resultado} y {UNIDADES[unidad]}"
    return resultado


def leer_centenas(numero):
    centena, decena = divmod(numero, 100)
    if numero == 0:
        resultado = 'CIEN'
    else:
        resultado = CIENTOS[centena]
        if decena > 0:
            decena_letras = leer_decenas(decena)
            resultado = f"{resultado} {decena_letras}"
    return resultado


def leer_miles(numero):
    millar, centena = divmod(numero, 1000)
    resultado = ''
    if millar == 1:
        resultado = ''
    if (millar >= 2) and (millar <= 9):
        resultado = UNIDADES[millar]
    elif (millar >= 10) and (millar <= 99):
        resultado = leer_decenas(millar)
    elif (millar >= 100) and (millar <= 999):
        resultado = leer_centenas(millar)
    resultado = f"{resultado} MIL"
    if centena > 0:
        centena_letras = leer_centenas(centena)
        resultado = f"{resultado} {centena_letras}"
    return resultado.strip()


def leer_millones(numero):
    millon, millar = divmod(numero, 1000000)
    resultado = ''
    if millon == 1:
        resultado = ' UN MILLON '
    if (millon >= 2) and (millon <= 9):
        resultado = UNIDADES[millon]
    elif (millon >= 10) and (millon <= 99):
        resultado = leer_decenas(millon)
    elif (millon >= 100) and (millon <= 999):
        resultado = leer_centenas(millon)
    if millon > 1:
        resultado = f"{resultado} MILLONES"
    if (millar > 0) and (millar <= 999):
        centena_letras = leer_centenas(millar)
        resultado = f"{resultado} {centena_letras}"
    elif (millar >= 1000) and (millar <= 999999):
        miles_letras = leer_miles(millar)
        resultado = f"{resultado} {miles_letras}"
    return resultado


def leer_millardos(numero):
    millardo, millon = divmod(numero, 1000000)
    miles_letras = leer_miles(millardo)
    millones_letras = leer_millones(millon)
    return f"{miles_letras} MILLONES {millones_letras}"


def numero_a_moneda_sunat(numero):
    numero_entero = int(numero)
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    moneda = MONEDA_SINGULAR if numero_entero == 1 else MONEDA_PLURAL

    letras = numero_a_letras(numero_entero)
    letras = letras.replace('UNO', 'UN')
    letras = f"{letras} Y {parte_decimal}/100 {moneda}"
    return letras

  
def __main__():
    cadena1 = numero_a_moneda_sunat(37.45)
    cadena2 = numero_a_moneda_sunat(1012121)
    cadena3 = numero_a_moneda_sunat(150)
    cadena4 = numero_a_moneda_sunat(1451.45)
    cadena5 = numero_a_moneda_sunat(1004501245)
    cadena6 = numero_a_letras(-25.00741)
    print(f'{cadena1}')
    print(f'{cadena2}')
    print(f'{cadena3}')
    print(f'{cadena4}')
    print(f'{cadena5}')
    print(f'{cadena6}')
    print(f'{cadena6.lower()}')
    
#esto es la version casi mejorada que esta en recursos
def porcentaje_decimal(doc):
    l=cantidad_lineas(doc)
    c='%'
    a=0
    lst = []
    for i in range(l):
        cond=True
        contador=1
        numero=''
        cadena=doc.paragraphs[i].text
        for pos,char in enumerate(cadena):
            if(char == c):
                lst.append(pos)
                a=a+1
        #r=cadena.find('%')#posicion en que se encuentra lo que se busca, en caso de no encontrarse devuelve -1
        if a!=0:
            for r in lst:
                while(cond):
                    valor=cadena[r-contador]#obtenner el valor
                    cond=valor.isdigit()#comprobar si es digito o no, si es falso acaba la ejecucion
                    if (cond):
                        numero=valor+numero#guardando el numero
                        contador=contador+1#avanza para el siguiente valor
                    elif(valor=='.'):# para casos → ab.cd%
                        contador=contador+1
                        numero='.'+numero
                        cond=True
                    elif (valor==' ' and numero==''):# para casos → ab.cd %
                        contador=contador+1
                        cond=True
                try:
                    r=numero.find('.')
                    if r!=a:
                        numero_separado=numero.split('.')
                        numeroconv1=int(numero_separado[0])#parte entera
                        numeroconv2=int(numero_separado[1])#parte decimal
                        porcentaje1=numero_to_letras(numeroconv1)
                        porcentaje2=numero_to_letras(numeroconv2)
                        porcentajeminus1=porcentaje1.lower()
                        porcentajeminus2=porcentaje2.lower()
                        nuevostring='%('+porcentajeminus1+' coma '+porcentajeminus2+' por ciento)'
                        nueva_cadena=cadena.replace('%',nuevostring)
                        #print(nuevostring)
                        doc.paragraphs[i].text=nueva_cadena
                    else:
                        numeroconv1=int(numero)
                        porcentaje1=numero_to_letras(numeroconv1)
                        porcentajeminus1=porcentaje1.lower()
                        nuevostring='%('+porcentajeminus1+' por ciento)'
                        nueva_cadena=cadena.replace('%',nuevostring)
                        #print(nuevostring)
                        doc.paragraphs[i].text=nueva_cadena
                except Exception as e:
                    print(e)
        lst.clear()
    doc.save('C:/Users/DELL/Desktop/pruebaword/pruebadecimales.docx')    
if __name__ == "__main__" :
    __main__()