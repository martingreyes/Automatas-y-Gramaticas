import re , os, platform
from datetime import datetime
from tabulate import tabulate
import xlwt
from xlwt import Workbook


def Valid_MAC_AP(string):
    reg = re.compile(r"^([0-9A-Fa-f]{2}[-]){5}([0-9A-Fa-f]{2})(:UM)$")

    if reg.match(string):
        return True
    else:
        return False


def Valid_MAC_Client(string):

    reg = re.compile(r"^([0-9A-Fa-f]{2}[-]){5}([0-9A-Fa-f]{2})$")
    
    if reg.match(string):
        return True
    else:
        return False


def Valid_Date(string):
    # * 28/08/2019 10:06 ejemplo
    # * 28/08/2030 10:06 ejemplo
        
    reg = re.compile(r"^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[012])/((19|20)\d\d) (0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$")

    if reg.match(string):
        return True
    else:
        return False


def Valid_User(string): 
    # ! Acepta: letra (minimo 3), numero, mayuscula, punto, minuscula, guion medio, barra (/)
    # * Hay un usario -> hest/Dip-Senz-02
    reg = re.compile(r"^(?=(.*[a-zA-Z]){3,})[a-zA-Z0-9/.-]*$")
    if reg.match(string):
        return True
    else:
        return False


def pasar_a_lista(archivo):
    
    # ! Leemos cada linea
    lineas_borrador = archivo.readlines()

    # ! Borramos la primera ya que es el nombre de las columnas.
    lineas_borrador.remove(lineas_borrador[0])

    # ! Por cada renglon en el txt, separamos cada elemento que tenga un ";" por una lista: "Hola;como;estas" --> ["hola", "como", "estas"]
    lineas = []
    for linea in lineas_borrador:
        linea = linea.split(";")

        # ! Validamos los datos para saber si está corrupto o no.
        if Valid_User(linea[1]) and Valid_Date(linea[2]) and Valid_Date(linea[3]) and linea[4].isdigit() and Valid_MAC_AP(linea[7]) and Valid_MAC_Client(linea[8]):
            
            # ! Se elimina el salto de linea al final de cada renglon y se pone en minuscula los nombres de usuarios.
            linea[-1] = linea[-1].strip("\n")
            linea[1] = linea[1].lower()

            lineas.append(linea)

    lineas_borrador.clear()

    return lineas


def users_list(lineas):

    # ! Se obtiene cada usario de la matriz y se guarda en una lista.
    usuarios = []
    for linea in lineas:
        if (linea[1] not in usuarios) and (Valid_User(linea[1])):
            usuarios.append(linea[1])
    print('\nLista de suarios:\n')
    
    # ! Para mostrar mas bonito por pantalla.
    i = 1
    for usr in usuarios:
        print(str(i)+')',usr)
        i += 1
    
    return usuarios


def convertir_segundos(segundos):
    # ! Esta funcion toma el valor de Session Time (tiempo en segundos de la conexion) y lo convierte en hh:mm:ss
    horas = int(segundos / 60 / 60)
    segundos -= horas*60*60
    minutos = int(segundos/60)
    segundos -= minutos*60
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"


def trasformar_fecha(fecha):
    # ! Esta funcion convierte la fecha ingresadama en str a formato datetime.
    fecha_dt = datetime.strptime(fecha, '%d/%m/%Y %H:%M')
    return fecha_dt


def show_verboso(lineas, usuarios, usuario, primer_fecha, segunda_fecha):
    # ! Muestra la linea cronologica de un usuario en forma de texto.
    contador = 1 
    for linea in lineas:

        if linea[1] == usuarios[usuario]:

            # ! Comprueba que la fecha de inicio esté en el rango que ingresa el usuario. 
            if (trasformar_fecha(linea[2]) >=  trasformar_fecha(primer_fecha)) and (trasformar_fecha(linea[2]) <= trasformar_fecha(segunda_fecha)):
                
                # ! Solo para el primer caso, ya que no se tiene un dispositivo anterior.
                if contador == 1:
                    string = "\n{}º El Usuario '{}' se conecto al AP {}, durante {}, con el dispositivo {}, desde {} hasta {}.".format(contador, linea[1], linea[7], convertir_segundos(int(linea[4])), linea[8], linea[2], linea[3])
                
                # ! El usuariocambia de dispositivo y de ubicacion.
                elif linea[8] != dispositivo and linea[7] != ubicacion:
                     string = "{}º El Usuario '{}' se conecto al AP {} ( ≠ ubicacion), durante {}, con el dispositivo {} ( ≠ dispositivo), desde {} hasta {}.".format(contador, linea[1], linea[7], convertir_segundos(int(linea[4])), linea[8], linea[2], linea[3])

                # ! El usuariocambia de dispositivo.
                elif linea[8] != dispositivo:  
                    string = "{}º El Usuario '{}' se conecto al AP {}, durante {}, con el dispositivo {} ( ≠ dispositivo), desde {} hasta {}.".format(contador, linea[1], linea[7], convertir_segundos(int(linea[4])), linea[8], linea[2], linea[3])
                
                # ! El usuariocambia de ubicacion.
                elif linea[7] != ubicacion: 
                    string = "{}º El Usuario '{}' se conecto al AP {} ( ≠ ubicacion), durante {}, con el dispositivo {}, desde {} hasta {}.".format(contador, linea[1], linea[7], convertir_segundos(int(linea[4])), linea[8], linea[2], linea[3])

                # ! El usuario se reconecta pero con el mismo dispositivo y ubicacion.
                else:
                    string = "{}º El Usuario '{}' se conecto al AP {}, durante {}, con el dispositivo {}, desde {} hasta {}.".format(contador, linea[1], linea[7], convertir_segundos(int(linea[4])), linea[8], linea[2], linea[3])
                
                dispositivo = linea[8]
                ubicacion = linea[7]
                contador = contador + 1
                print(string)


def show_table(lineas, usuarios, usuario, primer_fecha, segunda_fecha):
    # ! Muestra la linea cronologica de un usuario en forma de una tabla.
    contador = 1
    tabla = []

    for linea in lineas:
        if linea[1] == usuarios[usuario]:

            # ! Comprueba que la fecha de inicio esté en el rango que ingresa el usuario. 
            if (trasformar_fecha(linea[2]) >=  trasformar_fecha(primer_fecha)) and (trasformar_fecha(linea[2]) <= trasformar_fecha(segunda_fecha)):
                fila =  ["{}º".format(contador), linea[1] , linea[2], linea[3], convertir_segundos(int(linea[4])), linea[7], linea[8]]
                tabla.append(fila)
                contador = contador + 1

    return(tabulate(tabla,headers=["", "Usuario", "Inicio de Conexion", "Fin de Conexion", "Duracion" , "MAC AP", "MAC Cliente" ] ))


def to_excel(resultado, pwd):
    # ! Resive una lista formada por listas (lineas) para cada usuario, que estan formadas por listas (linea) de datos.   Resultado -> Lineas (por usuario) -> Linea(lista de los valores por columna) 
    wb = Workbook()
    sheet = wb.add_sheet('Resultado')

    # ! Cada numero representa un color diferente
    colores = [47,45,26,41,42,17,46,60,54,70,55]    
    
    headers = ["", "Usuario", "Inicio de Conexion", "Fin de Conexion", "Duracion" , "MAC AP", "MAC Cliente" ]

    # ! Agrega los nombres de las columnas en la primera fila del excel.
    for elemento in range(len(headers)):
        sheet.write(0, elemento, headers[elemento]) # ! row, col, data

    fila = 1
    columna = 0
    for usuario in resultado:
        columna = 0

        # ! Se agrega un color a todos los datos de un mismo usuario.
        st = xlwt.easyxf('pattern: pattern solid;')
        st.pattern.pattern_fore_colour = colores[0]
        # ! Se saca el color de la lista y se pone al final.
        colores.remove(colores[0])
        colores.append(st.pattern.pattern_fore_colour)

        # ! Escribe cada elemento de linea en una casilla del excel
        for lista in usuario:
            columna = 0

            for elemento in lista:
                sheet.write(fila, columna, elemento, style = st )
                columna = columna + 1
            
            fila = fila + 1

    if platform.system() == "Windows":
        wb.save('{}\Trabajo Final Grupo 1.xls'.format(pwd))
    else:
        wb.save('{}/Trabajo Final Grupo 1.xls'.format(pwd))


    
def main():
    # ! Se leer el archivo txt    
    pwd = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    archivo = open(os.path.join(pwd, 'acts-user1.txt'),'r')
    
    # ! Convierte la variable en una lista de listas.
    lineas = pasar_a_lista(archivo)
    
    # ! Guardamos en una lista a todos los usuarios. 
    usuarios = users_list(lineas)

    # ! Pedimos los datos al usuario.
    usuario = (input("\nIngrese el indice del usuario: "))
    while (not (usuario.isdigit())) or (int(usuario) < 1) or (int(usuario) > len(usuarios)):
                usuario = (input("Debe ingresar un numero entero. Ingrese el indice del usuario: "))
    usuario = int(usuario) - 1


    primer_fecha = input("\nIngrese primer fecha (dd/mm/yyyy hh:mm): ")
    while not Valid_Date(primer_fecha):
         primer_fecha = input("Debe ingresar una fecha con el formato dd/mm/yyyy hh:mm. Ingrese primer fecha: ")


    segunda_fecha = input("\nIngrese segunda fecha (dd/mm/yyyy hh:mm): ")
    while not Valid_Date(segunda_fecha):
        segunda_fecha = input("Debe ingresar una fecha con el formato dd/mm/yyyy hh:mm. Ingrese segunda fecha: ")


    verboso = input("\n¿Modo verboso? (y/n): ")
    while verboso.lower() != "y" and verboso.lower() != "n":
        verboso = input("Opcion invalida. ¿Modo verboso? (y/n): ")
    
    
    if verboso.lower() == "y":
        show_verboso(lineas, usuarios, usuario, primer_fecha, segunda_fecha)
    else:
        print("\n",show_table(lineas, usuarios, usuario, primer_fecha, segunda_fecha))


    # ! Devolvemos la tabla con los valores (como en show_table pero sin el encabezado) para luego guardarlo en el excel.
    contador = 1
    tabla = []
    for linea in lineas:
        if linea[1] == usuarios[usuario]:
            # ! Comprueba que la fecha de inicio esté en el rango que ingresa el usuario. 
            if (trasformar_fecha(linea[2]) >=  trasformar_fecha(primer_fecha)) and (trasformar_fecha(linea[2]) <= trasformar_fecha(segunda_fecha)):
                fila =  ["{}º".format(contador), linea[1] , linea[2], linea[3], convertir_segundos(int(linea[4])), linea[7], linea[8]]
                tabla.append(fila)
                contador = contador + 1
    return tabla
    

if __name__ == "__main__":

    # ! Resultado es una lista formada por listas (lineas) para cada usuario, que estan formadas por listas (linea) de datos.   Resultado -> Lineas (por usuario) -> Linea(lista de los valores por columna) 
    resultado = []
    while True:
            resultado.append(main())  
            continuar = input("\n¿Quiere continuar? (y/n): ")
            while continuar.lower() != "y" and continuar.lower() != "n":
                continuar = input("Opcion invalida. ¿Quiere continuar? (y/n): ")
            if continuar == "n".lower():
                break
    
    # ! Se exporta a excel la variable resultado.
    pwd = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    to_excel(resultado,pwd)
    print("\nYa esta disponible el archivo 'Trabajo Final Grupo 1.xls' (ubicado en: {}) con los resultados, donde cada color representa una iteracion.\n".format(pwd))
    