import datetime
import time
import sys 
from tabulate import tabulate
import pandas as pd
import sqlite3 as sql
from sqlite3 import Error
import csv
from openpyxl import Workbook
try:
    with sql.connect("Evidencia_3.db") as conn:
        mi_cursor = conn.cursor()
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Pacientes (clave INTEGER PRIMARY KEY, apellidoPaterno TEXT NOT NULL, apellidoMaterno TEXT, nombres TEXT NOT NULL, fechaNacimiento timestamp NOT NULL, sexoPaciente TEXT NOT NULL);")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Citas (folio INTEGER PRIMARY KEY AUTOINCREMENT, clave_paciente INTEGER, fechaCita timestamp NOT NULL, turno TEXT NOT NULL, edad INTEGER, peso REAL, estatura REAL, horaLlegada timestamp , presionArterial TEXT, diagnostico TEXT, FOREIGN KEY (clave_paciente) REFERENCES tb_Pacientes(clave));")
except Error as e:
    print (e)
except:
    print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
def menu_Principal():
    print()
    print('---------------------- CONSULTORIO ----------------------')
    print('---------------------- MENU PRINCIPAL ----------------------')
    print('1. Registro de pacientes')
    print('2. Citas')
    print('3. Consultas y reportes')
    print('4. Salir del sistema')
clave_paciente = None 
while True:
    while True:
        menu_Principal()
        tupla_opciones = (1, 2, 3, 4)
        opcion_menuPrincipal_str = input('Selecciona la opcion a realizar: ')
        if opcion_menuPrincipal_str.isdigit():
            opcion_menuPrincipal = int(opcion_menuPrincipal_str)
            if opcion_menuPrincipal in tupla_opciones:
                print('Se registro correctamente')
                print()
                print()
                break
            else:
                print('Error. Seleccionar una opcion valida del menu. Intente de neuvo.')
                print()
        else:
            print('Error. Seleccionar el numero de la opcion. Intente de nuevo')
            print()
    if opcion_menuPrincipal == 1:
        print()    
        print('----------------- REGISTRO DE PACIENTE -----------------')
        while True:
            while True:
                try:
                    primer_apellido_paciente = input('Proporciona tu apellido Paterno: ').strip().upper()
                    if primer_apellido_paciente == '':
                        print('== IMPORTANTE ==')
                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                        if cancelar == '*':
                            primer_apellido_paciente = '*'
                            break
                        else:
                            raise ValueError('\nError generado previamente: Error. El dato no puede ser omitido.\n')
                    elif not primer_apellido_paciente.replace(' ', '').isalpha():
                        print('== IMPORTANTE ==')
                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                        if cancelar == '*':
                            primer_apellido_paciente = '*'
                            break
                        raise ValueError(f'El dato ingresado "{primer_apellido_paciente}" no contiene únicamente letras o espacios.')
                    else:
                        print('El dato ingresado se registró correctamente.')
                        print()
                        break
                except ValueError as error:
                    print(error)
            if primer_apellido_paciente == '*':
                print('Se cancelo la operacion.\nRegreseando al menu principal...')
                break
            while True:
                segundo_apellido_paciente = input('Proporciona tu apellido Materno (o presiona Enter para omitir): ').strip().upper()
                if segundo_apellido_paciente == '':
                    segundo_apellido_paciente = 'N/A'
                    print('El dato se ha omitido.')
                    break  
                elif not segundo_apellido_paciente.replace(' ', '').isalpha():
                    print('== IMPORTANTE ==')
                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                    if cancelar == '*':
                        segundo_apellido_paciente = '*'
                        break
                    print(f'El dato ingresado "{segundo_apellido_paciente}" no contiene únicamente letras.')
                else:
                    print('El dato ingresado se registró correctamente.')
                    break
            if segundo_apellido_paciente == '*':
                print('Se cancelo la operacion.\nRegreseando al menu principal...')
                break
            while True:
                print()
                nombres_paciente = input('Proporciona tus nombres: ').strip().upper()
                if nombres_paciente == '':
                    print('Error: El dato no puede ser omitido.')
                    print('\n== IMPORTANTE ==')
                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                    if cancelar == '*':
                        nombres_paciente = '*'
                        break
                elif not nombres_paciente.replace(' ', '').isalpha():
                    print(f'El dato ingresado "{nombres_paciente}" no contiene únicamente letras.')
                    print('\n== IMPORTANTE ==')
                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                    if cancelar == '*':
                        nombres_paciente = '*'
                        break
                else:
                    print('El dato ingresado se registró correctamente.')
                    print()
                    break
            if nombres_paciente == '*':
                print('Se cancelo la operacion.\nRegreseando al menu principal...')
                break
            while True:
                try:
                    fecha_nacimiento_paciente = input('Ingrese su fecha de nacimiento en el formato mm/dd/aaaa: ')
                    fecha_nacimiento_paciente = datetime.datetime.strptime(fecha_nacimiento_paciente, '%m/%d/%Y').date()
                    fecha_actual = datetime.date.today()
                    if fecha_nacimiento_paciente >= fecha_actual:
                        print('Error. La fecha de nacimiento no es valida.')
                        print('== IMPORTANTE ==')
                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                        if cancelar == '*':
                            fecha_nacimiento_paciente = '*'
                            break  # Aquí también
                    else:
                        print('El dato ingresado se registró correctamente.')
                        print()
                        break
                except ValueError:
                    print('Error: El formato de fecha ingresado no es válido. Por favor, ingrese la fecha en el formato mm/dd/aaaa.')
                    print('== IMPORTANTE ==')
                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                    if cancelar == '*':
                        fecha_nacimiento_paciente = '*'
                        break
            if fecha_nacimiento_paciente == '*':
                print('Se cancelo la operacion.\nRegreseando al menu principal...')
                break
            while True:
                sexo_paciente = str(input('Ingrese su sexo (H, M o N): ')).upper()
                if sexo_paciente.isalpha():
                    if sexo_paciente == 'H' or sexo_paciente == 'M' or sexo_paciente == 'N':
                        if sexo_paciente == 'H':
                            sexo_paciente_final = 'HOMBRE'
                        elif sexo_paciente == 'M':
                            sexo_paciente_final = 'MUJER'
                        elif sexo_paciente == 'N' or '':
                            sexo_paciente_final = 'NO CONTESTO'
                        print('El dato ingresado se registró correctamente.')
                        break
                    else:
                        print('ERROR. Ingrese una opcion valida (H, M o N)')
                        print('== IMPORTANTE ==')
                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                        if cancelar == '*':
                            sexo_paciente = '*'
                            break
                else:
                    print(f'ERROR. El valor ingresado {sexo_paciente} no es valido. Intente de nuevo')
                    print('== IMPORTANTE ==')
                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                    if cancelar == '*':
                        sexo_paciente = '*'
                        break
            if sexo_paciente == '*':
                break        
            try:
                with sql.connect("Evidencia_3.db") as conn:
                    mi_cursor = conn.cursor()
                    datos = primer_apellido_paciente, segundo_apellido_paciente, nombres_paciente, fecha_nacimiento_paciente, sexo_paciente_final
                    mi_cursor.execute('INSERT INTO Pacientes (apellidoPaterno, apellidoMaterno, nombres, fechaNacimiento, sexoPaciente) VALUES (?,?,?,?,?)', datos)
                    mi_cursor.execute('SELECT last_insert_rowid()') 
                    clave_paciente = mi_cursor.fetchone()[0] 
            except sql.Error as e:
                print(e)
            except:
                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
            finally:
                conn.close()
            print('El paciente se ha registrado correctamente.')
            print()
            print('Datelles del paciente:')
            print(f'Clave: {clave_paciente}')
            print(f'Apellido Paterno: {primer_apellido_paciente}')
            print(f'Apellido Materno: {segundo_apellido_paciente}')
            print(f'Nombres: {nombres_paciente}')
            print(f'Fecha de nacimiento: {fecha_nacimiento_paciente}')
            print(f'Sexo del paciente: {sexo_paciente_final}')
            print()
            break
    if opcion_menuPrincipal == 2:
        while True:
            print()
            print('----------------- MENU DE CITAS -----------------')
            print('1. Programacion de citas')
            print('2. Realizacion de citas programadas')
            print('3. Cancelacion de citas')
            print('4. Salir del submenu')
            while True:
                try:
                    tupla_citas = (1, 2, 3, 4)
                    opcion_submenuCitas = int(input('Selecciona la opcion del menu. Solo el numero: '))
                    if opcion_submenuCitas in tupla_citas:
                        print('Opcion valida')
                        break
                    else:
                        print('ERROR. Selecciona una opcion valida del menu. Intenta de nuevo')
                except ValueError:
                    print('ERROR. El valor ingresado no es numerico. Intenta de nuevo')
            if opcion_submenuCitas == 1:
                print('----------------- PROGRAMACION DE CITAS -----------------\n')
                claves_in_list = []
                print('------------------------- CLAVES DISPONIBLES ------------------------- ')
                try:
                    with sql.connect("Evidencia_3.db") as conn:
                        mi_cursor = conn.cursor()
                        mi_cursor.execute('SELECT clave, apellidoPaterno, apellidoMaterno, nombres FROM Pacientes')
                        claves_disponibles = mi_cursor.fetchall()
                        if claves_disponibles:
                            print(f'{"Clave":<10} {"Apellido Paterno":<20} {"Apellido Materno":<20} {"Nombres":<20}')
                            print('-' * 70)
                            for clave, apellidoPaterno, apellidoMaterno, nombres in claves_disponibles:
                                print(f'{clave:<10} {apellidoPaterno:<20} {apellidoMaterno:<20} {nombres:<20}\n')
                                claves_in_list.append(clave)
                        else:
                            print('\nError. No hay pacientes registrados. \nVolviendo al menú')
                except sql.Error as e:
                    print(e)
                except:
                    print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                finally:
                    conn.close()
                if claves_in_list:
                    while True:
                        while True:
                            try:
                                clave_seleccionada = int(input('Ingrese la clave a registrar: '))
                                print()
                                if clave_seleccionada in claves_in_list:
                                    break
                                else:
                                    print('ERROR. Selecciona una clave existente')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        clave_seleccionada = '*'
                                        break
                            except ValueError:
                                print('ERROR. La clave debe ser numerica. Intente de nuevo.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    clave_seleccionada = '*'
                                    break
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                if cancelar == '*':
                                    clave_seleccionada = '*'
                                    break
                        if clave_seleccionada == '*':
                            print('\nSe cancelo la operacion.\nRegreseando al menu...')
                            break
                        dias_habiles = datetime.timedelta(days=60)
                        fecha_actual_cita = datetime.date.today()
                        while True:
                            try:
                                fecha_cita = input('Fecha de la cita (En el siguiente formato (mm/dd/aaaa)): ')
                                fecha_cita = datetime.datetime.strptime(fecha_cita, '%m/%d/%Y').date()
                                if fecha_cita <= fecha_actual_cita:
                                    print('La fecha no es válida. Debe ser posterior al día actual.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        fecha_cita = '*'
                                        break
                                    continue
                                if fecha_cita > fecha_actual_cita + dias_habiles:
                                    print('La fecha no puede ser mayor a 60 días.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        fecha_cita = '*'
                                        break
                                    continue
                                if fecha_cita.weekday() == 6:
                                    recorrer_cita_opcion = input('La fecha de la cita cae en domingo. ¿Desea recorrer la cita un dia antes, día Sabado? (SI/NO): ').upper()
                                    if recorrer_cita_opcion == 'SI':
                                        fecha_cita = fecha_cita - datetime.timedelta(days=1)
                                    else:
                                        continue
                                print('La fecha de la cita es válida:', fecha_cita)
                                print()
                                break
                            except ValueError:
                                print('Error: El formato de fecha ingresado no es válido. Por favor, ingrese la fecha en el formato mm/dd/aaaa.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    fecha_cita = '*'
                                    break
                        if fecha_cita == '*':
                            print('\nSe cancelo la operacion.\nRegreseando al menu...')
                            break
                        while True:
                            tupla_turno = {1, 2, 3}
                            print('Turno #1. Mañana')
                            print('Turno #2. Mediodía')
                            print('Turno #3. Tarde')
                            print()
                            turno_cita = input('Favor elegir el turno de la cita. (Solo el numero.): ')
                            if turno_cita == '':
                                print('Error. El valor no puede omitirse')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    turno_cita = '*'
                                    break
                            if turno_cita.isdigit():
                                turno_cita = int(turno_cita)
                                if turno_cita in tupla_turno:
                                    if turno_cita == 1:
                                        turno_str = 'Mañana'
                                    elif turno_cita == 2:
                                        turno_str = 'Mediodía'
                                    else:
                                        turno_str = 'Tarde'
                                    print(f'El turno de la cita es de {turno_str}.')
                                    break
                                else:
                                    print('Error. Ingrese un turno valido. Intente de nuevo. ')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        turno_cita = '*'
                                        break
                            else:
                                print('Error. El dato ingresado no es numerico. Intenta de nuevo.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    turno_cita = '*'
                                    break
                        if turno_cita == '*':
                            print('\nSe cancelo la operacion.\nRegreseando al menu...')
                            break
                        try:
                            with sql.connect("Evidencia_3.db") as conn:
                                mi_cursor = conn.cursor()
                                datos = (clave_seleccionada,fecha_cita, turno_str)
                                mi_cursor.execute('INSERT INTO Citas (clave_paciente, fechaCita, turno) VALUES (?,?,?)' , datos)
                        except sql.Error as e:
                            print(f'ERROR. Base de datos {e}')
                        except:
                            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                        finally:
                            conn.close()    
                        try:
                            with sql.connect('Evidencia_3.db') as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute('SELECT folio, clave_paciente, fechaCita, edad, turno FROM Citas WHERE clave_paciente = ? AND fechaCita = ? AND turno = ?', (clave_seleccionada, fecha_cita, turno_str))
                                datos_cita = mi_cursor.fetchone()
                                if datos_cita:
                                    folio, clave_paciente, fechaCita, edad, turno = datos_cita
                                    mi_cursor.execute('SELECT apellidoPaterno, apellidoMaterno, nombres, fechaNacimiento FROM Pacientes WHERE clave = ?', (clave_paciente,))
                                    detalles_paciente = mi_cursor.fetchone()
                                    apellido_paterno, apellido_materno, nombres, fecha_nacimiento = detalles_paciente
                                    print('La cita se ha registrado correctamente.')
                                    print()
                                    print('Detalles de la cita:')
                                    print(f'Clave del Paciente: {clave_paciente}')
                                    print(f'Folio de la cita: {folio}')
                                    print(f'Apellido Paterno: {apellido_paterno}')
                                    print(f'Apellido Materno: {apellido_materno}')
                                    print(f'Nombres: {nombres}')
                                    print(f'Fecha de Nacimiento: {fecha_nacimiento}')
                                    print(f'Fecha de Cita: {fechaCita}')
                                    print(f'Turno de Cita: {turno}')
                                    print()
                                else:
                                    print('\nNo se encontraron datos para la cita recién registrada.')
                        except sql.Error as e:
                            print(f'ERROR. Base de datos {e}')
                        except:
                            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                        finally:
                            conn.close()
                        break
                    else:
                        print()
            if opcion_submenuCitas == 2:
                folios_in_list = []
                try:
                    with sql.connect('Evidencia_3.db') as conn:
                        mi_cursor = conn.cursor()
                        mi_cursor.execute('SELECT folio FROM Citas')
                        folios_disponibles = mi_cursor.fetchall()  
                        if folios_disponibles:
                            print('Folios disponibles:')
                            for folio in folios_disponibles:
                                print(f'\t-{folio[0]}-')  
                                folios_in_list.append(folio[0])
                        else:
                            print('Error. No hay pacientes con cita.')
                except sql.Error as e:
                    print(f'ERROR. En la Base de datos: {e}')
                except:
                    print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                finally:
                    conn.close()
                if folios_in_list:
                    while True:
                        hora_actual = datetime.datetime.now()
                        hora_actual_str = hora_actual.strftime("%H:%M:%S")
                        while True:
                            try:
                                folio_seleccionado = int(input('Ingrese el folio del paciente: '))
                                if folio_seleccionado in folios_in_list:
                                    break
                                else:
                                    print('ERROR. Elige un folio existente')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        folio_seleccionado = '*'
                                        break
                            except ValueError:
                                print('ERROR. El folio debe ser numerico. Intente de nuevo.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    folio_seleccionado = '*'
                                    break
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    folio_seleccionado = '*'
                                    break
                        if folio_seleccionado == '*':
                            print('Se cancelo la operación. \nRegresando al menu...')
                            break
                        while True:
                            try:
                                peso_paciente = float(input('Ingrese el peso en kilogramos del paciente: '))
                                if peso_paciente <= 0:
                                    print('Error: El peso debe ser un valor numérico positivo.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        peso_paciente = '*'
                                        break
                                else:
                                    print(f'Peso registrado correctamente: {peso_paciente} ')
                                    print()
                                    break
                            except ValueError:
                                print('Error: Por favor, ingrese un valor numérico válido para el peso.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    peso_paciente = '*'
                                    break
                        if peso_paciente == '*':
                            print('Se cancelo la operación. \nRegresando al menu...')
                            break
                        while True:
                            try:
                                estatura_paciente = float(input('Ingrese la estatura en centímetros del paciente: '))
                                if estatura_paciente <= 0:
                                    print('Error: La estatura debe ser un valor numérico positivo.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        estatura_paciente = '*'
                                        break
                                else:
                                    print('Estatura registrada correctamente:', estatura_paciente)
                                    print()
                                    break
                            except ValueError:
                                print('Error: Por favor, ingrese un valor numérico válido para la estatura.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    estatura_paciente = '*'
                                    break 
                        if estatura_paciente == '*':
                            print('Se cancelo la operación. \nRegresando al menu...')
                            break
                        while True:
                            try:
                                sistolica = int(input('Ingrese el valor de la sistolica: '))
                                if sistolica < 210:
                                    print(f'El valor de la sistolica regirstrado correctamente: {sistolica} \n')
                                    break
                                else:
                                    print('Error. Debe ser un valor valido. Intente de nuevo.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        sistolica = '*'
                                        break 
                            except ValueError:
                                print('Error. El dato ingresado no es numerico. Intente de nuevo.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    sistolica = '*'
                                    break 
                        if sistolica == '*':
                            print('Se cancelo la operación. \nRegresando al menu...')
                            break
                        while True:
                            try:
                                asistolica = int(input('Ingresa el valor de la diastolica: '))
                                if asistolica < 210:
                                    print(f'Valor de las diastolica registrado correctamente: {asistolica} \n')
                                    break
                                else:
                                    print('Error. Debe ser un valor valido. Intente de nuevo.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        asistolica = '*'
                                        break     
                            except ValueError:
                                print('Error. El dato ingresado no es numerico. Intente de nuevo.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    asistolica = '*'
                                    break                     
                        if asistolica == '*':
                            print('Se cancelo la operación. \nRegresando al menu...')
                            break
                        sistolica = str(sistolica)
                        asistolica = str(asistolica)
                        presion_arterial = f'{sistolica.zfill(3)}/{asistolica.zfill(3)}'
                        while True:
                            print('El diagnostico no puede excederse de los 200 caracteres')
                            diagnostico = input('Diagnostico: ')
                            if len(diagnostico) == 0:
                                print('ERROR. El diagnostico no puede omitirse. Intente de nuevo.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    diagnostico = '*'
                                    break 
                                continue
                            if len(diagnostico) > 200:
                                print('ERROR. El diagnostico ingresado excede del limite. Intente de nuevo.')
                                print('\n== IMPORTANTE ==')
                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                print()
                                if cancelar == '*':
                                    diagnostico = '*'
                                    break
                                continue
                            print('Diagnostico valido')
                            break
                        if diagnostico == '*':
                            print('Se cancelo la operación. \nRegresando al menu...')
                            break
                        try:
                            with sql.connect('Evidencia_3.db') as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute('SELECT Pacientes.fechaNacimiento FROM Pacientes INNER JOIN Citas ON Pacientes.clave = Citas.clave_paciente WHERE Citas.folio = ?', (folio_seleccionado,))
                                fecha_nacimiento_str = mi_cursor.fetchone()[0]
                                fecha_nacimiento_paciente = datetime.datetime.strptime(fecha_nacimiento_str, '%Y-%m-%d').date()
                                fecha_actual = datetime.date.today()
                                edad = fecha_actual.year - fecha_nacimiento_paciente.year - ((fecha_actual.month, fecha_actual.day) < (fecha_nacimiento_paciente.month, fecha_nacimiento_paciente.day))
                                datos = (edad, peso_paciente, estatura_paciente, hora_actual_str, presion_arterial, diagnostico, folio_seleccionado)
                                mi_cursor.execute('UPDATE Citas SET edad=?, peso=?, estatura=?, horaLlegada=?, presionArterial=?, diagnostico=? WHERE folio=?', datos)
                        except sql.Error as e:
                            print(f'ERROR. En la Base de datos: {e}')
                        except:
                            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                        finally:
                            conn.close()
                        break
                else:
                    print('Regresando al menú...\n')
            if opcion_submenuCitas == 3:
                while True:
                    while True:
                        print('----------------- CANCELACION DE CITAS -----------------')
                        print('1. Busqueda por fecha.')
                        print('2. Busqueda por paciente')
                        print('3. Salir')
                        try:
                            tupla_cancelacion = (1, 2, 3)
                            opcion_cancelacion = int(input('Selecciona la opcion del menu. Solo el numero: '))
                            if opcion_cancelacion in tupla_cancelacion:
                                print('Opcion valida')
                                break
                            else:
                                print('ERROR. Selecciona una opcion valida del menu. Intenta de nuevo')
                        except ValueError:
                            print('ERROR. El valor ingresado no es numerico. Intenta de nuevo')
                    if opcion_cancelacion == 1:
                        print('----------------- BUSQUEDA POR FECHA -----------------')
                        while True:
                            while True:
                                try:
                                    busqueda_fecha_cancelacion = input('Ingrese la fecha de la cita a buscar, en el formato (mm/dd/yyyy): ')
                                    busqueda_fecha_cancelacion = datetime.datetime.strptime(busqueda_fecha_cancelacion, '%m/%d/%Y').date()
                                    break
                                except ValueError:
                                    print('ERROR. El formato de la fecha no es el solicitado. Intente de nuevo')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    print()
                                    if cancelar == '*':
                                        busqueda_fecha_cancelacion = '*'
                                        break
                            if busqueda_fecha_cancelacion == '*':
                                print('Se cancelo la operación. \nRegresando al menu...')
                                break
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute('''
                                        SELECT Citas.folio, Pacientes.apellidoPaterno, Pacientes.apellidoMaterno, Pacientes.nombres, Citas.fechaCita, Citas.turno
                                        FROM Citas
                                        INNER JOIN Pacientes ON Citas.clave_paciente = Pacientes.clave
                                        WHERE Citas.fechaCita = ? ''', (busqueda_fecha_cancelacion,))
                                    citas_en_fecha = mi_cursor.fetchall()
                                    if citas_en_fecha:
                                        print('Citas para la fecha proporcionada:')
                                        print('-' * 110)
                                        print('{:<10} {:<20} {:<20} {:<20} {:<20} {:<20}'.format('Folio', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Fecha de Cita', 'Turno'))
                                        print('-' * 110)
                                        for folio, apellidoPaterno, apellidoMaterno, nombres, fechaCita, turno in citas_en_fecha:
                                            print('{:<10} {:<20} {:<20} {:<20} {:<20} {:<20}'.format(folio, apellidoPaterno, apellidoMaterno, nombres, fechaCita, turno))
                                        print('-' * 110)
                                        while True:
                                            try:
                                                folio_cancelar = int(input('Ingrese el folio de la cita a eliminar: '))
                                                break
                                            except ValueError:
                                                print(f'ERROR. El dato ingresado no es numérico. Intente de nuevo')
                                                print('ERROR. El diagnostico ingresado excede del limite. Intente de nuevo.')
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                print()
                                                if cancelar == '*':
                                                    folio_cancelar = '*'
                                                    break
                                        mi_cursor.execute('DELETE FROM Citas WHERE folio=?', (folio_cancelar,))
                                        eliminacion = mi_cursor.fetchall()
                                        if eliminacion:
                                            print('Se ha eliminado la cita.')
                                        else:
                                            print('No se elimino ninguna cita.')
                                    else:
                                        print('No se encontraron citas para la fecha proporcionada.')
                            except sql.Error as e:
                                print(f'ERROR. En la Base de datos: {e}')
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            finally:
                                conn.close()
                            break
                            if folio_cancelar == '*':
                                print('Se cancelo la operación. \nRegresando al menu...')
                                break
                    if opcion_cancelacion == 2:
                        while True:
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute("""
                                        SELECT c.folio, p.clave, p.nombres, p.apellidoPaterno, p.apellidoMaterno, c.fechaCita
                                        FROM Citas c
                                        INNER JOIN Pacientes p ON c.clave_paciente = p.clave
                                        WHERE c.fechaCita > datetime('now')
                                    """)
                                    citas_pendientes = mi_cursor.fetchall()
                                    if citas_pendientes: 
                                        print("CITAS PENDIENTES")
                                        print("-" * 132)
                                        print("| {:<8} | {:<10} | {:<25} | {:<25} | {:<25} | {:<20} |".format("Folio", "Clave", "Nombre", "Apellido Paterno", "Apellido Materno", "Fecha de la Cita"))
                                        print("-" * 132)
                                        
                                        for cita in citas_pendientes:
                                            folio, clave, nombres, apellidoPaterno, apellidoMaterno, fechaCita = cita
                                            print("| {:<8} | {:<10} | {:<25} | {:<25} | {:<25} | {:<20} |".format(folio, clave, nombres, apellidoPaterno, apellidoMaterno, fechaCita))
                                            print()
                                        mi_cursor.execute("""
                                            SELECT c.folio
                                            FROM Citas c
                                            INNER JOIN Pacientes p ON c.clave_paciente = p.clave
                                            WHERE c.fechaCita > datetime('now')
                                        """)
                                        folios_pendientes = [cita[0] for cita in mi_cursor.fetchall()]
                                        
                                        while True:
                                            try:
                                                folio_eliminar = int(input('Ingrese el folio a elimninar: '))
                                                if folio_eliminar in folios_pendientes:
                                                    mi_cursor.execute('DELETE FROM Citas WHERE folio = ?', (folio_eliminar,))
                                                    print('La cita ha sido eliminada exitosamente')
                                                    break
                                                else:
                                                    print('ERROR. Ingrese un folio valido. Intente de nuevo')
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                print()
                                                if cancelar == '*':
                                                    folio_eliminar = '*'
                                                    break
                                            except ValueError:
                                                print('ERROR. El dato ingresado no es numerico. Intente de nuevo.')
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                print()
                                                if cancelar == '*':
                                                    folio_eliminar = '*'
                                                    break
                                    else:
                                        print('No hay citas pendientes')
                            except sql.Error as e:
                                print(f'ERROR. En la Base de datos: {e}')
                            finally:
                                conn.close()
                            if folio_eliminar == '*':
                                print('Se cancelo la operación. \nRegresando al menu...')
                                break
                            break
                    if opcion_cancelacion == 3:
                        print('Saliendo del submenu')
                        break
            if opcion_submenuCitas == 4:
                print('Saliendo del submenu de citas')
                print()
                break
    if opcion_menuPrincipal == 3:
        while True:
            print()
            print('---------------------- MENU REPORTES Y CONSULTAS ----------------------')
            print('1. Reportes por citas.')
            print('2. Reportes de pacientes.')
            print('3. Estadisticos Demograficos.')
            print('4. Salir del menu.')
            tupla_1 = (1, 2, 3, 4)
            while True:
                opcion_menu4_1 = input('Selecciona la opcion. Solo el numero:  ')
                if opcion_menu4_1.isdigit():
                    opcion_menu_4 = int(opcion_menu4_1)
                    if opcion_menu_4 in tupla_1:
                        print('Se registro correctamente.')
                        print()
                        break
                    else:
                        print('ERROR. Selecciona una opcion valida en el menu.')
                else:
                    print(f'ERROR. El dato {opcion_menu4_1} no es numerico. Intenta de nuevo')
            if opcion_menu_4 == 1:
                while True:
                    print('---------------------- SUBMENU REPORTE DE CITAS ----------------------')
                    print('1. Reporte por periodo.')
                    print('2. Reporte de pacientes.')
                    print('3. Salir del menu')
                    while True:
                        opcion_1 = input('Selecciona la opcion del menu. Solo el numero:  ')
                        if opcion_1.isdigit():
                            opcion_submenu_1 = int(opcion_1)
                            if opcion_submenu_1 in tupla_1:
                                print('Se registro correctamente.')
                                break
                            else:
                                print('ERROR. Selecciona una opcion valida del menu')
                        else:
                            print(f'ERROR. El dato {opcion_1} no es numerio. Intenta de nuevo.')
                    if opcion_submenu_1 == 1:
                        while True:
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    while True:
                                        try:
                                            fecha_inicial = input("Introduzca la fecha inicial (Valor de la fecha en el formato mm/dd/aaaa): ")
                                            fecha_inicial = datetime.datetime.strptime(fecha_inicial, "%m/%d/%Y").date()
                                            break
                                        except ValueError:
                                            print("El dato debe estar en el formato proporcionado")
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                fecha_inicial = '*'
                                                break
                                    if fecha_inicial == '*':
                                        print('Se cancelo la operacion.\nRegreseando al menu...')
                                        break
                                    while True:
                                        try:
                                            fecha_fin = input("Introduzca la fecha final (Valor de la fecha en el formato mm/dd/aaaa): ")
                                            fecha_fin = datetime.datetime.strptime(fecha_fin, "%m/%d/%Y").date()
                                            if fecha_fin > fecha_inicial:
                                                break
                                            else:
                                                print("La fecha inicial no debe ser mayor a la final. Intente de nuevo")
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                if cancelar == '*':
                                                    fecha_fin = '*'
                                                    break
                                        except ValueError:
                                            print("El dato debe estar en el formato proporcionado")
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                fecha_fin = '*'
                                                break
                                    if fecha_fin == '*':
                                        print('Se cancelo la operacion.\nRegreseando al menu...')
                                        break
                                    mi_cursor.execute("""
                                        SELECT c.folio, p.clave, p.apellidoPaterno, p.apellidoMaterno, p.nombres, p.fechaNacimiento, p.sexoPaciente, c.fechaCita, c.turno, c.edad, c.peso, c.estatura, c.horaLlegada, c.presionArterial
                                        FROM Citas c
                                        INNER JOIN Pacientes p ON c.clave_paciente = p.clave
                                        WHERE c.fechaCita BETWEEN ? AND ?
                                    """, (fecha_inicial, fecha_fin))
                                    datos_citas_periodo = mi_cursor.fetchall()
                                    if datos_citas_periodo:
                                                print("{:<10} {:<10} {:<20} {:<20} {:<20} {:<17} {:<10} {:<17} {:<10} {:<15} {:<15} {:<15} {:<15} {:<15}".format(
                                                    "Folio", "Clave", "Apellido Paterno", "Apellido Materno", "Nombres", "Fecha Nacimiento", "Sexo", "Fecha Cita", "Turno", "Edad", "Peso", "Estatura", "Hora Llegada", "Presión Arterial"))
                                                print('-' * 224)
                                                
                                                for fila in datos_citas_periodo:
                                                    fila_formateada = []
                                                    for valor in fila:
                                                        if valor is None:
                                                            fila_formateada.append("No disponible")
                                                        else:
                                                            fila_formateada.append(valor)
                                                    print("{:<10} {:<10} {:<20} {:<20} {:<20} {:<17} {:<10} {:<17} {:<10} {:<15} {:<15} {:<15} {:<15} {:<15}".format(*fila_formateada))
                                                
                                                print('\n1. Exportar Mediante CSV.')
                                                print('2. Exportar Mediante EXCEL')
                                                print('3. No exportar')
                                                while True:
                                                    try:
                                                        exportar = int(input('Seleccione la opción (solo el número): '))
                                                        if exportar in [1, 2, 3]:
                                                            break
                                                        else:
                                                            print('Error. Seleccione una opción válida.')
                                                            print('\n== IMPORTANTE ==')
                                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                            if cancelar == '*':
                                                                exportar = '*'
                                                                break
                                                    except ValueError:
                                                        print('Error. El dato ingresado no es numérico. Intenta de nuevo.')
                                                        print('\n== IMPORTANTE ==')
                                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                        if cancelar == '*':
                                                            exportar = '*'
                                                            break
                                                if exportar == '*':
                                                    print('Se cancelo la operacion.\nRegreseando al menu...')
                                                    break
                                                if exportar == 1:
                                                    try:
                                                        with open('Export_Pacientes_Citas_Periodo.csv', 'w', newline='') as archivo_csv:
                                                            escritor_csv = csv.writer(archivo_csv)
                                                            encabezados = ["Folio", "Clave", "Apellido Paterno", "Apellido Materno", "Nombres", "Fecha Nacimiento", "Sexo", "Fecha Cita", "Turno", "Edad", "Peso", "Estatura", "Hora Llegada", "Presión Arterial"]
                                                            escritor_csv.writerow(encabezados)
                                                            escritor_csv.writerows(datos_citas_periodo)
                                                        print("Los datos se han exportado a Export_Pacientes_Citas_Periodo.csv correctamente.")
                                                    except FileNotFoundError:
                                                        print('Error. Ocurrio un problema al generar Export_Pacientes_Citas_Periodo.csv')
                                                if exportar == 2:
                                                    try:
                                                        df = pd.DataFrame(datos_citas_periodo, columns=["Folio", "Clave", "Apellido Paterno", "Apellido Materno", "Nombres", "Fecha Nacimiento", "Sexo", "Fecha Cita", "Turno", "Edad", "Peso", "Estatura", "Hora Llegada", "Presión Arterial"])
                                                        df.to_excel("Export_Pacientes_Citas_Periodo.xlsx", index=False)
                                                        print("Los datos se han exportado a Export_Pacientes_Citas_Periodo.xlsx correctamente.")
                                                    except FileNotFoundError:
                                                        print('Error. Ocurrio un problema al generar Export_Pacientes_Citas_Periodo.xlsx')
                                                
                                                if exportar == 3:
                                                    print('No se exportó el reporte.')
                                    else:
                                        print('No hay pacientes con cita...')
                            except sql.Error as e:
                                print(e)
                            except Exception as ex:
                                print(f"Se produjo el siguiente error: {ex}")
                            break
                    if opcion_submenu_1 == 2:
                        print()
                        print('---------------------- PACIENTES CON CITA ----------------------')
                        print('Detalles de todas las citas:')
                        try:
                            with sql.connect('Evidencia_3.db') as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute("""
                                    SELECT c.folio, p.clave, p.apellidoPaterno, p.apellidoMaterno, p.nombres, p.fechaNacimiento, p.sexoPaciente, c.fechaCita, c.turno, c.edad, c.peso, c.estatura, c.horaLlegada, c.presionArterial
                                    FROM Citas c
                                    INNER JOIN Pacientes p ON c.clave_paciente = p.clave
                                """)
                                datos_citas = mi_cursor.fetchall()
                                
                                if datos_citas:
                                    print("{:<10} {:<10} {:<20} {:<20} {:<20} {:<17} {:<10} {:<17} {:<10} {:<15} {:<15} {:<15} {:<15} {:<15}".format(
                                        "Folio", "Clave", "Apellido Paterno", "Apellido Materno", "Nombres", "Fecha Nacimiento", "Sexo", "Fecha Cita", "Turno", "Edad", "Peso", "Estatura", "Hora Llegada", "Presión Arterial"))
                                    print('-' * 224)
                                    # Datos tabulados
                                    for fila in datos_citas:
                                        fila_formateada = []
                                        for valor in fila:
                                            if valor is None:
                                                fila_formateada.append("No disponible")
                                            else:
                                                fila_formateada.append(valor)
                                        print("{:<10} {:<10} {:<20} {:<20} {:<20} {:<17} {:<10} {:<17} {:<10} {:<15} {:<15} {:<15} {:<15} {:<15}".format(*fila_formateada))
                                    
                                    print('\n1. Exportar Mediante CSV.')
                                    print('2. Exportar Mediante EXCEL')
                                    print('3. No exportar')
                                    while True:
                                        try:
                                            exportar = int(input('Seleccione la opcion (unicamente el numero): '))
                                            if exportar in [1, 2, 3]:
                                                break
                                            else:
                                                print('Error. Seleccione una opcion valida.')
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                if cancelar == '*':
                                                    exportar = '*'
                                                    break
                                        except ValueError:
                                            print('Error. El dato ingresado no es numerico. Intenta de nuevo.')
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                exportar = '*'
                                                break
                                    if exportar == '*':
                                        print('Se cancelo la operacion.\nRegreseando al menu...')
                                        break
                                    if exportar == 1:
                                        try:
                                            with open('Export_Pacientes_Citas.csv', 'w', newline='') as archivo_csv:
                                                escritor_csv = csv.writer(archivo_csv)
                                                encabezados = ["Folio", "Clave", "Apellido Paterno", "Apellido Materno", "Nombres", "Fecha Nacimiento", "Sexo", "Fecha Cita", "Turno", "Edad", "Peso", "Estatura", "Hora Llegada", "Presión Arterial"]
                                                escritor_csv.writerow(encabezados)
                                                escritor_csv.writerows(datos_citas)
                                            print("Los datos se han exportado a Export_Pacientes_Citas.csv correctamente.")
                                        except FileNotFoundError:
                                            print('Error. Ocurrio un error al exportar Export_Pacientes_Citas.csv')
                                    if exportar == 2:
                                        try:
                                            df = pd.DataFrame(datos_citas, columns=["Folio", "Clave", "Apellido Paterno", "Apellido Materno", "Nombres", "Fecha Nacimiento", "Sexo", "Fecha Cita", "Turno", "Edad", "Peso", "Estatura", "Hora Llegada", "Presión Arterial"])
                                            df.to_excel("Export_Pacientes_Citas.xlsx", index=False)
                                            print("Los datos se han exportado a Export_Pacientes_Citas.xlsx correctamente.")
                                        except FileNotFoundError:
                                            print('Error. Ocurrio un error al exportar Export_Pacientes_Citas.xlsx')
                                    if exportar == 3:
                                        print('No se exporto el reporte.')
                                else:
                                    print('No hay pacientes con cita...')
                        except sql.Error as e:
                            print(e)
                        except Exception as ex:
                            print(f"Se produjo el siguiente error: {ex}")
                    if opcion_submenu_1 == 3:
                        print('Saliendo del menu...')
                        break                   
            if opcion_menu_4 == 2:
                while True:
                    print('---------------------- SUBMENU REPORTE DE PACIENTES ----------------------')
                    print('1. Listado completo de los pacientes.')
                    print('2. Busqueda por clave del paciente')
                    print('3. Busqueda por apellidos y nombres')
                    print('4. Salir del menu')
                    tupla_2 = (1, 2, 3, 4)
                    while True:
                        opcion_2 = input('Selecciona la opcion del menu. Solo el numero:  ')
                        if opcion_2.isdigit():
                            opcion_submenu_2 = int(opcion_2)
                            if opcion_submenu_2 in tupla_2:
                                print('Se registro correctamente')
                                print()
                                break
                            else:
                                print('ERROR. Seleccciona una opcion valida en el menu.')
                        else:
                            print(f'ERROR. El dato ({opcion_2}) no es numerico. Intenta de nuevo.')
                    if opcion_submenu_2 == 1:
                        print('---------------------- LISTADO DE PACIENTES ----------------------')
                        while True:
                            print('Detalles de todos los pacientes:')
                            try:
                                with sql.connect ('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute('SELECT * FROM Pacientes')
                                    pacientes = mi_cursor.fetchall()
                                    if pacientes:
                                        pacientes_df = pd.read_sql_query('SELECT * FROM Pacientes', conn)
                                        print("{:<10} {:<20} {:<20} {:<20} {:<20} {:<10}".format("Clave", "Apellido Paterno", "Apellido Materno", "Nombres", "Fecha de Nacimiento", "Sexo"))
                                        print("-"*110)
                                        for paciente in pacientes:
                                            clave, apellido_paterno, apellido_materno, nombres, fecha_nacimiento, sexo = paciente
                                            print("{:<10} {:<20} {:<20} {:<20} {:<20} {:<10}\n".format(clave, apellido_paterno, apellido_materno, nombres, fecha_nacimiento, sexo))
                                        
                                        print('1. Exportar reporte mediante CSV.')
                                        print('2. Exportar reporte mediante EXCEL.')
                                        print('3. NO exportar reporte.')
                                        while True:
                                            try:
                                                exportacion = int(input('Seleccione la opcion (Solo el número): '))
                                                if exportacion in (1, 2, 3):
                                                    break
                                                else:
                                                    print('ERROR. Seleccione una opcion valida. Intente de nuevo.')
                                                    print('\n== IMPORTANTE ==')
                                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                    if cancelar == '*':
                                                        exportacion = '*'
                                                        break
                                            except ValueError:
                                                print('ERROR. Ingrese solo el numero de la opcion. Intente de nuevo')
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                if cancelar == '*':
                                                    exportacion = '*'
                                                    break
                                        if exportacion == '*':
                                            print('Se cancelo la operacion.\nRegreseando al menu...')
                                            break
                                        if exportacion == 1:
                                            try:
                                                with open ('Export_Pacientes.csv', 'w', newline='') as archivo:
                                                    grabador_Pacientes = csv.writer(archivo)
                                                    grabador_Pacientes.writerow(('Clave', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Fecha de nacimiento', 'Sexo'))
                                                    grabador_Pacientes.writerows(pacientes)
                                                print('Los datos se han exportado ha Export_Pacientes.csv exitosamente.')
                                            except FileNotFoundError:
                                                print('ERROR. No se pudo exportar los datos del archivo Pacientes.csv')
                                            except csv.Error as e:
                                                print(f"Error CSV: {e}")   
                                        if exportacion == 2:
                                            try:
                                                pacientes_df.to_excel('Export_Pacientes.xlsx', index=False)
                                                print("Los datos se han exportado a Export_Pacientes.xlsx satisfactoriamente.")
                                            except FileNotFoundError:
                                                print('ERROR. No se pudo exportar los datos del archivo Pacientes.csv')
                                            except csv.Error as e:
                                                print(f"Error CSV: {e}") 
                                        if exportacion == 3:
                                            print('No se exporto el reporte.')
                                    else:
                                        print('-No hay pacientes registrados.-')
                            except Error as e:
                                print (e)
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            break
                    if opcion_submenu_2 == 2:
                        print()
                        print('---------------------- BUSQUEDA POR CLAVE ----------------------')
                        while True:
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    while True:
                                        try:
                                            clave_buscar = int(input('Ingrese la clave a buscar: '))
                                            break
                                        except ValueError:
                                            print('ERROR. El valor ingresado no es numerico. Intente de nuevo.')
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                clave_buscar = '*'
                                                break
                                    if clave_buscar == '*':
                                        print('Se cancelo la operacion.\nRegreseando al menu...')
                                        break
                                    mi_cursor.execute('SELECT * FROM Pacientes WHERE clave = ?', (clave_buscar,))
                                    paciente_Encontrado = mi_cursor.fetchall()
                                    if paciente_Encontrado:
                                        for clave, apellidoPaterno, apellidoMaterno, nombres, fechaNacimiento, sexo in paciente_Encontrado:
                                            print(f'Clave: {clave}')
                                            print(f'  Apellido Paterno: {apellidoPaterno}')
                                            print(f'  Apellido Materno: {apellidoMaterno}')
                                            print(f'  Nombres: {nombres}')
                                            print(f'  Fecha de Nacimiento: {fechaNacimiento}')
                                            print(f'  Sexo: {sexo}')
                                        while True:
                                            pregunta_expediente_clinico = input('Desea consultar el expediente clinico (SI/NO):').upper()
                                            if pregunta_expediente_clinico.isalpha():
                                                print()
                                                break
                                            else:
                                                print('ERROR. Ingresa solo "SI" o "NO". Intenta de nuevo.')
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                if cancelar == '*':
                                                    pregunta_expediente_clinico = '*'
                                                    break
                                        if pregunta_expediente_clinico == '*':
                                            print('Se cancelo la operacion.\nRegreseando al menu...')
                                            break
                                        if pregunta_expediente_clinico == 'SI':
                                            mi_cursor.execute('SELECT fechaCita ,edad, peso, estatura, horaLlegada, presionArterial, diagnostico FROM Citas WHERE clave_paciente = ?', (clave_buscar,))
                                            expediente_clinico = mi_cursor.fetchall()
                                            if expediente_clinico:
                                                for fechaCita, edad, peso, estatura, horaLlegada, presionArterial, diagnostico in expediente_clinico:
                                                    print(f'Fecha de la cita: {fechaCita if fechaCita is not None else "No disponible"}')
                                                    print('Hora de llegada:', horaLlegada if horaLlegada is not None else 'No disponible')
                                                    print('Edad:', edad if edad is not None else 'No disponible')
                                                    print('Peso del paciente:', peso if peso is not None else 'No disponible')
                                                    print('Estatura del paciente:', estatura if estatura is not None else 'No disponible')
                                                    print('Presion arterial:', presionArterial if presionArterial is not None else 'No disponible')
                                                    print('Diagnostico:', diagnostico if diagnostico is not None else 'No disponible')
                                            else:
                                                print('No tiene expediente clinico')
                                    if not paciente_Encontrado:
                                        print(f'No se econtro ningun paciente con la clave {clave_buscar}')
                            except Error as e:
                                print (e)
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            break
                    if opcion_submenu_2 == 3:
                        print('---------------------- BUSQUEDA POR APELLIDO Y NOMBRES ----------------------')
                        while True:
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    while True:
                                        primer_apellido_busqueda = input('Ingrese el primer apellido del paciente: ')
                                        if primer_apellido_busqueda.strip().replace(' ', '').isalpha():
                                            primer_apellido_busqueda = primer_apellido_busqueda.upper()  
                                            break
                                        else:
                                            print('Error: El primer apellido solo puede contener letras. Por favor, inténtelo de nuevo.')
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                primer_apellido_busqueda = '*'
                                                break
                                    if primer_apellido_busqueda == '*':
                                        print('Se cancelo la operacion.\nRegreseando al menu...')
                                        break
                                    while True:
                                        nombres_busqueda = input('Ingrese los nombres del paciente: ')
                                        if nombres_busqueda.strip().replace(' ', '').isalpha():
                                            nombres_busqueda = nombres_busqueda.upper() 
                                            break
                                        else:
                                            print('Error: Los nombres solo pueden contener letras. Por favor, inténtelo de nuevo.')
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                nombres_busqueda = '*'
                                                break
                                    if nombres_busqueda == '*':
                                        print('Se cancelo la operacion.\nRegreseando al menu...')
                                        break
                                    mi_cursor.execute('SELECT * FROM Pacientes WHERE apellidoPaterno = ? AND nombres = ?', (primer_apellido_busqueda, nombres_busqueda))
                                    pacientes_encontrados = mi_cursor.fetchall()
                                    if pacientes_encontrados:
                                        claves_expedientes = []
                                        for clave, apellidoPaterno, apellidoMaterno, nombres, fechaNacimiento, sexo in pacientes_encontrados:
                                            clave_expediente = clave
                                            print(f'Clave: {clave}')
                                            print(f'  Apellido Paterno: {apellidoPaterno}')
                                            print(f'  Apellido Materno: {apellidoMaterno}')
                                            print(f'  Nombres: {nombres}')
                                            print(f'  Fecha de Nacimiento: {fechaNacimiento}')
                                            print(f'  Sexo: {sexo}')
                                            claves_expedientes.append(clave)
                                        while True:
                                            pregunta_expediente_clinico = input('Desea consultar el expediente clinico (SI/NO): ').upper()
                                            if pregunta_expediente_clinico.isalpha():
                                                print()
                                                break
                                            else:
                                                print('ERROR. Ingresa solo "SI" o "NO". Intenta de nuevo.')
                                                print('\n== IMPORTANTE ==')
                                                print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                if cancelar == '*':
                                                    pregunta_expediente_clinico = '*'
                                                    break
                                        if pregunta_expediente_clinico == '*':
                                            print('Se cancelo la operacion.\nRegreseando al menu...')
                                            break
                                        if pregunta_expediente_clinico == 'SI':
                                            while True:
                                                try:
                                                    clave_a_expediente = int(input('Ingrese la clave del que desea el expediente: '))
                                                    if clave_a_expediente in claves_expedientes:
                                                        break
                                                    else:
                                                        print('Error. Ingrese una clave valida. Intente de nuevo')
                                                        print('\n== IMPORTANTE ==')
                                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                        if cancelar == '*':
                                                            clave_a_expediente = '*'
                                                            break
                                                except ValueError:
                                                    print('Error. El dato ingresado no debe ser numerico. Intente de nuevo.')
                                                    print('\n== IMPORTANTE ==')
                                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                                    if cancelar == '*':
                                                        clave_a_expediente = '*'
                                                        break
                                            if clave_a_expediente == '*':
                                                print('Se cancelo la operacion.\nRegreseando al menu...')
                                                break
                                            print('Expediente Clinico: ')
                                            mi_cursor.execute('SELECT fechaCita, edad, peso, estatura, horaLlegada, presionArterial, diagnostico FROM Citas WHERE clave_paciente = ?', (clave_a_expediente,))
                                            expediente = mi_cursor.fetchall()
                                            if expediente:
                                                for fechaCita, edad, peso, estatura, horaLlegada, presion, diagnostico in expediente:
                                                    print(f'  Fecha Cita: {fechaCita if fechaCita is not None else "No disponible"}')
                                                    print(f'  Edad: {edad if edad is not None else "No disponible"}')
                                                    print(f'  Peso del paciente: {peso if peso is not None else "No disponible"}')
                                                    print(f'  Estatura del paciente: {estatura if estatura is not None else "No disponible"}')
                                                    print(f'  Hora de llegada: {horaLlegada if horaLlegada is not None else "No disponible"}')
                                                    print(f'  Presión arterial: {presion if presion is not None else "No disponible"}')
                                                    print(f'  Diagnóstico: {diagnostico if diagnostico is not None else "No disponible"}')
                                            else:
                                                print('El paciente no tiene expediente clinico')
                                        else:
                                            print('Volviendo al menu....')
                                    if not pacientes_encontrados:
                                        print('No se encontro ningun paciente.')
                            except Error as e:
                                print (e)
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            break
                    if opcion_submenu_2 == 4:
                        print('Saliendo del menu...')
                        break
            if opcion_menu_4 == 3:
                while True:
                    print('---------------------- ESTADISTICOS DEMOGRAFICOS ----------------------')
                    print('1. Por edad.')
                    print('2. Por sexo.')
                    print('3. Por edad y sexo.')
                    print('4. Salir del submenu')
                    while True:
                        try:
                            opcion_stats = int(input('Seleccione la opcion (solo el numero): '))
                            if opcion_stats in [1, 2, 3, 4]:
                                break
                            else:
                                print('Error. Selecciona una opcion valida del menu.')
                        except ValueError:
                            print('Errro. Ingrese un dato numerico. Intente de nuevo.')
                    
                    if opcion_stats == 1:
                        print('\n--- ESTADISTICO DEMOGRAFICO POR EDAD ---')
                        while True:
                            while True:
                                try:
                                    edad_inicio = int(input('Ingrese desde que edad necesita: '))
                                    if edad_inicio > 0:
                                        break
                                    else:
                                        print('Error. El dato ingresado no acepta valores negativos. Intente de nuevo')
                                        print('\n== IMPORTANTE ==')
                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                        if cancelar == '*':
                                            edad_inicio = '*'
                                            break
                                except ValueError:
                                    print(f'El dato ingresado no es numerico. Intente de nuevo')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    if cancelar == '*':
                                        edad_inicio = '*'
                                        break
                            if edad_inicio == '*':
                                print('Se cancelo la operacion.\nRegreseando al menu...')
                                break
                            while True:
                                try:
                                    edad_final = int(input('Ingrese hasta que edad necesita: '))
                                    if edad_final > edad_inicio:
                                        if edad_final > 0:
                                            break
                                        else:
                                            print('Error. El dato ingresado no acepta valores numericos. Intente de nuevo')
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                edad_final = '*'
                                                break
                                    else:
                                        print('Error. La edad final debe ser mayor que la edad inicial. Intente de nuevo')
                                        print('\n== IMPORTANTE ==')
                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                        if cancelar == '*':
                                            edad_final = '*'
                                            break
                                except ValueError:
                                    print('Error. El dato ingresado no es numerico. Intenta de nuevo.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    if cancelar == '*':
                                        edad_final = '*'
                                        break
                            if edad_final == '*':
                                print('Se cancelo la operacion.\nRegreseando al menu...')
                                break
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute('SELECT folio, peso, estatura FROM Citas WHERE edad BETWEEN ? AND ?', (edad_inicio, edad_final))
                                    stats_edad = mi_cursor.fetchall()
                                    if stats_edad:
                                        df_edad = pd.DataFrame(stats_edad, columns=['Folio', 'Peso', 'Estatura'])
                                        df_edad.set_index('Folio', inplace=True)
                                        print("Estadísticas descriptivas para el peso:")
                                        print("Conteo:", df_edad['Peso'].count())
                                        print("Valor mínimo:", df_edad['Peso'].min())
                                        print("Valor máximo:", df_edad['Peso'].max())
                                        print("Media aritmética:", df_edad['Peso'].mean())
                                        print("Mediana:", df_edad['Peso'].median())
                                        print("Desviación estándar:", df_edad['Peso'].std())
                                        print("\nEstadísticas descriptivas para la estatura:")
                                        print("Conteo:", df_edad['Estatura'].count())
                                        print("Valor mínimo:", df_edad['Estatura'].min())
                                        print("Valor máximo:", df_edad['Estatura'].max())
                                        print("Media aritmética:", df_edad['Estatura'].mean())
                                        print("Mediana:", df_edad['Estatura'].median())
                                        print("Desviación estándar:", df_edad['Estatura'].std())
                                    else:
                                        print('No hay estadisticas para hacer el reporte.')
                            except Error as e:
                                    print (e)
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            break
                    if opcion_stats == 2:
                        print('\n--- ESTADISTICO DEMOGRAFICO POR SEXO ---')
                        while True:
                            while True:
                                print('H = HOMBRE')
                                print('M = MUJER')
                                print('N = NO CONTESTO')
                                sexo = input('Seleccione el sexo (H o M o N): ').upper()
                                if sexo.isalpha():
                                    if sexo in ['H', 'M', 'N']:
                                        print(f'Seleccionaste {sexo}.')
                                        break
                                    else:
                                        print('Error. Selecciona un sexo valido. Intenta de nuevo')
                                        print('\n== IMPORTANTE ==')
                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                        if cancelar == '*':
                                            sexo = '*'
                                            break
                                else:
                                    print('Error. Ingrese un unicamente la letra. Intente de nuevo ')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    if cancelar == '*':
                                        sexo = '*'
                                        break
                            if sexo == '*':
                                print('Se cancelo la operacion.\nRegreseando al menu...')
                                break
                            if sexo == 'H':
                                sexo = 'HOMBRE'
                            if sexo == 'M':
                                sexo = 'MUJER'
                            if sexo == 'N':
                                sexo = 'NO CONTESTO'
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute('''SELECT c.folio, c.peso, c.estatura FROM Citas c
                                                        INNER JOIN  Pacientes p ON p.clave = c.clave_paciente
                                                        WHERE p.sexoPaciente = ?''', (sexo,))
                                    stats_sexo = mi_cursor.fetchall()
                                    if stats_sexo:
                                        df_sexo = pd.DataFrame(stats_sexo, columns=['Folio', 'Peso', 'Estatura'])
                                        df_sexo.set_index('Folio', inplace=True)                         
                                        print(df_sexo)
                                        print("Estadísticas descriptivas para el peso:")
                                        print("Conteo:", df_sexo['Peso'].count())
                                        print("Valor mínimo:", df_sexo['Peso'].min())
                                        print("Valor máximo:", df_sexo['Peso'].max())
                                        print("Media aritmética:", df_sexo['Peso'].mean())
                                        print("Mediana:", df_sexo['Peso'].median())
                                        print("Desviación estándar:", df_sexo['Peso'].std())
                                        print("\nEstadísticas descriptivas para la estatura:")
                                        print("Conteo:", df_sexo['Estatura'].count())
                                        print("Valor mínimo:", df_sexo['Estatura'].min())
                                        print("Valor máximo:", df_sexo['Estatura'].max())
                                        print("Media aritmética:", df_sexo['Estatura'].mean())
                                        print("Mediana:", df_sexo['Estatura'].median())
                                        print("Desviación estándar:", df_sexo['Estatura'].std())
                                    else:
                                        print('No hay estadisticas para hacer el reporte.')
                            except Error as e:
                                print (e)
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            break
                    if opcion_stats == 3:
                        print('\n--- ESTADISTICO DEMOGRAFICO POR EDAD Y SEXO ---')
                        while True:
                            while True:
                                try:
                                    edad_inicio = int(input('Ingrese desde que edad necesita: '))
                                    if edad_inicio > 0:
                                        break
                                    else:
                                        print('Error. El dato ingresado no acepta valores negativos. Intente de nuevo')
                                        print('\n== IMPORTANTE ==')
                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                        if cancelar == '*':
                                            edad_inicio = '*'
                                            break
                                except ValueError:
                                    print(f'El dato ingresado no es numerico. Intente de nuevo')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    if cancelar == '*':
                                        edad_inicio = '*'
                                        break
                            if edad_inicio == '*':
                                print('Se cancelo la operacion.\nRegreseando al menu...')
                                break
                            while True:
                                try:
                                    edad_final = int(input('Ingrese hasta que edad necesita: '))
                                    if edad_final > edad_inicio:
                                        if edad_final > 0:
                                            break
                                        else:
                                            print('Error. El dato ingresado no acepta valores numericos. Intente de nuevo')
                                            print('\n== IMPORTANTE ==')
                                            print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                            cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                            if cancelar == '*':
                                                edad_final = '*'
                                                break
                                    else:
                                        print('Error. La edad final debe ser mayor que la edad inicial. Intente de nuevo')
                                        print('\n== IMPORTANTE ==')
                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                        if cancelar == '*':
                                            edad_final = '*'
                                            break
                                except ValueError:
                                    print('Error. El dato ingresado no es numerico. Intenta de nuevo.')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    if cancelar == '*':
                                        edad_final = '*'
                                        break
                            if edad_final == '*':
                                print('Se cancelo la operacion.\nRegreseando al menu...')
                                break
                            while True:
                                print('H = HOMBRE')
                                print('M = MUJER')
                                print('N = NO CONTESTO')
                                sexo = input('Seleccione el sexo (H o M o N): ').upper()
                                if sexo.isalpha():
                                    if sexo in ['H', 'M', 'N']:
                                        print(f'Seleccionaste {sexo}.')
                                        break
                                    else:
                                        print('Error. Selecciona un sexo valido. Intenta de nuevo')
                                        print('\n== IMPORTANTE ==')
                                        print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                        cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                        if cancelar == '*':
                                            sexo = '*'
                                            break
                                else:
                                    print('Error. Ingrese un unicamente la letra. Intente de nuevo ')
                                    print('\n== IMPORTANTE ==')
                                    print(' Al ingresar "*" se cancelara toda la operacion en curso')
                                    cancelar = input('Ingresa "*" para cancelar la opcion o presiona ENTER para continuar: ')
                                    if cancelar == '*':
                                        sexo = '*'
                                        break
                            if sexo == '*':
                                print('Se cancelo la operacion.\nRegreseando al menu...')
                                break
                            if sexo == 'H':
                                sexo = 'HOMBRE'
                            if sexo == 'M':
                                sexo = 'MUJER'
                            if sexo == 'N':
                                sexo = 'NO CONTESTO'
                            try:
                                with sql.connect('Evidencia_3.db') as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute('''SELECT c.folio, c.peso, c.estatura FROM Citas c
                                                        INNER JOIN  Pacientes p ON p.clave = c.clave_paciente
                                                        WHERE c.edad BETWEEN ? AND ? AND p.sexoPaciente = ?''', (edad_inicio, edad_final, sexo,))
                                    stats_sexo_edad = mi_cursor.fetchall()
                                    if stats_sexo_edad:
                                        df_sexo_edad = pd.DataFrame(stats_sexo_edad, columns=['Folio', 'Peso', 'Estatura'])
                                        df_sexo_edad.set_index('Folio', inplace=True)                         
                                        print("Estadísticas descriptivas para el peso:")
                                        print("Conteo:", df_sexo_edad['Peso'].count())
                                        print("Valor mínimo:", df_sexo_edad['Peso'].min())
                                        print("Valor máximo:", df_sexo_edad['Peso'].max())
                                        print("Media aritmética:", df_sexo_edad['Peso'].mean())
                                        print("Mediana:", df_sexo_edad['Peso'].median())
                                        print("Desviación estándar:", df_sexo_edad['Peso'].std())
                                        print("\nEstadísticas descriptivas para la estatura:")
                                        print("Conteo:", df_sexo_edad['Estatura'].count())
                                        print("Valor mínimo:", df_sexo_edad['Estatura'].min())
                                        print("Valor máximo:", df_sexo_edad['Estatura'].max())
                                        print("Media aritmética:", df_sexo_edad['Estatura'].mean())
                                        print("Mediana:", df_sexo_edad['Estatura'].median())
                                        print("Desviación estándar:", df_sexo_edad['Estatura'].std())
                                    else:
                                        print('No hay estadisticas para hacer el reporte.')
                            except Error as e:
                                print (e)
                            except:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            break
                    if opcion_stats == 4:
                        print('Volviendo al menu de Reportes y Consultas')
                        break
            if opcion_menu_4 == 4:
                print('Saliendo del menu...')
                print('\n')
                break
    if opcion_menuPrincipal == 4:
        print('Esta a punto de salir del sistema.')
        salir_Sistema = input('¿Está seguro que desea salir del sistema? (SI/NO): ').upper()
        if salir_Sistema == 'SI':
            print()
            print('Gracias por usar el sistema.')
            print('Vuelva pronto')
            print('Saliendo del sistema...')
            break
        elif salir_Sistema == 'NO':
            print('Volviendo al menú principal...')
            print()
        elif salir_Sistema == '':
            print('Error. No se puede omitir.')
        else:
            print('Por favor, responda con SI o NO.')
            continue