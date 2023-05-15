## LIBRARIES
# ---

# Selenium
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains

# For scraping
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

# Options driver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select

# Dataframes
import pandas as pd
import itertools
import os

# Simulating human behavior
import time
from time import sleep
import random

# Clear data
import unidecode

# Json files
import json
import re
import numpy as np
import itertools
from pandas import json_normalize

# To use explicit waits
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Download files
import urllib.request
import requests
from openpyxl import Workbook



## FUNCTION
# ---------

def scraper_contraloria( anios, tipo_servicio, grupo ):

    try:
        service = Service( ChromeDriverManager().install( ) )
        driver = webdriver.Chrome( service = service )
        driver.maximize_window()

        url = f'https://appbp.contraloria.gob.pe/BuscadorCGR/Informes/Avanzado.html'
        driver.get( url )
        
        wait = WebDriverWait( driver, 60 )
              
        # Seleccionar años: 
        try:    
            time.sleep( 2 )
            busqueda_periodo = wait.until( EC.element_to_be_clickable( ( By.XPATH, '//*[@id="aPeriodo"]' ) ) ).click()
 
            #creo un objeto con la tabla donde estarán los años
            tabla_anios = driver.find_element( By.XPATH, '//*[@id="lblmenuanioconclusion"]' )

            try: 
                for anio in anios:
                    time.sleep( 2 )
                    tabla_anios.location_once_scrolled_into_view
                    tabla_anios.find_element( By.XPATH, f".//label[contains(., { anio })]").click() 
                    print( anio )
            except:
                ( '\nno año\n' )                    

        except:
            print( '\naños no encontrados\n' )
            
        # Seleccionar municipalidades
        try:       
            time.sleep(2)
            busqueda_sector = wait.until( EC.element_to_be_clickable( ( By.XPATH, '//*[@id="aSector"]' ) ) ).click()

            # tabla donde estan los sectores
            tabla_sectores = driver.find_element( By.XPATH, '//*[@id="lblmenusector"]' )

            sectores = ['"MUNICIPALIDADES DISTRITALES"', '"MUNICIPALIDADES PROVINCIALES"']
            
            try:
                for sec in sectores:
                    time.sleep(2)
                    tabla_sectores.location_once_scrolled_into_view
                    tabla_sectores.find_element( By.XPATH, f".//label[contains(., { sec } ) ]").click()
                    print( sec )
            except:
                print( '\nno municipios\n' )
    
        except:
            print( '\nmunicipios no encontrados\n' )
            
        
        # Seleccionar Tipo de servicio
        try: 
            time.sleep( 2 )
            busqueda_tipo = wait.until( EC.element_to_be_clickable( ( By.XPATH, '//*[@id="aTipoServicio"]' ) ) ).\
                            click()

            # Tabla donde estan los tipos
            tabla_tipo = driver.find_element( By.XPATH, '//*[@id="lblmenuservicio"]' )
            tabla_tipo.find_element( By.XPATH, f".//label[contains(., { tipo_servicio })]").\
                       click()
            print( tipo_servicio )
            
        except:
            print( '\ntipos no encontrados\n' )       
                        
       # Creamos Carpeta 'scraper_contraloria'.
        try:
            os.mkdir('scraper_contraloria')
        except:
            pass
        
        
        # Creamos listas vacias para trabajar los loops.

        regions                                        =      []
        modalidad_de_servicio                          =      []
        num_de_inf                                     =      []
        entidad                                        =      []
        titulo_del_informe                             =      []
        evento                                         =      []
        operativo                                      =      []
        n_de_p_c_p_r                                   =      []
        tipo_de_responsabilidad                        =      []
        fecha_de_emision                               =      []
        fecha_de_conclusion                            =      []
        fecha_de_publicacion                           =      []
        link_de_ficha_de_resumen                       =      []
        link_de_informe                                =      []
        
        
        # Obtenemos el número de páginas para iterar
        try: 
            time.sleep( 2 )
            n_paginas = driver.find_element( By.XPATH, '//*[@id="lbltotalItems"]' ).\
                               text
            n_paginas = int( n_paginas.split( ' ' )[ - 1 ] )
            print( f'N. paginas: { n_paginas }' )
        except:
            print( 'No N. paginas' )  
            
        # Procesar todas las páginas
        try: 
            for i in range( 1, n_paginas + 1 ):
                print( f"Procesando página { i }..." )
                
                # Esperar hasta que la página se haya cargado completamente
                wait.until(EC.presence_of_element_located( ( By.XPATH, '//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[1]/td[1]' ) ) )

                # Obtener el número de filas en la página actual
                n_filas = len( driver.find_elements( By.XPATH, '//*[@id="tablaResultadosUltimosInformes"]/tbody/tr' ) )

                # Procesar todos los elementos de la página actual
                for x in range( 1, n_filas + 1 ):

                    # Esperar hasta que el elemento específico se haya cargado completamente
                    wait.until(EC.presence_of_element_located( ( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[1]') ) )
                    
                    reg           = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[1]' ).text ]   
                    mod_de_serv2  = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[2]' ).text ]
                    num_d_inf2    = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[3]' ).text ]
                    ent           = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[4]' ).text ] 
                    tit_d_i       = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[5]' ).get_attribute( 'textContent' ) ]              
                    ev            = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[6]' ).get_attribute( 'textContent' ) ]
                    op            = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[7]' ).get_attribute( 'textContent' ) ]
                    n_d_p_c_p_r   = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[8]' ).get_attribute( 'textContent' ) ]
                    t_d_r         = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[9]' ).get_attribute( 'textContent' ) if driver.\
                                             find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[9]' ).get_attribute( 'textContent' ).strip() else "None" ]
                    f_d_e         = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[10]' ).get_attribute( 'textContent' ) ]
                    f_d_c         = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[11]' ).get_attribute( 'textContent' ) ]
                    f_d_p         = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[12]' ).get_attribute( 'textContent' ) ]
                    link_d_f_d_r2 = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[13]' ).find_element( By.TAG_NAME, 'a' ).get_attribute( 'href' ) ]
                    link_d_i2     = [ driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[14]' ).find_element( By.TAG_NAME, 'a' ).get_attribute( 'href' ) ]

                    mod_de_serv   =   driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[2]' ).text 
                    num_d_inf     =   driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[3]' ).text               
                    link_d_f_d_r  =   driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[13]' ).find_element( By.TAG_NAME, 'a' ).get_attribute( 'href' )
                    link_d_i      =   driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[14]' ).find_element( By.TAG_NAME, 'a' ).get_attribute( 'href' )

                    
                
                    # Llenamos las listas vacias de arriba con datos por iteración.
                    regions                                        +=     reg 
                    modalidad_de_servicio                          +=     mod_de_serv2
                    num_de_inf                                     +=     num_d_inf2
                    entidad                                        +=     ent
                    titulo_del_informe                             +=     tit_d_i
                    evento                                         +=     ev 
                    operativo                                      +=     op 
                    n_de_p_c_p_r                                   +=     n_d_p_c_p_r 
                    tipo_de_responsabilidad                        +=     t_d_r 
                    fecha_de_emision                               +=     f_d_e 
                    fecha_de_conclusion                            +=     f_d_c 
                    fecha_de_publicacion                           +=     f_d_p 
                    link_de_ficha_de_resumen                       +=     link_d_f_d_r2
                    link_de_informe                                +=     link_d_i2                            
                    
                    
                    # Modificar los nobmres y números para quitar elementos extraños
                    try: 
                        mod_de_serv = mod_de_serv.strip().replace( ' ', '_' ).replace( '"', '' )
                        mod_de_serv = unidecode.unidecode( mod_de_serv )
                        num_d_inf   = num_d_inf.strip().replace( '/', '_' ).replace( '\\', '_' )
                        # print( f"Modalidad de servicio: { mod_de_serv }" )
                        # print( f"Número de Informe: { num_d_inf }" )
                        
                    except:
                        print( "No modificaciones" )        
                        

                    # Creamos las carpetas donde se ubicarán los pdf.

                    try:
                        tipo_previo = '"SERVICIO CONTROL PREVIO"'
                        tipo_simultaneo = '"SERVICIO CONTROL SIMULTANEO"'
                        tipo_posterior = '"SERVICIO CONTROL POSTERIOR"'
                        
                        modalidades_posterior = [ '"ACCION OFICIO POSTERIOR"', '"AUDITORIA CUMPLIMIENTO"', '"AUDITORIA DESEMPEÑO"', 
                                                  '"AUDITORIA FINANCIERA"',  '"SERVICIO DE CONTROL ESPECÍFICO A HECHOS CON PRESUNTA IRREGULARIDAD"' ]
                        modalidades_simultaneo = [ '"ACCIÓN SIMULTÁNEA"', '"CONTROL CONCURRENTE"', '"ORIENTACIÓN DE OFICIO"',
                                                   '"REPORTE DE AVANCE"', '"VISITA DE CONTROL"', '"VISITA PREVENTIVA"' ]
                        modalidades_previo = [ '"ASOCIACIÓN PÚBLICO PRIVADA"', '"ENDEUDAMIENTO INTERNO O EXTERNO"', '"OBRAS POR IMPUESTOS"',
                                                '"PRESTACIONES DE ADICIONALES DE OBRA"', '"PRESTACIONES DE ADICIONALES DE SUPERVISIÓN"' ]

                        
                        if tipo_servicio == '"SERVICIO CONTROL POSTERIOR"':
                            tipo = tipo_posterior
                            tipo = tipo.replace( ' ', '_' ).replace( '"', '' )
                            for modalidad in modalidades_posterior:
                                modalidad = modalidad.replace('"', '').replace( ' ', '_' )
                                modalidad = unidecode.unidecode( modalidad )
                                if modalidad == mod_de_serv:
                                    try:
                                        folder_path = os.path.join( 'scraper_contraloria', tipo, grupo, modalidad, num_d_inf )
                                        os.makedirs( folder_path, exist_ok = True )

                                        informe_path = os.path.join( folder_path, f'{ num_d_inf }-informe.pdf' )
                                        resumen_path = os.path.join( folder_path, f'{ num_d_inf }-resumen.pdf' )

                                        try:
                                            response_informe = requests.get( link_d_i )
                                            with open( informe_path, 'wb') as informe_file:
                                                informe_file.write( response_informe.content )
                                            print(f'\n{ num_d_inf }-informe.pdf' )

                                            response_resumen = requests.get( link_d_f_d_r )
                                            with open( resumen_path, 'wb' ) as resumen_file:
                                                resumen_file.write( response_resumen.content )
                                            print( f'{ num_d_inf }-resumen.pdf\n' )                                       
                                        except:
                                            print( f'No folder2-resumen.pdf' )                                               
                                    except:
                                        pass
                                        print( '\n no subdirectorios por numero \n' )

                        if tipo_servicio == '"SERVICIO CONTROL SIMULTANEO"':
                            tipo = tipo_simultaneo
                            tipo = tipo.replace( ' ', '_' ).replace( '"', '' )
                            for modalidad in modalidades_simultaneo:   
                                modalidad = modalidad.replace('"', '').replace( ' ', '_' )
                                modalidad = unidecode.unidecode( modalidad )                                
                                if modalidad == mod_de_serv:
                                    try:
                                        folder_path = os.path.join( 'scraper_contraloria', tipo, grupo, modalidad, num_d_inf )
                                        os.makedirs( folder_path, exist_ok = True )

                                        informe_path = os.path.join( folder_path, f'{ num_d_inf }-informe.pdf' )
                                        resumen_path = os.path.join( folder_path, f'{ num_d_inf }-resumen.pdf' )

                                        try:
                                            response_informe = requests.get( link_d_i )
                                            with open( informe_path, 'wb') as informe_file:
                                                informe_file.write( response_informe.content )
                                            print(f'\n{ num_d_inf }-informe.pdf' )

                                            response_resumen = requests.get( link_d_f_d_r )
                                            with open( resumen_path, 'wb' ) as resumen_file:
                                                resumen_file.write( response_resumen.content )
                                            print( f'{ num_d_inf }-resumen.pdf\n' )                                       
                                        except:
                                            print( f'No folder2-resumen.pdf' )                                                 
                                    except:
                                        pass
                                        print( '\n no subdirectorios por numero \n' )

                        if tipo_servicio == '"SERVICIO CONTROL PREVIO"':
                            tipo = tipo_previo
                            tipo = tipo.replace( ' ', '_' ).replace( '"', '' )
                            for modalidad in modalidades_previo:
                                modalidad = modalidad.replace('"', '').replace( ' ', '_' )
                                modalidad = unidecode.unidecode( modalidad )                                
                                if modalidad == mod_de_serv:
                                    try:
                                        folder_path = os.path.join( 'scraper_contraloria', tipo, grupo, modalidad, num_d_inf )
                                        os.makedirs( folder_path, exist_ok = True )

                                        informe_path = os.path.join( folder_path, f'{ num_d_inf }-informe.pdf' )
                                        resumen_path = os.path.join( folder_path, f'{ num_d_inf }-resumen.pdf' )

                                        try:
                                            response_informe = requests.get( link_d_i )
                                            with open( informe_path, 'wb') as informe_file:
                                                informe_file.write( response_informe.content )
                                            print(f'\n{ num_d_inf }-informe.pdf' )

                                            response_resumen = requests.get( link_d_f_d_r )
                                            with open( resumen_path, 'wb' ) as resumen_file:
                                                resumen_file.write( response_resumen.content )
                                            print( f'{ num_d_inf }-resumen.pdf\n' )                                       
                                        except:
                                            print( f'No folder2-resumen.pdf' )                                                
                                    except:
                                        pass
                                        print( '\n no subdirectorios por numero \n' )
                            
                    except:
                        print( '\n No subcarpetas por tipo de servicio \n' )                       
                        

                # Después de procesar todos los elementos de la página actual, avanzamos a la siguiente página
                try:
                    if i < n_paginas:
                        siguiente_pagina = driver.find_element( By.XPATH, '//*[@id="Li_Siguiente"]/a' )
                        actions = ActionChains( driver )
                        actions.move_to_element( siguiente_pagina ).click().perform()
                        wait.until( EC.staleness_of( driver.find_element( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[{ x }]/td[1]' ) ) )
                        wait.until( EC.presence_of_element_located( ( By.XPATH, f'//*[@id="tablaResultadosUltimosInformes"]/tbody/tr[1]/td[1]' ) ) )
                        print( f"SE AVANZA A LA PÁGINA { i + 1 }" )
                except:
                    print( "NO SE AVANZA A LA SIGUIENTE PÁGINA" )           
                  
        except:
            print( '0 documentos encontrados' )    
            
        # Guardar datos extraídos en un archivo Excel
        try: 
            
            tipo_servicio = tipo_servicio.replace( ' ', '_' ).replace( '"', '' )
                
            datos_extraidos = {
            'Region': regions,
            'Modalidad': modalidad_de_servicio,
            'Número de informe': num_de_inf,
            'Entidad': entidad,
            'Titulo del Informe': titulo_del_informe,
            'Evento': evento,
            'Operativo': operativo,
            'Numero de Detalle por CP/PR': n_de_p_c_p_r,
            'Tipo de Responsabilidad': tipo_de_responsabilidad,
            'Fecha de Emision': fecha_de_emision,
            'Fecha de Conclusion': fecha_de_conclusion,
            'Fecha de Publicacion': fecha_de_publicacion,
            'Enlace de Resumen': link_de_ficha_de_resumen,
            'Enlace de Informe': link_de_informe
            }
            
            de = pd.DataFrame( datos_extraidos )
            de.to_excel( f'scraper_contraloria/{ tipo_servicio }/{ tipo_servicio }_{ grupo }_informacion.xlsx' )
            print( "\nDATOS EXTRAIDOS EN EXCEL\n" )
        except:
            print( "\nNO DATOS EXTRAIDOS EN EXCEL\n" )          
                    
    except:
        driver.quit()
        print( '\nquit\n' )