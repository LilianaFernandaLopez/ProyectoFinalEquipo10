#Web scraper de las librerías Sótano, Gandhi y Porrúa
#Paqueterías necesarias:
import pandas as pd
import pandasql as ps
import time
import numpy as np
import re
from selenium import webdriver
import matplotlib.pyplot as plt
import tkinter as tk

#Creamos el Dataframe donde se van a almacenar los datos 
aux = pd.DataFrame()
aux.to_excel('df_librerias.xlsx', index=False)

#Web scraper Porrúa
def Buscador_Precios_Selenium_Porrua(producto):
    ### ingresamos a la pagina web 
    path ='C:\webdriver3/chromedriver'
    driver=webdriver.Chrome(path)
    url= "https://porrua.mx/catalogsearch/result/?q="+producto
    driver.get(url)
    #La página de Porrua tarda en cargar, por lo que es necesario ponerle 
    #esta pausa para que recopile los datos
    time.sleep(5)
    productos= driver.find_elements_by_class_name('product-item-info')
    
    ### accedemos a las urls almacenadas en la variable productos
    lista_urls=[]
    for i in range(len(productos)):
        try:
            lista_urls.append(productos[i].find_element_by_tag_name("a").get_attribute("href"))
        except:
            lista_urls.append(np.nan)
    lista_urls
            
    ### accedemos a los nombres de los productos
    lista_nombres=[]
    for i in range(len(productos)):
        try:
            lista_nombres.append(productos[i].find_elements_by_tag_name("a")[1].text)
        except:
            lista_nombres.append(np.nan)
    lista_nombres
            
    #En el caso de Porrúa no se muestran los autores, por lo que siempre 
    #aplicará la excepción
    lista_autores = []
    for i in range(len(productos)):
        try:
            lista_autores.append(productos[i].find_element_by_class_name('autor').text)
        except:
            lista_autores.append(np.nan)
            
    ### accedemos a los precios base y promo de los productos 
    precios_actuales=[]
    precios_antes=[]
    for i in range(len(productos)):
        try:
            precios_actuales.append(productos[i].find_elements_by_class_name("price")[0].text)
        except:
            precios_actuales.append(np.nan)
        try:
            precios_antes.append(productos[i].find_elements_by_class_name("old-price")[0].text)
        except:
            precios_antes.append(np.nan)
    precios_actuales

    ###Ponemos todos los datos en el Dataframe de las librerías 
    ###y lo convertimos en un excel
    df_librerias = pd.DataFrame({'Título':lista_nombres, 'Autor':lista_autores, 'url':lista_urls, 'Precio':precios_actuales, 'Promocion':precios_antes})
    df_librerias['Librería'] = 'Porrúa Librerías'
    df_librerias['Categoría'] = producto
    df_librerias['Fecha'] = time.strftime('%d/%m/%y')
    df_librerias = df_librerias[['Fecha', 'Librería', 'Título', 'Autor', 'Precio', 'Promocion', 'Categoría', 'url']]      
    datos_webscraper = pd.read_excel('df_librerias.xlsx')
    datos_webscraper = pd.concat([datos_webscraper, df_librerias], axis=0)
    datos_webscraper.to_excel('df_librerias.xlsx', index = False)            
    
    ###Salimos del navegador que abrimos
    driver.quit
    
    return df_librerias


#Web scraper de Sótano
def Buscador_Precios_Selenium_Sotano(producto):
    path = 'C:\webdriver3/chromedriver'
    driver = webdriver.Chrome(path)
    url = 'https://www.elsotano.com/busqueda/listaLibros.php?tipoBus=full&tipoArticulo=&palabrasBusqueda='+producto
    driver.get(url)  
    
    productos = driver.find_elements_by_class_name('item')
    
    lista_urls = []
    for i in range(len(productos)):
        try:
            lista_urls.append(productos[i].find_element_by_tag_name('a').get_attribute('href'))
        except:
            lista_urls.append(np.nan)

    lista_nombres = []
    for i in range(len(productos)):
        try:
            lista_nombres.append(productos[i].find_element_by_class_name('so-booktitle').text)
        except:
            lista_nombres.append(np.nan)
            
    lista_autores = []
    for i in range(len(productos)):
        try:
            lista_autores.append(productos[i].find_element_by_class_name('so-bookwriter').text)
        except:
            lista_autores.append(np.nan)
            
    lista_precios = []
    lista_promociones = []
    
    for i in range(len(productos)):
        prod = productos[i].text.split('\n')
        producto_aux = prod[-1].split(' ')
        if len(producto_aux)>1:
            try:
                precio_promo = producto_aux[0]
                precio_normal = producto_aux[1]
                lista_promociones.append(precio_promo)
                lista_precios.append(precio_normal)
            except: 
                precio_normal = producto_aux[0]
                lista_precios.append(precio_normal)
        
        else: 
            precio_normal = producto_aux[0]
            lista_precios.append(precio_normal)
            lista_promociones.append(np.nan)
            
    df_librerias = pd.DataFrame({'Título':lista_nombres, 'Autor':lista_autores, 'url':lista_urls, 'Precio':lista_precios, 'Promocion':lista_promociones})
    df_librerias['Librería'] = 'Sótano Librerías'
    df_librerias['Categoría'] = producto
    df_librerias['Fecha'] = time.strftime('%d/%m/%y')
    df_librerias = df_librerias[['Fecha', 'Librería', 'Título', 'Autor', 'Precio', 'Promocion', 'Categoría', 'url']]      
    datos_webscraper = pd.read_excel('df_librerias.xlsx')
    datos_webscraper = pd.concat([datos_webscraper, df_librerias], axis=0)
    datos_webscraper.to_excel('df_librerias.xlsx', index = False)            
    
    driver.quit
    
    return df_librerias


#Web scraper de Gandhi
def Buscador_Precios_Selenium_Gandhi(producto):
    path = 'C:\webdriver3/chromedriver'
    driver = webdriver.Chrome(path)
    url= "https://www.gandhi.com.mx/catalogsearch/result/?q="+producto+"+"
    driver.get(url)
    
    productos = driver.find_elements_by_class_name("product-item-info")
    
    lista_urls = []
    for i in range(len(productos)):
        try:
            lista_urls.append(productos[i].find_element_by_tag_name('a').get_attribute('href'))
        except:
            lista_urls.append(np.nan)
            
    lista_nombres = []
    for i in range(len(productos)):
        try:
            lista_nombres.append(productos[i].find_element_by_class_name('product name product-item-name').text)
        except:
            lista_nombres.append(np.nan)
            
    
    lista_autores = []
    for i in range(len(productos)):
        try:
            lista_autores.append(productos[i].find_element_by_class_name('autor').text)
        except:
            lista_autores.append(np.nan)
            
    lista_precios = []
    
    for i in range(len(productos)):
        try:
            lista_precios.append(productos[i].find_element_by_class_name('price').text)
        except:
            lista_precios.append(np.nan)
            
    lista_promociones = []       
    for i in range(len(productos)):
        try:
            lista_promociones.append(productos[i].find_element_by_class_name('special-price').text)
        except:
            lista_promociones.append(np.nan)
            
    df_librerias = pd.DataFrame({'Título':lista_nombres, 'Autor':lista_autores, 'url':lista_urls, 'Precio':lista_precios, 'Promocion':lista_promociones})
    df_librerias['Librería'] = 'Gandhi Librerías'
    df_librerias['Categoría'] = producto
    df_librerias['Fecha'] = time.strftime('%d/%m/%y')
    df_librerias = df_librerias[['Fecha', 'Librería', 'Título', 'Autor', 'Precio', 'Promocion', 'Categoría', 'url']]      
    datos_webscraper = pd.read_excel('df_librerias.xlsx')
    datos_webscraper = pd.concat([datos_webscraper, df_librerias], axis=0)
    datos_webscraper.to_excel('df_librerias.xlsx', index = False)            
    
    driver.quit
    
    return df_librerias


#Corremos los web scrapers para los productos que buscamos
for productos in ['matematicas', 'novela', 'filosofia']:
    Buscador_Precios_Selenium_Sotano(productos)
    Buscador_Precios_Selenium_Gandhi(productos)
    Buscador_Precios_Selenium_Porrua(productos) 


###Vemos el Excel que hemos hecho
df_librerias=pd.read_excel("df_librerias.xlsx")
df_librerias

###Convertimos los precios en floats para poder operar con ellos
def precios_floats(datos):

        
    #### eliminamos el signo de pesos de ambas columnas
    
    for i in range(len(datos['Precio'])):
        try:
            datos['Precio'].iloc[i]=datos['Precio'].iloc[i].strip('$')
        except:
            pass
        
    for i in range(len(datos['Promocion'])):
        try:
            datos['Promocion'].iloc[i]=datos['Promocion'].iloc[i].strip('$')
        except:
            pass
        
    
    ### quitamos la separacion de comas para miles
    
    datos['Precio']=datos['Precio'].replace(',','',regex=True)
    datos['Promocion']=datos['Promocion'].replace(',','',regex=True)
    
     
        
    ### convertimos los precios a numericos    
    datos['Precio'] = pd.to_numeric(datos['Precio'], errors='coerce')
    datos['Promocion'] = pd.to_numeric(datos['Promocion'], errors='coerce')

    
    
    datos.to_excel('df_librerias_limpio.xlsx',index=False)
        
     ### visualizamos los tipos de datos
    print(datos.dtypes)
    return datos

###Visualizamos el nuevo Dataframe con los valores float
precios_floats(df_librerias)

###Leemos el Excel con los valores floats
df_librerias=pd.read_excel("df_librerias_limpio.xlsx")
df_librerias

###Comenzaremos a hacer las consultas en SQL con la base de datos creada
#Tomamos los libros que tengan promoción
ps.sqldf("select * from df_librerias where Promocion is not null order by Promocion DESC")

#Buscamos los libros que tengan 'Volpi' en el autor
ps.sqldf("select * from df_librerias where Autor like '%Volpi%'")

#Tomamos el libro con el precio más alto
ps.sqldf("select MAX (Precio), Categoría, Título, Autor from df_librerias")

#Elegimos los libros que cuesten entre 300 y 400 pesos
ps.sqldf("select * from df_librerias where Precio between 300 and 400")

#Seleccionamos los libros con título 'El Alquimista'
ps.sqldf("select * from df_librerias where (Título = 'El Alquimista')")

#Ordena los libros por título
ps.sqldf("select * from df_librerias order by Título")

#Junta los libros que tengan el mismo precio, sin importar la librería
ps.sqldf("select * from df_librerias group by Precio")

#Cuenta el número de libros que tiene cada autor en el Dataframe
df = ps.sqldf("SELECT COUNT(url), Precio, Autor FROM df_librerias GROUP BY Autor order by COUNT(url)")
df = df.dropna(how='any',axis=0) 
df

###Comenzaremos ahora a poner las gráficas 
#Definimos el estilo para los títulos y labels
font1 = {'family':'serif','color':'sienna','size':17}
font2 = {'family':'serif','color':'dimgrey','size':13}

#Gráfica 1
df_promo = ps.sqldf("select * from df_librerias where Promocion is not null order by Precio DESC limit 10")
df_promo
z = df_promo.Promocion
a = df_promo.Categoría
plt.pie(z,labels=z)
plt.title('Promociones de los 10 libros más caros', fontdict = font1)
plt.show()

#Gráfica 2
df_titulo = ps.sqldf("select * from df_librerias group by Librería, Precio order by Precio")
df_titulo
x = df_titulo.Librería
y = df_titulo.Precio
plt.bar(x, y, color = ['olivedrab'])
plt.xlabel('Librería', fontdict = font2)
plt.ylabel('Precio', fontdict = font2)
plt.title('Precio del libro más caro de la librería', fontdict = font1)
plt.show()


#Gráfica 3
df_titulo = ps.sqldf("select * from df_librerias group by Librería, Precio order by Precio")
df_titulo
x = df_titulo.Categoría
y = df_titulo.Precio
plt.bar(x, y, color=['goldenrod'])
plt.xlabel('Categoría', fontdict = font2)
plt.ylabel('Precio', fontdict = font2)
plt.title('Precio del libro más caro por categoría', fontdict = font1)
plt.show()

#Gráfica 4
df_1 = ps.sqldf("SELECT COUNT(url), Precio, Autor FROM df_librerias GROUP BY Autor order by COUNT(url)") 
s = ps.sqldf("SELECT COUNT(url) as Numero_urls, Autor FROM df_librerias WHERE Librería = 'Sótano Librerías' GROUP BY Autor order by COUNT(url)")
s = s.dropna(how='any',axis=0) 
s.drop(s[s['Numero_urls'] < 2].index, inplace = True)
g = ps.sqldf("SELECT COUNT(url) as Numero_urls, Precio, Autor FROM df_librerias WHERE Librería = 'Gandhi Librerías' GROUP BY Autor order by COUNT(url)")
g = g.dropna(how='any',axis=0) 
g.drop(g[g['Numero_urls'] < 2].index, inplace = True)
s_x1 = s.Autor
s_y1 = s.Numero_urls
g_x1 = g.Autor
g_y1 = g.Numero_urls
plt.barh(s_x1, s_y1, color = 'royalblue')
plt.barh(g_x1, g_y1, color = 'gold')
plt.title('Número de libros por autor', fontdict = font1)
plt.xlabel('Número de libros', fontdict = font2)
plt.ylabel('Autor', fontdict = font2)
plt.legend(["Sótano Librerías", "Gandhi Librerías"])
plt.show()

