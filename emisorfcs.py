import os
import tkinter
from tkinter import filedialog
from tkinter import simpledialog
import csv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium import webdriver
import autoit
import posiciones_csv


def create_driver():
    driver = webdriver.Chrome()
    return driver

def login(cuit,pwd, driver):
    
    driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")

    username = WebDriverWait(driver, 10000).until(
        EC.presence_of_element_located((By.ID, "F1:username"))
    )

    username.clear()
    print("El cuit a poner es: " + cuit)
    username.send_keys(cuit)

    siguiente = WebDriverWait(driver, 10000).until(
        EC.presence_of_element_located((By.ID, "F1:btnSiguiente"))
    )
    siguiente.click()

    password = WebDriverWait(driver, 10000).until(
        EC.presence_of_element_located((By.ID, "F1:password"))
    ) 

    print("Las pass a poner: " + pwd)

    password.send_keys(pwd)

    ingresar = WebDriverWait(driver, 10000).until(
        EC.presence_of_element_located((By.ID, "F1:btnIngresar"))
    )
    ingresar.click()
    time.sleep(2)

def boton_si_no(var):
    si_no = tkinter.Toplevel()
    si_no.geometry("200x100")
    si = tkinter.Button(master=si_no, text="Si", width=5, command=lambda: var.set(True))
    no = tkinter.Button(master=si_no, text="No", width=5, command= lambda: var.set(0))
    label = tkinter.Label(master=si_no, text="Presione 'Sí' si se generó bien la factura, 'No' en caso contrario.", wraplength=200)
    label.pack(side="top")
    si.pack(side="left", padx=25)
    no.pack(side="right", padx=25)

    si.wait_variable(var)
    si_no.destroy()

def emitir_fc(Factura, driver):

    #Entro en comprobantes en línea
    cmpEnLinea = driver.find_elements_by_xpath('//*[@id="servicesContainer"]/div[8]/div/div/div/div[2]/h4')
    while not cmpEnLinea:
        cmpEnLinea = driver.find_elements_by_xpath('//*[@id="servicesContainer"]/div[8]/div/div/div/div[2]/h4')
    cmpEnLinea[0].click()

    time.sleep(2)

    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    time.sleep(1)


    btnEmpresa = driver.find_element_by_xpath("//input[@value='" + Factura[posiciones_csv.POS_EMPRESA] + "']")
    btnEmpresa.click()

    #Apreto emitir
    btnEmitir = driver.find_elements_by_xpath('//*[@id="btn_gen_cmp"]/span[2]')
    while not btnEmitir:
        btnEmitir = driver.find_elements_by_xpath('//*[@id="btn_gen_cmp"]/span[2]')

    btnEmitir[0].click()

    #Punto de Venta
    select_pto_venta = Select(driver.find_element_by_id("puntodeventa"))
    puntos_de_venta = select_pto_venta.options
    for pto_venta in puntos_de_venta:
        if " " + Factura[posiciones_csv.POS_PTO_DE_VENTA].zfill(5) == pto_venta.text[0: 6: 1]:
            boton_pto_venta = select_pto_venta.select_by_visible_text(pto_venta.text)

    time.sleep(1)

    #Selecciono tipo de comprobante
    select_comprobante = Select(driver.find_element_by_id("universocomprobante"))
    select_comprobante.select_by_visible_text(Factura[posiciones_csv.POS_TIPOFC])

    #Sigo post-pto de venta
    btnContinuar = driver.find_elements_by_xpath('//*[@id="contenido"]/form/input[2]')
    while not btnContinuar:
        btnContinuar = driver.find_elements_by_xpath('//*[@id="contenido"]/form/input[2]')

    btnContinuar[0].click()

    #Fecha emisión

    btnFecha = driver.find_elements_by_xpath('//*[@id="fc"]')
    while not btnFecha:
        btnFecha = driver.find_elements_by_xpath('//*[@id="fc"]')

    btnFecha[0].clear()
    btnFecha[0].send_keys(Factura[posiciones_csv.POS_FECHA_EMISION])

    #Pongo el servicio
    select_serv = Select(driver.find_element_by_id("idconcepto"))
    
    while not select_serv:
        select_serv = Select(driver.find_element_by_id("idconcepto"))

    select_serv.select_by_index(2)
    
    #Fecha Desde

    btnDesde = driver.find_elements_by_xpath('//*[@id="fsd"]')
    while not btnDesde:
        btnDesde = driver.find_elements_by_xpath('//*[@id="fsd"]')

    btnDesde[0].clear()
    btnDesde[0].send_keys(Factura[posiciones_csv.POS_FECHA_DESDE])

    #Fecha Hasta

    btnHasta = driver.find_elements_by_xpath('//*[@id="fsh"]')
    while not btnHasta:
        btnHasta = driver.find_elements_by_xpath('//*[@id="fsh"]')

    btnHasta[0].clear()
    btnHasta[0].send_keys(Factura[posiciones_csv.POS_FECHA_HASTA])

    #Fecha Vencimiento

    btnVto = driver.find_elements_by_xpath('//*[@id="vencimientopago"]')
    while not btnVto:
        btnVto = driver.find_elements_by_xpath('//*[@id="vencimientopago"]')

    btnVto[0].clear()
    btnVto[0].send_keys(Factura[posiciones_csv.POS_FECHA_VENC])

    autoit.win_activate("RCEL - Google Chrome") 

    btnContinuar = driver.find_elements_by_xpath('//*[@id="contenido"]/form/input[2]')
    while not btnContinuar:
        btnContinuar = driver.find_elements_by_xpath('//*[@id="contenido"]/form/input[2]')

    btnContinuar[0].click()

    #Condición IVA
    select_cond_iva = Select(driver.find_element_by_id('idivareceptor'))

    while not select_cond_iva:
        select_cond_iva = Select(driver.find_element_by_id('idivareceptor'))
    try:
        select_cond_iva.select_by_visible_text(" " + Factura[posiciones_csv.POS_CONDICION_IVA])
    except:
        select_cond_iva.select_by_visible_text(Factura[posiciones_csv.POS_CONDICION_IVA])

    if len(Factura[posiciones_csv.POS_CUIT_RECEPTOR]) < 11:
        #Es un DNI, no un CUIT
        btnCUIT = driver.find_elements_by_xpath('//*[@id="idtipodocreceptor"]')
        btnCUIT[0].click()
        ayuda_cuit = 0
        while ayuda_cuit != 6:
            autoit.send("{DOWN}")
            ayuda_cuit += 1
        
    btnCUIT = driver.find_elements_by_xpath('//*[@id="nrodocreceptor"]')
    while not btnCUIT:
        btnCUIT = driver.find_elements_by_xpath('//*[@id="nrodocreceptor"]')

    btnCUIT[0].clear()                 
    btnCUIT[0].send_keys(Factura[posiciones_csv.POS_CUIT_RECEPTOR])

    #Boton otro
    btnOtro = driver.find_elements_by_xpath('//*[@id="formadepago7"]')
    btnOtro[0].click()
    time.sleep(1)
    #Domicilio
    if int(Factura[posiciones_csv.POS_FORZAR_DOMICILIO]) == 1:
        btnDomicilio = driver.find_elements_by_xpath('//*[@id="domicilioreceptorcombo"]')
        if not btnDomicilio:
                btnDomicilio = driver.find_elements_by_xpath('//*[@id="domicilioreceptor"]')    
                btnDomicilio[0].clear()
                btnDomicilio[0].send_keys(Factura[posiciones_csv.POS_DOMICILIO])
                autoit.win_activate("RCEL - Google Chrome")
        else:
                select_domicilio = Select(driver.find_element_by_id("domicilioreceptorcombo"))
                select_domicilio.select_by_visible_text("Otro...") 
                btnDomicilio = driver.find_elements_by_xpath('//*[@id="domicilioreceptor"]')    

                btnDomicilio[0].clear()
                btnDomicilio[0].send_keys(Factura[posiciones_csv.POS_DOMICILIO])
                autoit.win_activate("RCEL - Google Chrome")
    else:
        btnDomicilio = driver.find_elements_by_xpath('//*[@id="domicilioreceptorcombo"]')
        if not btnDomicilio:
            print("No encontré el domicilio")
            btnDomicilio = driver.find_elements_by_xpath('//*[@id="domicilioreceptor"]')
            btnDomicilio[0].send_keys(Factura[posiciones_csv.POS_DOMICILIO])
            autoit.win_activate("RCEL - Google Chrome")
        else:
            pass


    btnContinuar = driver.find_elements_by_xpath('//*[@id="formulario"]/input[2]')
    while not btnContinuar:
        btnContinuar = driver.find_elements_by_xpath('//*[@id="formulario"]/input[2]')

    btnContinuar[0].click()

    #Pongo Concepto

    btnProducto = driver.find_elements_by_xpath('//*[@id="detalle_descripcion1"]')
    while not btnProducto:
        btnProducto = driver.find_elements_by_xpath('//*[@id="detalle_descripcion1"]')

    btnProducto[0].clear()
    btnProducto[0].send_keys(Factura[posiciones_csv.POS_CONCEPTO])

    autoit.win_activate("RCEL - Google Chrome")

    #Pongo la unidad 
    select_unidad = Select(driver.find_element_by_id('detalle_medida1'))
    select_unidad.select_by_value("98")

    #Pongo el precio 
    btnPrecio = driver.find_elements_by_xpath('//*[@id="detalle_precio1"]')

    btnPrecio[0].clear()
    btnPrecio[0].send_keys(Factura[posiciones_csv.POS_IMPORTE])

    #Alícuota IVA 
    btnIVA = driver.find_elements_by_xpath('//*[@id="detalle_tipo_iva1"]')

    if not btnIVA:
        pass
    else:
        select_iva = Select(driver.find_element_by_id("detalle_tipo_iva1"))
        select_iva.select_by_visible_text(" " + Factura[posiciones_csv.POS_ALICUOTA_IVA])


    btnContinuar = driver.find_elements_by_xpath('//*[@id="contenido"]/form/input[8]')
    try:
        btnContinuar[0].click()
    except: 
        btnContinuar = driver.find_elements_by_xpath('//*[@id="contenido"]/form/input[2]')
        btnContinuar[0].click()
    
    var = tkinter.BooleanVar(value=False)

    boton_si_no(var)

    if var.get() == True:

        btnConfirmar = driver.find_elements_by_xpath('//*[@id="btngenerar"]')
        btnConfirmar[0].click()

        time.sleep(1)
        autoit.win_activate("RCEL - Google Chrome")
        autoit.send("{ENTER}")

        time.sleep(3)   

        btnImprimir = driver.find_elements_by_xpath('//*[@id="botones_comprobante"]/input')

        btnImprimir[0].click()
    else:
        pass

    global ventana
    ventana.update()
    time.sleep(2)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])


def leerCSV(tempdir):
    facturas = []
    print(tempdir)
    with open(tempdir, 'r', encoding='utf-8') as file:
        reader = csv.reader(file, delimiter = ';')
        lineas = 0
        for row in reader:
            if lineas == 0:
                lineas += 1 
            else:
                facturas.append(row)

    return facturas

def emitir_fcs(Facturas, driver):
    ultimoCuitEmisor = 0
    i = 0
    for factura in Facturas:
        if ultimoCuitEmisor != int(Facturas[i][0]):
            print("Ultimo cuit: " + str(ultimoCuitEmisor) + "Nuevo: " + Facturas[i][0])
            ultimoCuitEmisor = int(Facturas[i][0])
            login(factura[0], factura[1], driver)
        i = i+1
        emitir_fc(factura, driver)

def leer_y_facturar(tempdir):
    facturas = leerCSV(tempdir)
    global driver
    driver = create_driver()
    emitir_fcs(facturas, driver)

def master_func():
    currdir = os.getcwd()
    global tempdir 
    tempdir = filedialog.askopenfilename(initialdir=currdir, title='Seleccione el CSV')
    leer_y_facturar(tempdir)
    
def main(): 
    global tempdir
    tempdir = ""
    global ventana
    ventana = tkinter.Tk()
    ventana.geometry("400x200")
    label = tkinter.Label(text="Seleccione el CSV que contenga la información de las facturas a emitir")
    label.grid(column=2, row=1)
    boton = tkinter.Button(text="Buscar", command = master_func)
    boton.grid(column=2, row=0, columnspan=5)
    ventana.mainloop()


if __name__ == '__main__':
    main()