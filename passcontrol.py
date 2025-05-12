'''análisis de registros en la aplicación pass gestión y notificación de tickets'''
# instalar requerimientos: pip install -r requirements.txt
# crear exe en carpeta dist:
  # python -m PyInstaller --onefile --noconsole passcontrol.py --version-file versionfile.txt
import os, sys, subprocess, io, base64, json
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge import service
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.schedulers import base as sched_base
from pystray import Icon as pyIcon, Menu as pyMenu, MenuItem as pyItem
import PIL.Image
import psutil
from passIcon import pass_icon

main_url = 'https://segsocial-smartit.onbmc.com/smartit/app/#/ticket-console'
last_ids = list() #se guardan los ids de los tickets anteriores
sched_seconds = 60 #intervalo scheduler
current_dir = os.getcwd() #directorio actual
tickets_solo_inss = True
sched = BackgroundScheduler() #cola de procesos
estadisticas = [] #distintos estados por empresa
jsonFile = "C:/temp/passcontrol.json" #archivo para guardar los settings del usuario

def get_default_options():
    '''activa las opciones por defecto'''
    os.environ.pop('HTTP_PROXY', None)
    os.environ.pop('HTTPS_PROXY', None)
    opt = webdriver.EdgeOptions() #Options()
    opt.use_chromium = True
    opt.add_experimental_option("detach", True)
    opt.binary_location = r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
    #options.add_experimental_option("excludeSwitches", ['enable-logging'])
    #options.ignore_local_proxy_environment_variables() #deprecated
    opt.add_argument("--no-sandbox")
    opt.add_argument("--disable-gpu")
    opt.add_argument('--disable-dev-shm-usage')
    #opt.add_argument("--user-data-dir=C:\\temp\\edge_data")
    #opt.add_argument("--log-level=3")
    opt.add_argument("--headless") #ventana oculta
    opt.add_argument("start-maximixed")
    return opt

#espera a que un elemento se encuentre cargado y lo devuelve
def wait_element(selector, patron):
    '''espera a que el elemento indicado se haya cargado'''
    wait = WebDriverWait(driver, timeout=20, poll_frequency=.2)
    wait.until(EC.presence_of_element_located((selector, patron)))
    return driver.find_element(selector, patron)

#captura las filas de items y guarda los valores
def get_items():
    '''captura los items y devuelve un array'''
    items = []
    #en caso de que no existan tickets, se producirá una excepcion
    try:
        print("entrando en get_items")
        viewPort = wait_element(By.CSS_SELECTOR, '.ngViewport .ng-scope')
        #vp_header = wait_element(By.CLASS_NAME, "ngHeaderScroller")
        vp_header = driver.find_element(By.CLASS_NAME, "ngHeaderScroller")
        #vp_header = wait_element(By.CLASS_NAME, "ngHeaderContainer")
        vp_header_items = vp_header.find_elements(By.CSS_SELECTOR, ".ngHeaderSortColumn > .ngHeaderText")
        print("vp_header_items count: %d" % len(vp_header_items))
        viewPort  = driver.find_element(By.CLASS_NAME, "ngViewport")
        vp_items = viewPort.find_elements(By.CSS_SELECTOR,'.ng-scope .ngRow')
        cabeceras = []
        #capturamos el nombre de las cabeceras y su posición
        for e in vp_header_items:
            hd_class = e.get_attribute("class")
            nombre_col = e.text.lower()
            indice_col  = [ x for x in hd_class.split(" ") if "colt" in x[0:4] ][0]
            cabeceras.append({ "nombre": nombre_col, "campo": indice_col })
        #print("cabeceras: %s" % cabeceras)
        contador = 0
        #filas de tickets- buscamos los campos segun el indice de las cabeceras
        for e in vp_items:
            contador += 1
            campos = {}
            for c in cabeceras:
                campos["fila"] = contador
                n_nombre = c["nombre"]
                match n_nombre:
                    case "mostrar id":
                        n_nombre = "Id"
                    case "fecha de creación":
                        n_nombre = "fecha"
                    case "nombre completo de cliente":
                        n_nombre = "nombre"
                    case "fecha de última modificación":
                        n_nombre = "fecha_mod"
                    case "usuario asignado":
                        n_nombre = "usu_asignado"
                campos[n_nombre] = e.find_element(By.CSS_SELECTOR, ".ng-scope .{}".format(c["campo"])).text.strip()
            #print("campos: %s" % campos)
            items.append(campos)
    except Exception as ex:
        print("excepcion en get_items: %s" % type(ex))
    get_estadisticas(items)
    return items

def get_estadisticas(items):
    '''carga la tabla estadisticas con los datos actuales'''
    global estadisticas
    dict_estadist = {}
    # items = get_items()
    dist_estados = set() #un conjunto no permite repetir items
    for i in items:
        #comprobamos si existe la empresa
        empr = i['empresa']
        estado = i['estado']
        dist_estados.add(estado)
        if empr not in dict_estadist.keys():
            dict_estadist[empr] = {}
            dict_estadist[empr][estado] = 1
        else:
            if estado not in dict_estadist[empr].keys():
                dict_estadist[empr][estado] = 1
            else:
                dict_estadist[empr][estado] = dict_estadist[empr][estado] + 1
    print("estadisticas: %s" % dict_estadist)
    estadisticas=[]
    hora = datetime.now()
    #estadisticas.append("Ult. comprobación: {}:{}:{}".format(hora.hour, hora.minute, hora.second))
    print("Ult. comprobación: {}:{}:{}".format(hora.hour, hora.minute, hora.second))
    for kit,vit in dict_estadist.items():
        tx_est = "{0:5s} - ".format(kit)
        for k,v in vit.items():
            tx_est += "{0:4s}:{1:2d} ".format(k[0:4], v)
        estadisticas.append(tx_est)
        print('est lineas: %s' % estadisticas)
    icon.update_menu() #actualizamos el menu con los nuevos items



#comprueba si aparece algún ticket nuevo y devuelve un array
#se devuelven los que estan en estado asignado pero sin usuario
def comprobar_tickets():
    '''comprueba si existen nuevos tickets'''
    global last_ids
    items = get_items()
    #print(items)
    ids = list(map(lambda i: i['Id'], items))
    print("numero de items: %d: %s" % (len(items), list(ids)))
    #new_ids = list(i for i in ids if i not in last_ids) #nuevos identificadores
    #nuevos tickets
    #new_tickets = list(x for x in items if x['Id'] in new_ids)
    #devolvemos los tickets estado="asignado" que no tengan usuario asignado
    active_tickets = list(x for x in items if (x['estado'] == "Asignado" and x['usu_asignado'] == ""))
    last_ids = ids[:]
    return active_tickets

def open_navigator():
    subprocess.Popen([options.binary_location, main_url])

def start_scheduler(seconds):
    '''arranca el scheduler con el intervalo indicado'''
    sched.add_job(main_loop, 'interval', seconds=seconds, id='job_id')
    sched.start()

# funciones para tray icon
def set_state_sched(sta):
    def inner(icon, item):
        global sched_seconds
        sched_seconds = sta
        print("sched_seconds: %d" % sta)
        sched.remove_job('job_id')
        sched.add_job(main_loop, 'interval', seconds=sta, id='job_id')
        icon.notify(title='Cambiado intervalo', message="nuevo valor: {} seg".format(sched_seconds))
        save_json() #guardamos los datos en el archivo json
    return inner

def get_state_sched(sta):
    def inner(item):
        return sched_seconds == sta
    return inner

def tray_sched(icon, item):
    global tickets_solo_inss
    print("tray_item: (%s)" % item)
    match item.text:
        case "Parar":
            print("sched state: %s" % sched.state)
            if sched.state != sched_base.STATE_RUNNING:
                print("arrancamos job")
                sched.resume()
                icon.notify(message="Estado actual: Funcionando")
            else:
                print("paramos scheduler")
                sched.pause()
                icon.notify(message="Estado actual: Pausado")

        case "Sólo INSS":
            tickets_solo_inss = not tickets_solo_inss
            print("tickets_solo_inss: %s" % tickets_solo_inss)
            if tickets_solo_inss:
                icon.notify(message="Mostrar únicamente tickets INSS")
            else:
                icon.notify(message="Mostrar todos tickets")
            save_json() #guardamos los cambios en el archivo json

def tray_quit(icon):
    '''paramos la cola, el driver de edge y cerramos el icono de systray'''
    if sched.running:
        print("quit: paramos scheduler")
        sched.shutdown()
    #driver.close()
    driver.quit() #preferible para liberar todos los recursos
    icon.stop()

def load_json():
    '''recuperamos el contenido de passcontrol.json'''
    global tickets_solo_inss, sched_seconds
    #si existe el archivo json leemos su contenido
    if os.path.exists(jsonFile):
        print("leer passcontrol.json")
        with open(jsonFile, "r", encoding='utf-8') as f:
            campos = json.loads(f.readline())
            tickets_solo_inss = campos['solo_INSS']
            sched_seconds = campos['intervalo']
            f.close()

def save_json():
    '''guardamos los valores actuales en el archivo passcontrol.json'''
    json_data = {'solo_INSS':tickets_solo_inss, 'intervalo': sched_seconds}
    with open (jsonFile, "w", encoding='utf-8') as f:
        f.write(json.dumps(json_data))
        f.close()
    print("datos guardados: %s" % json_data)

#bucle principal
def main_loop():
    '''se encarga de comprobar regularmente si hay que notificar nuevos tickets'''
    driver.refresh()
    #importante ajustar el zoom para que entren todas las columnas
    driver.execute_script("document.body.style.zoom='10%'")

    nuevos = comprobar_tickets()
    def ticket_formato(tit, mess):
        return "{}: {}\n".format(tit, mess)
    if len(nuevos) > 0:
        print("tickets nuevos: %s" % list(map(lambda i: i['Id'], nuevos)))
        rel = ""
        rel_inss = ""
        nuevos_inss = 0
        for nt in nuevos:
            rel += ticket_formato(nt['Id'], nt['remitente'])
            #comprobamos si pertenece al inss
            if (nt["empresa"] in ["INSS","GISS"]) or (nt["empresa"]=="SJSS" and nt["remitente"][2:3] == 'I'):
                nuevos_inss += 1
                rel_inss += ticket_formato(nt['Id'], nt['remitente'])
        print("tickets inss(%d): %s" % (nuevos_inss, rel_inss))
        #antes de notificar comprobamos que no estemos en la pantalla de bloqueo
        if check_process("LogonUI.exe") > 0:
            print("pantalla bloqueo detectada")
        else: #mostramos las notificaciones
            if tickets_solo_inss:
                if nuevos_inss > 0:
                    icon.notify(title='Nuevos tickets!', message=rel_inss)
            else:
                icon.notify(title='Nuevos tickets!', message=rel)
    else:
        print("no hay nuevos tickets")
    # driver.refresh()
    # #importante ajustar el zoom para que entren todas las columnas
    # driver.execute_script("document.body.style.zoom='20%'")

def check_process(nomproc):
    '''devolver el numero de instancias del proceso indicado'''
    count = 0
    for proc in psutil.process_iter():
        if proc.name() == nomproc:
            count += 1
    return count

def gen_stat_items():
    return  (pyItem( '%s' % e, action=None)
            for e in estadisticas)

if __name__ == "__main__":
    #comprobamos si el programa ya está corriendo
    print("instancias de passcontrol: %d" % check_process("passcontrol.exe"))
    #cargamos los datos de usuario del archivo json
    load_json()
    #if check_run_program():
    if check_process("passcontrol.exe") > 2:
        print("passcontrol ya está ejecutándose")
        sys.exit(0)
    #creamos la imagen desde el codigo base64
    buffer = io.BytesIO(base64.b64decode(pass_icon))
    image = PIL.Image.open(buffer)

    icon = pyIcon('pass menu', image, 'passControl', menu=pyMenu(
        pyItem('Parar', tray_sched, checked=lambda item: sched.state != sched_base.STATE_RUNNING),
        pyItem('Sólo INSS', tray_sched, checked=lambda item: tickets_solo_inss),
        # pyItem('Estadísticas', pyMenu( lambda: (
        #     pyItem( '%s' % e, action=None)
        #     for e in estadisticas),
        # )),
        pyItem('Estadísticas', pyMenu(gen_stat_items)),
        pyItem('Tiempo(seg)', pyMenu( lambda: (
            pyItem(
                '%d' % i,
                set_state_sched(i),
                checked=get_state_sched(i),
                radio=True)
            for i in [30,60,120,300]),
        )),
        pyItem('Salir', tray_quit)
    ))

    #abrimos ventana url principal
    options = get_default_options()
    serv = service.Service(current_dir + '\\msedgedriver.exe')
    driver = webdriver.Edge(options=options, service=serv)
    #driver = webdriver.Edge(options=options)
    driver.implicitly_wait(15) #tiempo de espera implicito
    driver.get(main_url)
    #cambiamos el zoom para que entren todas las columnas en el viewport
    #driver.execute_script("document.body.style.zoom='20%'")

    print("titulo ventana principal: %s" % driver.title)
    main_loop() #primera ejecución al arrancar
    #act_estadisticas()

    #arrancamos la cola
    #sched = BackgroundScheduler()
    start_scheduler(sched_seconds)

    #arranca el loop principal
    icon.run()
