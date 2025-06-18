'''análisis de registros en la aplicación pass gestión y notificación de tickets'''
# instalar requerimientos: pip install -r requirements.txt
# crear exe en carpeta dist:
# python -m PyInstaller --onefile --noconsole passcontrol.py --version-file versionfile.txt
import os, sys, io, base64, json
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
from comun import pass_icon, main_url, lista_ambitos
from abrir_edge import abrir_edge

#main_url = 'https://segsocial-smartit.onbmc.com/smartit/app/#/ticket-console'
last_ids = list() #se guardan los ids de los tickets anteriores
sched_seconds = 60 #intervalo scheduler
current_dir = os.getcwd() #directorio actual
sched = BackgroundScheduler() #cola de procesos
estadisticas = [] #distintos estados por empresa
ambito = [] #relación de ambitos
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
    estadisticas=[]
    #comprobamos si no hay incidencias
    if len(items) == 0:
        estadisticas.append("sin incidencias")
        icon.update_menu()
        return

    for i in items:
        #comprobamos si existe la empresa
        empr = i['empresa']
        estado = i['estado']
        if empr not in dict_estadist.keys():
            dict_estadist[empr] = {}
            dict_estadist[empr][estado] = 1
        else:
            if estado not in dict_estadist[empr].keys():
                dict_estadist[empr][estado] = 1
            else:
                dict_estadist[empr][estado] += 1 #incrementamos si la key existe
    print("estadisticas: %s" % dict_estadist)

    hora = datetime.now()
    print("Ult. comprobación: {0}:{1:02}:{2:02}".format(hora.hour, hora.minute, hora.second))
    for kit,vit in dict_estadist.items():
        tx_est = "{0:5s} - ".format(kit)
        for k,v in vit.items():
            tx_est += "{0:4s}:{1:2d} ".format(k[0:4], v)
        estadisticas.append(tx_est)
    icon.update_menu() #actualizamos el menu con los nuevos items



#comprueba si aparece algún ticket nuevo y devuelve un array
#se devuelven los que estan en estado asignado pero sin usuario
def comprobar_tickets():
    '''comprueba si existen nuevos tickets'''
    global last_ids
    items = get_items()
    ids = list(map(lambda i: i['Id'], items))
    print("numero de items: %d: %s" % (len(items), list(ids)))
    active_tickets = list(x for x in items if (x['estado'] == "Asignado" and x['usu_asignado'] == ""))
    last_ids = ids[:]
    return active_tickets

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

def set_ambito(amb):
    def inner(icon, item):
        #global ambito
        if amb not in ambito:
            ambito.append(amb)
        else:
            ambito.remove(amb)
        save_json()
        print(f"set_ambito: {ambito}")
        icon.notify(title='Ambito actualizado', message="Ambito: {}".format(ambito))
    return inner

def get_ambito(amb):
    def inner(item):
        return amb in ambito
    return inner


def tray_sched(icon, item):
    # global tickets_solo_inss
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
        case 'Abrir navegador Edge':
            abrir_edge(options.binary_location)

def tray_quit(icon):
    '''paramos la cola, el driver de edge y cerramos el icono de systray'''
    if sched.running:
        print("quit: paramos scheduler")
        sched.shutdown()
    #driver.close()
    driver.quit() #preferible para liberar todos los recursos
    icon.stop()

#gestión archivo json
def load_json():
    '''recuperamos el contenido de passcontrol.json'''
    global sched_seconds, ambito
    #si existe el archivo json leemos su contenido
    if os.path.exists(jsonFile):
        print("leer passcontrol.json")
        with open(jsonFile, "r", encoding='utf-8') as f:
            campos = json.loads(f.readline())
            #tickets_solo_inss = campos['solo_INSS']
            sched_seconds = campos['intervalo']
            ambito = campos['ambito']
            f.close()

def save_json():
    '''guardamos los valores actuales en el archivo passcontrol.json'''
    json_data = {'intervalo': sched_seconds, 'ambito': ambito}
    with open (jsonFile, "w", encoding='utf-8') as f:
        f.write(json.dumps(json_data))
        f.close()
    print("datos guardados: %s" % json_data)

#bucle principal
def main_loop():
    '''se encarga de comprobar regularmente si hay que notificar nuevos tickets'''
    if len(ambito) == 0:
        icon.notify(title='Aviso', message="No se monstrará ningún ticket, seleccione al menos un ámbito")
        
    driver.refresh()
    #importante ajustar el zoom para que entren todas las columnas
    driver.execute_script("document.body.style.zoom='10%'")

    tickets_asignados = comprobar_tickets()
    def ticket_formato(tit, mess):
        return "{}: {}\n".format(tit, mess)
    if len(tickets_asignados) > 0:
        print("tickets asignados: %s" % list(map(lambda i: i['Id'], tickets_asignados)))
        rel = ""
        tickets_filtrados = 0
        for nt in tickets_asignados:
            ambito_usuario = "INSS" if nt["remitente"][2:3] == 'I' else None
            ambito_usuario = "TGSS" if nt["remitente"][2:3] == 'T' else None      
            if(nt["empresa"] in ambito) or (nt["empresa"] == "SJSS" and ambito_usuario in ambito):
                tickets_filtrados += 1
                rel += ticket_formato(nt['Id'], nt['remitente'])
            print(f"tickets_filtrados: {tickets_filtrados} {rel}")

        #antes de notificar comprobamos que no estemos en la pantalla de bloqueo    
        if check_process("LogonUI.exe") > 0:
            print("pantalla bloqueo detectada")
        else: #mostramos las notificaciones
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


if __name__ == "__main__":
    #comprobamos si el programa ya está corriendo
    print("instancias de passcontrol: %d" % check_process("passcontrol.exe"))
    #cargamos los datos de usuario del archivo json
    load_json()
    if check_process("passcontrol.exe") > 2:
        print("passcontrol ya está ejecutándose")
        sys.exit(0)
    #creamos la imagen desde el codigo base64
    buffer = io.BytesIO(base64.b64decode(pass_icon))
    image = PIL.Image.open(buffer)

    icon = pyIcon('pass menu', image, 'passControl', menu=pyMenu(
        pyItem('Abrir navegador Edge', tray_sched),
        pyItem('Parar', tray_sched, checked=lambda item: sched.state != sched_base.STATE_RUNNING),
        pyItem('Ambito', pyMenu( lambda: (
            pyItem(
                text = '%s' % i,
                action = set_ambito(i), # lambda i: ambito.append(i) if i not in ambito else ambito.remove(i),
                checked = get_ambito(i))
            for i in lista_ambitos)
        )),
        pyItem('Estadísticas', pyMenu( lambda: (
            pyItem( '%s' % e, action=None)
            for e in estadisticas),
        )),
        pyItem('Tiempo(seg)', pyMenu( lambda: (
            pyItem(
                '%s' % i,
                set_state_sched(i),
                checked=get_state_sched(i),
                radio=True)
            for i in [30,60,120,300]),
        )),
        pyItem('Salir', tray_quit)
    ))
    
    #abrimos ventana url principal
    serv = service.Service(current_dir + '\\msedgedriver.exe')
    options = get_default_options()
    driver = webdriver.Edge(options=options, service=serv)
    #import edgedriver_autoinstaller
    #edgedriver_autoinstaller.install() #instala el driver de edge si no existe
    #options.executable_path = "msedgedriver.exe"
    #driver = webdriver.Edge(options=options)

    driver.implicitly_wait(15) #tiempo de espera implicito
    driver.get(main_url)
    #cambiamos el zoom para que entren todas las columnas en el viewport
    #driver.execute_script("document.body.style.zoom='20%'")
    driver.execute_script('document.title = "pass_background_control"')
    print("titulo ventana principal: %s" % driver.title)
    main_loop() #primera ejecución al arrancar

    #arrancamos la cola
    start_scheduler(sched_seconds)

    #arranca el loop principal
    icon.run()
