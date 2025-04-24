# instalar requerimientos: pip install -r requirements.txt
# crear exe en carpeta dist: pyinstaller --hide-console hide-late --onefile .\passcontrol.py
import os, sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge import service
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.schedulers import base as sched_base
#from plyer import notification
from pystray import Icon as pyIcon, Menu as pyMenu, MenuItem as pyItem
import PIL.Image

main_url = 'https://segsocial-smartit.onbmc.com/smartit/app/#/ticket-console'
last_ids = list() #se guardan los ids de los tickets anteriores
sched_seconds = 10 #intervalo scheduler
current_dir = os.getcwd() #directorio actual

def get_default_options():
    """activa las opciones por defecto"""
    os.environ.pop('HTTP_PROXY', None)
    os.environ.pop('HTTPS_PROXY', None)
    opt = webdriver.EdgeOptions() #Options()
    opt.use_chromium = True
    opt.add_experimental_option("detach", True)
    opt.binary_location = r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
    #options.add_experimental_option("excludeSwitches", ['enable-logging'])
    #options.ignore_local_proxy_environment_variables() #deprecated
    opt.add_argument("--no-sandbox")
    opt.add_argument("--log-level=3")
    opt.add_argument("--headless")
    opt.add_argument("start-maximized")
    return opt

#espera a que un elemento se encuentre cargado y lo devuelve
def wait_element(selector, patron):
    '''espera a que el elemento indicado se haya cargado'''
    wait = WebDriverWait(driver, timeout=15, poll_frequency=.2)
    wait.until(EC.visibility_of_element_located((selector, patron)))
    return driver.find_element(selector, patron)

#captura las filas de items y guarda los valores
def get_items():
    '''captura los items y devuelve un array'''
    items = []
    #en caso de que no existan tickets, se producirá una excepcion
    try:
        #viewPort = wait_element(By.CSS_SELECTOR, '.ngViewport .ng-scope')
        vp_header = wait_element(By.CLASS_NAME, "ngHeaderContainer")
        viewPort  = driver.find_element(By.CLASS_NAME, "ngViewport")
        vp_header_items = vp_header.find_elements(By.CSS_SELECTOR, ".ngHeaderText")
        vp_items = viewPort.find_elements(By.CSS_SELECTOR,'.ng-scope .ngRow')
        cabeceras = []
        #capturamos el nombre de las cabeceras y su posición
        for e in vp_header_items:
            hd_class = e.get_attribute("class")
            nombre_col = e.text
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
                    case "Mostrar ID":
                        n_nombre = "Id"
                    case "Remitente":
                        n_nombre = "remitente"
                    case "Fecha de creación":
                        n_nombre = "fecha"
                campos[n_nombre] = e.find_element(By.CSS_SELECTOR, ".ng-scope .{}".format(c["campo"])).text.strip()
            #print("campos: %s" % campos)
            items.append(campos)
    except Exception as ex:
        print("excepcion en get_items: %s" % type(ex))
    return items
    
#comprueba si aparece algún ticket nuevo y devuelve un array
def comprobar_tickets():
    '''comprueba si existen nuevos tickets'''
    global last_ids
    items = get_items()
    #print(items)
    ids = list(map(lambda i: i['Id'], items))
    print("numero de items: %d: %s" % (len(items), list(ids)))
    new_ids = list(i for i in ids if i not in last_ids) #nuevos identificadores
    #nuevos tickets
    new_tickets = list(x for x in items if x['Id'] in new_ids)
    last_ids = ids[:]
    return new_tickets

# def ver_notif(notif):
#     '''muestra una notificación'''
#     img_path = "logo_pass32.ico"
#     notification.notify(
#             title = "Nuevos tickets!!",
#             message = notif,
#             #we need to download a icon of ico file format
#             app_icon = img_path,
#             # the notification stays for 50sec
#             timeout  = 50
#     )

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
        icon.notify(title='cambiado intervalo', message="nuevo valor: {} seg".format(sched_seconds))
    return inner

def get_state_sched(sta):
    def inner(item):
        return sched_seconds == sta
    return inner

def tray_sched(icon, item):
    print("sched state: %s" % sched.state)
    if( sched.state != sched_base.STATE_RUNNING ):
        print("arrancamos job")
        sched.resume()
        icon.notify(message="estado actual: funcionando")
    else:
        print("paramos scheduler")
        sched.pause()
        icon.notify(message="estado actual: pausado")


def tray_quit(icon):
    '''paramos la cola, el driver de edge y cerramos el icono de systray'''
    if sched.running:
        print("quit: paramos scheduler")
        sched.shutdown()
    driver.close()
    icon.stop()


#bucle principal
def main_loop():
    '''se encarga de comprobar regularmente si hay que notificar nuevos tickets'''
    nuevos = comprobar_tickets()
    if( len(nuevos) > 0 ):
        print("tickets nuevos: %s" % list(map(lambda i: i['Id'], nuevos)))
        rel = ""
        for nt in nuevos:
            rel += "%s: %s\n" % (nt['Id'], nt['remitente'])
        #ver_notif(rel)
        icon.notify(title='nuevos tickets!', message=rel)
    else:
        print("no hay nuevos tickets")
    driver.refresh()


if __name__ == "__main__":
    #definimos icono para la bandeja del sistema
    image = PIL.Image.open(current_dir+"\\logo_pass32.ico")
    icon = pyIcon('pass menu', image, 'tickets pass', menu=pyMenu(
        #pyItem('parar', tray_sched, checked=lambda item: not sched.running),
        pyItem('parar', tray_sched, checked=lambda item: sched.state != sched_base.STATE_RUNNING),
        pyItem('tiempo(seg)', pyMenu( lambda: (
            pyItem(
                '%d' % i,
                set_state_sched(i),
                checked=get_state_sched(i),
                radio=True)
            for i in [10,30,60,90,120,300]),
        )),
        pyItem('Salir', tray_quit)
    ))

    #abrimos ventana url principal
    options = get_default_options()
    serv = service.Service(current_dir + '\\msedgedriver.exe')
    #service = webdriver.EdgeService(log_output="edgedriver.log")
    driver = webdriver.Edge(options=options, service=serv)
    driver.implicitly_wait(10) #tiempo de espera implicito
    driver.get(main_url)

    print("titulo ventana principal: %s" % driver.title)
    #arrancamos la cola
    sched = BackgroundScheduler()
    start_scheduler(sched_seconds)

    #run_program = True
    #arranca el loop principal
    icon.run()
    #sys.exit(0)
