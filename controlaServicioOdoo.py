import sys, os, logging, csv, subprocess
from threading import Timer,Thread,Event
from logging.handlers import RotatingFileHandler
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

"""Script para detectar la parada del servicio Odoo
"""

#Variables
segundosComprobarServicio = 15
rutaLog = '/root/log/'
if not os.path.isdir(rutaLog):
   os.mkdir(rutaLog)

prefijoNomFichero = "Log_Odoo_Monitor"
maxLogSize = 20

#Configuracion del fichero de log de sistema -> logInterno
log_formatter = logging.Formatter('%(asctime)s %(levelname)s %(funcName)s(%(lineno)d) %(message)s')
logInternoFichero = rutaLog + prefijoNomFichero + ".txt"
my_handler = RotatingFileHandler(logInternoFichero, mode='a', maxBytes=maxLogSize*1024*1024, backupCount=20, encoding=None, delay=0)
my_handler.setFormatter(log_formatter)
my_handler.setLevel(logging.INFO)
logInterno = logging.getLogger('root')
logInterno.setLevel(logging.INFO)
logInterno.addHandler(my_handler)


#Funcion que controla los temporizadores
class gestionaTimers():

   def __init__(self,t,hFunction):
      self.t=t
      self.hFunction = hFunction
      self.thread = Timer(self.t,self.handle_function)

   def handle_function(self):
      self.hFunction()
      self.thread = Timer(self.t,self.handle_function)
      self.thread.start()

   def start(self):
      self.thread.start()

   def cancel(self):
      self.thread.cancel()


def compruebaServicioOdoo():
   readStatusCmd = 'systemctl status odoo | grep Active'
   try:
      # os.system(readStatusCmd)
      service_state_input = subprocess.check_output(readStatusCmd, shell=True)

   except subprocess.CalledProcessError as e:
      logInterno.info("Error al ejecutar comando. " + e.output)
   logInterno.info("Salida Shell: " + service_state_input )      
   if service_state_input.find('running') < 0:
      logInterno.info("Error: Servicio Odoo Parado")
      # Arrancamos el servicio
      startOdooCmd = 'systemctl start odoo'
      try:
         output = subprocess.check_output(startOdooCmd, shell=True)
         logInterno.info("Servicio: " + output)
      except subprocess.CalledProcessError as e:  
         logInterno.info("Error Arrancando Odoo")       


   


t1 = gestionaTimers((segundosComprobarServicio), compruebaServicioOdoo) 
t1.start()

#Instalacion del servicio
#Crearmos un archivo /lib/systemd/system/controlaRaid.service
"""
[Unit]
Description=Servicio para detectar la parada del servicio de Odoo
After=multi-user.target
Conflicts=getty@tty1.service

[Service]
Type=simple
ExecStart=/usr/bin/python /root/controlaServicioOdoo.py
StandardInput=tty-force

[Install]
WantedBy=multi-user.target
"""

#Instalar y arrancar el servicio
"""
systemctl daemon-reload
systemctl enable controlaServicioOdoo.service
systemctl start controlaServicioOdoo.service
"""
