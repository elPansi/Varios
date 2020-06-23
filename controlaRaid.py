import sys, os, logging, csv, subprocess
from threading import Timer,Thread,Event
from logging.handlers import RotatingFileHandler
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

"""Script para detectar la rotura de un HDD del RAID HP, gestionado por hpaucli
"""

#Variables
minutosComprobarRAID = 1
rutaLog = '/root/log/'
prefijoNomFichero = "Log_RAID"
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

#Preparamos el EMAIL
usuarioSMTP = 'email@origen.com'
emailRecepcion = ['email@destino.com']
servidorSMTP = 'smtp.office365.com:587'
emailEnvio = usuarioSMTP
passSMTP = 'passCorreo'

mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = 'Notificacion de disco duro Roto'
cuerpoEmail = """Atencion requerida, se ha roto un disco duro del servidor principal"""

mail.attach(MIMEText(cuerpoEmail, 'plain'))


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


def compruebaRoturaDisco():
    #command1 = "hpacucli ctrl all show config | grep '(8.2 TB, RAID 5, OK)' > /root/errorRaid"
    command1 = "hpacucli ctrl all show config | grep Failed > /root/errorRaid"
    command2 = "wc < /root/errorRaid | awk '{print $1}'"
    salida = ''
    try:
        os.system(command1)
        salida = subprocess.check_output(command2, shell=True)
    except subprocess.CalledProcessError as e:
        logInterno.info("Error al ejecutar comando. " + e.output)
    logInterno.info("Salida: " + salida )
    if(int(salida)>0): #Se ha roto un disco
        #mandamos el email
        logInterno.info("Rotura de disco detectada. Notificando a traves de email")
        servidor = smtplib.SMTP(servidorSMTP)
        servidor.starttls()
        servidor.login(usuarioSMTP, passSMTP)
        servidor.sendmail(mail['From'], mail['To'], mail.as_string())
        servidor.quit()

t1 = gestionaTimers((minutosComprobarRAID * 60), compruebaRoturaDisco) 
t1.start()

#Instalacion del servicio
#Crearmos un archivo /lib/systemd/system/controlaRaid.service
"""
[Unit]
Description=Servicio para detectar la rotura de un HDD del Raid y notificar por email
After=multi-user.target
Conflicts=getty@tty1.service

[Service]
Type=simple
ExecStart=/usr/bin/python /root/controlaRaid.py
StandardInput=tty-force

[Install]
WantedBy=multi-user.target
"""

#Instalar y arrancar el servicio
"""
systemctl daemon-reload
systemctl enable controlaRaid.service
systemctl start controlaRaid.service
"""