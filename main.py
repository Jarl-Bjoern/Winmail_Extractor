# Rainer Herold
# Version 0.2
# 06.08.2021 Vers. 0.1
# 07.09.2021 Vers. 0.2

# Bibliotheken_Implementierung
import email, os, sys, time, tnefparse
if (os.name == 'nt'):
    from win32console import GetConsoleWindow
    from win32con import SW_HIDE, WM_CLOSE
    from win32gui import PostMessage, ShowWindow
    
    # WIN_ID_Ermitteln
    ID_CMD = GetConsoleWindow()
    ShowWindow(ID_CMD, SW_HIDE)
else:
    sys.exit()
from atexit import register
from email import policy
from datetime import datetime

# Variablen_Bereich
Datum_Formatiert, Endung, Go, n, Pfad_Mail = "", "", 0, 1, r'C:\Temp'
#Pfad_Ausgang = os.path.expanduser(r'~\Desktop')
#Pfad_Mail = os.path.expanduser(r'~\Desktop\mail')

# Funktionen
def Mail_Extrahieren():
    os.path.exists(Pfad_Mail) or os.makedirs(Pfad_Mail)
    output_count = 0
    for root, dirs, files in os.walk(Pfad_Ausgang, topdown=False):
        for file in files:
            if (file.endswith('.eml')):
                try:
                    with open(os.path.join(root, file), "r") as f:
                        msg = email.message_from_file(f, policy=policy.default)
                        for attachment in msg.iter_attachments():
                            output_filename = attachment.get_filename()
                             # If no attachments are found, skip this file
                            if output_filename:
                                with open(os.path.join(Pfad_Mail, output_filename), "wb") as of:
                                    of.write(attachment.get_payload(decode=True))
                                    output_count += 1
                        if output_count == 0:
                            print(f"Es wurde kein Anhang für die Datei {f.name} gefunden!")
                 # this should catch read and write errors
                except IOError:
                    print(f"Es wurde ein Problem mit der Datei {f.name} festgestellt!")
                return 1, output_count

@register
def Programm_Schliessen():
    PostMessage(ID_CMD, WM_CLOSE, 0, 0)

# Aufruf_Bereich
if __name__ == "__main__":
    #Mail_Extrahieren()
    #if (len(sys.argv) > 1):
        #IMAPClientID = sys.argv[1]
        IMAPClientID = 'test'
        for root, dirs, files in os.walk(Pfad_Mail, topdown=False):
            for file in files:
                if (file.endswith('.dat')):
                    with open(os.path.join(root, file), 'rb') as tneffile:
                        Datum = os.path.getmtime(os.path.join(root, file))
                        Datum_Konverter = datetime.fromtimestamp(Datum)
                        tnefobj = tnefparse.tnef.TNEF(tneffile.read())
                        for i in tnefobj.attachments:
                            Datei_Name = i.name
                            for j in range(0, len(Datei_Name)):
                                if (Datei_Name[j] == '.'):
                                    Go = 1
                                if (Go == 1):
                                    Endung += Datei_Name[j]
                            for k in range(0, len(str(Datum_Konverter))):
                                if (not str(Datum_Konverter)[k] == '-' and not str(Datum_Konverter)[k] == ' ' and not str(Datum_Konverter)[k] == ':' and not str(Datum_Konverter)[k] == '.'):
                                    Datum_Formatiert += str(Datum_Konverter)[k]
                            with open(os.path.join(Pfad_Mail, Datei_Name), 'wb') as afp:
                                afp.write(i.data)
                            afp.close()
                            try:
                                os.rename(os.path.join(Pfad_Mail, Datei_Name), os.path.join(Pfad_Mail, f'imap_mail.{IMAPClientID}.{Datum_Formatiert}_0.part{n}{Endung}'))
                            except:
                                os.remove(os.path.join(Pfad_Mail, f'imap_mail.{IMAPClientID}.{Datum_Formatiert}_0.part1{Endung}'))
                                os.rename(os.path.join(Pfad_Mail, Datei_Name), os.path.join(Pfad_Mail, f'imap_mail.{IMAPClientID}.{Datum_Formatiert}_0.part{n}{Endung}'))
                            Go, Endung, Datum_Formatiert = 0, "", ""
                    tneffile.close()
