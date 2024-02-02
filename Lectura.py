# Importamos librerias para lecturas
import cv2
import pyqrcode
import png
from datetime import datetime
from pyzbar.pyzbar import decode
import numpy as np
import openpyxl as xl

# Creamos la videocaptura
cap = cv2.VideoCapture(0)

#Variables
mañana = []
tarde = []
noche = []

#Horario
def infhora():
    inf = datetime.now()
    fecha = inf.strftime('%Y:%m:%d')
    hora = inf.strftime('%H:%M:%S')

    return hora, fecha
# Empezamos
while True:
    # Leemos los frames
    ret, frame = cap.read()

    #interfaz
    cv2.putText(frame, 'Locate the QR Code',(160, 80), cv2.FONT_HERSHEY_SIMPLEX, 1,(0, 255, 0), 2)
    cv2.rectangle(frame, (170, 100), (400,470),(0,255,0), 2)

    #Extraer fecha y hora
    hora, fecha = infhora()
    diasem = datetime.today().weekday()

    print (diasem)

    a, me, d = fecha [0:4], fecha [5:7], fecha [8:10]
    h, m, s = int (hora[0:2]), int(hora[3:5]), int(hora[6:8])

    #Creamos Archivo
    nomar = str(a) + '-' + str(me) + '-' + str(d)
    texth = str(h) + ':' + str(m) + ':' + str(s)
    print(texth)
    print(nomar)
    wb = xl.Workbook()

    # Leemos los codigos QR
    for codes in decode(frame):
        # Extraemos info
        #info = codes.data

        # Decodidficamos
        info = codes.data.decode('utf-8')

        # Tipo de persona LETRA
        tipo = info[0:2]
        tipo = int(tipo)
        letr = chr(tipo)

        #numero
        num = info[2:]

        # Extraemos coordenadas
        pts = np.array([codes.polygon], np.int32)
        xi, yi = codes.rect.left, codes.rect.top

        # Redimensionamos
        pts = pts.reshape((-1,1,2))

        #ID Completo
        codigo = letr + num

        if 4 >= diasem >= 0:

            if 20 >= h >= 17:
                cv2.polylines(frame, [pts], True, (255, 255, 0), 5)

                if codigo not in mañana:

                    pos = len(mañana)
                    mañana.append(codigo)

                    hojam = wb.create_sheet("Mañana")
                    datos = hojam.append(mañana)
                    wb.save(nomar + '.xlsx')

                    cv2.putText(frame, letr + '0' + str(num), (xi, yi - 15), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 55, 0), 2)

                elif codigo in mañana:
                    cv2.putText(frame,'EL ID' + str(codigo),
                                (xi - 65, yi - 45), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                    cv2.putText(frame, 'Fue registrado',
                                (xi - 65, yi - 15), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                    print(mañana)

    # Mostramos FPS
    cv2.imshow(" LECTOR DE QR", frame)
    # Leemos teclado
    t = cv2.waitKey(5)
    if t == 27:
        break

cv2.destroyAllWindows()
cap.release()