import json
import os
import time
import wave
from datetime import datetime
import locale
import clr
import webbrowser
import requests 

from array import array
from shutil import move
from struct import pack
from sys import byteorder

import numpy as np
import pyaudio
import speech_recognition as sr
#import win32serviceutil
from gtts import gTTS
from matplotlib import pylab
from pydub import AudioSegment 

import subprocess
import threading


tiempo = "https://www.el-tiempo.net/api/json/v2/provincias/29"

openhardwaremonitor_hwtypes = ['Mainboard','SuperIO','CPU','RAM','GpuNvidia','GpuAti','TBalancer','Heatmaster','HDD']
openhardwaremonitor_sensortypes = ['Voltage','Clock','Temperature','Load','Fan','Flow','Control','Level','Factor','Power','Data','SmallData']
temperatura = []
componente = []

THRESHOLD = 800
CHUNK_SIZE = 2048
CHUNK_SIZE_PLAY = 2000
FORMAT = pyaudio.paInt16
RATE = 44100
MAXIMUM = 12000
r = sr.Recognizer()

dt = datetime.now()
locale.setlocale(locale.LC_ALL, 'es-ES')
#print(dt.strftime("%A %d %B %Y %I:%M"))
#print(dt.strftime("%A %d %B del %Y - %H:%M"))

#FUNCIONES TEMPERATURAS
def initialize_openhardwaremonitor():
    file = 'OpenHardwareMonitor'
    clr.AddReference(file)

    from OpenHardwareMonitor import Hardware

    handle = Hardware.Computer()
    handle.MainboardEnabled = False
    handle.CPUEnabled = True
    handle.RAMEnabled = False
    handle.GPUEnabled = True
    handle.HDDEnabled = False
    handle.Open()
    return handle

def fetch_stats(handle):
    for i in handle.Hardware:
        i.Update()
        for sensor in i.Sensors:
            parse_sensor(sensor)
        for j in i.SubHardware:
            j.Update()
            for subsensor in j.Sensors:
                parse_sensor(subsensor)


def parse_sensor(sensor):
    global temperatura
    global componente
    if sensor.Value is not None:
        if type(sensor).__module__ == 'OpenHardwareMonitor.Hardware':
            sensortypes = openhardwaremonitor_sensortypes
            hardwaretypes = openhardwaremonitor_hwtypes
        else:
            return

        if sensor.SensorType == sensortypes.index('Temperature'):
            #print(u"%s %s Temperature Sensor #%i %s - %s\u00B0C" % (hardwaretypes[sensor.Hardware.HardwareType], sensor.Hardware.Name, sensor.Index, sensor.Name, sensor.Value))

            temperatura.append("%s" % (sensor.Value))
            componente.append("%s" % (sensor.Hardware.Name))

#FUNCIONES RECONOCIMIENTO DE VOZ

def is_silent(snd_data):
   
    """Devuelve True si el sonido ambiente está por debajo de un rango concreto."""
   
    return max(snd_data) < THRESHOLD


def normalize(snd_data):
    """Normaliza el volumen de una pista de audio"""
    times = float(MAXIMUM) / max(abs(i) for i in snd_data)

    r = array('h')
    for i in snd_data:
        r.append(int(i * times))
    return r


def trim(snd_data):
    """Corta los silencios al principio y al final"""

    def _trim(sound_data):
        snd_started = False
        r = array('h')

        for i in sound_data:
            if not snd_started and abs(i) > THRESHOLD:
                snd_started = True
                r.append(i)

            elif snd_started:
                r.append(i)
        return r

    snd_data = _trim(snd_data)
    snd_data.reverse()
    snd_data = _trim(snd_data)
    snd_data.reverse()
    return snd_data


def add_silence(snd_data, seconds):
    
    r = array('h', [0 for i in range(int(seconds * RATE))])
    r.extend(snd_data)
    r.extend([0 for i in range(int(seconds * RATE))])
    return r


def record():
    """ Graba el audio usando el micrófono """
    p = pyaudio.PyAudio()
    stream = p.open(format=FORMAT, channels=1, rate=RATE,
                    input=True, output=True,
                    frames_per_buffer=CHUNK_SIZE)

    num_silent = 0
    snd_started = False

    r = array('h')

    while 1:

        snd_data = array('h', stream.read(CHUNK_SIZE))
        if byteorder == 'big':
            snd_data.byteswap()
        r.extend(snd_data)

        silent = is_silent(snd_data)

        if silent and snd_started:
            num_silent += 1
        elif not silent and not snd_started:
            snd_started = True

        if snd_started and num_silent > 30:
            break

    sample_width = p.get_sample_size(FORMAT)
    stream.stop_stream()
    stream.close()
    p.terminate()

    r = normalize(r)
    r = trim(r)
    r = add_silence(r, 0.5)
    return sample_width, r


def record_to_file(path):
    """ Usando la función record, crea un fichero wav en el directorio del programa """
    sample_width, data = record()
    data = pack('<' + ('h' * len(data)), *data)

    wf = wave.open(path, 'wb')
    wf.setnchannels(1)
    wf.setsampwidth(sample_width)
    wf.setframerate(RATE)
    wf.writeframes(data)
    wf.close()


def soundplot(audiofile):
    wf = wave.open(audiofile, 'rb')
    p = pyaudio.PyAudio()
    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True)

    data = wf.readframes(CHUNK_SIZE_PLAY)
    t1 = time.time()
    counter = 1
    pylab.style.use('dark_background')
    pylab.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

    while len(data) > 0:
        stream.write(data)
        if counter % 2 == 0:
            np_data = np.fromstring(data, dtype=np.int16)
            pylab.plot(np_data, color="#41FF00")
            pylab.axis('off')
            pylab.axis([0, len(data) / 2, -2 ** 16 / 2, 2 ** 16 / 2])
            pylab.savefig("03.svg", format="svg", transparent=True)
            move("03.svg", "plot.svg")
            pylab.close('all')

        data = wf.readframes(CHUNK_SIZE_PLAY)
        counter += 1
    stream.close()
    p.terminate()

 
def speak(text):
    tts = gTTS(text=text, lang='es-ES')
    tts.save('files/hello.mp3')
    AudioSegment.from_mp3("files/hello.mp3").export("files/hello.wav", format="wav")
    os.remove('files/hello.mp3')
    soundplot('files/hello.wav')

#FUNCIONES PARA EJECUTAR PROGRAMAS
def spotify():
    
    subprocess.run('C:/Users/Antonio/AppData/Roaming/Spotify/Spotify.exe')

def lol():

     subprocess.run('C:/Riot Games/Riot Client/RiotClientServices.exe --launch-product=league_of_legends --launch-patchline=live')

def bloc():

     subprocess.run('notepad.exe')


if __name__ == '__main__':

    webbrowser.open_new_tab("header.html")
    HardwareHandle = initialize_openhardwaremonitor()
    fetch_stats(HardwareHandle)
    speak(' Iniciando sistemas... Buenos días!')
   

    while True:
        print("Háblale al micrófono")
        record_to_file('demo.wav')
        print("Grabado! Escrcito a demo.wav")
        voice = sr.AudioFile('demo.wav')
        print("Abriendo fichero de audio")
        with voice as source:
            audio = r.record(source)
        try:
            print("Reconociendo audio...")
            # Aquí usamos Google Speech Recognizer para reconocer audio en español a texto
            
            a = r.recognize_google(audio, language='es-ES')
            print(a)
            if "proyecto" in a:
             
                if "cuál es" in a and "propósito" in a:
                    speak("he sido creado para que Miguel y Antonio puedan aprobar ")

                elif "quién es" in a and "creador" in a:
                    speak("mis creadores son Miguel Angel y Antonio")
                
                elif "a que"  in a or "a qué" in a and "estamos" in a:
                    speak(dt.strftime("Estamos a %A %d %B de %Y y son las %I:%M"))

                elif "cuál es" in a and "temperatura" in a and "procesador" in a:
                    speak("La temperatura del procesador es: " + temperatura[0] + " grados")
                
                elif "cuál es" in a and "temperatura" in a and "gráfica" in a:
                    speak("La temperatura de la grafica es: " + temperatura[3] + " grados")
                
                elif "cuáles son" in a or "cuales son" in a and "componentes" in a:
                    speak("El procesador es un " + componente[0] + "y la grafica es una " + componente[3])
                
                elif "suma" in a:

                    cortar_frase = a.split() #2 y 4
                    resultado = 0
                    #print(cortar_frase)
                    #num1 = cortar_frase[2]
                    #num2 = cortar_frase[4]
                    #resultado = int(num1) + int(num2)
                    #print(cortar_frase)
                    
                    #print(type(cortar_frase[2]))
                   
                    for elemento in cortar_frase:
                        
                        if elemento.isdigit():
                            
                            resultado = resultado + int(elemento)
                            
                    speak("El resultado es %s" % (resultado))

                elif "resta" in a:

                    cortar_frase = a.split()
                    entrar = False
                   
                    for elemento in cortar_frase:
                        
                        if elemento.isdigit():
                            if entrar == True:
                                resultado = resultado - int(elemento)
                            else:
                                resultado = int(elemento)
                                entrar = True

                            
                    speak("El resultado es %s" % (resultado))
                
                elif "multiplica" in a:

                    cortar_frase = a.split() 
                    resultado = 1
                   
                    for elemento in cortar_frase:
                        
                        if elemento.isdigit():
                            
                            resultado = resultado * int(elemento)
                            
                    speak("El resultado es %s" % (resultado))                    
                
                elif "divide" in a:

                    cortar_frase = a.split()
                    entrar = False
                   
                    for elemento in cortar_frase:
                        
                        if elemento.isdigit():
                            if entrar == True:
                                resultado = resultado / float(elemento)
                            else:
                                resultado = float(elemento)
                                entrar = True
                    if(resultado - int(resultado) == 0.5):
                        resultado = str(int(resultado)) + " y medio"

                    speak("El resultado es %s" % (resultado))
                elif "busca" in a:

                    cortar_frase = a.split()
                    cortar_frase.pop(0)
                    cortar_frase.pop(0)
                    busqueda = " ".join(cortar_frase)

                    speak("Obteniendo busqueda de " + busqueda)
                    webbrowser.open_new_tab("https://www.google.com/search?&q=%s" % busqueda)

                elif "qué" in a and "tiempo" in a:
                    
                    respuesta = requests.get(tiempo)
                    lista = json.loads(respuesta.text)

                    speak("En " + lista["ciudades"][0]["name"] + " tenemos el cielo " + lista["ciudades"][0]["stateSky"]["description"] + " con temperaturas entre " + lista["ciudades"][0]["temperatures"]["min"] + " y " + lista["ciudades"][0]["temperatures"]["max"] + " grados.")

                elif "quiero" in a and "música" in a:
                    
                    speak("Abriendo Spotify")
                    
                    spoty = threading.Thread(target=spotify)
                    spoty.start()

                elif "quiero" in a and "LoL" in a or "lol" in a:

                    speak("Abriendo League of legends")

                    liga = threading.Thread(target=lol)
                    liga.start()

                elif "quiero" in a and "vídeo" in a:
                    speak("Abriendo YouTube")

                    webbrowser.open_new_tab("https://www.youtube.com/")   

                elif "quiero" in a and "peli" in a or "pelí" in a or "serie" in a:

                    speak("Abriendo Netflix")
                    webbrowser.open_new_tab("https://www.netflix.com/")

                elif "tengo" in a and "apuntar" in a:
                    speak("Abriendo Bloc de Notas")

                    notas = threading.Thread(target=bloc)
                    notas.start()
                    
                elif "ciérrate" in a or "cierra" in a:
                    
                    speak("Hasta luego máquina")
                    
                    break;

        
        except Exception as e:
            print(e)
        print("Reconocimiento terminado")



