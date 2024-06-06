import os
import xml.etree.ElementTree as ET
import PySimpleGUI as sg
from time import sleep
diretorio = ''
cfops = {}
valores = {}

def obter_path():
    
    sg.set_options(font=('Trebuchet', 14))
    layout = [
        [sg.Text('Selecione a Pasta com os XMLS',background_color="#000000",pad=(103,5),text_color="#eeeeee")],
        [sg.Input(key="-FILE-", enable_events=True,background_color="#eeeeee",size=(36)), sg.FolderBrowse(button_text='Procurar',button_color="#2a3990",auto_size_button=(False),size=(7,1))],
        [sg.Image(r"C:\app\logomarca.png\\",pad=(120,10),)],
        [sg.Button('Confirmar',button_color="#2a3990")]
    ]

    window = sg.Window('Selecionar Pasta', layout, icon='C:\app\\logo.ico' ,background_color='#000000',size=(520,350))

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED:
            window.close()
      
        elif event == '-FILE-':
            file_path = values["-FILE-"]
            window["-FILE-"].update(file_path)
        elif event == 'Confirmar':
            file_path = values["-FILE-"]
            if file_path:
                window.close() 
                return file_path
            else:
                sg.popup('Por favor, selecione uma Pasta', title='Aviso', keep_on_top=True)        

def busca_cfop(CA,path):
    namespaces = {
        'nfe': 'http://www.portalfiscal.inf.br/nfe',
    }
    cfops = {}
    
    for path, subdirs, files in os.walk(path):
        for name in files:
            try:
                xml_path = os.path.join(path, name)
                tree = ET.parse(xml_path)
                root = tree.getroot()
                
                chave = root.find(".//{http://www.portalfiscal.inf.br/nfe}chNFe").text
                
                if chave in CA and chave not in cfops:
                    cfop_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}CFOP", namespaces)
                    value_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}vNF", namespaces)
                    if cfop_element is not None:
                        cfop = cfop_element.text
                        cfops[chave] = cfop
                    if value_element is not None:
                        valor = value_element.text
                        valores[chave] = valor
                else:
                    print(name)
                if len(cfops) == len(CA):
                    return cfops
            except Exception as e:
                pass
    return cfops, valores


