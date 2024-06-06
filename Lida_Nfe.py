import PySimpleGUI as sg
import pandas as pd

def lerArquivo_NFE():
    sg.set_options(font=('Trebuchet', 14))
    layout = [
        [sg.Text('Selecione o arquivo com as NFS',background_color="#000000",pad=(103,5),text_color="#eeeeee")],
        [sg.Input(key="-FILE-", enable_events=True,background_color="#eeeeee",size=(36)), sg.FileBrowse(button_text='Procurar',button_color="#2a3990",auto_size_button=(False),size=(7,1),file_types=(("Excel","*.xlsx"),))],
        [sg.Image(r"C:\app\logomarca.png\\",pad=(120,10),)],
        [sg.Button('Confirmar',button_color="#2a3990")]
    ]

    window = sg.Window('Selecionar Arquivo', layout, icon='C:\app\\logo.ico' ,background_color='#000000',size=(520,350))

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED:
            window.close()
            return
        elif event == '-FILE-':
            file_path = values["-FILE-"]
            window["-FILE-"].update(file_path)
        elif event == 'Confirmar':
            file_path = values["-FILE-"]
            if file_path:
                df = pd.read_excel(file_path, converters={'Chave de Acesso': str, 'Documento Destinatário': str})
                df["Sistema"] = ""
                df["Emitente"] = ""
                df["CFOP"] = ""
                df["Valor Total"] = ""
                window.close() 
                return df
            else:
                sg.popup('Por favor, selecione um arquivo Excel!', title='Aviso', keep_on_top=True)         

    