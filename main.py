import PySimpleGUI as sg
from bsc_cfop import obter_path, busca_cfop
from Lida_Cte import lerArquivo_CTE
from Lida_Nfe import lerArquivo_NFE
from salva_arquivo import salvar_xlsx
from trata_dados import buscarNotas, buscarCnpj_Nfe, buscarcc, buscarEmitente, buscarCnpj_Cte, chave_acesso

def main():
    layout = [
        [sg.Text('Selecione a Rotina desejada:', pad=(160, 5), background_color="#000000")],
        [sg.Button('CTE', pad=(100, 30), auto_size_button=False, button_color="#2a3990"), sg.Button('NFE', auto_size_button=False, button_color="#2a3990")],
        [sg.Image(r"C:\app\logomarca.png\\",pad=(120,10),)]
    ]

    window = sg.Window('Selecionar Rotina', layout, icon='C:\app\logo.ico', background_color='#000000', size=(520, 350))

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Cancel':
            break

        elif event == 'CTE':
            window.hide()
            df = lerArquivo_CTE()
            print(df)
            salvar_xlsx(
                tipo='CTE',
                df=df,
                emitente=buscarEmitente(df),
                cnpjs=buscarCnpj_Cte(df),
                result=buscarNotas(df),
                ccs=None,
                s_cfop=None,
                s_valor=None
            )
            break

        elif event == 'NFE':
            window.hide()
            df = lerArquivo_NFE()
            print(buscarcc(df))
            lista= chave_acesso(df)
            path=obter_path()
            cfop, valor=busca_cfop(CA=lista,path=path)
            salvar_xlsx(
                tipo= 'NFE',
                df=df,
                result=buscarNotas(df),
                emitente=buscarEmitente(df),
                cnpjs=buscarCnpj_Nfe(df),
                ccs=buscarcc(df),
                s_cfop=(cfop),
                s_valor=(valor)
            )
            break
    window.close()

if __name__ == '__main__':
    main()
