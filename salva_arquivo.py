from datetime import datetime
import PySimpleGUI as sg

def salvar_xlsx(df,emitente,cnpjs,result,ccs,tipo,s_cfop,s_valor )            :
    if tipo == "CTE":
        for i, infos in df.iterrows():

            if infos[1] in emitente.keys(): # Busca CNPJ emitente no xlsx e preenche o nome do fornecedor.
                df.loc[i,"Emitente"] = emitente[infos[1]]

            if infos[3] in cnpjs.keys(): # Busca CNPJ destinatário e preenche o nome do destinatário.
                df.loc[i,"Destinatário"] = cnpjs[infos[3]]

            if infos[0] in result.keys(): # Busca data digitação da nota, se contém data foi dado entrada.
                df.loc[i,"Sistema"] = datetime.strptime(result[infos[0]], "%Y%m%d").strftime("%d/%m/%Y")

        output_file = sg.popup_get_file("Salvar Arquivo Excel", save_as=True, file_types=(("Arquivos Excel", "*.xlsx"),))
        if output_file:
            df.to_excel(output_file, index=False)
            sg.popup("Arquivo Excel salvo com sucesso!")
            return
        else:
            sg.popup_error("Nenhum local de arquivo selecionado.")
    elif tipo == "NFE":
        for i, infos in df.iterrows():

            if infos[1] in cnpjs.keys(): # Busca Doc destinatário no xls e preenche o nome 
                df.loc[i,"Emitente"] = cnpjs[infos[1]]

            if infos[0] in result.keys(): # Busca data digitação da nota, se contém data foi dado entrada.
                df.loc[i,"Sistema"] = datetime.strptime(result[infos[0]], "%Y%m%d").strftime("%d/%m/%Y")

            if infos[0] in s_cfop.keys(): #Pega dicionário do arquivo bsc_cfop e itera.
                df.loc[i,"CFOP"] = s_cfop[infos[0]]

            if infos[0] in s_valor.keys(): #Pega dicionário do arquivo bsc_cfop e itera.
                df.loc[i,"Valor Total"] = s_valor[infos[0]]

            if infos[1] in ccs.keys():  # Query que busca último C.C que esse fornecedor deu entrada.
                df.loc[i,"Ultimo C.C Fornecedor"] = ccs[infos[1]]
        output_file = sg.popup_get_file("Salvar Arquivo Excel", save_as=True, file_types=(("Arquivos Excel", "*.xlsx"),))

        if output_file:
            df.to_excel(output_file, index=False)
            sg.popup("Arquivo Excel salvo com sucesso!")
            return
        else:
            sg.popup_error("Nenhum local de arquivo selecionado.")