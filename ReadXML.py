import win32com.client
from datetime import datetime, timedelta


import xml.etree.ElementTree as ET
import re
import os

# --- CONFIGURAÇÕES ---
# A pasta de download continua a mesma
DOWNLOAD_FOLDER = r'C:\Users\joao.vitor\Desktop\Automação\Automação XML\xml_abastecimentos'

def baixar_anexos_do_dia_anterior():
    """
    Controla o Outlook instalado no Windows para buscar e-mails de ontem
    e baixar anexos XML, sem precisar de senha de aplicativo.
    """
    print("Iniciando automação do Outlook Desktop...")
    try:
        download_path_absoluto = os.path.abspath(DOWNLOAD_FOLDER)

        if not os.path.exists(download_path_absoluto):
            os.makedirs(download_path_absoluto)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        inbox = outlook.GetDefaultFolder(6)
        print(f"Acessando a Caixa de Entrada: '{inbox.Name}'")

        messages = inbox.Items

        ontem = datetime.now() - timedelta(days=1)
        inicio_dia = ontem.replace(hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M')
        fim_dia = ontem.replace(hour=23, minute=59, second=59).strftime('%Y-%m-%d %H:%M')
        filtro = f"[ReceivedTime] >= '{inicio_dia}' AND [ReceivedTime] <= '{fim_dia}'"
        
        print(f"Procurando por e-mails recebidos em {ontem.strftime('%d/%m/%Y')}...")
        mensagens_filtradas = messages.Restrict(filtro)

        if len(mensagens_filtradas) == 0:
            print("Nenhum e-mail encontrado para a data de ontem.")
            return

        print(f"Encontrados {len(mensagens_filtradas)} e-mails. Verificando anexos XML...")

        for message in mensagens_filtradas:
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    if attachment.FileName.lower().endswith('.xml'):
                        # CORREÇÃO 1 (continuação): Usa o caminho absoluto para salvar
                        caminho_arquivo = os.path.join(download_path_absoluto, attachment.FileName)
                        print(f"  - Baixando anexo: '{attachment.FileName}' do e-mail: '{message.Subject}'")
                        attachment.SaveAsFile(caminho_arquivo)
        
        print("\nDownload dos anexos via Outlook concluído.")

    except Exception as e:
        print(f"\nOcorreu um erro ao tentar automatizar o Outlook: {e}")
        print("Dicas: Verifique se o Outlook está instalado e se você o abriu pelo menos uma vez.")

def extrair_placa_km(infCpl_texto):
    if not infCpl_texto: return None, None 
    placa_match = re.search(r'Placa[:\-]?\s*([A-Z]{3}[0-9][A-Z0-9][0-9]{2})', infCpl_texto, re.IGNORECASE)
    km_match = re.search(r'KM[:\-]?\s*(\d+)', infCpl_texto, re.IGNORECASE)
    placa = placa_match.group(1) if placa_match else None
    km = km_match.group(1) if km_match else None
    return placa, km

def ler_abastecimentos_de_arquivo(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        if '}' in root.tag:
            ns = {'ns': root.tag.split('}')[0].strip('{')}
        else:
            ns = {'ns': ''} 

        dados = []
        for det in root.findall('.//ns:det', namespaces=ns):
            infCpl_node = root.find('.//ns:infCpl', namespaces=ns)
            infCpl_text = infCpl_node.text if infCpl_node is not None else ''
            placa, km = extrair_placa_km(infCpl_text)

            dados.append({
                'Produto': det.findtext('.//ns:xProd', namespaces=ns),
                'Quantidade': det.findtext('.//ns:qCom', namespaces=ns),
                'Valor Unitário': det.findtext('.//ns:vUnCom', namespaces=ns),
                'Data de Emissão': root.findtext('.//ns:dhEmi', namespaces=ns),
                'Número da Nota': root.findtext('.//ns:nNF', namespaces=ns),
                'Placa': placa,
                'KM': km
            })
        return dados
    except ET.ParseError as e:
        print(f"Erro ao fazer o parse do arquivo XML: {xml_path}. Erro: {e}")
        return []

def ler_todos_os_abastecimentos(diretorio_xml):
    todos_dados = []
    diretorio_absoluto = os.path.abspath(diretorio_xml)
    if not os.path.exists(diretorio_absoluto):
        print(f"Diretório '{diretorio_absoluto}' não encontrado. Nenhum arquivo para processar.")
        return []
        
    for arquivo in os.listdir(diretorio_absoluto):
        if arquivo.lower().endswith('.xml'):
            caminho_arquivo = os.path.join(diretorio_absoluto, arquivo)
            print(f"Processando: {caminho_arquivo}")
            dados_arquivo = ler_abastecimentos_de_arquivo(caminho_arquivo)
            todos_dados.extend(dados_arquivo)
    
    print("\n--- RESULTADO FINAL ---")
    for dado in todos_dados:
        print('---')
        for chave, valor in dado.items():
            print(f'{chave}: {valor}')
    return todos_dados

if __name__ == "__main__":
    baixar_anexos_do_dia_anterior()
    
    print("\nIniciando a leitura dos arquivos XML baixados...")
    ler_todos_os_abastecimentos(DOWNLOAD_FOLDER)