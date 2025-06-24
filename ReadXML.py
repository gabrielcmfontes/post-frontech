import win32com.client
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
import re
import os
import json
import requests

DOWNLOAD_FOLDER = r'C:\Users\joao.vitor\Desktop\Automação\Automação XML\xml_abastecimentos'


def limpar_pasta_xml(diretorio):
    """
    Apaga todos os arquivos .xml de um diretório específico de forma segura.
    """
    print("Iniciando limpeza da pasta de XMLs antigos...")
    
    dir_absoluto = os.path.abspath(diretorio)

    if not os.path.exists(dir_absoluto):
        print(f"A pasta '{dir_absoluto}' não existe. Nada a limpar.")
        return

    if os.path.basename(dir_absoluto) != 'xml_abastecimentos':
        print(f"ERRO DE SEGURANÇA: A limpeza foi cancelada pois o diretório '{os.path.basename(dir_absoluto)}' não é a pasta esperada ('xml_abastecimentos').")
        return

    arquivos_deletados = 0
    for arquivo in os.listdir(dir_absoluto):
        if arquivo.lower().endswith('.xml'):
            caminho_completo = os.path.join(dir_absoluto, arquivo)
            try:
                os.remove(caminho_completo)
                print(f"  - Deletado: {arquivo}")
                arquivos_deletados += 1
            except Exception as e:
                print(f"  - Erro ao deletar {arquivo}: {e}")
    
    if arquivos_deletados == 0:
        print("Nenhum arquivo .xml encontrado para limpar.")
    else:
        print(f"Limpeza concluída. {arquivos_deletados} arquivos foram deletados.")


def baixar_anexos_periodo(data_inicio, data_fim):
    """
    Baixa anexos XML do Outlook para o período especificado (datas inclusivas).
    """
    print(f"\nIniciando automação do Outlook Desktop para o período de {data_inicio.strftime('%d/%m/%Y')} até {data_fim.strftime('%d/%m/%Y')}...")
    try:
        download_path_absoluto = os.path.abspath(DOWNLOAD_FOLDER)
        if not os.path.exists(download_path_absoluto):
            os.makedirs(download_path_absoluto)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        print(f"Acessando a Caixa de Entrada: '{inbox.Name}'")

        messages = inbox.Items
        inicio_dia = data_inicio.replace(hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M')
        fim_dia = data_fim.replace(hour=23, minute=59, second=59).strftime('%Y-%m-%d %H:%M')
        filtro = f"[ReceivedTime] >= '{inicio_dia}' AND [ReceivedTime] <= '{fim_dia}'"
        print(f"Data interpretada início: {data_inicio}")
        print(f"Data interpretada fim: {data_fim}")
        print(f"Filtro usado: {filtro}")
        print(f"Procurando por e-mails recebidos entre {data_inicio.strftime('%d/%m/%Y')} e {data_fim.strftime('%d/%m/%Y')}...")
        mensagens_filtradas = messages.Restrict(filtro)

        if len(mensagens_filtradas) == 0:
            print("Nenhum e-mail novo encontrado para o período informado.")
            return

        print(f"Encontrados {len(mensagens_filtradas)} e-mails. Baixando anexos XML...")
        for message in mensagens_filtradas:
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    if attachment.FileName.lower().endswith('.xml'):
                        caminho_arquivo = os.path.join(download_path_absoluto, attachment.FileName)
                        print(f"  - Baixando anexo: '{attachment.FileName}'")
                        attachment.SaveAsFile(caminho_arquivo)
        print("\nDownload dos anexos concluído.")
    except Exception as e:
        print(f"\nOcorreu um erro ao tentar automatizar o Outlook: {e}")
        print("Dicas: Verifique se o Outlook está instalado e se você o abriu pelo menos uma vez.")


def extrair_placa_km(infCpl_texto):
    if not infCpl_texto: return None, None
    placa_match = re.search(r'Placa[:\-]?\s*([A-Z]{3}[0-9][A-Z0-9][0-9]{2})', infCpl_texto, re.IGNORECASE)
    km_match = re.search(r'KM[:\-]?\s*(\d+)', infCpl_texto, re.IGNORECASE)
    placa = placa_match.group(1).upper() if placa_match else None
    km = km_match.group(1) if km_match else None
    return placa, km

def ler_abastecimentos_de_arquivo(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        if '}' in root.tag:
            ns = {'ns': root.tag.split('}')[0].strip('{')}
        else:
            ns = {'ns': ''} if not root.tag.startswith('{') else {'ns': root.tag.split('}')[0].strip('{')}

        ide_node = root.find('.//ns:ide', namespaces=ns)
        nNF_node = ide_node.find('./ns:nNF', namespaces=ns) if ide_node is not None else None
        if nNF_node is not None and nNF_node.text is not None and nNF_node.text.isdigit():
            invoice_id = int(nNF_node.text)
        else:
            invoice_id = None
        dhEmi_node = ide_node.find('./ns:dhEmi', namespaces=ns) if ide_node is not None else None
        invoice_date = dhEmi_node.text if dhEmi_node is not None else None
        
        emit_node = root.find('.//ns:emit', namespaces=ns)
        xNome_node = emit_node.find('./ns:xNome', namespaces=ns) if emit_node is not None else None
        issuer = xNome_node.text if xNome_node is not None else None

        infCpl_node = root.find('.//ns:infAdic/ns:infCpl', namespaces=ns)
        infCpl_text = infCpl_node.text if infCpl_node is not None else ''
        placa, km = extrair_placa_km(infCpl_text)
        if km is not None and km.isdigit():
            kilometers = int(km)
        else:
            kilometers = None

        dados_json_list = []
        for det in root.findall('.//ns:det', namespaces=ns):
            prod_node = det.find('./ns:prod', namespaces=ns)
            xProd_node = prod_node.find('./ns:xProd', namespaces=ns) if prod_node is not None else None
            fuel_type = xProd_node.text if xProd_node is not None else None
            qCom_node = prod_node.find('./ns:qCom', namespaces=ns) if prod_node is not None else None
            try:
                quantity = float(qCom_node.text) if qCom_node is not None and qCom_node.text is not None else None
            except ValueError:
                quantity = None
            vUnCom_node = prod_node.find('./ns:vUnCom', namespaces=ns) if prod_node is not None else None
            try:
                unit_cost = float(vUnCom_node.text) if vUnCom_node is not None and vUnCom_node.text is not None else None
            except ValueError:
                unit_cost = None
            vProd_node = prod_node.find('./ns:vProd', namespaces=ns) if prod_node is not None else None
            try:
                total_cost = float(vProd_node.text) if vProd_node is not None and vProd_node.text is not None else None
            except ValueError:
                total_cost = None

            json_output = {
                "invoiceId": invoice_id, "issuer": issuer, "invoiceDate": invoice_date, "date": invoice_date,
                "plate": placa, "kilometers": kilometers, "fuelType": fuel_type, "quantity": quantity,
                "unitCost": unit_cost, "totalCost": total_cost
            }
            dados_json_list.append(json_output)
            
        return dados_json_list
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao ler o arquivo {xml_path}. Erro: {e}")
        return []

def ler_todos_os_abastecimentos(diretorio_xml):
    # ... (código inalterado) ...
    todos_dados = []
    diretorio_absoluto = os.path.abspath(diretorio_xml)
    if not os.path.exists(diretorio_absoluto):
        print(f"A pasta '{diretorio_absoluto}' está vazia ou não foi encontrada. Nenhum arquivo para processar.")
        return []
    arquivos_xml = [f for f in os.listdir(diretorio_absoluto) if f.lower().endswith('.xml')]
    if not arquivos_xml:
        print(f"Nenhum arquivo .xml encontrado em '{diretorio_absoluto}' para processar.")
        return []
    for arquivo in arquivos_xml:
        caminho_arquivo = os.path.join(diretorio_absoluto, arquivo)
        print(f"Processando: {caminho_arquivo}")
        dados_arquivo = ler_abastecimentos_de_arquivo(caminho_arquivo)
        todos_dados.extend(dados_arquivo)
    print("\n--- RESULTADO FINAL (JSON LOCAL) ---")
    print(json.dumps(todos_dados, indent=2, ensure_ascii=False))
    return todos_dados


# --- FUNÇÃO DE POST CORRIGIDA PARA ENVIAR UM ITEM DE CADA VEZ ---
def postAbastecimentos(lista_de_abastecimentos):
    """
    Envia cada abastecimento da lista para a URL, um de cada vez, mostrando detalhes da requisição e resposta para debug.
    """
    url = "http://192.168.11.95:3012/fuel"
    headers = { "Content-Type": "application/json" }
    respostas_servidor = []

    print("\nIniciando o post dos abastecimentos, um por um...")
    
    for abastecimento_individual in lista_de_abastecimentos:
        invoice_id = int(abastecimento_individual.get("invoiceId", "ID Desconhecido"))
        print(f"  - Enviando abastecimento ID: {invoice_id}...")
        print(f"    Corpo enviado: {json.dumps(abastecimento_individual, ensure_ascii=False)}")
        try:
            response = requests.post(url, headers=headers, json=abastecimento_individual, timeout=10)
            print(f"    Resposta bruta: {response.status_code} {response.reason}")
            print(f"    Conteúdo da resposta: {response.text}")
            response.raise_for_status() 
            print(f"    - Sucesso! Status: {response.status_code}")
            respostas_servidor.append(response.json())
        except requests.exceptions.HTTPError as http_err:
            print(f"    - ERRO HTTP do servidor ao enviar ID {invoice_id}: {http_err}")
            print(f"      Resposta do servidor: {response.text}")
            respostas_servidor.append({'error_id': invoice_id, 'details': response.text})
        except requests.exceptions.RequestException as e:
            print(f"    - ERRO DE CONEXÃO ao enviar ID {invoice_id}: {e}")
            respostas_servidor.append({'error_id': invoice_id, 'details': str(e)})

    return respostas_servidor
    
def main(data_str):
    """
    Função principal para processar os abastecimentos de um único dia.
    Recebe a data como string no formato 'DD/MM/YYYY'.
    """
    try:
        data = datetime.strptime(data_str, "%d/%m/%Y")
    except ValueError:
        print(f"Data inválida: {data_str}. Use o formato DD/MM/YYYY.")
        return

    limpar_pasta_xml(DOWNLOAD_FOLDER)
    baixar_anexos_periodo(data, data)  # Apenas um dia

    print("\nIniciando a leitura dos arquivos XML baixados...")
    dados_json_list = ler_todos_os_abastecimentos(DOWNLOAD_FOLDER)

    print("\nIniciando o post dos abastecimentos...")
    postAbastecimentos(dados_json_list)

    print("\n--- Processo Concluído ---")

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python ReadXML.py DD/MM/YYYY")
    else:
        main(sys.argv[1])