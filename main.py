import time
import pandas as pd
from datetime import datetime, timedelta, date

import win32com.client as win32
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

### gerenciador do chromedriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def refresh_excel_workbook(excel_path):
    ### abre o arquivo excel no caminho especificado, atualiza e salva
    print("iniciando excel para atualizar a planilha...")
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True
    try:
        wb = excel_app.Workbooks.Open(excel_path)

        ### atualiza conexoes de dados
        for conn in wb.Connections:
            conn.Refresh()
        ### ou poderiamos usar: wb.RefreshAll()

        ### pausa para garantir a conclusao da atualizacao
        time.sleep(45)

        ### salva e fecha
        wb.Save()
        wb.Close()
    except Exception as e:
        print("erro ao atualizar excel:", e)
    finally:
        excel_app.Quit()
    print("planilha atualizada e salva com sucesso!")

def login_bASF(driver):
    ### realiza login no portal basf/neogrid e fecha popup inicial
    url = "https://basf.neogrid.com/basf/login/basf"
    driver.get(url)

    ### aguarda campo de email e insere credenciais
    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='frmLogin:fldEmail']"))
    )
    email_input.send_keys("email@email.com")

    ### aguarda campo de senha e insere credenciais
    senha_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='frmLogin:fldPassword']"))
    )
    senha_input.send_keys("password")

    ### clica em login
    botao_login = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='frmLogin:btnLogin']"))
    )
    botao_login.click()

    ### fecha o popup inicial
    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='_homePopupHidelink']"))
    ).click()

def consultar_cte(driver, documento_busca):
    ### consulta situacoes de ct-e para o documento_busca, retorna numero de linhas encontradas e lista de situacoes
    situacoes_cte = []
    cte_count = 0

    ### torna visivel o menu oculto para ct-e
    menu_oculto_cte = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ngMenuForm:m33_menu']"))
    )
    driver.execute_script("arguments[0].style.display = 'block';", menu_oculto_cte)

    ### clica na opcao do menu ct-e
    menu_button_cte = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='ngMenuForm:m34:anchor']"))
    )
    menu_button_cte.click()

    ### insere o numero de doc e envia
    input_campo_cte = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='filterForm:fldDocTransporte']"))
    )
    input_campo_cte.clear()
    input_campo_cte.send_keys(str(documento_busca))
    input_campo_cte.send_keys(Keys.ENTER)

    ### pequena pausa
    time.sleep(2)
    try:
        ### checa se "nenhum registro encontrado"
        msg_erro_cte = driver.find_element(By.XPATH, "//*[@id='j_id254']").text.strip()
        if "Nenhum registro encontrado" in msg_erro_cte:
            return (0, ["Sem registro de CT-e"])
    except:
        pass

    try:
        ### aguarda a tabela dos resultados
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='ngFindListForm:tblDataTable']"))
        )
        time.sleep(2)

        ### pega as linhas e conta quantos ct-es
        linhas_cte = driver.find_elements(By.XPATH, "//*[@id='ngFindListForm:tblDataTable']/tbody/tr")
        cte_count = len(linhas_cte)

        for index in range(cte_count):
            try:
                ### localiza situacao
                situacao_xpath = f"//*[@id='ngFindListForm:tblDataTable:{index}:j_id174']"
                situacao_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, situacao_xpath))
                )
                situacao_texto = situacao_element.text.strip()

                ### se situacao e 'recebido basf', adiciona, senao, abre popup divergencias
                if situacao_texto == "Recebido BASF":
                    situacoes_cte.append("Recebido BASF")
                else:
                    icone_xpath = f"//*[@id='ngFindListForm:tblDataTable:{index}:j_id210']"
                    icone_element = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, icone_xpath))
                    )
                    icone_element.click()

                    ### aguarda popup e analisa divergencias
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            "//*[@id='divergenciasBatimentoModal_form:tblDataTableDivergencias']"
                        ))
                    )
                    time.sleep(2)

                    linhas_div = driver.find_elements(
                        By.XPATH,
                        "//*[@id='divergenciasBatimentoModal_form:tblDataTableDivergencias']/tbody/tr"
                    )
                    for idx_div in range(len(linhas_div)):
                        descricao_xpath = f"//*[@id='divergenciasBatimentoModal_form:tblDataTableDivergencias:{idx_div}:j_id719']"
                        try:
                            descricao_element = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, descricao_xpath))
                            )
                            descricao_texto = descricao_element.text.strip()
                        except:
                            descricao_texto = "Erro ao capturar descrição"

                        situacoes_cte.append(f"{situacao_texto} - {descricao_texto}")

                    ### fecha popup
                    fechar_popup = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='divergenciasBatimentoModal_hidelink']"))
                    )
                    fechar_popup.click()
                    time.sleep(2)

            except:
                situacoes_cte.append("Erro ao capturar situação CT-e")
    except:
        situacoes_cte.append("Erro ao capturar CT-e")

    return (cte_count, situacoes_cte)

def consultar_notfis(driver, documento_busca):
    ### consulta situacoes de notfis para documento_busca, retorna lista de strings
    situacoes_notfis = []

    ### torna visivel o menu oculto para notfis
    menu_oculto_notfis = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ngMenuForm:m37_menu']"))
    )
    driver.execute_script("arguments[0].style.display = 'block';", menu_oculto_notfis)

    ### clica na opcao do menu notfis
    menu_button_notfis = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='ngMenuForm:m38:anchor']"))
    )
    menu_button_notfis.click()

    ### insere o numero e envia
    input_campo_notfis = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='filterForm:fldDocTransporte']"))
    )
    input_campo_notfis.clear()
    input_campo_notfis.send_keys(str(documento_busca))
    input_campo_notfis.send_keys(Keys.ENTER)

    ### pequena pausa
    time.sleep(2)
    try:
        msg_erro_notfis = driver.find_element(By.XPATH, "//*[@id='j_id216']").text.strip()
        if "Nenhum registro encontrado" in msg_erro_notfis:
            return ["Sem registro de NOTFIS"]
    except:
        pass

    try:
        ### aguarda a tabela de resultados
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='ngFindListForm:tblDataTable']"))
        )
        time.sleep(2)

        ### obtem as linhas de notfis
        linhas_notfis = driver.find_elements(
            By.XPATH,
            "//*[@id='ngFindListForm:tblDataTable']/tbody/tr"
        )
        for idx in range(len(linhas_notfis)):
            situacao_xpath_notfis = f"//*[@id='ngFindListForm:tblDataTable:{idx}:j_id164']"
            try:
                sit_element_notfis = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, situacao_xpath_notfis))
                )
                situacao_notfis_texto = sit_element_notfis.text.strip()
                situacoes_notfis.append(situacao_notfis_texto)
            except:
                situacoes_notfis.append("Erro ao capturar situação NOTFIS")
    except:
        situacoes_notfis.append("Erro ao capturar NOTFIS")

    return situacoes_notfis

def main():
    ### caminho do arquivo excel
    pasta_base = r"S:\Publico\Dep. Financeiro - Faturamento\1. Basf"
    nome_planilha_base = "00. Base Basf.xlsx"
    caminho_arquivo = os.path.join(pasta_base, nome_planilha_base)

    ### atualiza excel
    refresh_excel_workbook(caminho_arquivo)

    ### le a planilha apos atualizacao
    df = pd.read_excel(caminho_arquivo, sheet_name="queryBasf")
    print("colunas encontradas:", df.columns.tolist())

    ### checa se coluna cnpj existe
    if "CNPJ" not in df.columns:
        raise ValueError("A coluna 'CNPJ' não existe na planilha 'queryBasf'!")

    ### verifica colunas obrigatorias
    colunas_obrigatorias = [
        "CONTROLE", "NUM DOC", "FILIAL", "DATA VENCIMENTO", "VALOR",
        "DATA EMISSAO", "DIA SEMANA"
    ]
    for needed_col in colunas_obrigatorias:
        if needed_col not in df.columns:
            raise ValueError(f"A coluna '{needed_col}' não existe na planilha!")

    ### remove linhas sem controle
    df = df.dropna(subset=["CONTROLE"])
    ### padroniza controle para string
    df["CONTROLE"] = df["CONTROLE"].astype(str)

    ### converte data emissao, se necessario
    df["DATA EMISSAO"] = pd.to_datetime(df["DATA EMISSAO"]).dt.date

    ### prepara filtro pelo dia de ontem
    ontem = date.today() - timedelta(days=1)
    hoje = date.today()

    ### checa se hoje e segunda
    if hoje.weekday() == 0:
        ### se segunda, carrega todos
        df_filtrado = df.copy()
        print("hoje é segunda-feira. carregando todos os registros.")
    else:
        ### senao, pega apenas data emissao de ontem
        df_filtrado = df[df["DATA EMISSAO"] == ontem]
        print(
            f"ontem foi {ontem.strftime('%d/%m/%Y')}. "
            f"carregando apenas registros com data emissao = ontem. "
            f"total de linhas filtradas: {len(df_filtrado)}"
        )

    ### se vazio, encerra
    if df_filtrado.empty:
        print("nenhum registro para processar após o filtro. encerrando script.")
        return

    ### agrupa por controle
    grouped = df_filtrado.groupby("CONTROLE", as_index=False)

    ### prepara listas de saida
    a_faturar_data = []
    erros_data = []

    consultados = set()

    ### inicia driver
    options = Options()
    options.add_argument("--start-maximized")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        ### login no basf
        login_bASF(driver)
        print("login realizado com sucesso!")

        ### consulta cada controle
        for controle_valor, group_df in grouped:
            if controle_valor in consultados:
                print(f"controle {controle_valor} já consultado. pulando.")
                continue
            consultados.add(controle_valor)

            num_docs_excel = len(group_df)
            print(f"\niniciando consulta para controle: {controle_valor}")

            ### chama consulta cte e notfis
            cte_count, situacoes_cte = consultar_cte(driver, controle_valor)
            situacoes_notfis = consultar_notfis(driver, controle_valor)

            ### checa se contagem de ct-e e igual no excel
            if cte_count < num_docs_excel:
                erros_data.append({
                    "CONTROLE": controle_valor,
                    "SIT CTE": "Contagem de CT-e diferente",
                    "SIT NOTFIS": "; ".join(situacoes_notfis) if situacoes_notfis else ""
                })
                continue

            ### verifica se todas situacoes sao recebido basf
            all_cte_ok = all(sit == "Recebido BASF" for sit in situacoes_cte)
            all_notfis_ok = all(sit == "Recebido BASF" for sit in situacoes_notfis)

            if not situacoes_cte or not situacoes_notfis or (not all_cte_ok or not all_notfis_ok):
                ### erro, adiciona em erros
                str_cte = "; ".join(situacoes_cte) if situacoes_cte else ""
                str_notfis = "; ".join(situacoes_notfis) if situacoes_notfis else ""
                erros_data.append({
                    "CONTROLE": controle_valor,
                    "SIT CTE": str_cte,
                    "SIT NOTFIS": str_notfis
                })
            else:
                ### tudo certo, prepara para faturar
                for _, row in group_df.iterrows():
                    a_faturar_data.append({
                        "CGCPAGADOR": row.get("CNPJ", ""),
                        "FILIALDOC": row["FILIAL"],
                        "DOCUMENTO": row.get("NUM DOC", ""),
                        "SERIE": row.get("SERIE", ""),
                        "ID": row["CONTROLE"],
                        "VENCIMENTODOC": row["DATA VENCIMENTO"],
                        "FRETETOTAL": row["VALOR"],
                    })

            time.sleep(1)

        print("\nconsulta concluída para todos os controles.")

    except Exception as e:
        print(f"erro na automação: {e}")
    finally:
        time.sleep(3)
        driver.quit()

    ### gera arquivo excel final
    data_hoje = datetime.now().strftime("%d.%m.%y")
    nome_arquivo_saida = f"{data_hoje} - Basf.xlsx"
    caminho_saida = os.path.join(pasta_base, nome_arquivo_saida)

    df_faturar = pd.DataFrame(a_faturar_data)
    if not df_faturar.empty:
        ### formata data de vencimento
        df_faturar["VENCIMENTODOC"] = pd.to_datetime(df_faturar["VENCIMENTODOC"]).dt.strftime("%d/%m/%Y")

    df_erros = pd.DataFrame(erros_data)

    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        df_faturar.to_excel(writer, sheet_name="A Faturar", index=False)
        df_erros.to_excel(writer, sheet_name="Erros", index=False)

    print(f"relatório gerado com sucesso em '{caminho_saida}'.")

    ### gera csv para a aba a faturar
    csv_folder = os.path.join(pasta_base, "00. CSV")
    os.makedirs(csv_folder, exist_ok=True)
    nome_arquivo_csv = f"{data_hoje} - Basf.csv"
    caminho_csv = os.path.join(csv_folder, nome_arquivo_csv)

    df_faturar.to_csv(caminho_csv, index=False, sep=";")
    print(f"arquivo csv gerado com sucesso em: {caminho_csv}")

    print("script finalizado.")

if __name__ == "__main__":
    main()
