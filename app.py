from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import os

# Carrega variáveis do .env (obrigatório para rodar)
load_dotenv()

# Pastas de destino/origem
def strip_quotes(val):
    if val is None:
        return None
    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
        return val[1:-1]
    return val

# Leia as variáveis do ambiente SEM aplicar strip_quotes ainda
destino_analitico_raw = os.getenv("DESTINO_ANALITICO")
destino_estoque_raw = os.getenv("DESTINO_ESTOQUE")
destino_venda_raw = os.getenv("DESTINO_VENDA")
empresa_nome = os.getenv("EMPRESA_NOME")
url_login = os.getenv("URL_LOGIN")
usuario = os.getenv("USUARIO")
senha = os.getenv("SENHA")

# Validação obrigatória das variáveis de ambiente (antes de strip_quotes)
if not all([url_login, usuario, senha, destino_analitico_raw, destino_estoque_raw, destino_venda_raw, empresa_nome]):
    raise RuntimeError(
        "Variáveis de ambiente obrigatórias não definidas. "
        "Verifique seu arquivo .env ou as variáveis do sistema."
    )

# Agora aplique strip_quotes
destino_analitico = strip_quotes(destino_analitico_raw)
destino_estoque = strip_quotes(destino_estoque_raw)
destino_venda = strip_quotes(destino_venda_raw)

# Funções
def autenticar(driver, url, usuario, senha):
    status = []
    from datetime import datetime
    from pathlib import Path
    import glob
    import shutil

    # --- Checagem inicial dos relatórios ---
    now = datetime.now()
    data_hoje = now.strftime("%d/%m/%Y")
    ano_atual = now.strftime("%Y")
    mes_atual = now.strftime("%m")

    # Cálculo do mês anterior e ano correto para o arquivo analítico anterior
    if mes_atual == "01":
        mes_anterior = "12"
        ano_anterior_arquivo = str(int(ano_atual) - 1)
    else:
        mes_anterior = f"{int(mes_atual)-1:02d}"
        ano_anterior_arquivo = ano_atual

    # O nome do analítico anterior deve ser sempre do mês anterior, mas o ano só muda se o mês anterior for 12
    nome_analitico_anterior = f"{ano_anterior_arquivo}_{mes_anterior}_Rel_Venda_Anali_André_Sborz_01-{mes_anterior}-{ano_anterior_arquivo}.xlsx"
    caminho_analitico_anterior = os.path.join(destino_analitico, nome_analitico_anterior)

    nome_estoque = "Estoque Atual.xlsx"
    caminho_estoque = os.path.join(destino_estoque, nome_estoque)

    nome_venda = f"{ano_atual}_Rel_Venda_Sint_Andre_Sborz.xlsx"
    caminho_venda = os.path.join(destino_venda, nome_venda)

    nome_analitico = f"{ano_atual}_{mes_atual}_Rel_Venda_Anali_André_Sborz_01-{mes_atual}-{ano_atual}.xlsx"
    caminho_analitico_final = os.path.join(destino_analitico, nome_analitico)

    # Adicione para o relatório de venda ano anterior:
    ano_passado = str(int(ano_atual) - 1)
    nome_venda_ano_passado = f"{ano_passado}_Rel_Venda_Sint_Andre_Sborz.xlsx"
    caminho_venda_ano_passado = os.path.join(destino_venda, nome_venda_ano_passado)

    relatorios_a_gerar = []
    rels_status = {}

    # Checa cada relatório e monta lista do que precisa gerar
    # Estoque
    if os.path.exists(caminho_estoque) and datetime.fromtimestamp(os.path.getmtime(caminho_estoque)).strftime("%d/%m/%Y") == data_hoje:
        status.append("Estoque Atual.xlsx - já gerado hoje")
        rels_status["estoque"] = "já gerado hoje"
    else:
        relatorios_a_gerar.append("estoque")

    # Venda Sintético
    if os.path.exists(caminho_venda) and datetime.fromtimestamp(os.path.getmtime(caminho_venda)).strftime("%d/%m/%Y") == data_hoje:
        status.append(f"{nome_venda} - já gerado hoje")
        rels_status["venda"] = "já gerado hoje"
    else:
        relatorios_a_gerar.append("venda")

    # Analítico mês atual
    if os.path.exists(caminho_analitico_final) and datetime.fromtimestamp(os.path.getmtime(caminho_analitico_final)).strftime("%d/%m/%Y") == data_hoje:
        status.append(f"{nome_analitico} - já gerado hoje")
        rels_status["analitico"] = "já gerado hoje"
    else:
        relatorios_a_gerar.append("analitico")

    # Analítico mês anterior: só não gera se já existe e foi gerado neste mês (YYYY-MM do arquivo == YYYY-MM atual)
    if os.path.exists(caminho_analitico_anterior):
        data_geracao_anterior = datetime.fromtimestamp(os.path.getmtime(caminho_analitico_anterior))
        if data_geracao_anterior.strftime("%Y-%m") == now.strftime("%Y-%m"):
            status.append(f"{nome_analitico_anterior} - já gerado este mês")
            rels_status["analitico_anterior"] = "já gerado este mês"
        else:
            relatorios_a_gerar.append("analitico_anterior")
    else:
        relatorios_a_gerar.append("analitico_anterior")

    # Venda Sintético Ano Anterior: só não gera se já existe e foi gerado neste ano (YYYY do arquivo == YYYY atual)
    if os.path.exists(caminho_venda_ano_passado):
        data_geracao_venda_ano_passado = datetime.fromtimestamp(os.path.getmtime(caminho_venda_ano_passado))
        if data_geracao_venda_ano_passado.strftime("%Y") == now.strftime("%Y"):
            status.append(f"{nome_venda_ano_passado} - já gerado este ano")
            rels_status["venda_ano_passado"] = "já gerado este ano"
        else:
            relatorios_a_gerar.append("venda_ano_passado")
    else:
        relatorios_a_gerar.append("venda_ano_passado")

    # Se todos já foram gerados, retorna imediatamente
    if not relatorios_a_gerar:
        return status

    # --- Processo Selenium só se houver relatório a gerar ---
    driver.get(url)
    try:
        usuario_element = driver.find_element(By.ID, "usuario")
        usuario_element.send_keys(usuario)
        status.append("Login: OK")
    except Exception as e:
        status.append(f"Login: Erro ao preencher usuário - {e}")
        return status

    try:
        senha_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "senha"))
        )
        if senha_element.is_displayed():
            senha_element.send_keys(senha)
            status.append("Senha: OK")
        else:
            status.append("Senha: Campo 'senha' não está visível.")
            return status
    except Exception as e:
        status.append(f"Senha: Erro ao preencher senha - {e}")
        return status

    try:
        entrar_element = driver.find_element(By.ID, "continuar")
        entrar_element.click()
        sleep(5)
        status.append("Entrar: OK")
    except Exception as e:
        status.append(f"Entrar: Erro ao clicar no botão 'continuar' - {e}")
        return status

    try:
        empresa_modal = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "ui-id-2"))
        )
        empresa_element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[@id='ui-id-2' and contains(text(), '{empresa_nome}')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", empresa_element)
        empresa_element.click()
        sleep(2)
        status.append("Selecionar empresa: OK")
    except Exception as e:
        status.append(f"Selecionar empresa: Erro - {e}")
        driver.quit()
        sleep(5)
        import sys
        import subprocess
        subprocess.Popen([sys.executable] + sys.argv)
        sys.exit(1)
        return status

    try:
        confirmar_element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'selecionarEmpresa')]"))
        )
        confirmar_element.click()
        sleep(5)
        status.append("Confirmar empresa: OK")
    except Exception as e:
        status.append(f"Confirmar empresa: Erro - {e}")
        return status

    # Gera apenas os relatórios necessários
    # --- Relatório de Estoque ---
    if "estoque" in relatorios_a_gerar:
        nome_estoque = "Estoque Atual.xlsx"
        caminho_estoque = os.path.join(destino_estoque, nome_estoque)
        gerar_estoque = True
        data_hoje = datetime.now().strftime("%d/%m/%Y")

        # Verifica se já existe o arquivo de estoque do dia
        if os.path.exists(caminho_estoque):
            data_mod = datetime.fromtimestamp(os.path.getmtime(caminho_estoque)).strftime("%d/%m/%Y")
            if data_mod == data_hoje:
                status.append("Estoque Atual.xlsx - Já gerado hoje")
                gerar_estoque = False

        if gerar_estoque:
            try:
                dashboard_url = "https://hering.myrp.app/app/gerencial/relatorios/relatoriosAvancados/gerar/?tipo=estoque"
                driver.execute_script(f"window.open('{dashboard_url}', '_blank');")
                driver.switch_to.window(driver.window_handles[-1])
                print("Dashboard aberto em nova aba com sucesso!")
                sleep(5)
                status.append("Acessar dashboard: OK")
            except Exception as e:
                status.append(f"Acessar dashboard: Erro - {e}")
                return status

            # Remover até 5 relatórios antigos (estoque)
            try:
                for _ in range(5):
                    # Sempre busca o primeiro "Remover" disponível
                    remover = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'upper') and contains(@class, '_mlxs') and text()='Remover']"))
                    )
                    remover.click()
                    sleep(1)
                    try:
                        remover_confirm_btn = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn orange') and text()='REMOVER']"))
                        )
                        remover_confirm_btn.click()
                        sleep(1)
                    except Exception as e:
                        status.append(f"Remover: Erro ao confirmar - {e}")
                status.append("Remover relatórios antigos: OK (até 5 removidos)")
            except Exception as e:
                status.append(f"Remover relatórios antigos: Erro - {e}")

            try:
                gerar_btn = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, '_c-btn--primary') and contains(text(), 'Gerar Relatório')]"))
                )
                gerar_btn.click()
                print("Botão 'Gerar Relatório' clicado.")
                sleep(2)
                status.append("Gerar Relatório: OK")
            except Exception as e:
                status.append(f"Gerar Relatório: Erro - {e}")
                return status

            # Botão "ENTENDI" é opcional
            try:
                entendi_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn blue') and contains(text(), 'ENTENDI')]"))
                )
                entendi_btn.click()
                sleep(1)
                status.append("Entendi: OK")
            except Exception as e:
                status.append(f"Entendi: Não exibido ou não clicável - {e}")

            try:
                encontrou_data = False
                while True:
                    try:
                        data_element = driver.find_element(By.XPATH, f"//p[contains(text(), '{data_hoje}')]")
                        if data_element.is_displayed():
                            encontrou_data = True
                            break
                    except:
                        pass
                    try:
                        refresh_btn = driver.find_element(By.XPATH, "//i[contains(@class, 'material-icons') and text()='refresh']")
                        refresh_btn.click()
                    except Exception as e:
                        status.append(f"Refresh: Erro ao clicar - {e}")
                    sleep(3)
                if encontrou_data:
                    status.append("Data relatório: OK")
                else:
                    status.append("Data relatório: Não encontrada")
            except Exception as e:
                status.append(f"Data relatório: Erro - {e}")
                return status

            try:
                baixar_btn = driver.find_element(By.XPATH, f"//p[contains(text(), '{data_hoje}')]/following::a[contains(@href, 'blob.core.windows.net/relatorios')][1]")
                arquivo_url = baixar_btn.get_attribute("href")
                baixar_btn.click()
                print("Relatório baixado com sucesso!")
                sleep(5)
                status.append("Baixar relatório: OK")
            except Exception as e:
                status.append(f"Baixar relatório: Erro - {e}")
                return status

            try:
                download_dir = str(Path.home() / "Downloads")
                destino = r"G:\Meu Drive\Myrp\Relatórios Grupos Loja\Estoque"
                timeout = 60
                arquivo_baixado = None
                for _ in range(timeout):
                    arquivos = glob.glob(os.path.join(download_dir, "*.xlsx"))
                    if arquivos:
                        arquivo_baixado = max(arquivos, key=os.path.getctime)
                        if not arquivo_baixado.endswith(".crdownload"):
                            break
                    sleep(1)
                if not arquivo_baixado or arquivo_baixado.endswith(".crdownload"):
                    status.append("Arquivo baixado: Não encontrado ou download não finalizado.")
                    return status
                shutil.copy2(arquivo_baixado, destino)
                estoque_atual_path = os.path.join(destino, "Estoque Atual.xlsx")
                shutil.copy2(arquivo_baixado, estoque_atual_path)
                status.append("Arquivo movido/copias: OK")
                # Remove o arquivo baixado da pasta de downloads
                try:
                    os.remove(arquivo_baixado)
                except Exception as e:
                    status.append(f"Erro ao remover arquivo dos downloads: {e}")
            except Exception as e:
                status.append(f"Arquivo movido/copias: Erro - {e}")
                return status
        status.append("Estoque Atual.xlsx - ok")
        rels_status["estoque"] = "ok"

    # --- Relatório de Venda Sintético ---
    if "venda" in relatorios_a_gerar:
        now = datetime.now()
        ano_atual = now.strftime("%Y")
        mes_atual = now.strftime("%m")
        nome_venda = f"{ano_atual}_Rel_Venda_Sint_Andre_Sborz.xlsx"
        caminho_venda = os.path.join(destino_venda, nome_venda)
        gerar_venda = True

        if os.path.exists(caminho_venda):
            data_mod = datetime.fromtimestamp(os.path.getmtime(caminho_venda)).strftime("%d/%m/%Y")
            if data_mod == data_hoje:
                status.append(f"{nome_venda} - Já gerado hoje")
                gerar_venda = False

        # Sempre acessa a página de venda para garantir o fluxo do relatório analítico
        try:
            venda_url = "https://hering.myrp.app/app/gerencial/relatorios/relatoriosAvancados/gerar/?tipo=venda"
            driver.get(venda_url)
            sleep(5)
            status.append("Acessar dashboard venda: OK")
        except Exception as e:
            status.append(f"Acessar dashboard venda: Erro - {e}")
            return status

        if gerar_venda:
            # Remover até 5 relatórios antigos de venda
            try:
                for _ in range(5):
                    remover = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'upper') and contains(@class, '_mlxs') and text()='Remover']"))
                    )
                    remover.click()
                    sleep(1)
                    try:
                        remover_confirm_btn = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn orange') and text()='REMOVER']"))
                        )
                        remover_confirm_btn.click()
                        sleep(1)
                    except Exception as e:
                        status.append(f"Venda - Remover: Erro ao confirmar - {e}")
                status.append("Venda - Remover relatórios antigos: OK (até 5 removidos)")
            except Exception as e:
                status.append(f"Venda - Remover relatórios antigos: Erro - {e}")

            # Seleciona o radio "vendedor"
            try:
                # Aguarda o radio estar presente e visível
                radio_vendedor = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "vendedor"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", radio_vendedor)
                # Usa JavaScript para disparar o evento de clique corretamente
                driver.execute_script("""
                    var radio = arguments[0];
                    if (!radio.checked) {
                        radio.click();
                        var evt = document.createEvent('HTMLEvents');
                        evt.initEvent('change', true, true);
                        radio.dispatchEvent(evt);
                    }
                """, radio_vendedor)
                status.append("Venda - Radio 'vendedor': OK")
                sleep(1)
            except Exception as e:
                status.append(f"Venda - Radio 'vendedor': Erro - {e}")
                return status

            # Garante que o botão "Gerar Relatório" está visível e clica via JavaScript
            try:
                gerar_btn_venda = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(@class, '_c-btn--primary') and contains(text(), 'Gerar Relatório')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", gerar_btn_venda)
                sleep(1)
                driver.execute_script("arguments[0].click();", gerar_btn_venda)
                print("Venda - Botão 'Gerar Relatório' clicado via JS.")
                sleep(2)
                status.append("Venda - Gerar Relatório: OK")
            except Exception as e:
                status.append(f"Venda - Gerar Relatório: Erro - {e}")
                return status

            # Após clicar em Gerar Relatório, tenta clicar em ENTENDI se aparecer
            try:
                entendi_btn_venda = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn blue') and contains(text(), 'ENTENDI')]"))
                )
                entendi_btn_venda.click()
                sleep(1)
                status.append("Venda - Entendi: OK")
            except Exception as e:
                status.append(f"Venda - Entendi: Não exibido ou não clicável - {e}")

            try:
                encontrou_data = False
                while True:
                    try:
                        data_element = driver.find_element(By.XPATH, f"//p[contains(text(), '{data_hoje}')]")
                        if data_element.is_displayed():
                            encontrou_data = True
                            break
                    except:
                        pass
                    try:
                        refresh_btn = driver.find_element(By.XPATH, "//i[contains(@class, 'material-icons') and text()='refresh']")
                        refresh_btn.click()
                    except Exception as e:
                        status.append(f"Venda - Refresh: Erro ao clicar - {e}")
                    sleep(3)
                if encontrou_data:
                    status.append("Venda - Data relatório: OK")
                else:
                    status.append("Venda - Data relatório: Não encontrada")
            except Exception as e:
                status.append(f"Venda - Data relatório: Erro - {e}")
                return status

            # Baixa o relatório de venda
            try:
                baixar_btn_venda = driver.find_element(By.XPATH, f"//p[contains(text(), '{data_hoje}')]/following::a[contains(@href, 'blob.core.windows.net/relatorios')][1]")
                arquivo_url_venda = baixar_btn_venda.get_attribute("href")
                baixar_btn_venda.click()
                print("Venda - Relatório baixado com sucesso!")
                sleep(5)
                status.append("Venda - Baixar relatório: OK")
            except Exception as e:
                status.append(f"Venda - Baixar relatório: Erro - {e}")
                return status

            # Move e renomeia o arquivo baixado para a pasta de venda
            try:
                download_dir = str(Path.home() / "Downloads")
                timeout = 60
                arquivo_baixado_venda = None
                for _ in range(timeout):
                    arquivos = glob.glob(os.path.join(download_dir, "*.xlsx"))
                    if arquivos:
                        arquivo_baixado_venda = max(arquivos, key=os.path.getctime)
                        if not arquivo_baixado_venda.endswith(".crdownload"):
                            break
                    sleep(1)
                if not arquivo_baixado_venda or arquivo_baixado_venda.endswith(".crdownload"):
                    status.append("Venda - Arquivo baixado: Não encontrado ou download não finalizado.")
                    return status
                nome_destino = os.path.join(destino_venda, nome_venda)
                shutil.copy2(arquivo_baixado_venda, nome_destino)
                status.append(f"Venda - Arquivo movido/renomeado: OK ({nome_venda})")
                try:
                    os.remove(arquivo_baixado_venda)
                except Exception as e:
                    status.append(f"Venda - Erro ao remover arquivo dos downloads: {e}")
            except Exception as e:
                status.append(f"Venda - Arquivo movido/renomeado: Erro - {e}")
                return status
        status.append(f"{nome_venda} - ok")
        rels_status["venda"] = "ok"

    # --- Relatório de Venda Analítico mês atual ---
    if "analitico" in relatorios_a_gerar:
        nome_analitico = f"{ano_atual}_{mes_atual}_Rel_Venda_Anali_André_Sborz_01-{mes_atual}-{ano_atual}.xlsx"
        caminho_analitico = os.path.join(destino_venda, nome_analitico)
        caminho_analitico_final = os.path.join(destino_analitico, nome_analitico)
        gerar_analitico = True

        if os.path.exists(caminho_analitico_final):
            data_mod = datetime.fromtimestamp(os.path.getmtime(caminho_analitico_final)).strftime("%d/%m/%Y")
            if data_mod == data_hoje:
                status.append(f"{nome_analitico} - Já gerado hoje")
                gerar_analitico = False

        if gerar_analitico:
            # Remove até 2 relatórios antigos (analítico)
            try:
                for _ in range(2):
                    remover = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'upper') and contains(@class, '_mlxs') and text()='Remover']"))
                    )
                    remover.click()
                    sleep(1)
                    remover_confirm_btn = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn orange') and text()='REMOVER']"))
                    )
                    remover_confirm_btn.click()
                    sleep(1)
                status.append("Venda - Remover relatórios analíticos antigos: OK (2 removidos)")
            except Exception as e:
                status.append(f"Venda - Remover relatórios analíticos antigos: Erro - {e}")

            # Seleciona "Analítico" no select
            try:
                select = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//select[@id][@name='tipoRelatorio']"))
                )
                driver.execute_script("arguments[0].value = '2'; arguments[0].dispatchEvent(new Event('change'));", select)
                sleep(1)
                status.append("Venda - Select Analítico: OK")
            except Exception as e:
                status.append(f"Venda - Select Analítico: Erro - {e}")
                return status

            # Seleciona "Esse mês" no select de período (corrigido para qualquer id dinâmico)
            try:
                select_mes = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//select[contains(@class, 'browser-default') and contains(@class, '_bglightGray') and option[@value='este_mes']]"))
                )
                driver.execute_script("""
                    var sel = arguments[0];
                    sel.value = 'este_mes';
                    for (var i = 0; i < sel.options.length; i++) {
                        sel.options[i].selected = sel.options[i].value === 'este_mes';
                    }
                    var evt = document.createEvent('HTMLEvents');
                    evt.initEvent('change', true, true);
                    sel.dispatchEvent(evt);
                """, select_mes)
                sleep(1)
                status.append("Venda - Select período 'Esse mês': OK")
            except Exception as e:
                status.append(f"Venda - Select período 'Esse mês': Erro - {e}")
                return status

            # Clica no botão "Gerar Relatório" analítico
            try:
                gerar_btn_analitico = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(@class, '_c-btn--primary') and contains(text(), 'Gerar Relatório')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", gerar_btn_analitico)
                sleep(1)
                driver.execute_script("arguments[0].click();", gerar_btn_analitico)
                print("Venda - Botão 'Gerar Relatório' Analítico clicado via JS.")
                sleep(2)
                status.append("Venda - Gerar Relatório Analítico: OK")
            except Exception as e:
                status.append(f"Venda - Gerar Relatório Analítico: Erro - {e}")
                return status

            # Após clicar em Gerar Relatório, tenta clicar em ENTENDI se aparecer
            try:
                entendi_btn_analitico = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn blue') and contains(text(), 'ENTENDI')]"))
                )
                entendi_btn_analitico.click()
                sleep(1)
                status.append("Venda - Entendi Analítico: OK")
            except Exception as e:
                status.append(f"Venda - Entendi Analítico: Não exibido ou não clicável - {e}")

            # Aguarda o relatório analítico do mês atual aparecer e baixa o arquivo
            try:
                encontrou_data = False
                data_hoje = now.strftime("%d/%m/%Y")
                baixar_btn_analitico = None
                for _ in range(40):  # tenta por até ~2 minutos
                    try:
                        baixar_btn_analitico = driver.find_element(By.XPATH, f"//p[contains(text(), '{data_hoje}')]/following::a[contains(@href, 'blob.core.windows.net/relatorios')][1]")
                        if baixar_btn_analitico.is_displayed():
                            encontrou_data = True
                            break
                    except:
                        pass
                    try:
                        refresh_btn = driver.find_element(By.XPATH, "//i[contains(@class, 'material-icons') and text()='refresh']")
                        refresh_btn.click()
                    except Exception as e:
                        status.append(f"Venda - Refresh Analítico: Erro ao clicar - {e}")
                    sleep(3)
                if not encontrou_data:
                    status.append("Venda - Data relatório Analítico: Não encontrada")
                    return status
                status.append("Venda - Data relatório Analítico: OK")
                # Baixa o relatório analítico
                baixar_btn_analitico.click()
                print("Venda - Relatório Analítico baixado com sucesso!")
                sleep(5)
                status.append("Venda - Baixar relatório Analítico: OK")
            except Exception as e:
                status.append(f"Venda - Baixar relatório Analítico: Erro - {e}")
                return status

            # Move e renomeia o arquivo baixado para a pasta correta com nome correto
            try:
                # Garante que download_dir está definido
                download_dir = str(Path.home() / "Downloads")
                timeout = 60
                arquivo_baixado_analitico = None
                for _ in range(timeout):
                    arquivos = glob.glob(os.path.join(download_dir, "*.xlsx"))
                    if arquivos:
                        arquivo_baixado_analitico = max(arquivos, key=os.path.getctime)
                        if not arquivo_baixado_analitico.endswith(".crdownload"):
                            break
                    sleep(1)
                if not arquivo_baixado_analitico or arquivo_baixado_analitico.endswith(".crdownload"):
                    status.append("Venda - Arquivo analítico baixado: Não encontrado ou download não finalizado.")
                    return status
                # Gera o nome correto YYYY_MM_Rel_Venda_Anali_André_Sborz_01-MM-YYYY.xlsx com base na data de geração
                data_arquivo = datetime.now()
                ano_corrente = data_arquivo.strftime("%Y")
                mes_corrente = data_arquivo.strftime("%m")
                nome_analitico_corrigido = f"{ano_corrente}_{mes_corrente}_Rel_Venda_Anali_André_Sborz_01-{mes_corrente}-{ano_corrente}.xlsx"
                caminho_analitico_final = os.path.join(destino_analitico, nome_analitico_corrigido)
                # Copia o arquivo baixado para a pasta correta com o nome correto (sempre sobrescreve)
                try:
                    if os.path.exists(caminho_analitico_final):
                        os.remove(caminho_analitico_final)
                    shutil.copyfile(arquivo_baixado_analitico, caminho_analitico_final)
                    os.remove(arquivo_baixado_analitico)
                    status.append(f"Venda - Arquivo analítico movido/renomeado: OK ({nome_analitico_corrigido})")
                except Exception as e:
                    status.append(f"Venda - Erro ao copiar/renomear arquivo analítico: {e}")
                    return status
            except Exception as e:
                status.append(f"Venda - Arquivo analítico movido/renomeado: Erro - {e}")
                return status
        status.append(f"{nome_analitico} - ok")
        rels_status["analitico"] = "ok"

    # --- Relatório de Venda Analítico mês anterior ---
    if "analitico_anterior" in relatorios_a_gerar:
        nome_analitico = nome_analitico_anterior
        caminho_analitico_final = caminho_analitico_anterior

        # Define ano_anterior para uso no restante do fluxo
        ano_anterior = ano_anterior_arquivo

        # Acessa a página do relatório analítico (igual ao mês atual)
        try:
            venda_url = "https://hering.myrp.app/app/gerencial/relatorios/relatoriosAvancados/gerar/?tipo=venda"
            driver.get(venda_url)
            sleep(5)
            status.append("Acessar dashboard analítico mês anterior: OK")
        except Exception as e:
            status.append(f"Acessar dashboard analítico mês anterior: Erro - {e}")
            return status

        # Remove até 2 relatórios antigos (analítico mês anterior)
        try:
            for _ in range(2):
                remover = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'upper') and contains(@class, '_mlxs') and text()='Remover']"))
                )
                remover.click()
                sleep(1)
                remover_confirm_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn orange') and text()='REMOVER']"))
                )
                remover_confirm_btn.click()
                sleep(1)
            status.append("Venda - Remover relatórios analíticos antigos mês anterior: OK (2 removidos)")
        except Exception as e:
            status.append(f"Venda - Remover relatórios analíticos antigos mês anterior: Erro - {e}")

        # Seleciona "Analítico" no select
        try:
            select = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//select[@id][@name='tipoRelatorio']"))
            )
            driver.execute_script("arguments[0].value = '2'; arguments[0].dispatchEvent(new Event('change'));", select)
            sleep(1)
            status.append("Venda - Select Analítico mês anterior: OK")
        except Exception as e:
            status.append(f"Venda - Select Analítico mês anterior: Erro - {e}")
            return status

        # Seleciona "Mês passado" no select de período (value="mes_passado")
        try:
            select_mes = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//select[contains(@class, 'browser-default') and contains(@class, '_bglightGray')]"))
            )
            driver.execute_script("""
                var sel = arguments[0];
                sel.value = 'mes_passado';
                for (var i = 0; i < sel.options.length; i++) {
                    sel.options[i].selected = sel.options[i].value === 'mes_passado';
                }
                var evt = document.createEvent('HTMLEvents');
                evt.initEvent('change', true, true);
                sel.dispatchEvent(evt);
            """, select_mes)
            sleep(1)
            status.append("Venda - Select período 'Mês passado': OK")
        except Exception as e:
            status.append(f"Venda - Select período 'Mês passado': Erro - {e}")
            return status

        # Clica no botão "Gerar Relatório" analítico mês anterior
        try:
            gerar_btn_analitico = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(@class, '_c-btn--primary') and contains(text(), 'Gerar Relatório')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", gerar_btn_analitico)
            sleep(1)
            driver.execute_script("arguments[0].click();", gerar_btn_analitico)
            print("Venda - Botão 'Gerar Relatório' Analítico mês anterior clicado via JS.")
            sleep(2)
            status.append("Venda - Gerar Relatório Analítico mês anterior: OK")
        except Exception as e:
            status.append(f"Venda - Gerar Relatório Analítico mês anterior: Erro - {e}")
            return status

        # Após clicar em Gerar Relatório, tenta clicar em ENTENDI se aparecer
        try:
            entendi_btn_analitico = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn blue') and contains(text(), 'ENTENDI')]"))
            )
            entendi_btn_analitico.click()
            sleep(1)
            status.append("Venda - Entendi Analítico mês anterior: OK")
        except Exception as e:
            status.append(f"Venda - Entendi Analítico mês anterior: Não exibido ou não clicável - {e}")


        encontrou_data = False
        baixar_btn_analitico = None
        data_hoje = now.strftime("%d/%m/%Y")  # Adicione esta linha para garantir a data correta
        for _ in range(40):  # tenta por até ~2 minutos
            try:
                baixar_btn_analitico = driver.find_element(
                    By.XPATH,
                    f"//p[contains(text(), '{data_hoje}')]/following::a[contains(@href, 'blob.core.windows.net/relatorios')][1]"
                )
                if baixar_btn_analitico.is_displayed():
                    encontrou_data = True
                    break
            except:
                pass
            try:
                refresh_btn = driver.find_element(By.XPATH, "//i[contains(@class, 'material-icons') and text()='refresh']")
                refresh_btn.click()
            except Exception as e:
                status.append(f"Venda - Refresh Analítico mês anterior: Erro ao clicar - {e}")
            sleep(3)
        if not encontrou_data:
            status.append("Venda - Data relatório Analítico mês anterior: Não encontrada")
            return status
        status.append("Venda - Data relatório Analítico mês anterior: OK")
        # Baixa o relatório analítico mês anterior (igual ao analítico atual)
        try:
            baixar_btn_analitico.click()
            print("Venda - Relatório Analítico mês anterior baixado com sucesso!")
            sleep(5)
            status.append("Venda - Baixar relatório Analítico mês anterior: OK")
        except Exception as e:
            status.append(f"Venda - Baixar relatório Analítico mês anterior: Erro ao clicar - {e}")
            return status
        # Move e renomeia o arquivo baixado para a pasta correta com nome correto do mês anterior
        try:
            download_dir = str(Path.home() / "Downloads")
            timeout = 60
            arquivo_baixado_analitico = None
            for _ in range(timeout):
                arquivos = glob.glob(os.path.join(download_dir, "*.xlsx"))
                if arquivos:
                    arquivo_baixado_analitico = max(arquivos, key=os.path.getctime)
                    if not arquivo_baixado_analitico.endswith(".crdownload"):
                        break
                sleep(1)
            if not arquivo_baixado_analitico or arquivo_baixado_analitico.endswith(".crdownload"):
                status.append("Venda - Arquivo analítico mês anterior baixado: Não encontrado ou download não finalizado.")
                return status
            nome_analitico_corrigido = f"{ano_anterior}_{mes_anterior}_Rel_Venda_Anali_André_Sborz_01-{mes_anterior}-{ano_anterior}.xlsx"
            caminho_analitico_final = os.path.join(destino_analitico, nome_analitico_corrigido)
            try:
                if os.path.exists(caminho_analitico_final):
                    os.remove(caminho_analitico_final)
                shutil.copyfile(arquivo_baixado_analitico, caminho_analitico_final)
                os.remove(arquivo_baixado_analitico)
                status.append(f"{nome_analitico_corrigido} - ok")
            except Exception as e:
                status.append(f"Venda - Erro ao copiar/renomear arquivo analítico mês anterior: {e}")
                return status
        except Exception as e:
            status.append(f"Venda - Arquivo analítico mês anterior movido/renomeado: Erro - {e}")
            return status

        rels_status["analitico_anterior"] = "ok"

    # --- Relatório de Venda Sintético Ano Anterior ---
    if "venda_ano_passado" in relatorios_a_gerar:
        try:
            venda_url = "https://hering.myrp.app/app/gerencial/relatorios/relatoriosAvancados/gerar/?tipo=venda"
            driver.get(venda_url)
            sleep(5)
            status.append("Acessar dashboard venda ano anterior: OK")
        except Exception as e:
            status.append(f"Venda - Acessar dashboard venda ano anterior: Erro - {e}")
            return status

        # Remove até 2 relatórios antigos (opcional)
        try:
            for _ in range(2):
                remover = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'upper') and contains(@class, '_mlxs') and text()='Remover']"))
                )
                remover.click()
                sleep(1)
                remover_confirm_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn orange') and text()='REMOVER']"))
                )
                remover_confirm_btn.click()
                sleep(1)
            status.append("Venda - Remover relatórios antigos ano anterior: OK (2 removidos)")
        except Exception as e:
            status.append(f"Venda - Remover relatórios antigos ano anterior: Erro - {e}")

        # Seleciona o radio "vendedor"
        try:
            radio_vendedor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "vendedor"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", radio_vendedor)
            driver.execute_script("""
                var radio = arguments[0];
                if (!radio.checked) {
                    radio.click();
                    var evt = document.createEvent('HTMLEvents');
                    evt.initEvent('change', true, true);
                    radio.dispatchEvent(evt);
                }
            """, radio_vendedor)
            status.append("Venda - Radio 'vendedor' ano anterior: OK")
            sleep(1)
        except Exception as e:
            status.append(f"Venda - Radio 'vendedor' ano anterior: Erro - {e}")
            return status

        # Seleciona "Ano passado" no select de período
        try:
            select_mes = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//select[contains(@class, 'browser-default') and contains(@class, '_bglightGray')]"))
            )
            driver.execute_script("""
                var sel = arguments[0];
                sel.value = 'ano_passado';
                for (var i = 0; i < sel.options.length; i++) {
                    sel.options[i].selected = sel.options[i].value === 'ano_passado';
                }
                var evt = document.createEvent('HTMLEvents');
                evt.initEvent('change', true, true);
                sel.dispatchEvent(evt);
            """, select_mes)
            sleep(1)
            status.append("Venda - Select período 'Ano passado': OK")
        except Exception as e:
            status.append(f"Venda - Select período 'Ano passado': Erro - {e}")
            return status

        # Clica no botão "Gerar Relatório"
        try:
            gerar_btn_venda = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(@class, '_c-btn--primary') and contains(text(), 'Gerar Relatório')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", gerar_btn_venda)
            sleep(1)
            driver.execute_script("arguments[0].click();", gerar_btn_venda)
            print("Venda - Botão 'Gerar Relatório' ano anterior clicado via JS.")
            sleep(2)
            status.append("Venda - Gerar Relatório ano anterior: OK")
        except Exception as e:
            status.append(f"Venda - Gerar Relatório ano anterior: Erro - {e}")
            return status

        # Após clicar em Gerar Relatório, tenta clicar em ENTENDI se aparecer
        try:
            entendi_btn_venda = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn blue') and contains(text(), 'ENTENDI')]"))
            )
            entendi_btn_venda.click()
            sleep(1)
            status.append("Venda - Entendi ano anterior: OK")
        except Exception as e:
            status.append(f"Venda - Entendi ano anterior: Não exibido ou não clicável - {e}")

        # Usa a mesma data do relatório de venda atual
        data_hoje_venda_atual = data_hoje

        encontrou_data = False
        baixar_btn_venda_ano_passado = None
        for _ in range(40):  # tenta por até ~2 minutos
            try:
                baixar_btn_venda_ano_passado = driver.find_element(
                    By.XPATH,
                    f"//p[contains(text(), '{data_hoje_venda_atual}')]/following::a[contains(@href, 'blob.core.windows.net/relatorios')][1]"
                )
                if baixar_btn_venda_ano_passado.is_displayed():
                    encontrou_data = True
                    break
            except:
                pass
            try:
                refresh_btn = driver.find_element(By.XPATH, "//i[contains(@class, 'material-icons') and text()='refresh']")
                refresh_btn.click()
            except Exception as e:
                status.append(f"Venda - Refresh ano anterior: Erro ao clicar - {e}")
            sleep(3)
        if not encontrou_data:
            status.append("Venda - Data relatório ano anterior: Não encontrada")
            return status
        status.append("Venda - Data relatório ano anterior: OK")
        # Baixa o relatório (igual ao venda atual)
        try:
            arquivo_url_venda_ano_passado = baixar_btn_venda_ano_passado.get_attribute("href")
            baixar_btn_venda_ano_passado.click()
            print("Venda - Relatório ano anterior baixado com sucesso!")
            sleep(5)
            status.append("Venda - Baixar relatório ano anterior: OK")
        except Exception as e:
            status.append(f"Venda - Baixar relatório ano anterior: Erro - {e}")
            return status

        # Move e renomeia o arquivo baixado para a pasta correta com nome correto do ano anterior
        try:
            download_dir = str(Path.home() / "Downloads")
            timeout = 60
            arquivo_baixado_venda_ano_passado = None
            for _ in range(timeout):
                arquivos = glob.glob(os.path.join(download_dir, "*.xlsx"))
                if arquivos:
                    arquivo_baixado_venda_ano_passado = max(arquivos, key=os.path.getctime)
                    if not arquivo_baixado_venda_ano_passado.endswith(".crdownload"):
                        break
                sleep(1)
            if not arquivo_baixado_venda_ano_passado or arquivo_baixado_venda_ano_passado.endswith(".crdownload"):
                status.append("Venda - Arquivo ano anterior baixado: Não encontrado ou download não finalizado.")
                return status
            nome_venda_corrigido = f"{ano_passado}_Rel_Venda_Sint_Andre_Sborz.xlsx"
            caminho_venda_ano_passado = os.path.join(destino_venda, nome_venda_corrigido)
            try:
                if os.path.exists(caminho_venda_ano_passado):
                    os.remove(caminho_venda_ano_passado)
                shutil.copyfile(arquivo_baixado_venda_ano_passado, caminho_venda_ano_passado)
                os.remove(arquivo_baixado_venda_ano_passado)
                status.append(f"{nome_venda_corrigido} - ok")
            except Exception as e:
                status.append(f"Venda - Erro ao copiar/renomear arquivo venda ano anterior: {e}")
                return status
        except Exception as e:
            status.append(f"Venda - Arquivo venda ano anterior movido/renomeado: Erro - {e}")
            return status

        rels_status["venda_ano_passado"] = "ok"

    # --- Relatórios do Grupo Lojas ---
    # ...código dos relatórios do Grupo Lojas...

    # Após os relatórios do Grupo Lojas, volta para a página principal para gerar por empresa
    grupo_lojas_ok = True
    if not os.path.exists(os.path.join(destino_estoque, nome_estoque)):
        grupo_lojas_ok = False
    else:
        data_mod_estoque = datetime.fromtimestamp(os.path.getmtime(os.path.join(destino_estoque, nome_estoque))).strftime("%d/%m/%Y")
        if data_mod_estoque != datetime.now().strftime("%d/%m/%Y"):
            grupo_lojas_ok = False
    if not os.path.exists(os.path.join(destino_venda, nome_venda)):
        grupo_lojas_ok = False
    else:
        data_mod_venda = datetime.fromtimestamp(os.path.getmtime(os.path.join(destino_venda, nome_venda))).strftime("%d/%m/%Y")
        if data_mod_venda != datetime.now().strftime("%d/%m/%Y"):
            grupo_lojas_ok = False
    if not os.path.exists(os.path.join(destino_venda, nome_analitico)):
        grupo_lojas_ok = False
    else:
        data_mod_analitico = datetime.fromtimestamp(os.path.getmtime(os.path.join(destino_venda, nome_analitico))).strftime("%d/%m/%Y")
        if data_mod_analitico != datetime.now().strftime("%d/%m/%Y"):
            grupo_lojas_ok = False

    if grupo_lojas_ok:
        try:
            driver.get("https://hering.myrp.app/ERP/Dashboard?alterandoEmpresa=1")
            sleep(3)
            status.append("Acessou página principal para relatórios por empresa.")
        except Exception as e:
            status.append(f"Erro ao acessar página principal para relatórios por empresa: {e}")
            return status

        try:
            empresa_modal = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.ID, "ui-id-2"))
            )
            empresa_element = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, f"//div[@id='ui-id-2' and contains(text(), '{empresa_nome}')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", empresa_element)
            empresa_element.click()
            sleep(2)
            status.append("Selecionar empresa para relatórios individuais: OK")
        except Exception as e:
            status.append(f"Selecionar empresa para relatórios individuais: Erro - {e}")
            return status

        try:
            confirmar_element = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'selecionarEmpresa')]"))
            )
            confirmar_element.click()
            sleep(5)
            status.append("Confirmar empresa para relatórios individuais: OK")
        except Exception as e:
            status.append(f"Confirmar empresa para relatórios individuais: Erro - {e}")
            return status

        # ...adicione aqui o fluxo para relatórios por empresa...
    return status

# Main
if __name__ == "__main__":
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # Maximiza a janela
    # chrome_options.add_argument("--headless") # Executa em background
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    driver.implicitly_wait(10)  # Espera implicitamente por 10 segundos
    try:
        resultado = autenticar(driver, url_login, usuario, senha)
        from datetime import datetime
        now = datetime.now()
        ano_atual = now.strftime("%Y")
        mes_atual = now.strftime("%m")
        # Resumo final dos relatórios
        print("\nResumo final dos relatórios:")
        def print_status(nome):
            linha = next((l for l in resultado if l.startswith(nome)), None)
            if linha:
                print(linha)
            else:
                print(f"{nome} - erro")
        print_status("Estoque Atual.xlsx")
        print_status(f"{ano_atual}_Rel_Venda_Sint_Andre_Sborz.xlsx")
        print_status(f"{ano_atual}_{mes_atual}_Rel_Venda_Anali_André_Sborz_01-{mes_atual}-{ano_atual}.xlsx")
        # Só mostra o analítico mês anterior se foi gerado agora
        nome_analitico_anterior = f"{int(ano_atual) if mes_atual != '01' else int(ano_atual)-1}_{'12' if mes_atual == '01' else f'{int(mes_atual)-1:02d}'}_Rel_Venda_Anali_André_Sborz_01-{'12' if mes_atual == '01' else f'{int(mes_atual)-1:02d}'}-{int(ano_atual) if mes_atual != '01' else int(ano_atual)-1}.xlsx"
        if any(l for l in resultado if l.startswith(nome_analitico_anterior) and "- ok" in l):
            print_status(nome_analitico_anterior)
        # Só mostra o venda ano anterior se foi gerado agora
        nome_venda_ano_passado = f"{int(ano_atual)-1}_Rel_Venda_Sint_Andre_Sborz.xlsx"
        if any(l for l in resultado if l.startswith(nome_venda_ano_passado) and "- ok" in l):
            print_status(nome_venda_ano_passado)
        # Explicação dos erros encontrados
        erros = [
            linha for linha in resultado
            if (
                ("Erro" in linha or "não encontrado" in linha or "não finalizado" in linha or "Não encontrada" in linha)
                and "Remover relatórios antigos" not in linha
                and "Remover relatórios analíticos antigos" not in linha
                and "Created TensorFlow Lite XNNPACK delegate for CPU." not in linha
                and "Refresh ano anterior: Erro ao clicar" not in linha  # ignora erro irrelevante de refresh
            )
        ]
        if erros:
            print("\nDetalhes dos erros encontrados:")
            for linha in erros:
                print(linha)
        sleep(20)  # Mantém o navegador aberto por 20 segundos após o login
    except Exception as e:
        print(f"Erro inesperado na execução principal: {e}")
    finally:
        driver.quit()
