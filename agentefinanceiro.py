import asyncio
from playwright.async_api import async_playwright, TimeoutError
import os
from datetime import datetime

LOGIN_URL = "https://www.obraprimaweb.com.br/login"
REPORT_URL = "https://www.obraprimaweb.com.br/compras/ordens-de-compra"
EMAIL = "dados.caicara@gmail.com"
PASSWORD = "ObraPrim@2025#."
DOWNLOAD_DIR = r"C:\Users\CAICARA\Documents\teste"

async def take_screenshot(page, step_name):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"DEBUG_{step_name}_{timestamp}.png"
    await page.screenshot(path=filename)
    print(f"📸 Captura de tela salva: {filename}")
    return filename

async def close_modals(page):
    """Fechamento robusto dos modais"""
    print("\n🔵 INICIANDO FECHAMENTO DE MODAIS")
    
    try:
        modal1_selector = "a#pushActionRefuse"
        await page.wait_for_selector(modal1_selector, timeout=10000)
        print("🟡 Fechando Modal 1 (Não, obrigado)...")
        await page.click(modal1_selector)
        #await take_screenshot(page, "modal_1_fechado")
        await asyncio.sleep(3)
    except Exception as e:
        print(f"🔴 Erro no Modal 1: {str(e)}")

    try:
        modal2_selector = "a.lnkNotShowAgain"
        await page.wait_for_selector(modal2_selector, timeout=10000)
        print("🟡 Fechando Modal 2 (Não exibir novamente)...")
        await page.click(modal2_selector)
        #await take_screenshot(page, "modal_2_fechado")
        await asyncio.sleep(2)
    except Exception as e:
        print(f"🔴 Erro no Modal 2: {str(e)}")

    print("🟢 Modais tratados\n")

async def select_checkboxes(page):
    """Seleção das checkboxes usando os labels fornecidos"""
    print("\n🔵 SELECIONANDO CHECKBOXES VIA LABELS")
    
    checkbox_labels = {
        #"Recebida": "label[for='cphPesquisa_ucSBRS_lvPurchaseOrderStatus_chkPurchaseStatus_6']",
        #"Recusada": "label[for='cphPesquisa_ucSBRS_lvPurchaseOrderStatus_chkPurchaseStatus_7']",
        #"Cancelada": "label[for='cphPesquisa_ucSBRS_lvPurchaseOrderStatus_chkPurchaseStatus_8']",
        #"Não aprovado": "label[for='cphPesquisa_ucSBRS_lvPurchaseOrderStatus_chkPurchaseStatus_9']"
    }

    for nome, selector in checkbox_labels.items():
        try:
            print(f"🟡 Marcando: {nome}")
            await page.click(selector)
            await asyncio.sleep(1)
            
            checkbox_id = selector.split("'")[1]
            is_checked = await page.is_checked(f"#{checkbox_id}")
            
            if is_checked:
                print(f"🟢 {nome} marcado com sucesso")
                #await take_screenshot(page, f"checkbox_{nome.lower()}_sucesso")
            else:
                print(f"🟡 Tentando método alternativo para {nome}")
                await page.evaluate("""selector => {
                    const label = document.querySelector(selector);
                    if (label) {
                        const checkbox = document.getElementById(label.htmlFor);
                        if (checkbox) {
                            checkbox.checked = true;
                            const event = new Event('change', { bubbles: true });
                            checkbox.dispatchEvent(event);
                        }
                    }
                }""", selector)
                await asyncio.sleep(1)
                
                if await page.is_checked(f"#{checkbox_id}"):
                    print(f"🟢 {nome} marcado com método alternativo")
                else:
                    print(f"🔴 Falha ao marcar {nome}")

        except Exception as e:
            print(f"🔴 Erro ao marcar {nome}: {str(e)}")
            #await take_screenshot(page, f"checkbox_{nome.lower()}_erro")

async def fill_date_field(page, selector, date):
    """Preenche um campo de data com tratamento robusto"""
    print(f"🟡 Processando campo {selector}...")
    
    try:
        # Clicar no campo para ativá-lo
        await page.click(selector)
        await asyncio.sleep(0.3)
        
        # Limpar o campo
        await page.fill(selector, "")
        
        # Preencher com a data
        await page.type(selector, date, delay=50)
        await asyncio.sleep(0.5)
        
        # Verificar preenchimento
        filled_date = await page.input_value(selector)
        if filled_date == date:
            print(f"🟢 {selector} preenchido com sucesso")
            return True
        
        # Fallback com JavaScript se o método normal falhar
        print(f"🟡 Usando fallback para {selector}")
        await page.evaluate("""([sel, dt]) => {
            const field = document.querySelector(sel);
            if (field) {
                field.value = dt;
                ['input', 'change', 'blur'].forEach(ev => {
                    field.dispatchEvent(new Event(ev, { bubbles: true }));
                });
            }
        }""", [selector, date])
        await asyncio.sleep(0.5)
        
        # Verificação final
        filled_date = await page.input_value(selector)
        if filled_date != date:
            raise Exception(f"Falha ao preencher {selector}. Esperado: {date}, Obtido: {filled_date}")
        
        print(f"🟢 {selector} preenchido via fallback")
        return True
        
    except Exception as e:
        print(f"🔴 Erro ao preencher {selector}: {str(e)}")
        #await take_screenshot(page, f"erro_preencher_{selector.replace('#', '')}")
        raise

async def setup_date_filter(page):
    """Configura os filtros de data inicial e final com a data atual"""
    print("\n🔵 CONFIGURANDO FILTROS DE DATA")
    
    try:
        # Elementos e data
        date_from_selector = "#cphPesquisa_ucSBRS_txtPurchaseDateFrom"
        date_to_selector = "#cphPesquisa_ucSBRS_txtPurchaseDateEnd"
        checkbox_selector = "#cphPesquisa_ucSBRS_chkPurchaseDateDelivered"
        current_date = datetime.now().strftime("%d/%m/%Y")
        
        # 1. Ativar filtro por data
        print("🟡 Ativando filtro por data...")
        await page.evaluate("""selector => {
            const cb = document.querySelector(selector);
            if (cb) {
                cb.checked = true;
                cb.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }""", checkbox_selector)
        await asyncio.sleep(1)
        
        # 2. Preencher data inicial
        await fill_date_field(page, date_from_selector, current_date)
        
        # 3. Preencher data final
        await fill_date_field(page, date_to_selector, current_date)
        
        #await take_screenshot(page, "filtros_data_configurados")
        print("🟢 Filtros de data configurados com sucesso")
        await asyncio.sleep(1)
        
    except Exception as e:
        print(f"🔴 Erro crítico nos filtros de data: {str(e)}")
        #await take_screenshot(page, "erro_filtros_data")
        raise

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            slow_mo=1000,
            args=["--start-maximized"]
        )
        context = await browser.new_context(viewport=None)
        page = await context.new_page()

        try:
            # ETAPA 1: LOGIN
            print("🔵 ETAPA 1/6 - LOGIN")
            await page.goto(LOGIN_URL, timeout=60000)
            #await take_screenshot(page, "1_pagina_login")

            await page.fill("input[placeholder=\"E-mail\"]", EMAIL)
            await page.fill("input[placeholder=\"Senha\"]", PASSWORD)
            #await take_screenshot(page, "2_credenciais_preenchidas")

            await page.click("text=\"ACESSAR\"")
            await page.wait_for_url("https://www.obraprimaweb.com.br/inicio", timeout=30000)
            #await take_screenshot(page, "3_login_sucesso")

            # ETAPA 2: RELATÓRIOS
            print("\n🔵 ETAPA 2/6 - ACESSANDO RELATÓRIOS")
            await page.goto(REPORT_URL, timeout=60000)
            #await take_screenshot(page, "4_pagina_relatorios")

            # ETAPA 3: MODAIS
            print("\n🔵 ETAPA 3/6 - FECHANDO MODAIS")
            await close_modals(page)

            # ETAPA 4: FILTROS
            print("\n🔵 ETAPA 4/6 - ABRINDO FILTROS")
            await page.click("span.barra-filtro-pesquisa-texto")
            await asyncio.sleep(3)
            #await take_screenshot(page, "5_filtros_abertos")

            # ETAPA 5: CHECKBOXES
            await select_checkboxes(page)

            # ETAPA 5.5: FILTRO DE DATA
            await setup_date_filter(page)

            # ETAPA 6: PESQUISA E EXPORTAÇÃO
            print("\n🔵 ETAPA 6/6 - EXPORTANDO")
            await page.click("a.lnkSearch")
            #await take_screenshot(page, "6_clique_pesquisar")

            try:
                await page.wait_for_selector("#cphConteudo_lnkExportarOrdemCompra", state="visible", timeout=25000)
                #await take_screenshot(page, "7_resultados_prontos")
            except TimeoutError:
                print("🟡 Tempo de carregamento excedido - continuando...")
                #await take_screenshot(page, "7_timeout_resultados")

            await page.click("#cphConteudo_lnkExportarOrdemCompra")
            await asyncio.sleep(2)
            #take_screenshot(page, "8_menu_exportar")

            async with page.expect_download(timeout=60000) as download_info:
                await page.click("#cphConteudo_lnkGridExportCsvOrdemCompra")
            
            download = await download_info.value
            os.makedirs(DOWNLOAD_DIR, exist_ok=True)
            filename = os.path.join(DOWNLOAD_DIR, "ordens_de_compra.csv")
            await download.save_as(filename)
            
            print(f"\n🟢 PROCESSO CONCLUÍDO! Arquivo salvo como: {filename}")

        except Exception as e:
            print(f"\n🔴 ERRO: {str(e)}")
            #await take_screenshot(page, "ERRO_FINAL")
        finally:
            await browser.close()

if __name__ == '__main__':
    asyncio.run(main())