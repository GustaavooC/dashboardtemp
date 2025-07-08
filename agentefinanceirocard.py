import asyncio
from playwright.async_api import async_playwright
import os
from datetime import date

# Configurações
LOGIN_URL = "https://www.obraprimaweb.com.br/login"
FINANCE_URL = "https://www.obraprimaweb.com.br/financeiro/documentos"
EMAIL = "dados.caicara@gmail.com"
PASSWORD = "ObraPrim@2025#."
DOWNLOAD_DIR = r"C:\Users\CAICARA\Documents\teste"


async def fill_date_field(page, selector, date_value):
    """Preenche um campo de data com fallback em JS."""
    print(f"🟡 Preenchendo campo {selector} com {date_value}...")
    try:
        await page.click(selector)
        await asyncio.sleep(0.3)
        await page.fill(selector, "")
        await page.type(selector, date_value, delay=50)
        await asyncio.sleep(0.5)

        filled = await page.input_value(selector)
        if filled == date_value:
            print(f"🟢 {selector} preenchido com sucesso.")
            return True

        # Fallback JS
        await page.evaluate("""([sel, dt]) => {
            const el = document.querySelector(sel);
            if (el) {
                el.value = dt;
                ['input','change','blur'].forEach(ev => el.dispatchEvent(new Event(ev, { bubbles: true })));
            }
        }""", [selector, date_value])
        await asyncio.sleep(0.5)

        filled = await page.input_value(selector)
        if filled != date_value:
            raise Exception(f"Falha ao preencher {selector}. Esperado: {date_value}, Obtido: {filled}")

        print(f"🟢 {selector} preenchido via fallback.")
        return True

    except Exception as e:
        print(f"🔴 Erro ao preencher {selector}: {str(e)}")
        raise


async def close_modals(page):
    """Fechamento apenas do Modal 1"""
    print("\n🔵 INICIANDO FECHAMENTO DE MODAIS")

    try:
        modal1_selector = "a#pushActionRefuse"
        await page.wait_for_selector(modal1_selector, timeout=10000)
        print("🟡 Fechando Modal 1 (Não, obrigado)...")
        await page.click(modal1_selector)
        await asyncio.sleep(3)
    except Exception as e:
        print(f"🔴 Erro no Modal 1: {str(e)}")

    print("🟢 Modais tratados\n")


async def desmarcar_checkbox_vencimento(page):
    """Desmarca o checkbox 'Vencimento' se estiver marcado."""
    print("🟡 Verificando checkbox 'Vencimento' para desmarcar...")
    try:
        is_checked = await page.is_checked("#cphPesquisa_chkDocumentosDataVencimento")
        if is_checked:
            print("🟡 Vencimento está marcado. Desmarcando...")
            await page.click("label[for='cphPesquisa_chkDocumentosDataVencimento']")
            await asyncio.sleep(1)
        else:
            print("🟢 Vencimento já está desmarcado.")
    except Exception as e:
        print(f"🔴 Erro ao desmarcar Vencimento: {str(e)}")


async def marcar_checkbox_emissao(page):
    """Marca o checkbox 'Emissão' clicando no label ou via fallback JS."""
    print("🟡 Tentando marcar o checkbox 'Emissão'...")
    try:
        await page.click("label[for='cphPesquisa_chkDocumentosDataEmissao']")
        await asyncio.sleep(1)
        print("🟢 Checkbox marcado via label.")
    except Exception as e:
        print(f"🟡 Falhou o clique no label. Usando fallback JS. Erro: {str(e)}")
        await page.evaluate("""
            () => {
                const cb = document.getElementById('cphPesquisa_chkDocumentosDataEmissao');
                if (cb) {
                    cb.checked = true;
                    ['input','change','blur'].forEach(ev => cb.dispatchEvent(new Event(ev, { bubbles: true })));
                }
            }
        """)
        await asyncio.sleep(1)
        print("🟢 Checkbox marcado via fallback JS.")


async def selecionar_cartao_bv(page, termo_digitado, primeira_vez):
    """
    Primeira vez: TAB 16x até o campo, ENTER para abrir, digita termo, ENTER para selecionar.
    Segunda vez: direto digita termo, seta para baixo, ENTER para selecionar.
    """
    if primeira_vez:
        print(f"🟡 Navegando com TAB até o campo Cartão BV Elo...")
        for _ in range(16):
            await page.keyboard.press("Tab")
            await asyncio.sleep(0.1)

        await page.keyboard.press("Enter")
        await asyncio.sleep(0.5)

    print(f"🟡 Digitando '{termo_digitado}' no campo de busca...")
    await page.keyboard.type(termo_digitado)
    await asyncio.sleep(0.5)

    if primeira_vez:
        await page.keyboard.press("Enter")
    else:
        await page.keyboard.press("ArrowDown")
        await asyncio.sleep(0.2)
        await page.keyboard.press("Enter")

    await asyncio.sleep(1.5)
    print(f"🟢 '{termo_digitado}' selecionado com sucesso via teclado.")


async def baixar_financeiro_documentos(page):
    print("\n🔵 ACESSANDO GUIA FINANCEIRO / DOCUMENTOS")
    current_date = date.today().strftime("%d/%m/%Y")

    # 1. Acessar a página
    await page.goto(FINANCE_URL, timeout=60000)
    await asyncio.sleep(2)

    # 2. Fechar eventuais modais
    await close_modals(page)

    # 3. Abrir filtros
    print("🟡 Abrindo filtros...")
    await page.click("span.barra-filtro-pesquisa-texto")
    await asyncio.sleep(2)

    # 4. Desmarcar Vencimento e marcar Emissão
    await desmarcar_checkbox_vencimento(page)
    await marcar_checkbox_emissao(page)
    await asyncio.sleep(1)

    # 5. Preencher datas
    print("🟡 Preenchendo datas...")
    await fill_date_field(page, "#cphPesquisa_txtDocumentosDataInicial", current_date)
    await fill_date_field(page, "#cphPesquisa_txtDocumentosDataFinal", current_date)
    await asyncio.sleep(1)

    # 6. Scroll até o fim do filtro
    print("🟡 Fazendo scroll até o final...")
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await asyncio.sleep(2)

    # 7. Selecionar cartões via teclado
    await selecionar_cartao_bv(page, "BV", primeira_vez=True)
    await selecionar_cartao_bv(page, "BV", primeira_vez=False)

    # 8. Clicar em Pesquisar
    print("🟡 Clicando em Pesquisar...")
    await page.click("a.lnkPesquisar")
    await asyncio.sleep(5)

    # 9. Exportar → CSV
    print("🟡 Clicando em Exportar...")
    await page.click("a.btn_exportar")
    await asyncio.sleep(2)

    print("🟡 Selecionando CSV...")
    async with page.expect_download(timeout=60000) as download_info:
        await page.click("#cphConteudo_lnkExportCsvFinanceiro")
    download = await download_info.value

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    filename = os.path.join(DOWNLOAD_DIR, "documentos_financeiro.csv")
    await download.save_as(filename)
    print(f"\n🟢 DOWNLOAD CONCLUÍDO! Arquivo salvo em: {filename}\n")


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
            # LOGIN
            print("🔵 LOGIN")
            await page.goto(LOGIN_URL, timeout=60000)
            await page.fill("input[placeholder='E-mail']", EMAIL)
            await page.fill("input[placeholder='Senha']", PASSWORD)
            await page.click("text='ACESSAR'")
            await page.wait_for_url("https://www.obraprimaweb.com.br/inicio", timeout=30000)
            print("🟢 Login realizado com sucesso.")

            # BAIXAR DOCUMENTOS FINANCEIROS
            await baixar_financeiro_documentos(page)

        except Exception as e:
            print(f"\n🔴 ERRO CRÍTICO: {str(e)}")
        finally:
            await browser.close()


if __name__ == '__main__':
    asyncio.run(main())
