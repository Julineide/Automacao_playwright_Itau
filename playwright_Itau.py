
from dotenv import load_dotenv 
from playwright.sync_api import sync_playwright
import os

load_dotenv() # carrega o .env automaticamente

def baixar_itau_base(caminho_vdo, caminho_goal):
    """
    Baixa o arquivo Itaú_base.xlsx da plataforma usando Playwright.
    O arquivo será salvo no caminho_final.
    """

    with sync_playwright() as pw:

        navegador = pw.chromium.launch(headless=True)
        contexto = navegador.new_context(accept_downloads=True)
        
        pagina = contexto.new_page()
        pagina.goto("https://fleet.vdo-web.com")

        usuario = os.getenv("VDO_EMAIL")
        senha = os.getenv("VDO_SENHA")

        pagina.get_by_role("textbox", name="Endereço de E-mail").fill(usuario)
        pagina.get_by_role("textbox", name="Senha").fill(senha)
        pagina.get_by_role("button", name="Entrar").click()

        pagina.locator('xpath =//*[@id="retail-nav"]/a/span/i').click()
        pagina.locator('xpath =//*[@id="retail-nav"]/div/ul/li[1]/div/input').fill("itaú unibanco")
        pagina.locator('xpath =//*[@id="retail-nav"]/div/ul/li[2]/a').click()
        pagina.locator('xpath =//*[@id="vehiclevolumn-dropdown"]').click()

        with pagina.expect_download() as download_info:
            pagina.locator('xpath = //*[@id="page-content"]/div/div[1]/div[1]/div/div/div[1]/div/ul/li/a').click()

        download = download_info.value
        download.save_as(caminho_vdo)
        print(f"Arquivo VDO baixado com sucesso: {caminho_vdo}")

        pagina2 = contexto.new_page()
        pagina2.goto("https://goal.inpaas.com")

        usuario2 = os.getenv("GOAL_EMAIL")
        senha2 = os.getenv("GOAL_SENHA")

        pagina2.get_by_role("textbox", name="Nome de Usuário").fill(usuario2)
        pagina2.get_by_role("textbox", name="Senha").fill(senha2)
        pagina2.get_by_role("button", name="Entrar").click()

        pagina2.get_by_role("link").filter(has_text="Ordens de Serviço").click()
        pagina2.locator("iframe").content_frame.locator("#cmb-finder-key").select_option("70f8496b-8e8f-4917-993e-afb517d4f984")

        with pagina2.expect_download() as download_info2:
            pagina2.locator("iframe").content_frame.get_by_role("button", name="export.xlsx").click()

        download2 = download_info2.value
        download2.save_as(caminho_goal)
        print(f"Arquivo GOAL baixado com sucesso: {caminho_goal}")

        navegador.close()
