from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import time

# Configurar o Selenium
def configurar_navegador():
    chromedriver_path = "C:\\Users\\demat\\Documents\\Natã\\chromedriver.exe"  # Caminho correto do ChromeDriver
    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service)
    driver.get("https://www.amazon.com.br")
    return driver

# Realizar pesquisa na Amazon
def pesquisar_livros(driver, termo_pesquisa):
    try:
        barra_pesquisa = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
        )
        barra_pesquisa.send_keys(termo_pesquisa)
        barra_pesquisa.send_keys(Keys.RETURN)
        time.sleep(3)  # Aguardar carregamento da página
    except TimeoutException:
        print("Erro: Barra de pesquisa não carregou.")

# Extrair informações dos livros
def extrair_dados(driver):
    livros = []
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//div[@data-component-type='s-search-result']"))
        )
        elementos_livros = driver.find_elements(By.XPATH, "//div[@data-component-type='s-search-result']")

        for livro in elementos_livros[:8]:  # Limitar a 8 livros
            titulo = autor = preco = nota = avaliacoes = "Indisponível"

            try:
                titulo = livro.find_element(By.XPATH, ".//h2/a/span").text.strip()
            except NoSuchElementException:
                pass

            try:
                autor = livro.find_element(By.XPATH, ".//div[@class='a-row a-size-base a-color-secondary']/span[2]").text.strip()
            except NoSuchElementException:
                pass

            try:
                preco_parte1 = livro.find_element(By.XPATH, ".//span[@class='a-price-whole']").text.strip()
                preco_parte2 = livro.find_element(By.XPATH, ".//span[@class='a-price-fraction']").text.strip()
                preco = f"R${preco_parte1},{preco_parte2}"
            except NoSuchElementException:
                pass

            try:
                nota = livro.find_element(By.XPATH, ".//span[@class='a-icon-alt']").text.strip()
            except NoSuchElementException:
                pass

            try:
                avaliacoes = livro.find_element(By.XPATH, ".//span[@class='a-size-base s-underline-text']").text.strip()
            except NoSuchElementException:
                pass

            livros.append({
                "Título": titulo,
                "Autor": autor,
                "Preço": preco,
                "Nota": nota,
                "Avaliações": avaliacoes
            })

    except TimeoutException:
        print("Erro: Não foi possível carregar os resultados.")
    
    return livros

# Salvar os dados em um arquivo Excel
def salvar_dados(dados):
    if not dados:
        print("Nenhum dado disponível para salvar.")
        return
    
    df = pd.DataFrame(dados)

    # Remover 'Indisponível' apenas onde possível
    df.replace("Indisponível", "", inplace=True)

    # Ordenar os dados por Título
    df = df.sort_values(by="Título", ascending=True)

    # Salvar no Excel
    df.to_excel('dados_corrigidos.xlsx', index=False)
    print("Dados salvos no arquivo 'dados_corrigidos.xlsx'.")

# Principal
def main():
    termo_pesquisa = "livros sobre automação"  # Termo da pesquisa
    driver = configurar_navegador()
    try:
        pesquisar_livros(driver, termo_pesquisa)
        dados = extrair_dados(driver)
        salvar_dados(dados)
    finally:
        driver.quit()
        print("Execução concluída!")

if __name__ == "__main__":
    main()
