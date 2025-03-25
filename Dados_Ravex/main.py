import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

# Configurar o driver (exemplo com Chrome)
driver = webdriver.Chrome()

# Acessar o site
driver.get("https://app.ravex.com.br/logistica/")
sleep(3)  # Esperar o site carregar

# Credenciais
email = 'jonathansantos175@gmail.com'
senha = 'Omega.2020'

# Preencher o campo de usuário
usuario_pesquisa = driver.find_element(By.XPATH, "//input[@id='edtLogin']")
usuario_pesquisa.send_keys(email)
sleep(1.5)

# Preencher o campo de senha
senha_pesquisa = driver.find_element(By.XPATH, "//input[@id='edtSenha']")
senha_pesquisa.send_keys(senha)

sleep(1.5)

# Clicar no botão de login
botao_login = driver.find_element(By.XPATH, "//button[@id='btnAutenticar']")
botao_login.click()
sleep(10)

# Esperar até que as linhas <tr> com a classe 'k-master-row k-state-selected' estejam visíveis
wait = WebDriverWait(driver, 10)
try:
    # Esperar por até 10 segundos até que a linha com a classe 'k-master-row' esteja visível
    linhas = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tr[contains(@class, 'k-master-row')]")))

    # Preparar dados para exportação
    dados_exportados = []

    # Iterar por cada linha <tr> com a classe 'k-master-row'
    for linha in linhas:
        # Obter todas as células <td> dentro da linha
        colunas = linha.find_elements(By.XPATH, ".//td")

        # Extrair o conteúdo de cada <td> (tanto números quanto palavras)
        dados = []
        for coluna in colunas:
            texto_coluna = coluna.text.strip()  # Remover espaços extras
            if texto_coluna != '':  # Garantir que não estamos pegando células vazias
                dados.append(texto_coluna)

        # Adicionar os dados extraídos à lista
        if dados:  # Adicionar à lista se houver dados
            dados_exportados.append(dados)

    # Se houver dados para exportar, criar um DataFrame com o Pandas e salvar como CSV
    if dados_exportados:
        # Criar um DataFrame a partir dos dados extraídos
        df = pd.DataFrame(dados_exportados)

        # Exportar o DataFrame para um arquivo CSV
        df.to_csv('dados_extraidos.csv', index=False, header=False, encoding='utf-8')
        print("Dados exportados com sucesso para 'dados_extraidos.csv'")

except Exception as e:
    print(f"Erro ao tentar localizar os dados: {e}")

# Fechar o navegador
driver.quit()
