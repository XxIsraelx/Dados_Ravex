import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# Função para remover pontuação
def remover_pontuacao(texto):
    return re.sub(r'[^\w\s]', '', texto)

# Configurar o driver
driver = webdriver.Chrome()
driver.get("https://app.ravex.com.br/logistica/")
sleep(2)  # Espera o site carregar

# Credenciais
email = 'jonathansantos175@gmail.com'
senha = 'Omega.2020'

# Login
driver.find_element(By.ID, "edtLogin").send_keys(email)
sleep(0.5)
driver.find_element(By.ID, "edtSenha").send_keys(senha)
sleep(0.5)
driver.find_element(By.ID, "btnAutenticar").click()
sleep(18)  # Tempo necessário para carregar a página pós-login

# Definir cabeçalhos personalizados e remover pontuação
cabecalhos = [
    "Status", "Entrega", "Sequencia", "Nota", "Tipo de operacao", "Serie", "Pedido", "Cliente", "Endereco", "Cidade",
    "Bairro", "Peso", "Peso Liquido", "Valor", "Est Entrega", "Inicio Entrega NF", "Inicio Descarga", "Fim Descarga",
    "Tempo Descarga", "Fim Entrega NF", "Tempo Entrega", "Qtde Itens", "Cod Ano", "Anomalia",
    "Id do vendedor", "Vendedor", "Fone Vendedor", "Email Vendedor", "Supervisor", "Fone Supervisor",
    "Email Supervisor", "Gerente", "Fone Gerente", "Email Gerente", "Promotor", "Fone Promotor", "Email Promotor",
    "Ocorrencia externa", "Senha", "Baixa", "Peso Devolvido", "Viagem de Origem", "Viagem de Destino",
    "Tipo de Cliente", "Remetente", "Metros da Entrega", "Nota Aderente", "CTE", "Valor CTE", "Emissao CTE",
    "Imagem", "Serie Pernoite", "Identificador de agrupamento", "Local de entrega", "Projeto", "Tipo de pedido",
    "Tipo de carga", "Divisao de negocio", "Data do faturamento", "Data de criacao"
]
cabecalhos = [remover_pontuacao(cab) for cab in cabecalhos]

# Esperar até que o tbody e as linhas sejam carregados
wait = WebDriverWait(driver, 10)
dados_completos = []
contagem_sucesso = 0  # Contador de <tr> com dados extraídos

try:
    tbody = wait.until(EC.presence_of_element_located((By.TAG_NAME, "tbody")))
    linhas_principais = tbody.find_elements(By.XPATH, "./tr")  # Melhorando a busca dentro do tbody

    for linha_principal in linhas_principais:
        try:
            linha_principal.click()
            div_element = wait.until(EC.element_to_be_clickable((By.ID, "btnNotasItens")))
            div_element.click()
            sleep(2)  # Mantendo o sleep necessário

            # Coletar apenas as linhas visíveis após a interação
            linhas = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tr[contains(@class, 'k-master-row')]")))

            extraiu_dados = False  # Flag para verificar se ao menos um dado foi extraído

            for linha in linhas:
                colunas = linha.find_elements(By.XPATH, ".//td")

                # Remover colunas indesejadas diretamente com list comprehension + enumerate (mais eficiente)
                dados = [col.text.strip() if col.text.strip() else "0" for i, col in enumerate(colunas) if i >= 2 and i not in {23, 27}]

                if dados:
                    dados_completos.append(dados)
                    extraiu_dados = True

            if extraiu_dados:
                contagem_sucesso += 1  # Contabiliza o <tr> que teve extração de dados bem-sucedida

        except Exception as e:
            print(f"Erro ao processar linha: {e}")

    # Criar único arquivo Excel com todos os dados
    if dados_completos:
        df = pd.DataFrame(dados_completos, columns=cabecalhos[:len(dados_completos[0])])

        # Criar planilha
        wb = Workbook()
        ws = wb.active
        ws.append(df.columns.tolist())

        # Adicionar os dados
        for row in df.itertuples(index=False):
            ws.append(row)

        # Formatar cabeçalhos
        for cell in ws[1]:  
            cell.font = Font(name="Aharoni", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Ajustar largura das colunas automaticamente
        for col in ws.columns:
            max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

        # Salvar arquivo Excel
        wb.save("dados_extraidos.xlsx")
        print("Dados exportados com sucesso para 'dados_extraidos.xlsx'")

    print(f"Total de <tr> processados com sucesso: {contagem_sucesso}")

except Exception as e:
    print(f"Erro ao tentar localizar os dados: {e}")

driver.quit()  # Fecha o navegador após a execução
