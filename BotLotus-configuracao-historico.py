from ntpath import join
import os
import shutil
import time
from datetime import datetime, timedelta
import configLotus
import unidecode

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException


class Bot():

    def __init__(self):
        self.reset_ambiente()
        self.driver = self.get_driver()
        self.dia_anterior_contratos = (datetime.now() - timedelta (1)).strftime('%d/%m/%Y')
        self.dia_anterior = (datetime.now() - timedelta (1)).strftime('%d-%m-%Y') 
        self.proximo_mes = (datetime.now() + timedelta (30)).strftime('%d/%m/%Y') 
        self.data= datetime.today().strftime('%d-%m-%Y')               
        self.run()
   
    def get_driver(self):
        chrome_options = Options()
        #chrome_options.add_argument('--headless')
        chrome_options.add_experimental_option("prefs", {"download.default_directory": configLotus.DIRETORIO_ARQUIVOS_TEMP})
        s=Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s,options=chrome_options)
        driver.maximize_window()
        return driver


    def cria_diretorio(self, caminho):
        if not os.path.isdir(caminho):
            os.mkdir(caminho)


    def reset_ambiente(self):
        list_caminhos = [
            configLotus.DIRETORIO_ARQUIVOS_TEMP
                   ]
        for caminho in list_caminhos:
            if os.path.isdir(caminho):
                shutil.rmtree(caminho)
            self.cria_diretorio(caminho)


    def login(self, usuario, senha):
        self.driver.get(configLotus.URL_SISTEMA)
        self.retorna_elemento('ID', 'j_username').send_keys(usuario)
        self.retorna_elemento('ID', 'j_password').send_keys(senha)
        self.driver.find_element(By.XPATH,'//*[@id="submit"]').click()
        #tempo necessário pois o carregamento do site demora
        time.sleep(3)
        

    def aguarda_download(self):
        seconds = 1
        dl_wait = True
        while dl_wait and seconds < 60:
            time.sleep(2)
            dl_wait = False
            for fname in os.listdir(configLotus.DIRETORIO_ARQUIVOS_TEMP):
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds += 1
        return seconds

    def renomar_arquivo(self, original, alterar):
        self.aguarda_download()
        #os.rename(original, alterar)
        shutil.move(original, alterar)
        
    
    def cria_diretorio(self, caminho):
        if not os.path.isdir(caminho):
            os.mkdir(caminho)

    
    def retorna_elemento(self, funcao, path):
        self.aguardar_elemento(funcao, path)
        return self.driver.find_element(getattr(By,funcao), path)


    def aguardar_elemento(self, funcao, path):
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((getattr(By,funcao), path)))
   
           
    def fecha_tela(self):
        self.aguarda_download
        telas = self.driver.window_handles
        time.sleep(2)
        self.driver.switch_to.window(telas[1])
        self.driver.close()
        self.driver.switch_to.window(telas[0])
    
    
    def relatorio_orcamento(self):
        
        self.driver.get('https://lotuscidade.sienge.com.br/sienge/ORC/filterRelatorioEmissaoOrcamento.do')
        self.retorna_elemento('XPATH', '//*[@id="searchObraForCadastroInsumoOrcamento"]/tbody/tr/td[6]/img[1]').click()
        iframe_orcamento = self.retorna_elemento('XPATH', '//*[@id="layerFormConsulta"]')
        self.driver.switch_to.frame(iframe_orcamento)
        obras_orcamento = self.driver.find_elements(By.ID, 'tabelaConsulta')
        string = ''
        for obra in obras_orcamento:
            string = obra.text
        print(string)
        import unidecode
        string_SA = unidecode.unidecode(string)
        list = string_SA.replace('/', '-').split('\n')
        lista_obra_orcamento = list
        self.driver.switch_to.default_content()
        
        for obra in lista_obra_orcamento:
            code_obra = obra.split(' ')[0]
            nome_obra = ' '.join(obra.split(' ')[1:])   
            try:       
                #desloca para o iframe de orçamento
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/ORC/filterRelatorioEmissaoOrcamento.do')
                #preenche codigo da obra
                campo = self.retorna_elemento('ID', 'filter.empreend.cdEmpreendView')
                campo.click()
                campo.send_keys(Keys.DELETE)
                campo.send_keys(code_obra, Keys.TAB)
                time.sleep(3)
                #baixa relatório
                self.retorna_elemento('CLASS_NAME', 'spwButtonMain').click()
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                        print('não há dados aqui') 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'orcamento' + '-' + obra + '.xlsx'
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim) 
                            self.fecha_tela()                            
                        else:
                            print('tentando de novo')
                            time.sleep(5)
                            tentativa+=1   
            except UnexpectedAlertPresentException:
                print('alarme')  
                continue
            except ElementNotInteractableException:
                print(code_obra)
                self.relatorio_orcamento()
            except TimeoutException:
                print('erro ao achar algum elemento')
                print(nome_obra)
                self.relatorio_orcamento()
                
    
    def relatorio_desembolso(self):
        for obra in self.lista_obras:  
            code_obra = obra.split(' ')[0]
            nome_obra = ' '.join(obra.split(' ')[1:]) 
            try:   
                #acessa iframe do relatório
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/SGO/filterAnaliticoApropObra.do')
                time.sleep(1)
                #digita data inicial
                di = self.retorna_elemento('ID', 'analise.periodoInicio')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos
                #digita data final
                df = self.retorna_elemento('ID', 'analise.periodoFim')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior_contratos)
                #digita o código da obra
                self.retorna_elemento('ID', 'cdEmpreendView').send_keys(code_obra, Keys.TAB)
                time.sleep(1)
                #seleciona seleção por data de pagamento
                self.retorna_elemento('XPATH', '//*[@id="holderConteudo2"]/form/table[1]/tbody/tr[10]/td[2]/select/option[3]').click()
                #desmarca caixa de previsão
                self.retorna_elemento('ID', 'analise.consDocPrev').click()
                #marca caixa para relatório detalhado
                self.retorna_elemento('ID', 'analise.imprimirDadosEmColunasNaoMescladas').click()
                #baixa relatório
                self.retorna_elemento('ID', 'visualizarButton').click()
                time.sleep(4)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                    print('não há dados aqui') 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'desembolso' + '-' + obra + '.xlsx'
                    tentativa=0
                    while tentativa<3:
                        if(os.path.isfile(caminho_inicio)):
                            print('renommeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim) 
                            self.fecha_tela()
                        else:
                            print('tentando de novo')
                            time.sleep(5)
                            tentativa+=1                       
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
            except IndexError:
                print('erro de tela')
            except ElementNotInteractableException:
                print(code_obra)
                self.relatorio_desembolso()
            except TimeoutException:
                print('erro ao achar algum elemento')
                print(nome_obra)
                self.relatorio_desembolso()
            
    
    def estoque(self):
        for obra in self.lista_obras:
            code_obra = obra.split(' ')[0]
            nome_obra = ' '.join(obra.split(' ')[1:])
            try:
                #acessando iframe do relatório
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/EST/filterPosicaoEstoqueAtual.do')
                #colocando código da obra
                self.retorna_elemento('ID', 'entity.obra.empreend.cdEmpreendView').send_keys(code_obra, Keys.TAB)
                time.sleep(1)
                #baixando relatório 
                self.retorna_elemento('CLASS_NAME', 'spwBotao').click()
                time.sleep(2)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                        print('não há dados aqui') 
                else:
                    self.aguarda_download
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'estoque' + '-' + obra + '.xlsx'
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim) 
                            self.fecha_tela()
                        else:
                            print('tentnado de novo')
                            time.sleep(5)
                            tentativa+=1                     
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
            except ElementNotInteractableException:
                print(code_obra)
                self.estoque()
            except TimeoutException:
                print('erro ao achar algum elemento')
                print(nome_obra)
                self.estoque()
            
            
    def contas_a_pagar(self):
        for empresa in self.lista_empresas:
            code_empresa = empresa.split(' ')[0]
            nome_empresa = ' '.join(empresa.split(' ')[1:])
            try:
                #acessa iframe do relatorio
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/CPG/filterContaPagar.do')
                #preenche data
                di = self.retorna_elemento('ID', 'dtEmissaoInicio')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos
                #preenche EMPRESA
                self.retorna_elemento('ID','entity.empresa.cdEmpresaView').send_keys(code_empresa, Keys.TAB)
                time.sleep(1)
                #abra parametros avançados        
                self.retorna_elemento('XPATH', '//*[@id="holderConteudo2"]/form/div[4]/span/img').click()
                time.sleep(1)
                #desmarca caixa de previsão
                self.retorna_elemento('ID', 'incluirDocs').click()
                time.sleep(1)
                #clica para baixar o relatório
                self.retorna_elemento('XPATH', '/html/body/div/div/form/p[1]/span[1]/span/input').click()
                time.sleep(2)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                        print('não há dados aqui') 
                else:
                    self.aguarda_download
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'contas_a_pagar' + '-' + empresa + '.xlsx'
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim) 
                            self.fecha_tela()
                        else:
                            print('tentando de novo')
                            time.sleep(5)
                            tentativa+=1 
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
            except ElementNotInteractableException:
                print(code_empresa)
                self.contas_a_pagar()
            except TimeoutException:
                print('erro ao achar algum elemento')
                print(nome_empresa)
                self.contas_a_pagar()
                
        
    def saldo_de_contratos(self):
        for obra in self.lista_obras: 
            code_obra = obra.split(' ')[0]
            nome_obra = ' '.join(obra.split(' ')[1:])
            try:           
                
                #acessa iframe do relatório
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/MED/filterRelatorioSaldoContrato.do')

                #coloca DATA INICIAL
                di = self. retorna_elemento('ID', 'filter.dtInicioPeriodo')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos

                #coloca data final
                df = self.retorna_elemento('ID', 'filter.dtFimPeriodo')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior_contratos)             
                
                #clica no campo de OBRA e envia informação
                self.retorna_elemento('ID', 'filter.obra.empreend.cdEmpreendView').send_keys(code_obra, Keys.TAB)
                time.sleep(1)
                '''Processo para escolher despesas diretas'''
                #clica na lupa das unidades contrutivas
                self.retorna_elemento('XPATH', '//*[@id="consUnidadeObra"]/tbody/tr/td[3]/img[1]').click()

                #mudando de iframe
                iframe_caixa = self.retorna_elemento('XPATH', '//*[@id="layerFormConsulta"]')
                self.driver.switch_to.frame(iframe_caixa)
                lista_caixa = self.driver.find_elements(By.TAG_NAME, 'tr')
                for nome in lista_caixa:
                    if("DESPESAS DIRETAS" in nome.text):
                        nome.click()
                        print(nome.text)

                #cica em selecionar
                self.retorna_elemento('ID', 'pbSelecionar').click()

                #retorna ao Iframe normal
                self.driver.switch_to.default_content()

                #clica para baixar o relatório
                self.retorna_elemento('XPATH', '/html/body/div/div/form/p/span[1]/span/input').click()            
                time.sleep(2)

                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                        print('não há dados aqui') 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'saldo_de_contratos' + '-' + obra + '.xlsx'
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim)
                            self.fecha_tela() 
                        else:
                            print('tentando de novo')
                            time.sleep(5)
                            tentativa+=1       
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue   
        
        
    def saldo_de_pedidos(self):
        for obra in self.lista_obras: 
            code_obra = obra.split(' ')[0]
            nome_obra = ' '.join(obra.split(' ')[1:])  
            try:        
                #acessa iframe do relatório
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/ADC/abrirRelatorioSaldosPedidos.do')
                time.sleep(1)
                #clica em DATA INICIAL
                di = self. retorna_elemento('XPATH', '/html/body/div/div/form/table/tbody/tr[5]/td[2]/table[1]/tbody/tr/td[1]/input')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos
                #coloca data final
                df = self.retorna_elemento('XPATH', '//*[@id="filter.dtFimPedido"]')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior_contratos)
                #clica em OBRAS e envia informacoes
                self.retorna_elemento('XPATH', '//*[@id="filter.obra.empreend.cdEmpreendView"]').send_keys(code_obra, Keys.TAB)
                time.sleep(1)
                #clica na CAIXA para HABILITAR somente pedidos autorizados
                self.retorna_elemento('XPATH', '/html/body/div/div/form/table/tbody/tr[7]/td[2]/input').click()
                #clica para BAIXAR relatório
                self.retorna_elemento('XPATH', '/html/body/div/div/form/p/span[1]/span/input').click()
                time.sleep(2)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                    print('não há dados aqui') 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'saldo_de_pedidos' + '-' + obra + '.xlsx'
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim)
                            self.fecha_tela() 
                        else:
                            print('tentando de novo')
                            time.sleep(5)
                            tentativa+=1      
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
            except ElementNotInteractableException:
                print(code_obra)
                self.saldo_de_pedidos()
            except TimeoutException:
                print('erro ao achar algum elemento')
                print(nome_obra)
                self.saldo_de_pedidos()
        
        
    def receitas_liquidas_brutas(self):
        for empresa in self.lista_empresas:
            code_empresa = empresa.split(' ')[0]
            nome_empresa = ' '.join(empresa.split(' ')[1:])
            try:
                #acesso iframe do relatório
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/CRC/filterContasRecebidas.do')
                #Coloca Data Inicial
                di = self.retorna_elemento('ID', 'dtRectoInicio')
                di.click()
                di.send_keys('01/01/2000', Keys.TAB, self.proximo_mes, Keys.TAB, '01/01/2000', Keys.TAB, self.proximo_mes, Keys.TAB, '01/01/2000',Keys.TAB, self.proximo_mes  )
                #clica em EMPRESA e recebe dados
                self.retorna_elemento('ID', 'cdEmpresaView').send_keys(code_empresa, Keys.TAB)
                time.sleep(1)            
                #expande parametros avancados
                self.retorna_elemento("XPATH", '/html/body/div[1]/div[2]/form/div[4]/span/img').click()        
                time.sleep(1)
                #Habilita segunda informacao no relatorio - CONTRATOS
                ct = self.retorna_elemento('XPATH', '/html/body/div[1]/div[2]/form/div[5]/table/tbody/tr[12]/td[2]/table/tbody/tr/td[1]/input')            
                ct.click()
                ct.send_keys(Keys.BACKSPACE)
                ct.send_keys(Keys.BACKSPACE)
                ct.send_keys('ct', Keys.TAB)
                time.sleep(1)
                #muda de iframe
                iframe = self.retorna_elemento('XPATH', '//*[@id="layerFormConsulta"]')
                self.driver.switch_to.frame(iframe)
                #clica em ct
                lista_documentos = self.driver.find_elements(By.TAG_NAME, 'tr')
                for documentos in lista_documentos:
                    if('CT CONTRATO' in documentos.text):
                        documentos.click()
                #clica em selecionar
                self.retorna_elemento('ID', 'pbSelecionar').click()
                time.sleep(1)            
                #volta pro iframe normal
                self.driver.switch_to.default_content()  
                time.sleep(1) 
                #clica para BAIXAR relatorio
                self.retorna_elemento('ID', 'btFiltrar').click()
                time.sleep(2)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                    print('não há dados aqui') 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'receitas_liquidas_brutas' + '-' + empresa + '.xlsx'
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim) 
                            self.fecha_tela()
                        else:
                            print('tentnado de novo')
                            time.sleep(5)
                            tentativa+=1    
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
            
                
        
    def obras_centro_de_custo(self): 
        for obra in self.lista_obras: 
            code_obra = obra.split(' ')[0]
            nome_obra = ' '.join(obra.split(' ')[1:])     
            try: 
                #acesso ao iframe do relatorio       
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/CAD/reportObraCentroCusto.do')
                #Altera a DATA
                di=self.retorna_elemento('XPATH', '//*[@id="empreendFilter.dataInicialStr"]')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos
                #coloca data final
                df = self.retorna_elemento('XPATH', '//*[@id="empreendFilter.dataFinalStr"]')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior_contratos)
                #clica em OBRAS e insere informações
                self.retorna_elemento('ID', 'entity.cdEmpreendView').send_keys(code_obra, Keys.TAB)
                time.sleep(1)
                #clica em BAIXAR arquivo
                self.retorna_elemento('XPATH', '/html/body/div/div/form/p/span[1]/span/input').click()
                time.sleep(2)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                    print('não há dados aqui') 
                else:
                    time.sleep(3)
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'            
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'obras_centro_de_custo' + '-' + obra + '.xlsx'           
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim) 
                            self.fecha_tela()
                        else:
                            print('tentando de novo')
                            time.sleep(5)
                            tentativa+=1     
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
            except TimeoutException:
                print('erro ao achar algum elemento')
                print(nome_obra)
                self.relatorio_orcamento()
                
        
    def extrato_conciliado(self):               
        for empresa in self.lista_empresas:
            code_empresa = empresa.split(' ')[0]
            nome_empresa = ' '.join(empresa.split(' ')[1:])
            time.sleep(1)
            try:
                #acessa iframe do relatório
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/CXA/filterExtratoConciliado.do')
                time.sleep(2)

                #Clica em EMPRESA e passa informação
                e = self.retorna_elemento('ID', 'entity.contaCorrente.empresa.cdEmpresaView')       
                e.click()
                e.send_keys(Keys.DELETE)
                print(code_empresa)
                e.send_keys(code_empresa, Keys.ENTER)     
                    
                #Seleciona DATA INICIAL
                time.sleep(3)
                di=self.retorna_elemento('XPATH', '//*[@id="dtInicio"]')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos

                #coloca data final
                df = self.retorna_elemento('ID', 'dtFim')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior_contratos)
                time.sleep(1)

                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                    print('não há dados aqui')
                    continue 
                else:
                    #clica para baixar o relatorio
                    self.retorna_elemento('XPATH', '//*[@id="holderConteudo2"]/form/p/span[1]/span/input').click()
                    time.sleep(2)

                    if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                        print('não há dados aqui')  
                        continue 
                    else:
                        self.aguarda_download()
                        caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'
                        caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'extrato_conciliado' + '-' + nome_empresa + '.xlsx'                
                        tentativa=0
                        while tentativa<1:
                            if(os.path.isfile(caminho_inicio)):
                                print('renomeando arquivo')
                                self.renomar_arquivo(caminho_inicio, caminho_fim)
                                self.fecha_tela()
                            else:
                                print('Tentando de novo')
                                time.sleep(5)
                                tentativa+=1     
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
            except TimeoutException:
                print('erro ao achar algum elemento')
                print(nome_empresa)
                self.relatorio_orcamento()
                
    def custo_por_nivel(self):
        for obra in self.lista_obras: 
            code_obra = obra.split(' ')[0]
            nome_obra = ' '.join(obra.split(' ')[1:])  
            try:
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/SGO/filterCustoNivel.do')
                #preenche código da obra
                self.retorna_elemento('ID', 'analise.obra.empreend.cdEmpreendView').send_keys(code_obra, Keys.TAB)
                time.sleep(3)
                #clica na lupa
                self.retorna_elemento('XPATH', '//*[@id="consUnidadeObra"]/tbody/tr/td[3]/img[1]').click()
                #mudando de iframe
                iframe_caixa = self.retorna_elemento('XPATH', '//*[@id="layerFormConsulta"]')
                self.driver.switch_to.frame(iframe_caixa)
                lista_caixa = self.driver.find_elements(By.TAG_NAME, 'tr')
                for nome in lista_caixa:
                    if("DESPESAS DIRETAS" in nome.text):
                        nome.click()
                        print(nome.text)
                #cica em selecionar
                self.retorna_elemento('ID', 'pbSelecionar').click()
                #retorna ao Iframe normal
                self.driver.switch_to.default_content()
                #seleciona data de pagamento
                self.retorna_elemento('XPATH', '//*[@id="analise.selecao"]/option[3]').click()
                #clica para baixar o relatório
                self.retorna_elemento('CLASS_NAME', 'spwBotaoDefault').click()
                time.sleep(2)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                    print('não há dados aqui')  
                    continue 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'custo_por_nivel' + '-' + obra + '.xlsx'                
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim)
                            self.fecha_tela()
                        else:
                            print('Tentando de novo')
                            time.sleep(5)
                            tentativa+=1     
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
                
    def extrato_estoque(self):        
        for obra in self.lista_obras: 
            code_obra = obra.split(' ')[0]
            try:
                #acessa site
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/EST/filterRelatorioExtratoEstoque.do')      
                #coloca codico centro de custo
                self.retorna_elemento('ID', 'centroCusto.empreend.cdEmpreendView').send_keys(code_obra, Keys.TAB)
                time.sleep(2)
                #colocando data
                di = self.retorna_elemento('ID', 'dtInicio')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos
                time.sleep(1)

                #seleciona relatório detalhado
                self.retorna_elemento('ID', 'detalhado').click()

                #clica em baixar o relatório
                self.retorna_elemento('CLASS_NAME', 'spwButtonMain').click()
                time.sleep(3)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                    print('não há dados aqui')  
                    continue 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'extrato_estoque' + '-' + obra + '.xlsx'                
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim)
                            self.fecha_tela()
                        else:
                            print('Tentando de novo')
                            time.sleep(5)
                            tentativa+=1 
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue
    
    
    def emissao_contratos(self):
        for obra in self.lista_obras: 
            code_obra = obra.split(' ')[0]
            try:
                #Acessa o Iframe do relatório
                self.driver.get('https://lotuscidade.sienge.com.br/sienge/MED/filterRelatorioEmissaoContrato.do')            
                #insere código da obra
                self.retorna_elemento('ID', 'filter.obra.empreend.cdEmpreendView').send_keys(code_obra, Keys.TAB)
                time.sleep(2)
                #insere data inicial
                di = self.retorna_elemento('ID', 'filter.dtInicioPeriodo' )
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",di, '01/01/2000') #self.dia_anterior_contratos
                #insere data final
                df = self.retorna_elemento('ID', 'filter.dtFimPeriodo')
                self.driver.execute_script("arguments[0].setAttribute('value',arguments[1])",df, self.dia_anterior_contratos)
                #clica em baixar relatório
                self.retorna_elemento('CLASS_NAME', 'spwButtonMain').click()
                time.sleep(3)
                if(self.driver.find_elements(By.CLASS_NAME, 'spwAlertaAviso')):
                        print('não há dados aqui')  
                        continue 
                else:
                    self.aguarda_download()
                    caminho_inicio = configLotus.DIRETORIO_ARQUIVOS_TEMP + '\\' + 'relatorio.xlsx'
                    caminho_fim = configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior + '\\' + 'emissao_contratos' + '-' + obra + '.xlsx'                
                    tentativa=0
                    while tentativa<1:
                        if(os.path.isfile(caminho_inicio)):
                            print('renomeando arquivo')
                            self.renomar_arquivo(caminho_inicio, caminho_fim)
                            self.fecha_tela()
                        else:
                            print('Tentando de novo')
                            time.sleep(5)
                            tentativa+=1 
                    continue           
            except UnexpectedAlertPresentException:
                print('Alert Text: Nenhum registro encontrado.')
                continue    
        
     
    def mudar_obra_empresa_unidadeconstrutiva(self):
        #cria lista empresas
        self.driver.get('https://lotuscidade.sienge.com.br/sienge/CAD/reportObraCentroCusto.do')
        self.retorna_elemento('XPATH', '//*[@id="searchEmpresa"]/tbody/tr/td[4]/img[1]').click()
        iframe_obra = self.retorna_elemento('XPATH', '//*[@id="layerFormConsulta"]')
        self.driver.switch_to.frame(iframe_obra)
        tabela = self.driver.find_elements(By.ID, 'tabelaResultado') 
        lista=''
        for i in tabela:
            lista = i.text
        emp = unidecode.unidecode(lista)
        emp2 = emp.replace('/', '-').split('\n') 
        self.lista_empresas = emp2
        self.driver.switch_to.default_content()
        
        #cria lista obras
        self.driver.get('https://lotuscidade.sienge.com.br/sienge/CAD/reportObraCentroCusto.do')
        self.retorna_elemento('XPATH', '//*[@id="searchObraCentroCusto"]/tbody/tr/td[3]/img[1]').click()
        iframe_obra = self.retorna_elemento('XPATH', '//*[@id="layerFormConsulta"]')
        self.driver.switch_to.frame(iframe_obra)
        tabela2 = self.driver.find_elements(By.ID, 'tabelaResultado')
        lista2 = ''
        for i in tabela2:
            lista2 = i.text
        obr = unidecode.unidecode(lista2)
        obr2 = obr.replace('/', '-').split('\n')
        self.lista_obras = obr2
        self.driver.switch_to.default_content() 
        
        time.sleep(1)
        
        
    def run(self):
        self.cria_diretorio(configLotus.DIRETORIO_ARQUIVOS + '\\' + self.dia_anterior)
        for credencial in configLotus.CREDENCIAIS:
            self.login(credencial['usuario'], credencial['senha'])  
            time.sleep(3)
            self.mudar_obra_empresa_unidadeconstrutiva()            
            self.relatorio_orcamento()
            self.relatorio_desembolso()            
            #self.estoque()
            self.contas_a_pagar()
            self.saldo_de_contratos()
            self.saldo_de_pedidos()
            self.receitas_liquidas_brutas()
            self.obras_centro_de_custo()
            self.extrato_conciliado()
            self.custo_por_nivel()
            self.extrato_estoque()
            self.emissao_contratos()
        self.driver.close()
        self.driver.quit()
   

if __name__ == '__main__':
    bot = Bot()
