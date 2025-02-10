from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import openpyxl
from openpyxl import Workbook

def save_to_excel(data_list, filename="imoveis.xlsx"):
    workbook = Workbook()
    sheet = workbook.active

    # Escreve o cabeçalho
    sheet.append(list(data_list[0].keys()))

    # Escreve os dados
    for data in data_list:
        sheet.append(list(data.values()))

    # Salva o arquivo Excel
    workbook.save(filename)

def process_ads(houses):
    ads_data = []
    for index, house in enumerate(houses):
        try:
            house_location = get_house_location(house)
            house_info = get_house_info(house)
            house_pricing = get_house_pricing(house)

            data_type = house.get_attribute('data-type')
            data_position = house.get_attribute('data-position')

            house_data = {
                'tipo': data_type,
                'posicao': data_position,
                **house_location,
                **house_info,
                **house_pricing
            }

            ads_data.append(house_data)

        except Exception as e:
            print(f"An error occurred for ad {index + 1}: {e}")
    
    return ads_data

def get_house_location(house):
    house_info = {}

    try:
        card_location = house.find_elements(By.XPATH, './/section[contains(@class, "card__location")]')
        if card_location:
            try:
                house_info['titulo'] = card_location[0].find_element(By.XPATH, './/h2[contains(@class, "l-text")]').text.strip()
            except:
                house_info['titulo'] = 'Título não disponível'
            try:
                house_info['endereco'] = card_location[0].find_element(By.XPATH, './/p[contains(@class, "l-text")]').text.strip()
            except:
                house_info['endereco'] = 'Endereço não disponível'
        else:
            house_info['titulo'] = 'Título não disponível'
            house_info['endereco'] = 'Endereço não disponível'
    except:
        house_info['titulo'] = 'Erro'
        house_info['endereco'] = 'Erro'

    return house_info

def get_house_info(house):
    house_info = {}

    try:
        amenities_section = house.find_elements(By.XPATH, './/section[contains(@class, "Amenities_card-amenities__kpLh7")]')
        if amenities_section:
            try:
                house_info['metragem'] = amenities_section[0].find_element(By.XPATH, './/span[@aria-label="Tamanho do imóvel"]/..').text.strip()
            except:
                house_info['metragem'] = 'Metragem não disponível'
            
            try:
                house_info['quartos'] = amenities_section[0].find_element(By.XPATH, './/span[@aria-label="Quantidade de quartos"]/..').text.strip()
            except:
                house_info['quartos'] = 'Quartos não disponível'
            
            try:
                house_info['banheiros'] = amenities_section[0].find_element(By.XPATH, './/span[@aria-label="Quantidade de banheiros"]/..').text.strip()
            except:
                house_info['banheiros'] = 'Banheiros não disponível'
            
            try:
                house_info['vagas'] = amenities_section[0].find_element(By.XPATH, './/span[@aria-label="Quantidade de vagas de garagem"]/..').text.strip()
            except:
                house_info['vagas'] = 'Vagas não disponível'
        else:
            house_info['metragem'] = 'Metragem não disponível'
            house_info['quartos'] = 'Quartos não disponível'
            house_info['banheiros'] = 'Banheiros não disponível'
            house_info['vagas'] = 'Vagas não disponível'
    except Exception as e:
        print(f"Erro ao tentar extrair informações do imóvel: {e}")
        house_info['metragem'] = 'Erro'
        house_info['quartos'] = 'Erro'
        house_info['banheiros'] = 'Erro'
        house_info['vagas'] = 'Erro'
    
    return house_info

def get_house_pricing(house):
    house_info = {}

    try:
        pricing_div = house.find_elements(By.XPATH, './/div[contains(@class, "ListingCard_result-card__wrapper__6osq8")]')
        
        if pricing_div:
            try:
                house_info['preco'] = pricing_div[0].find_element(By.XPATH, './/p[contains(@class, "l-text--variant-heading-small")]').text.strip()
            except:
                house_info['preco'] = 'Preço não disponível'
            
            try:
                financial_info = pricing_div[0].find_element(By.XPATH, './/p[contains(@class, "l-u-color-neutral-44")]').text.strip()
                financial_values = financial_info.split("|")
                house_info['condominio'] = financial_values[0].replace("Cond.", "").strip() if len(financial_values) > 0 else 'Condomínio não disponível'
                house_info['iptu'] = financial_values[1].replace("IPTU", "").strip() if len(financial_values) > 1 else 'IPTU não disponível'
            except:
                house_info['condominio'] = 'Condomínio não disponível'
                house_info['iptu'] = 'IPTU não disponível'
        else:
            house_info['preco'] = 'Preço não disponível'
            house_info['condominio'] = 'Condomínio não disponível'
            house_info['iptu'] = 'IPTU não disponível'
    except Exception as e:
        print(f"Erro ao tentar extrair informações de preço/condomínio/IPTU: {e}")
        house_info['preco'] = 'Erro'
        house_info['condominio'] = 'Erro'
        house_info['iptu'] = 'Erro'

    return house_info

def scroll_page():
    scroll_pause_time = 1
    last_height = driver.execute_script("return document.body.scrollHeight")
    max_scroll_attempts = 10
    scroll_attempts = 0

    while True:
        actions = ActionChains(driver)
        for _ in range(10):
            actions.send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(0.5)

        time.sleep(scroll_pause_time)
        
        new_height = driver.execute_script("return document.body.scrollHeight")
        
        if new_height > last_height:
            last_height = new_height
            scroll_attempts = 0
        else:
            scroll_attempts += 1
            if scroll_attempts >= max_scroll_attempts:
                houses = driver.find_elements(By.CSS_SELECTOR, 'div[data-type]')
                return houses  # Retorna os anúncios processados
                
def find_next_button():
    try:
        paginator_section = driver.find_element(By.CLASS_NAME, 'listing-wrapper__pagination')
        buttons = paginator_section.find_elements(By.CLASS_NAME, 'l-button--context-primary')
        
        for button in buttons:
            button_text = button.text.strip().lower()
            
            if not button_text:
                button_text = button.get_attribute('aria-label').strip().lower()
            
            if "próxima" in button_text:
                return button
    except Exception as e:
        print(f"Erro ao encontrar o botão de próxima página: {e}")
    
    return None

def load_all_pages():
    all_ads_data = []  # Lista para armazenar dados de todas as páginas
    while True:
        houses = scroll_page()
        if houses:
            ads_data = process_ads(houses)
            all_ads_data.extend(ads_data)  # Acumula os dados da página atual

            button = find_next_button()
            if button:
                try:
                    time.sleep(3)
                    button.click()
                    time.sleep(5)
                except Exception as e:
                    print(f"Erro ao clicar no botão de próxima página: {e}")
                    break
            else:
                break
        else:
            break

    # Salva todos os dados acumulados de uma vez só
    save_to_excel(all_ads_data)

# Inicialize o driver e inicie o processo
url = "https://www.zapimoveis.com.br/venda/cobertura/rj+rio-de-janeiro+zona-oeste+recreio-dos-bandeirantes/?__ab=sup-hl-pl:newC,exp-aa-test:B,super-high:new,olx:control,off-no-hl:new,TOP-FIXED:card-b,pos-zap:control,ngt-new:test,new-rec:b,lgpd-ldp:test&transacao=venda&onde=,Rio%20de%20Janeiro,Rio%20de%20Janeiro,Zona%20Oeste,Recreio%20Dos%20Bandeirantes,,,neighborhood,BR%3ERio%20de%20Janeiro%3ENULL%3ERio%20de%20Janeiro%3EZona%20Oeste%3ERecreio%20Dos%20Bandeirantes,-23.017871,-43.464748,&tipos=cobertura_residencial&pagina=1&areaMinima=350&ordem=Menor%20pre%C3%A7o"

driver = webdriver.Chrome()
driver.get(url)

load_all_pages()

driver.quit()
