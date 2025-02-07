import pandas as pd  
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import re



def cleanhtml(raw_html):
  cleantext = re.sub(re.compile('<.*?>'), '', raw_html)
  if "]:  " in cleantext:
      cleantext = cleantext.replace("]:  ","")
  if ")" in cleantext:
      cleantext = cleantext.replace(")","")
  if ":" in cleantext:
      cleantext = cleantext.replace(":","")
  return cleantext.strip()

def cleanhtml_abt(raw_html):
  cleantext = re.sub(re.compile('<.*?>'), '', raw_html)
  return cleantext.strip()

def findCardName(card):
    pattern = r"ชื่อการ์ด(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardID(card):
    pattern = r"ID(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardType(card):
    if "Type: " in card:
        card = card.replace('Type: ','ประเภทการ์ด')
    pattern = r"ประเภทการ์ด(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardRarity(card):
    if "Rarity: " in card:
        card = card.replace('Rarity: ','ความหายาก')
    pattern = r"ความหายาก(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardColor(card):
    if "Color: " in card:
        card = card.replace('Color: ','สี')
    pattern = r"สี(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardGroup(card):
    if "ลักษณะเฉพาะ" in card:
        card = card.replace('ลักษณะเฉพาะ','ลักษณะ')
    pattern = r"ลักษณะ(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardLV(card):
    pattern = r"เลเวล(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardAP(card):
    pattern = r"แอคชั่นพอยท์(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardLP(card):
    if "LP(Logic Point): " in card:
        card = card.replace('LP(Logic Point): ','โลจิคพอยท์')
    pattern = r"โลจิคพอยท์(.*?)(<br/>|$)"
    match = re.search(pattern, card)
    if match:
        return cleanhtml(match.group(1)) 
    return ''

def findCardAbility(card):
    soup = BeautifulSoup(card, "html.parser")
    start_p = None
    for p in soup.find_all("p"):
        if p.find("img"):  # ถ้า <p> มี <img>
            start_p = p
            break
    found_start = False
    found_stop = False
    results = []
    for p in soup.find_all("p"):
        if p == start_p:
            found_start = True
        if p.find("b", string="FAQ:"):
            break
        if found_start:
            text = p.get_text(" ", strip=True) 
            results.append(text)
        
    ability = "\n".join(results)
    return ability

def findCardPic(card):
    pattern = r'src="([^"]*)"'
    match = re.search(pattern, card)
    return "https://www.undercutgames.com"+match.group(1)

df = pd.DataFrame(columns=["card_pic", "card_pic_url", "card_name","card_id","card_type", "card_lv","card_rarity","card_color","card_group","card_ap", "card_lp", "card_ability"])

cards=[]
flag = True;

driver = webdriver.Chrome()
driver.get("https://www.undercutgames.com/product/single/conan?single-selected=11&fbclid=IwY2xjawIRgi1leHRuA2FlbQIxMAABHQ2sckOnT2KfMMATaQ-GE-Sfrj0in_PJTFXlmCRE1Mkf1-bBhVhDf7bCVg_aem_zluTJJOZPrIV9b0EmSVYCQ")
actions = ActionChains(driver)

driver.find_element(By.CLASS_NAME, "single-card-name").click()
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, "single-displayer-bg"))
)

while flag==True:
    soup = BeautifulSoup(driver.page_source, "html.parser")
    cards.append(soup.find("div", class_="single-displayer-bg"))
    element = driver.find_element(By.CLASS_NAME, "justify-content-center")
    actions.move_to_element(element).move_by_offset(1, 1).click().perform()
    try: 
        driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/div[1]/div[3]/i").click()
        WebDriverWait(driver, 1)
    except:
        flag=False;


for card in cards:
    card_pic = findCardPic(str(card))
    card_name = findCardName(str(card))
    card_id = findCardID(str(card))
    card_type= findCardType(str(card))
    card_rarity = findCardRarity(str(card))
    card_color = findCardColor(str(card))
    card_lv = findCardLV(str(card))
    if card_type=='ตัวละคร':
        card_group = findCardGroup(str(card))
        card_ap = findCardAP(str(card))
        card_lp = findCardLP(str(card))
    else:
        card_group = ''
        card_ap = ''
        card_lp = ''
    card_ability = findCardAbility(str(card))
    new_row = {'card_pic_url': card_pic,"card_name":card_name,
               "card_id":card_id,"card_type":card_type,
               "card_rarity":card_rarity,"card_color":card_color,
               "card_group":card_group,"card_ap":card_ap,
               "card_lp":card_lp, "card_ability":card_ability,
               "card_lv":card_lv,'card_pic':''}
    df = df._append(new_row, ignore_index=True)

driver.quit()
df.to_excel("conan_cards.xlsx", index=False)
print('✅ completed')