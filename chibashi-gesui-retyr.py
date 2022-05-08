'''
Created on 2021/08/26

@author: kentoo
'''
import time
import chromedriver_binary
from selenium import webdriver 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select,WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
import os 
import retry
import json
import mojimoji 
import common 

#半角であれば全角へ ＋ 何丁目が-の場合変換

def h_to_z(address):
    S = mojimoji.han_to_zen(address)
    
    if int(S.count('－'))== 2:
        S = S.replace('－','丁目',1)
    
    return S
    


#住所を分割

def address_separate(address):
    
    address = h_to_z(address)
    ad = ["県","市","区","町","目","－"]
    ans = []
    
    for i in range(6):
        target = ad[i]
        idx = address.find(target)
        if ad[i] == "町":
            if idx == -1:
                i = 0
                idx = 0
                while address[i].isdigit()==False:
                    idx+=1
                    i+=1
                ans.append(address[:idx])
                address = address.replace(address[:idx],"")
            else:
                ans.append(address[:idx+1])
                address = address.replace(address[:idx+1],"")
        else:
            ans.append(address[:idx+1])
            #print(address[:idx+1])
            address = address.replace(address[:idx+1],"")
            #print(address)
        
    ans[5]=ans[5].replace("－","番")
    
    ans.append(address)
    
    return ans

def chiba_gesui(address):
    
    #現在のディレクトリ取得
    iDir = os.getcwd()
    
    
    
    
    #印刷用設定
    appState = {
        "recentDestinations": [
            {
                "id": "Save as PDF",
                "origin": "local",
                "account":""
            }
            
        ],
        "selectedDestinationId": "Save as PDF",
        "version":2,
        "pageSize":"A4"
    }
    
    #ダウンロード先を指定
    options = webdriver.ChromeOptions()
    
    options.add_argument('--kiosk-printing')
    
    options.add_experimental_option("prefs",{
    "printing.print_preview_sticky_settings.appState":
    json.dumps(appState),
    "download.default_directory": '~/Downloads'
    })
    #options.add_experimental_option("prefs", {"download.default_directory": iDir })
    
    #'~/Downloads'
    browser = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)
    
    
    
    #ブラウザ
    url='http://s-page.tumsy.com/chibagesui/index.html'
    browser.get(url)
    

    
    #町名
    chiba_address  = address 
    
    
    
    #県,市,区,町,丁目,番地,号
    separate = address_separate(chiba_address)
    
    
    
    #印刷縮尺選択 1/500 "btnScale500" 1/750 "btnScale500" 1/1000 "btnScale1000"
    scale = "btnScale500"
    
    
    
    #フレーム移動
    find_css_retry = common.make_retriable(browser.find_element_by_css_selector)
    browser.switch_to.frame(find_css_retry("frame"))

    
    #同意ボタン
    agree_button = common.find_id_with_retry(browser, "LinkButton1")
    common.click_with_retry(agree_button)

    browser.switch_to_default_content()

    
    
    
    #フレーム移動
    browser.switch_to.frame(find_css_retry("frame"))
    
    
    #区を入力
    select_id = common.find_id_with_retry(browser, "ELM_CMB_LEV1")
    common.select_with_retry(select_id,separate[2])

    
    
    #町名を入力
    select_id2 = common.find_id_with_retry(browser, "ELM_CMB_LEV2")
    common.select_with_retry(select_id2,separate[3])
    

    
    
    #～丁目を入力（～丁目がない場合は丁目なしを入力）

        
    
    select_id3 = common.find_id_with_retry(browser, "ELM_CMB_LEV3")
    if separate[4]=="":
        common.select_with_retry(select_id3,"丁目なし")
    else:
        common.select_with_retry(select_id3,separate[4])

    
   
     
    print("1")
    
    #番地を入力 
    try:
        '''
        select_id4 = common.find_id_with_retry(browser, "ELM_CMB_LEV4")
        print(select_id4)
        common.select_with_retry(select_id4,separate[5])
        '''
        time.sleep(1)
        select_id4 = browser.find_element_by_id('ELM_CMB_LEV4')
        banchi = Select(select_id4)
        banchi.select_by_visible_text(separate[5])
        
    except NoSuchElementException :
        
        select_id4 = common.find_id_with_retry(browser, "ELM_CMB_LEV4")
        common.select_with_retry(select_id4,separate[5]+"地")
  
        
    
    
    #検索ボタン
    search = common.find_id_with_retry(browser, "btnAddSchDlgOK")
    common.click_with_retry(search)

    
    #印刷ボタン
    time.sleep(2)
    printing = common.find_id_with_retry(browser, "OpenPrintPage")
    common.click_with_retry(printing)
    browser.switch_to_default_content()

    
    #タブ移動
    #time.sleep(2)
    
    window = browser.window_handles[1]
    browser.switch_to.window(window)
    
    
    #print(browser.title)
    
    
    #PDFスケール選択
    
    PDF_scale  = common.find_id_with_retry(browser, scale)
    common.click_with_retry(PDF_scale)

    
    #PDF印刷
    time.sleep(2)
    PDF = browser.find_element_by_id('btnPdf')
    PDF.click()
    
    #PDFページに移動
    #time.sleep(10)
    window = browser.window_handles[2]
    browser.switch_to.window(window)
    
    #印刷
    #time.sleep(2)
    browser.execute_script(f'document.title = "{chiba_address}"')
    #time.sleep(2)
    browser.execute_script('return window.print();')
    
    
    #ファイル名変更
    time.sleep(5)
    path_b=r"C:\Users\kentoo\Downloads\CreatePDF.pdf"
    path_a = fr"C:\Users\kentoo\Downloads\下水_{chiba_address}.pdf"
    os.rename(path_b,path_a)
    print(os.path.exists(path_a))
    
    time.sleep(5)
    browser.quit()
    
if __name__ == '__main__':
    AD = '千葉県千葉市中央区青葉町１２６５-２'
    AD2 = '千葉県千葉市花見川区瑞穂１丁目１８-１'
    AD3 = '千葉県千葉市若葉区桜木北２丁目１-１'
    AD4 = '千葉県千葉市若葉区桜木２丁目１-１３'
    AD5 = '千葉県千葉市稲毛区穴川２丁目１０-２'
    AD6 = '千葉県千葉市稲毛区穴川町３７７-１'
    AD7 = '千葉県千葉市花見川区瑞穂１-１８-１'
    AD8 = '千葉県千葉市中央区新千葉1-1-1'
    AD9 = '千葉県千葉市中央区亥鼻１丁目６－１'
    AD10 = '千葉県千葉市稲毛区稲毛１丁目１６-１２'
    AD11 = '千葉県千葉市美浜区ひび野2-6-1'
    chiba_gesui(AD11)
    #print(address_separate(AD4))
    
