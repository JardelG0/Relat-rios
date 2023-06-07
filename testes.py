
import pandas as pd


arquivo_excel1 = "INFORMATIVO 5° GRE.xlsx"
dados = pd.read_excel(arquivo_excel1, sheet_name="Sheet1")     

for i in range(4):
    print(dados['DIAS'][i],'=>', type(dados['DIAS'][i]))
# val =dados.loc[1,'DIAS']
# str(val)
# val += "1"
# dados.loc[1,'DIAS'] = val
# print(dados.loc[1,'DIAS'])


# from selenium import webdriver

# import json

# chrome_options = webdriver.ChromeOptions()
# settings = {
#        "recentDestinations": [{
#             "id": "Save as PDF",
#             "origin": "local",
#             "account": "",
#         }],
#         "selectedDestinationId": "Save as PDF",
#         "version": 2
#     }
# prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
# chrome_options.add_experimental_option('prefs', prefs)
# # chrome_options.add_argument('--kiosk-printing')
# CHROMEDRIVER_PATH = '/usr/local/bin/chromedriver'
# driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=CHROMEDRIVER_PATH)
# driver.get("https://google.com")
# driver.execute_script('window.print();')
# from selenium.webdriver.common.print_page_options import PrintOptions

# print_options = PrintOptions()
# print_options.page_ranges = ['1']

# driver.get("printPage.html")

# base64code = driver.print_page(print_options)




# # import os
# # from selenium import webdriver
# # from selenium.webdriver.common.keys import Keys
# # import json
# # import time

# # chrome_options = webdriver.ChromeOptions()

# # settings = {
# #     "recentDestinations": [{
# #         "id": "Save as PDF",
# #         "origin": "local",
# #         "account": ""
# #     }],
# #     "selectedDestinationId": "Save as PDF",
# #     "version": 2,
# #     "isHeaderFooterEnabled": False,
# #     "mediaSize": {
# #         "height_microns": 210000,
# #         "name": "ISO_A5",
# #         "width_microns": 148000,
# #         "custom_display_name": "A5"
# #     },
# #     "customMargins": {},
# #     "marginsType": 2,
# #     "scaling": 175,
# #     "scalingType": 3,
# #     "scalingTypePdf": 3,
# #     "isCssBackgroundEnabled": True
# # }

# # mobile_emulation = { "deviceName": "Nexus 5" }
# # chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)
# # chrome_options.add_argument('--enable-print-browser')
# # #chrome_options.add_argument('--headless')

# # prefs = {
# #     'printing.print_preview_sticky_settings.appState': json.dumps(settings),
# #     'savefile.default_directory': '<path>'
# # }
# # chrome_options.add_argument('--kiosk-printing')
# # chrome_options.add_experimental_option('prefs', prefs)

# # for dirpath, dirnames, filenames in os.walk('<source path>'):
# #     for fileName in filenames:
# #         print(fileName)
# #         driver = webdriver.Chrome("./chromedriver", options=chrome_options)
# #         driver.get(f'file://{os.path.join(dirpath, fileName)}')
# #         time.sleep(7)
# #         driver.execute_script('window.print();')
# #         driver.close()