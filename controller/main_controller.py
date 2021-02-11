from excel.model import *
from bs4 import BeautifulSoup
import requests

excel = ExcelModel()
excel.readExcel("C:/Users/eotlr/project/portal/SpacTrack_list.xlsx")
target_list = excel.read_work_book("fail")
target_list = target_list[1:]

class SECCrawler(object):

    def __init__(self):
        self.save_path = "C:/Users/eotlr/project/portal/html/"

    def download_S_1(self, target_list):
        basic ="https://www.sec.gov/"
        for spac in target_list:
            try:
                html = requests.get(spac[1]).text
                soup = BeautifulSoup(html, "html.parser")
                link = soup.select("table.tableFile2 > tr ")
                for tr in link:
                    if "S-1" in tr.text:
                        strTr = str(tr)
                        strTr = strTr[strTr.find("/Archive"):]
                        strTr = strTr[:strTr.find('"')]
                first_archive = basic + strTr
                first_archive_html = requests.get(first_archive).text
                soup = BeautifulSoup(first_archive_html, "html.parser")
                archive = soup.select("table.tableFile > tr > td > a")
                target = requests.get(basic + archive[0]['href']).text
                file_path = self.save_path + spac[0] + ".txt"
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(target)
                    f.close()
            except Exception as ex:
                print(spac[0])
                print(ex)


sec_crawler = SECCrawler()
sec_crawler.download_S_1(target_list)