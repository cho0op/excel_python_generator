import urllib.request
from bs4 import BeautifulSoup

from openpyxl import load_workbook


def get_link(cell):
    return cell.value


def main():
    workbook = load_workbook(filename="excel_for_python.xlsx")
    sheet = workbook.active
    try:
        for i in range(1, 1445):
            description_cell = sheet["C"][i]
            link_cell = sheet["E"][i]
            if description_cell.value is None:
                req = urllib.request.Request(link_cell.value, headers={'User-Agent': 'Mozilla/5.0'})
                html = urllib.request.urlopen(req).read()
                htmlParse = BeautifulSoup(html, 'html.parser')
                for div in htmlParse.find_all("div", {"class": "offers-description__specs"}):
                    p_text = div.find("p").text
                    description_cell.value = p_text
                    print(p_text)
    except:
        workbook.save("excel_for_python_ready.xlsx")

    workbook.save("excel_for_python_ready.xlsx")


if __name__ == '__main__':
    main()
