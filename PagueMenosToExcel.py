#! python
# This program searchs product name and it's price and
# creates an excel with the information.

import requests
import bs4
import os
import openpyxl
import time
from openpyxl.styles import Font

os.chdir('C:\\Users\\mathe\\Desktop\\Scrapping')
url = 'https://www.superpaguemenos.com.br'
print('Acessando ' + url)
res = requests.get(url)
try:
    res.raise_for_status()
except Exception as exc:
    print('There was a problem: %s' % (exc))

soup = bs4.BeautifulSoup(res.text, 'lxml')
linkElem = soup.select('.level-3 > a')
linkList = []
dataBase = {}

# Gets de url for type of products list inside categories/subcategories.
for i in range(len(linkElem)):
    linkList.append(linkElem[i].get('href'))

for link in linkList:
    n = 0
    found = []
    if link == None:
        continue
    while len(found) == 0:
        # Downloads type of products page.
        if link.startswith('/'):
            if link.endswith('c'):
                url = ('https://www.superpaguemenos.com.br' +
                       link + '/&p=' + str(n + 1))
            elif link.endswith('/?p=1'):
                url = ('https://www.superpaguemenos.com.br' +
                       link.split('?p=1')[0] + '?p=' + str(n + 1))
            elif link.endswith('/'):
                url = ('https://www.superpaguemenos.com.br' +
                       link + '?p=' + str(n + 1))
            else:
                url = ('https://www.superpaguemenos.com.br' +
                       link + '/?p=' + str(n + 1))
        else:
            if link.endswith('c'):
                url = ('https://www.superpaguemenos.com.br/' +
                       link + '/&p=' + str(n + 1))
            elif link.endswith('/?p=1'):
                url = ('https://www.superpaguemenos.com.br/' +
                       link.split('?p=1')[0] + '?p=' + str(n + 1))
            elif link.endswith('/'):
                url = ('https://www.superpaguemenos.com.br/' +
                       link + '?p=' + str(n + 1))
            else:
                url = ('https://www.superpaguemenos.com.br/' +
                       link + '/?p=' + str(n + 1))

        # Downloads products page.
        resLink = requests.get(url)
        try:
            resLink.raise_for_status()
        except Exception as exc:
            print('There was a problem: %s' % (exc))

        # Get categoria, subcategoria, tipo de produto.
        if '33947-achocolatado-em-po' in link:
            break
        elif '10836-pos-banho-infantil' in link:
            break
        elif 'cuidado-intimo' in link:
            break
        elif '10838-shampoo-infantil' in link:
            break
        elif 'brinquedo' in link:
            break
        elif 'jardinagem' in link:
            break
        elif 'alimentos-funcionais' in link:
            break
        elif link.startswith('r'):
            break
        elif link.startswith('/') and (link.endswith('1') or link.endswith('c')):
            startSpace, categoria, subcategoria, tipoProduto, endSpace = link.split(
                '/')
        elif link.startswith('/') and (not link.endswith('/')):
            startSpace, categoria, subcategoria, tipoProduto = link.split('/')
        elif link.startswith('/') and len(link.split('/')) == 5:
            startSpace, categoria, subcategoria, tipoProduto, endSpace = link.split(
                '/')
        elif link.startswith('/') and len(link.split('/')) == 4:
            break
        else:
            categoria, subcategoria, tipoProduto, endSpace = link.split('/')

        # Adiciona categoria, subcategoria, tipo de produto à dataBase.
        dataBase.setdefault(categoria, {})
        dataBase[categoria].setdefault(subcategoria, {})
        dataBase[categoria][subcategoria].setdefault(tipoProduto, {})

        # Parse the page for productName and Price.
        print("Downloading product's name and price in " + url)
        soup = bs4.BeautifulSoup(resLink.text, 'lxml')
        productName = soup.select('.title > a > span')
        price = soup.select('.price')

        # Add all products and price from url to lists.
        productList = []
        priceList = []

        for i in range(len(price)):
            productList.append(productName[i].getText().strip())
            priceList.append(price[i].getText())

        # Add products and prices to dataBase.
        for k in range(len(productList)):
            dataBase[categoria][subcategoria][tipoProduto].setdefault(
                productList[k], ' ')
            dataBase[categoria][subcategoria][tipoProduto][productList[k]
                                                           ] = priceList[k]
        found = soup.select('.pd2 > h1')
        n += 1

# Open a new excel file and write the contents of dataBase to it.
print('Escrevendo dados...')
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'PagueMenos'
arial16BoldFont = Font(name='Arial', size=16, bold=True)
arial14BoldFont = Font(name='Arial', size=14, bold=True)
arial12BoldFont = Font(name='Arial', size=12, bold=True)
arial12Font = Font(name='Arial', size=12)

# Create a variable with starting row.
c = 1
for category in dataBase.keys():
    sheet['A' + str(c)] = 'Categoria'
    sheet['A' + str(c)].font = arial12BoldFont
    sheet['B' + str(c)] = ' '.join(category.upper().split('-')[1:])
    sheet['B' + str(c)].font = arial12BoldFont
    c += 1
    for subcategory in dataBase[category].keys():
        sheet['A' + str(c)] = 'Subcategoria'
        sheet['A' + str(c)].font = arial12BoldFont
        sheet['B' + str(c)] = ' '.join(subcategory.upper().split('-'))
        sheet['B' + str(c)].font = arial12BoldFont
        c += 1
        for productType in dataBase[category][subcategory].keys():
            sheet['A' + str(c)] = 'Tipo de Produto'
            sheet['A' + str(c)].font = arial12BoldFont
            sheet['B' + str(c)] = ' '.join(productType.upper().split('-'))
            sheet['B' + str(c)].font = arial12BoldFont
            c += 1
            for produto, preço in dataBase[category][subcategory][productType].items():
                sheet['B' + str(c)] = produto
                sheet['B' + str(c)].font = arial12Font
                sheet['C' + str(c)] = preço
                sheet['C' + str(c)].font = arial12Font
                c += 1

sheet.column_dimensions['A'].width = 25
sheet.column_dimensions['B'].width = 90
sheet.column_dimensions['C'].width = 11
wb.save('produtoEpreço.xlsx')
print('Done.')
