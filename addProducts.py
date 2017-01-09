# Idea was thought of back in June of 2016 in regards to creating a way for qvpython script to easily manipulate an excel spreadsheet. After obtaining help from Vaso Vasich, this script was actually made in an effort by Vaso Vasich under the SDEP business, which is a subset of Arce Enterprises.
# 2016
# First Author: Vaso Vasich
# Second Author: David Arce

# Program is designed to make a fluid way for product to be added to an excel sheet with minimal effort in order to make it quicker to add mass amounts of product. This excel sheet is then uploaded to an ecommerce website with no changes needed.

part_types = {
    'Rocker Arms': {
        'specs': [
            'Stud Diameter(in.):',
            'Ratio:'
        ],
        'set': 16,
        'images': {
        # still need to add more
            'Ultra Pro Magnum': [
                'https://s25.postimg.org/vd6zj9x8v/ultra_pro_magnum_rocker_arms_1.jpg',
                'https://s25.postimg.org/l4eidg973/ultra_pro_magnum_rocker_arms_2.jpg'
            ],
            'Ultra Gold Aluminum': [
                'https://s25.postimg.org/k73hqu8b3/ultra_gold_rocker_arms_1.jpg'
            ],
        }
    }

}


def inven():
    import openpyxl
    from openpyxl.styles import Alignment
    wb = openpyxl.load_workbook('testSheet.xlsx')
    sheet = wb.active

    #VARIABLES
    running = True
    placer = 0
    brand = ""
    productNum = ""
    prodPrice = 0
    prodName = ""
    trackInventory = "TRUE"
    quantity = 4
    backorder1 = "FALSE"
    prodWeight = ""
    taxable= "TRUE"
    hidden = "FALSE"
    prodCategory = ""
    prodCategoryWeb = "http://sandiegoengineparts.com/t/"
    prodBrandWeb = "http://sandiegoengineparts.com/t/"
    prodNameWeb = "http://sandiegoengineparts.com/products?keywords="
    prodImageUrl = ""
    imageIndex = 0

    while(running):
        question = input("New Product to add! \n Is new product similar to last? y or n: ")
        if(question == 'y'):
            placer = int(placer)
            placer += 1
            placer = str(placer)
            productNum = input("Part #: ")
            prodPrice = float(input("Price: "))
            prodPrice = round(prodPrice,2)
            imageIndex = checkImages(part_types[prodType]['images'][prodLine], prodType, imageIndex)
            prodImageUrl = part_types[prodType]['images'][prodLine][imageIndex]
            imageIndex += 1
            sheet['A'+placer].alignment = Alignment(wrapText=True)
            sheet['A'+placer].value = prodName + " " + productNum
            sheet['B'+placer].alignment = Alignment(wrapText=True)
            sheet['B'+placer].value = productNum
            sheet['C'+placer].alignment = Alignment(wrapText=True)
            sheet['C'+placer].value = prodPrice
            sheet['F'+placer].alignment = Alignment(wrapText=True)
            sheet['F'+placer].value = trackInventory
            sheet['G'+placer].alignment = Alignment(wrapText=True)
            sheet['G'+placer].value = quantity
            sheet['H'+placer].alignment = Alignment(wrapText=True)
            sheet['H'+placer].value = backorder1
            appliesTo = applies()
            prodSpecs = getSpecifications(prodType)
            sheet['E'+placer].alignment = Alignment(wrapText=True)
            sheet['E'+placer].value ="""<div>
  <h2>{1} {2} Specifications</h2>
  <ul>
    <li><b>Brand: </b><a href="{6}">{3}</a></li>
    <li><b>Manufacturer's Part Number: </b>{1}</li>
    <li><b>Part Type: </b><a href="{7}">{2}</a></li>
    <li><b>Product Line: </b><a href="{8}">{0}</a></li>
    {9}
    <li><b>Sold in Set of {4}</b></li>
  </ul>

  <h2>Applies to:</h2>
  <ul>
{5}
  </ul>
</div>""".format(prodName,productNum,prodType,prodBrand,prodSet,appliesTo,prodBrandWeb,prodCategoryWeb,prodNameWeb,prodSpecs)

            sheet['I'+placer].alignment = Alignment(wrapText=True)
            sheet['I'+placer].value = prodWeight
            sheet['J'+placer].alignment = Alignment(wrapText=True)
            sheet['J'+placer].value = taxable
            sheet['K'+placer].alignment = Alignment(wrapText=True)
            sheet['K'+placer].value = hidden
            sheet['L'+placer].alignment = Alignment(wrapText=True)
            sheet['L'+placer].value = prodCategory
            sheet['M'+placer].alignment = Alignment(wrapText=True)
            sheet['M'+placer].value = prodImageUrl
            sheet['N'+placer].alignment = Alignment(wrapText=True)
            sheet['N'+placer].value = prodName + " " + productNum
            sheet['O'+placer].alignment = Alignment(wrapText=True)
            sheet['O'+placer].value = "Find " + prodName + " " + productNum +" at San Diego Engine Parts where high quality engine parts meet low costs."
            wb.save('testSheet.xlsx')

        # New Product
        else:
            if(question == 'n'):
                prodType = input("Type of Part(eg. Rocker Arms): ")
                # If (prodType in part_types) do all this...don't need check necessarily now
                prodBrand = input("Brand (eg. COMP Cams): ")
                prodLine = input("Product Line (eg. Ultra Pro Magnum): ")
                productNum = input("Part #: ")
                prodCategory = makeCategories(prodType, prodBrand)

                prodCategoryWeb += prodType.replace(' ', '-').lower()
                # prodCategoryWeb += input("Type of Part URL: ")
                prodBrandWeb += prodBrand.replace(' ', '-').lower()
                # prodBrandWeb += input("Brand URL: ")
                prodName = prodBrand + ' ' + prodLine + ' ' + prodType
                # prodName = input("(Brand - Product Line - Type): ")
                prodNameWeb += prodName.replace(' ', '+').lower()
                # prodNameWeb += input("Product Line URL: ")
                prodWeight = int(input("Weight: "))
                prodSet = part_types[prodType]['set']
                placer = input("Row: ")
                # prodRAdiameter = input("Diameter: ")
                # prodRAratio = input("Ratio: ")
                prodPrice = float(input("Price: "))
                prodPrice = round(prodPrice,2)
                if (prodLine in part_types[prodType]['images']):
                    prodImageUrl = part_types[prodType]['images'][prodLine][0]
                    imageIndex += 1
                else:
                    print('No Image was found for ' + prodLine)
                # prodImageUrl = input("Image Url: ")
                sheet['A'+placer].alignment = Alignment(wrapText=True)
                sheet['A'+placer].value = prodName + " " + productNum
                sheet['B'+placer].alignment = Alignment(wrapText=True)
                sheet['B'+placer].value = productNum
                sheet['C'+placer].alignment = Alignment(wrapText=True)
                sheet['C'+placer].value = prodPrice
                sheet['F'+placer].alignment = Alignment(wrapText=True)
                sheet['F'+placer].value = trackInventory
                sheet['G'+placer].alignment = Alignment(wrapText=True)
                sheet['G'+placer].value = quantity
                sheet['H'+placer].alignment = Alignment(wrapText=True)
                sheet['H'+placer].value = backorder1
                # running = True
                appliesTo = applies()
                prodSpecs = getSpecifications(prodType)
                sheet['E'+placer].alignment = Alignment(wrapText=True)
                sheet['E'+placer].value = """<div>
  <h2>{1} {2} Specifications</h2>
  <ul>
    <li><b>Brand: </b><a href="{6}">{3}</a></li>
    <li><b>Manufacturer's Part Number: </b>{1}</li>
    <li><b>Part Type: </b><a href="{7}">{2}</a></li>
    <li><b>Product Line: </b><a href="{8}">{0}</a></li>
    {9}
    <li><b>Sold in Set of {4}</b></li>
  </ul>

  <h2>Applies to:</h2>
  <ul>
{5}
  </ul>
</div>""".format(prodName,productNum,prodType,prodBrand,prodSet,appliesTo,prodBrandWeb,prodCategoryWeb,prodNameWeb,prodSpecs)
# .format(prodName,productNum,prodType,prodBrand,prodSet,appliesTo,prodBrandWeb,prodCategoryWeb,prodNameWeb,prodRAratio,prodRAdiameter)

                sheet['I'+placer].alignment = Alignment(wrapText=True)
                sheet['I'+placer].value = prodWeight
                sheet['J'+placer].alignment = Alignment(wrapText=True)
                sheet['J'+placer].value = taxable
                sheet['K'+placer].alignment = Alignment(wrapText=True)
                sheet['K'+placer].value = hidden
                sheet['L'+placer].alignment = Alignment(wrapText=True)
                sheet['L'+placer].value = prodCategory
                sheet['M'+placer].alignment = Alignment(wrapText=True)
                sheet['M'+placer].value = prodImageUrl
                sheet['N'+placer].alignment = Alignment(wrapText=True)
                sheet['N'+placer].value = prodName + " " + productNum
                sheet['O'+placer].alignment = Alignment(wrapText=True)
                sheet['O'+placer].value = "Find " + prodName + " " + productNum +" at San Diego Engine Parts where high quality engine parts meet low costs."
                wb.save('testSheet.xlsx')
            else:
              if(question == 'x'):
                running = False
        wb.save('testSheet.xlsx')

def applies():
    running = True
    appliesTo = ""
    while(running):
        question = input("Continue adding Applies to? y or n: ")
        if(question == 'y'):
            appliesTo = appliesTo + '    <li>' + input("Applies to: ") + '</li>\n'
        else:
            if(question == 'n'):
                return appliesTo

def getSpecifications(product_type):
    running = True
    prodSpecs = ""
    while(running):
        for k in part_types[product_type]['specs']:
            v = input('Get ' + k)
            prodSpecs += '    <li><b>' + k + '</b> ' + v + '</li>\n'
        return prodSpecs

def makeCategories(product_type, product_brand):    
    categories = product_type + ', Brands/' + product_brand + ', Brands/' + product_brand + '/' + product_brand + ' ' + product_type
    return categories

def checkImages(product_line, product_type, index):
    imageGroupLength = len(part_types[product_type]['images'])
    if (imageGroupLength == 1):
        index = 0
        return index
    else:
        if (index < imageGroupLength):
            return index
        else:
            index = 0
            return index

inven()