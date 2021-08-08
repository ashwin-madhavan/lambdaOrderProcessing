import csv
import json
import re

import numpy as numpy
import xlsxwriter

import sys
import csv

# the below handles reading in large csv files
maxInt = sys.maxsize
while True:
    # decrease the maxInt value by factor 10
    # as long as the OverflowError occurs.
    try:
        csv.field_size_limit(maxInt)
        break
    except OverflowError:
        maxInt = int(maxInt / 10)

# read in order and quotes file into data list
# TODO: "orders.csv" and "quotes.csv" to your respective orders and quotes files that need to be read in
ordersDatafile = "orders.csv"
ordersData = list(csv.reader(open(ordersDatafile, encoding='utf-8')))

quotesDatafile = "quotes.csv"
quotesData = list(csv.reader(open(quotesDatafile, encoding='utf-8')))

# CREATE XLSX WRITER
# TODO: rename output file if desired
workbook = xlsxwriter.Workbook("order+quotes_processed.xlsx")
worksheet = workbook.add_worksheet()


# FUNCTIONS

# adds data in a list to a row in xlsx
def add_to_xlsx(list):
    for x in range(len(list)):
        num = x + 65
        if num < 91:
            character = chr(num)
            worksheet.write(character + str(row_num_excel), list[x])
        else:
            character = 'A' + chr(num - 26)
            worksheet.write(character + str(row_num_excel), list[x])


# takes header name in orders or quotes file and finds index that header is in
def find_index(headerStr, data):
    try:
        return data[0].index(headerStr)
    except:
        return -1


# returns raw data found in a specified index from a row in orders or quotes file
def raw(row, num):
    try:
        data = row[num]
        return data
    except:
        return 'N/A'


# in the quotes file header "docdoc_line_items" gives product information generated in quote in JSON form. This function
# parses that JSON into a list with each index containing product information
def parseJSON(customerType, JSONString):
    data = json.loads(JSONString)
    productsInQuote = []
    try:
        for item in data:
            # parse the product information in data and add to productList list
            # the below information being appended can be added raw
            productList = []
            productList.append(item['title'])
            productList.append(int(item['quantity']))
            productList.append(item['product_line'])
            productList.append(float(item['unit_price']))
            subItem = item['subitems']

            OS = 'no OS'
            processor = 'no processor'
            GPU = 'no GPU'
            GPUProductNumber = 'None'
            CPUProductNumber = 'None'
            numGPUs = 0
            numProcessors = 0

            # The below finds if OS, Processor, or GPU is in the product
            for subItems in subItem:
                if 'operating system' == subItems['title'].lower() or 'operating systems' == subItems['title'].lower():
                    OS = subItems['description']

                if 'processor' in subItems['title'].lower() or 'processors' == subItems['title'].lower():
                    processor = subItems['description']
                    try:
                        partitionedString = processor.split(' ')
                        indexAMD = partitionedString.index('AMD')
                        numProcessorWithX = partitionedString[indexAMD - 1]
                        numProcessors = int(numProcessorWithX[0:len(numProcessorWithX) - 1])
                    except:
                        numProcessors = 0

                if 'gpu' == subItems['title'].lower() or 'gpus' == subItems['title'].lower():
                    GPU = subItems['description']
                    try:
                        partitionedString = GPU.partition('x')
                        numGPUs = int(str(partitionedString[0]))
                    except:
                        numGPUs = 0

                # TODO: Add new GPUs to search for if needed...
                # First index is GPU strings to check for... add "*" to end of strings that need to be check as nothing
                # more or less( RTX 6000 vs RTX A6000 both contain "6000" but for RTX 6000 you just want to check if
                # "6000" is present so input into the list as : "6000*")
                GPUStrList = [

                    ['RTX 6000*', 'VCQ-RTX6000-BLK', 'VCQ-RTX6000-EDU'],
                    ['RTX A4000*', 'hold', 'hold'],
                    ['RTX 8000', 'VCQ-RTX8000-BLK', 'VCQ-RTX8000-EDU'],
                    ['A100 80 GB', '935-23587-0000-200', '935-23587-0000-200'],
                    ['A100 40 GB', '935-23587-0000-200', '935-23587-0000-200'],
                    ['A100 PCLE', '900-21001-0000-000', '900-21001-0000-000'],
                    ['RTX 5000', 'VCQ-RTX5000-BLK', 'VCQ-RTX5000-EDU'],
                    ['RTX A6000', 'VCQ-RTX6000-BLKA6', 'VCQ-RTX6000-EDUA6'],
                    ['RTX 3090', 'hold3090', 'hold3090'],
                    ['RTX 3080', 'hold3080', 'hold3080'],
                    ['RTX 3070', 'hold3070', 'hold3070']

                ]

                # The below checks if GPU strings(index 0) in GPUStrList are present, if so a GPU Product Number is
                # generated
                for GPUInfo in GPUStrList:
                    split = GPUInfo[0].split(" ")
                    count = 0
                    indexStr = 0
                    for str in split:
                        if '*' in str:
                            gpuSplit = GPU.upper().split(" ")
                            for gpuStr in gpuSplit:
                                if str[0:len(str) - 1] == gpuStr:
                                    count += 1
                                    break
                        else:
                            if str.upper() in GPU.upper():
                                count += 1
                            else:
                                check = False
                                indexStr = GPUInfo
                                break
                    if count == len(split):
                        indexStr = GPUStrList.index(GPUInfo)
                        break
                    else:
                        indexStr = -1
                try:
                    if indexStr == -1:
                        GPUProductNumber = "None"
                    else:
                        if customerType == 'EDU':
                            GPUProductNumber = GPUStrList[indexStr][2]
                        else:
                            GPUProductNumber = GPUStrList[indexStr][1]
                except:
                    GPUProductNumber = 'None'

                # TODO: Add new CPUs to search for if needed...
                CPUStrList = [
                    ['THREADRIPPER 3960X', '100-000000010'],
                    ['THREADRIPPER 3990X', '100-000000163'],
                    ['THREADRIPPER PRO 3995WX', '100-000000087'],
                    ['THREADRIPPER 3970X', '100-000000011'],
                    ['THREADRIPPER PRO 3975WX', '100-000000086'],
                    ['THREADRIPPER PRO 3955WX', '100-0000000167'],
                ]

                # The below checks if CPU strings(index 0) in CPUStrList are present, if so a GPU Product Number is
                # generated
                for CPUInfo in CPUStrList:
                    split = CPUInfo[0].split(" ")
                    count = 0
                    indexStr = 0
                    for str in split:
                        if '*' in str:
                            cpuSplit = processor.upper().split(" ")
                            for cpuStr in cpuSplit:
                                if str[0:len(str) - 1] == gpuStr:
                                    count += 1
                                    break

                        if str.upper() in processor.upper():
                            count += 1
                        else:
                            check = False
                            indexStr = CPUInfo
                            break
                    if count == len(split):
                        indexStr = CPUStrList.index(CPUInfo)
                        break
                    else:
                        indexStr = -1
                try:
                    if indexStr == -1:
                        CPUProductNumber = "None"
                    else:
                        CPUProductNumber = CPUStrList[indexStr][1]
                except:
                    CPUProductNumber = 'None'

            # adds information to ProductList list
            productList.append(processor)
            productList.append(numProcessors * item['quantity'])
            productList.append(GPU)
            productList.append(numGPUs * item['quantity'])
            productList.append(OS)
            productList.append(GPUProductNumber)
            productList.append(CPUProductNumber)
            totalPrice = float(item['quantity']) * float(item['unit_price'])
            totalPrice = round(totalPrice, 2)
            productList.append(totalPrice)

            # add productList as an individual product to productsInQuote List
            productsInQuote.append(productList)
    except:
        print("error in quote: " + row[0])
    return productsInQuote


# returns if customer is a EDU or standard
def customerTypeFxn(row, num):
    try:
        email = row[num]
        if 'edu' in email.lower() or 'univ' in email.lower() or 'college' in email.lower():
            return 'EDU'
        else:
            return 'Standard'
    except:
        return 'Standard'


# maps sales ID to specific sales person
def salesMappingFxn(row, num):
    try:
        salesPersonID = str(row[num])
        toAppend = sales_mapping[salesPersonID]
        return toAppend
    except:
        return 'none'


# truncates date information to specified bounds
def dateFxn(row, num, firstBound, secondBound):
    try:
        string = row[num]
        substring = string[firstBound:secondBound - 1]
        return substring
    except:
        return "N/A"


# global variables
# TODO: edit sales ID and salesman if needed
sales_mapping = {
    '24982193515f4c62bcaed3c59d029e69': 'Cole',
    '561393b56dd248a3bbe2fa7ec5bc8a44': 'Toby',
    '83857ba9fd5a40a2a7277331247be17d': 'Clouse',
    'c8622520857c46fa9a00772f96f5126c': 'Mitesh',
    'e0fda6345b2d4b248cbfac6aefcdc04e': 'Tejas',
    'f27c8f54e7da464eb0cb431b9bebb1eb': 'Robert',
    'fa1f93d21c754fa1affab910e32be717': 'Ryan',
    'fe9f5f88bde243a3b059f2f579f2cadb': 'Alex',
    '9ffd838ec03c4b44b601d182d097c7b4': 'Cathleen',
    '': ''
}
quotes_mapping = {}
orders_mapping = {}
row_num_excel = 1
serialNum = 1

# TODO: can edit/add any information needed to be gathered from a specified column in quotes.csv before "Product(
#  s) Information"
quotesItemsToAddTuples = [
    ("Quote ID", "id", raw, int),
    ("Organization", "organization", raw, str),
    ("Quote Generation Date", "to_timestamp", dateFxn, str),
    ("First Name", "first_name", raw, str),
    ("Last Name", "last_name", raw, str),
    ("Email", "email", raw, str),
    ("Customer Type", "email", customerTypeFxn, str),
    ("Zip Code", "zipcode", raw, int),
    ("Country Code", "country_code", raw, str),
    ("Product(s) Information", "docdoc_line_items", parseJSON, str),
    ("Sales Mapping", "created_by_id", salesMappingFxn, str)
]

# TODO: can edit/add any information needed to be gathered from a specified column in orders.csv
ordersItemsToAddTuples = [
    ("Quotes Id", "quote_id", raw, int),
    ("Invoice Id", "id", raw, int),
    ("Discount", "discount", raw, int),
    ("Shipping", "shipping_and_handling", raw, float),
    ("Taxes", "taxes", raw, float)

]

quotesItemsArray = numpy.asarray(quotesItemsToAddTuples)
ordersItemsArray = numpy.asarray(ordersItemsToAddTuples)

# Parse Quotes and Orders Tuples
quotesHead = []
quotesIndexNums = []
quotesPerformActionOnItem = []
quotesItemDataType = []
for x in range(0, len(quotesItemsToAddTuples)):
    quotesHead.append(quotesItemsArray[(x, 0)])
    quotesIndexNums.append(find_index(quotesItemsArray[x, 1], quotesData))
    quotesPerformActionOnItem.append(quotesItemsArray[x, 2])
    quotesItemDataType.append(quotesItemsArray[x, 3])

ordersHead = []
ordersIndexNums = []
ordersPerformActionOnItem = []
ordersItemDataType = []
for x in range(0, len(ordersItemsToAddTuples)):
    ordersHead.append(ordersItemsArray[(x, 0)])
    ordersIndexNums.append(find_index(ordersItemsArray[x, 1], ordersData))
    ordersPerformActionOnItem.append(ordersItemsArray[x, 2])
    ordersItemDataType.append(ordersItemsArray[x, 3])

# TODO: can use quotesHead and ordersHead to create headers for output excel... however currently needs to be
#  manually typed into this headers list
headers = ["Quote ID", "Organization", "Quote Generation Date", "First Name", "Last Name", "Email", "Customer Type",
           "Zip Code", "Country Code", "Sales Mapping", "Product", "Quantity", "Product Line", "Unit Price", "CPU",
           "CPU Quantity", "GPU", "GPU Quantity", "Operating System", "GPU Product Number", "CPU Product Number",
           "Total Price", "Invoice ID", "Discount", "Shipping", "Taxes", "Quote Matched/Filled with Order"]
add_to_xlsx(headers)
row_num_excel += 1
TOTAL_INCORRECT_FORMATTED_QUOTES_DATA = 0

# for loop processes information in quotes.csv into quotesMapping list
for x in range(1, len(quotesData)):
    row = quotesData[x]
    quoteToBeMapped = []
    for y in range(0, len(quotesPerformActionOnItem)):
        try:
            if quotesPerformActionOnItem[y] == salesMappingFxn or quotesPerformActionOnItem[y] == raw or \
                    quotesPerformActionOnItem[y] == customerTypeFxn:
                toAppend = quotesPerformActionOnItem[y](row, quotesIndexNums[y])
            elif quotesPerformActionOnItem[y] == dateFxn:
                toAppend = quotesPerformActionOnItem[y](row, quotesIndexNums[y], 0, 11)
            else:
                toAppend = quotesPerformActionOnItem[y](quoteToBeMapped[6], row[quotesIndexNums[y]])

            if quotesItemDataType[y] == int:
                try:
                    toAppend = int(toAppend)
                except:
                    toAppend = ''
            quoteToBeMapped.append(toAppend)
        except:
            # print(">>>ERROR with processing performed action: " + str(quotesPerformActionOnItem[y]) + " on Row: " + str(
            #     x))
            TOTAL_INCORRECT_FORMATTED_QUOTES_DATA += 1

    quotes_mapping[quoteToBeMapped[0]] = quoteToBeMapped

# for loop processes information in orders.csv into ordersMapping list
for x in range(1, len(ordersData)):
    row = ordersData[x]
    ordersToBeMapped = []
    for y in range(0, len(ordersPerformActionOnItem)):
        if ordersPerformActionOnItem[y] == raw:
            toAppend = ordersPerformActionOnItem[y](row, ordersIndexNums[y])

        try:
            if ordersItemDataType[y] == int:
                toAppend = int(toAppend)

            elif ordersItemDataType[y] == float:
                toAppend = float(toAppend)
        except:
            toAppend = ''
        ordersToBeMapped.append(toAppend)
    orders_mapping[ordersToBeMapped[0]] = ordersToBeMapped

# Combining quotes.csv and orders.csv
for x in quotes_mapping:

    temp = quotes_mapping[x]
    try:
        if len(temp[9]) == 0:
            temp[9] = str(temp[9])
            temp[10] = temp[9]
            temp[9] = ""

            try:
                quoteID = quotes_mapping[x][0]
                ordersArray = orders_mapping[quoteID]
                ordersArray = ordersArray[1:len(ordersArray)]

                for space in range(0, 11):
                    temp.append("")
                for order in ordersArray:
                    temp.append(order)
            except:
                pass

            add_to_xlsx(temp)
            row_num_excel += 1
        else:
            for temp2 in temp[9]:
                temp2 = temp[0:9] + temp[10:11] + temp2

                try:
                    quoteID = quotes_mapping[x][0]
                    ordersArray = orders_mapping[quoteID]
                    ordersArray = ordersArray[1:len(ordersArray)]

                    for order in ordersArray:
                        temp2.append(order)
                    temp2.append("TRUE")
                except:
                    pass

                add_to_xlsx(temp2)
                row_num_excel += 1
    except:
        print("Error Matching Order")

workbook.close()
print("finished!")
