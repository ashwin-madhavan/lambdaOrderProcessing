import csv
import json
import re

import numpy as numpy
import xlsxwriter

import sys
import csv

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

# RENAME "orders.csv" and "quotes.csv" to your respective orders and quotes files
# ordersDatafile = "C:/Users/ashwi/Desktop/orders.csv"
ordersDatafile = "orders.csv"
ordersData = list(csv.reader(open(ordersDatafile, encoding='utf-8')))

quotesDatafile = "quotes.csv"
quotesData = list(csv.reader(open(quotesDatafile, encoding='utf-8')))

# create xlsx writer
workbook = xlsxwriter.Workbook("order+quotes_processed2.xlsx")
worksheet = workbook.add_worksheet()


# functions
def add_to_xlsx(list):
    for x in range(len(list)):
        worksheet.write(chr(x + 65) + str(row_num_excel), list[x])


def find_index(headerStr, data):
    try:
        return data[0].index(headerStr)
    except:
        return -1


def raw(row, num):
    try:
        data = row[num]
        return data
    except:
        return 'N/A'


def parseJSON(customerType, JSONString):
    data = json.loads(JSONString)
    productsInQuote = []
    try:
        for item in data:
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
                        # partitionedString = processor.partition('x')
                        # numProcessors = int(str(partitionedString[0]))
                    except:
                        numProcessors = 0

                if 'gpu' == subItems['title'].lower() or 'gpus' == subItems['title'].lower():
                    GPU = subItems['description']
                    try:
                        partitionedString = GPU.partition('x')
                        numGPUs = int(str(partitionedString[0]))
                    except:
                        numGPUs = 0

                GPUStrList = [
                    ['RTX 6000*', 'VCQ-RTX6000-BLK', 'VCQ-RTX6000-EDU'],
                    ['RTX A4000*', 'hold', 'hold'],
                    ['RTX 8000', 'VCQ-RTX8000-BLK', 'VCQ-RTX8000-EDU'],
                    ['A100 80 GB', '935-23587-0000-200', '935-23587-0000-200'],
                    ['A100 40 GB', '935-23587-0000-200', '935-23587-0000-200'],
                    ['A100 PCLE', '900-21001-0000-000', '900-21001-0000-000'],
                    ['RTX 5000', 'VCQ-RTX5000-BLK', 'VCQ-RTX5000-EDU'],
                    ['RTX A6000', 'VCQ-RTX6000-BLK', 'VCQ-RTX6000-EDU']
                ]
                for GPUInfo in GPUStrList:
                    split = GPUInfo[0].split(" ")
                    count = 0
                    indexStr = 0
                    for str in split:
                        if '*' in str:
                            gpuSplit = GPU.upper().split(" ")
                            for gpuStr in gpuSplit:
                                if str[0:len(str)-1] == gpuStr:
                                    count += 1
                                    break

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
                    if customerType == 'EDU':
                        GPUProductNumber = GPUStrList[indexStr][2]
                    else:
                        GPUProductNumber = GPUStrList[indexStr][1]
                except:
                    GPUProductNumber = 'No GPU Product Number'
                #############################

                if 'THREADRIPPER' in processor.upper() and '3960X' in processor.upper():
                    CPUProductNumber = '100-000000010'
                elif 'THREADRIPPER' in processor.upper() and '3990X' in processor.upper():
                    CPUProductNumber = '100-000000163'
                elif 'THREADRIPPER' in processor.upper() and 'PRO' in processor.upper() and '3995WX' in processor.upper():
                    CPUProductNumber = '100-000000087'
                elif 'THREADRIPPER' in processor.upper() and '3970X' in processor.upper():
                    CPUProductNumber = '100-000000011'
                elif 'THREADRIPPER' in processor.upper() and 'PRO' in processor.upper() and '3975WX' in processor.upper():
                    CPUProductNumber = '100-000000086'
                elif 'THREADRIPPER' in processor.upper() and 'PRO' in processor.upper() and '3955WX' in processor.upper():
                    CPUProductNumber = '100-000000167'

            productList.append(processor)
            productList.append(numProcessors * item['quantity'])
            productList.append(GPU)
            productList.append(numGPUs * item['quantity'])
            productList.append(OS)
            productList.append(GPUProductNumber)
            productList.append(CPUProductNumber)
            # productList.append(float(item['quantity']) * float(item['unit_price']))
            totalPrice = float(item['quantity']) * float(item['unit_price'])
            totalPrice = round(totalPrice, 2)
            productList.append(totalPrice)

            productsInQuote.append(productList)
    except:
        print("error in quote: " + row[0])
    return productsInQuote


def customerTypeFxn(row, num):
    try:
        email = row[num]
        if 'edu' in email.lower() or 'univ' in email.lower() or 'college' in email.lower():
            return 'EDU'
        else:
            return 'Standard'
    except:
        return 'Standard'


def salesMappingFxn(row, num):
    try:
        salesPersonID = str(row[num])
        toAppend = sales_mapping[salesPersonID]
        return toAppend
    except:
        return 'none'


def dateFxn(row, num, firstBound, secondBound):
    try:
        string = row[num]
        substring = string[firstBound:secondBound]
        return substring
    except:
        return "N/A"


# global variables
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

headers = quotesHead
for x in ordersHead:
    headers.append(x)

add_to_xlsx(headers)
row_num_excel += 1

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
            print(">>>ERROR with processing performed action: " + quotesPerformActionOnItem[y] + " on Row: " + x)

    quotes_mapping[quoteToBeMapped[0]] = quoteToBeMapped

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
count = 0

for x in quotes_mapping:
    try:
        quoteID = quotes_mapping[x][0]
        temp = orders_mapping[quoteID]
        quotes_mapping[x].append(temp[1])
        quotes_mapping[x].append(temp[2])
        quotes_mapping[x].append(temp[3])
        quotes_mapping[x].append(temp[4])
        quotes_mapping[x].append("true")

        count += 1
    except:
        continue

# print("Total quotes matched to an order: " + count)

for x in quotes_mapping:
    temp = quotes_mapping[x]
    temp[9] = str(temp[9])
    add_to_xlsx(temp)
    row_num_excel += 1

workbook.close()
print("finished!")
