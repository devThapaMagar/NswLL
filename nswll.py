"""
Application developed to get the unstructed data from the pdf to csv using the co ordinates of the blocks 
"""
from logging import exception
import re
import pdfplumber
import pandas as pd
import collections
from collections import namedtuple
import datetime
from datetime import date
from datetime import datetime
import fitz
import os

"""
Main Class Start
"""


class NswLL:
    Line = namedtuple(
        "line", "A B C D E F G H I J K L M N O P Q R S T U V W X")
    #HeaderLine = namedtuple("line","record_type invoice_no client_id customer_id customer_name date_invoice time_invoice date_delivered delivery_address1 delivery_address2 delivery_suburb delivery_pcode delivery_Instruction customer_order_no license_number invoice_image_path docket_image_path R S T U V W X")
    #DescriptionLine = namedtuple("line","record_type invoice_no product_code order_type order_quantity order_amount wet_amount gst_amount reference_no product_description original_invoice_no delivery_debtor_name debtor_code dlelivery_address1 delivery_address2 suburb post_code sell_in_units pack_size keg_serial_numbers customer_order_no_wms batch_nunmber bbd vintage")

    lines = []

    """
    Function to retrieve the string apart from the postcode
    returns string
    """

    def extractApartFromPostCode(self, s):
        arg = ""
        match = re.split(r'\d{4}', s)
        if match:
            arg = match[0]

        return arg
    """
    Function to remove unwanted string from the ABN
    returns string
    """

    def removeUnwantedStrFromAbn(self, str):
        str = str.replace(" ", "")
        str = str.replace(":", "")
        return str

    """
    Function to retrieve four digit from the string
    return string 
    """

    def extractFourDigit(self, s):
        match = re.search(r'\d{4}', s)
        digit = ""
        if match is not None:
            digit = match.group()
        return digit
    """
    Function to retrieve suburb state and postcode from the string
    return suburb, state, postcode
    """

    def extractSuburbStatePostcode(self, s):
        match = re.search(r'\d{4}', s)
        postCode = ""
        state = ""
        suburb = ""
        if match:
            postCode = match.group()
        stateList = ["NSW", "NT", "QLD", "SA", "TAS", "VIC", "WA"]

        for x in stateList:
            str = re.split(x, s)
            if len(str) > 1:
                suburb = str[0]
                state = x

        return suburb, state, postCode

    """
    Function to add the lines into the array
    return void
    """

    def addLines(self, a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x):
        self.lines.append(self.Line(a, b, c, d, e, f, g, h, i,
                          j, k, l, m, n, o, p, q, r, s, t, u, v, w, x))

    """
    Function to extract data from the pdf file
    return void
    """

    def extractData(self, file):
        abn_no = ""
        try:
            with pdfplumber.open(file) as pdf:
                pages = pdf.pages
                for page in pdf.pages:
                    table = page.extract_table()
                    text = page.extract_text()
                    if text is None:
                        doc = fitz.open(file)
                        text = doc[0].getText()

                    textArr = text.split('\n')

                    specialChar = "[@_!#$%^&*()<>?/|}{~:. ]"
                    r = re.compile('A'+specialChar+'B'+specialChar+'N')

                    abnRes = list(
                        filter(lambda x: x.startswith("ABN"), textArr))
                    if abnRes:
                        abnStr = abnRes[0]
                    else:
                        filtered_list = list(filter(r.match, textArr))
                        if filtered_list:
                            abnStr = ''.join(
                                char for char in filtered_list[0] if char.isalnum())

                    if abnStr:
                        abn_no = self.removeUnwantedStrFromAbn(
                            abnStr.split("ABN")[1])

        except Exception as e:
            print(e)

        if(abn_no == "58135579958"):
            self.daylesFordPdf(file, abn_no)
        elif(abn_no == "72097637575"):
            self.singleVineyardSellersPdf(file, abn_no)
        elif(abn_no == "57166237510"):
            self.barricaPdf(file, abn_no)
        elif(abn_no == "27612547742"):
            self.affinityPdf(file, abn_no)

    """
    Function to retrieve the data from the Dayles Ford pdf
    return void
    """

    def daylesFordPdf(self, file, abn_no):
        try:
            doc = fitz.open(file)
            billToRect = fitz.Rect(
                22.092514038085938, 176.5294189453125, 130.0463409423828, 217.53567504882812)
            shipToRect = fitz.Rect(
                381.3424987792969, 178.95936584472656, 463.36248779296875, 213.03358459472656)
    #         salesPersonRect = fitz.Rect(67.75199890136719, 298.8512268066406, 121.27948760986328, 308.9059143066406)
            invoiceNoRect = fitz.Rect(
                114.88580322265625, 158.8877410888672, 154.74566650390625, 168.9424285888672)
            dateRect = fitz.Rect(
                528.4000854492188, 62.573577880859375, 573.9400634765625, 73.64779663085938)
            deliveryInstructionRect = fitz.Rect(
                18.507869720458984, 484.541259765625, 152.44435119628906, 504.9459228515625)

            billTo = doc[0].get_textbox(billToRect)
            shipTo = doc[0].get_textbox(shipToRect)
            salesPerson = ""
            invoiceNo = doc[0].get_textbox(invoiceNoRect).split("#")[1]
            dateDelivered = doc[0].get_textbox(dateRect)
            deliveryInstruction = doc[0].get_textbox(deliveryInstructionRect)

            deliveryFullAddress = shipTo.split("\n")
            suburbStatePostcode = self.extractSuburbStatePostcode(
                deliveryFullAddress[2])
            deliveryPCode = suburbStatePostcode[2]
            deliverySuburb = suburbStatePostcode[0]
            customerName = deliveryFullAddress[0]
            deliveryAddress1 = deliveryFullAddress[1]
            deliveryAddress2 = ""

            self.addLines("H", invoiceNo, abn_no, "00", customerName, "", "", dateDelivered, deliveryAddress1, deliveryAddress2,
                          deliverySuburb, deliveryPCode, deliveryInstruction, "", "", "", file, "", "", "", "", "", "", "")

            count = 0
            while count < doc.page_count:
                blocks = doc[count].getText("blocks")
                itemList = []
                for x in blocks:

                    if(x[0] == 31.307479858398438 and (x[2] > 573 or x[2] < 575)):
                        itemList.append(x)

                for y in itemList:
                    items = y[4].split("\n")
                    itemCode = items[1]
                    orderType = ""
                    qty = items[0]
                    description = items[2]

                    self.addLines("D", invoiceNo, itemCode, orderType, qty, "", " ", " ", " ", description, " ", customerName,
                                  " ", deliveryAddress1, deliveryAddress2, deliverySuburb, deliveryPCode, " ", " ", " ", " ", " ", " ", "")
                count += 1

        except Exception as e:
            print(e)

    """
    Function to retrieve the data from affinity pdf
    return void
    """

    def affinityPdf(self, file, abn_no):

        try:
            doc = fitz.open(file)
            billToRect = fitz.Rect(
                86.47200012207031, 187.10723876953125, 184.0244903564453, 229.56192016601562)
            shipToRect = fitz.Rect(
                324.0719909667969, 187.10723876953125, 421.62451171875, 229.56192016601562)
            salesPersonRect = fitz.Rect(
                67.75199890136719, 298.8512268066406, 121.27948760986328, 308.9059143066406)
            invoiceNoRect = fitz.Rect(
                384.4800109863281, 103.3712387084961, 472.1179504394531, 113.4259262084961)
            dateRect = fitz.Rect(
                497.9519958496094, 299.4992370605469, 527.9881591796875, 309.5539245605469)
            abnRect = fitz.Rect(
                40.5359992980957, 115.10723876953125, 130.97476196289062, 125.16192626953125)

            billTo = doc[0].get_textbox(billToRect)
            shipTo = doc[0].get_textbox(shipToRect)
            salesPerson = doc[0].get_textbox(salesPersonRect)
            invoiceNo = doc[0].get_textbox(invoiceNoRect).split("\n")[1]
            dateDelivered = doc[0].get_textbox(dateRect)
        #     abnNo = doc[0].get_textbox(abnRect).split("A.B.N.")[1]

            deliveryFullAddress = shipTo.split("\n")
            suburbStatePostcode = self.extractSuburbStatePostcode(
                deliveryFullAddress[3])
            deliveryPCode = suburbStatePostcode[2]
            deliverySuburb = suburbStatePostcode[0]
            customerName = deliveryFullAddress[0]
            deliveryAddress1 = deliveryFullAddress[1]
            deliveryAddress2 = deliveryFullAddress[2]

            self.addLines("H", invoiceNo, abn_no, "00", customerName, "", "", dateDelivered, deliveryAddress1,
                          deliveryAddress2, deliverySuburb, deliveryPCode, " ", "", "", "", file, "", "", "", "", "", "", "")

            count = 0
            while count < doc.page_count:
                blocks = doc[count].getText("blocks")
                itemList = []
                for x in blocks:
                    if(x[0] == 60.98400115966797 and x[2] == 555.350830078125):
                        itemList.append(x)

                for y in itemList:
                    items = y[4].split("\n")
                    itemCode = items[1]
                    orderType = items[4]
                    qty = items[0]
                    description = items[2]

                    self.addLines("D", invoiceNo, itemCode, orderType, qty, "", " ", " ", " ", description, " ", customerName,
                                  " ", deliveryAddress1, deliveryAddress2, deliverySuburb, deliveryPCode, " ", " ", " ", " ", " ", " ", "")
                count += 1
        except Exception as e:
            print(e)
        finally:
            doc.close()

    """
    Function to retrieve data from Barrica PDF
    return void
    """

    def barricaPdf(self, file, abn_no):
        try:
            doc = fitz.open(file)

            billToRect = fitz.Rect(
                468.75, 60.07500076293945, 573.780029296875, 114.44100189208984)
            invoiceNoRect = fitz.Rect(
                192.75, 66.07499694824219, 785.274169921875, 84.51300048828125)
            dateRect = fitz.Rect(192.75, 81.07499694824219,
                                 785.2862548828125, 96.51300048828125)

            billTo = doc[0].get_textbox(billToRect)
            invoiceNo = doc[0].get_textbox(invoiceNoRect).split("\n")[3]
            dateDelivered = doc[0].get_textbox(dateRect).split("\n")[3]
            deliveryFullAddress = billTo.split("\n")

            suburbStatePostcode = self.extractSuburbStatePostcode(
                deliveryFullAddress[2])
            deliveryPCode = suburbStatePostcode[2]
            deliverySuburb = suburbStatePostcode[0]
            customerName = deliveryFullAddress[0]
            deliveryAddress1 = deliveryFullAddress[1]
            deliveryAddress2 = ""
            self.addLines("H", invoiceNo, abn_no, "00", customerName, "", "", dateDelivered, deliveryAddress1,
                          deliveryAddress2, deliverySuburb, deliveryPCode, " ", "", "", "", file, "", "", "", "", "", "", "")

            count = 0
            while count < doc.page_count:
                blocks = doc[count].getText("blocks")

                itemList = []
                for x in blocks:
                    if(x[0] == 43.5 and (x[2] > 781 or x[2] < 783)):
                        itemList.append(x)

                for y in itemList:
                    items = y[4].split("\n")
                    itemCode = items[2]
                    orderType = items[1]
                    qty = items[0]
                    description = items[3]

                    self.addLines("D", invoiceNo, itemCode, orderType, qty, "", " ", " ", " ", description, " ", customerName, " ", deliveryAddress1,
                                  deliveryAddress2, deliverySuburb, deliveryPCode, " ", " ", " ", " ", " ", " ", self.extractFourDigit(description))

                count += 1
        except Exception as e:
            print(e)
        finally:
            doc.close()

    """
    Function to retrieve data from singleVineyarSellers PDF
    return void
    """

    def singleVineyardSellersPdf(self, file, abn_no):
        try:
            doc = fitz.open(file)
            blocks = doc[0].getText("blocks")
            billToRect = fitz.Rect(
                29.20001220703125, 249.99481201171875, 137.2489776611328, 298.94488525390625)
            shipToRect = fitz.Rect(
                457.658203125, 249.994873046875, 565.7073364257812, 298.9449462890625)
    #         salesPersonRect = fitz.Rect(67.75199890136719, 298.8512268066406, 121.27948760986328, 308.9059143066406)
            invoiceNoDateRect = fitz.Rect(
                30.0, 193.22491455078125, 564.91162109375, 207.22491455078125)
    #         dateRect = fitz.Rect(528.4000854492188, 62.573577880859375, 573.9400634765625, 73.64779663085938)
            deliveryInstructionRect = fitz.Rect(
                210.73341369628906, 249.994873046875, 258.8125, 260.994873046875)

            billTo = doc[0].get_textbox(billToRect)
            shipTo = doc[0].get_textbox(shipToRect)
            salesPerson = ""

            invoiceNoDate = doc[0].get_textbox(invoiceNoDateRect).split("\n")
            invoiceNo = invoiceNoDate[0].split("#")[1]
            dateDelivered = invoiceNoDate[2].split(":")[1]
            deliveryInstruction = doc[0].get_textbox(deliveryInstructionRect)

            deliveryFullAddress = shipTo.split("\n")
            suburbStatePostcode = self.extractSuburbStatePostcode(
                deliveryFullAddress[2])
            deliveryPCode = suburbStatePostcode[2]
            deliverySuburb = suburbStatePostcode[0]
            customerName = deliveryFullAddress[0]
            deliveryAddress1 = deliveryFullAddress[1]
            deliveryAddress2 = ""

            self.addLines("H", invoiceNo, abn_no, "00", customerName, "", "", dateDelivered, deliveryAddress1, deliveryAddress2,
                          deliverySuburb, deliveryPCode, deliveryInstruction, "", "", "", file, "", "", "", "", "", "", "")

            count = 0
            while count < doc.page_count:
                blocks = doc[count].getText("blocks")
                itemList = []
                for x in blocks:
                    if(x[0] == 40.629791259765625 and (x[2] > 560 or x[2] < 575)):
                        itemList.append(x)

                for y in itemList:
                    items = y[4].split("\n")
                    itemCode = items[2]
                    orderType = ""
                    qty = items[0]
                    description = items[3]

                    self.addLines("D", invoiceNo, itemCode, orderType, qty, "", " ", " ", " ", description, " ", customerName,
                                  " ", deliveryAddress1, deliveryAddress2, deliverySuburb, deliveryPCode, " ", " ", " ", " ", " ", " ", "")
                count += 1

        except Exception as e:
            print(e)


"""
End of Main Class
"""

"""
File Path 
"""
path = ''  # file Path

files = os.listdir(path)
nswll = NswLL()

"""
Loop start for all the files in the directory
"""
for f in files:
    if f.endswith(".pdf"):
        filePath = path+f
        nswll.extractData(filePath)

"""
Loop end
"""
df = pd.DataFrame(nswll.lines)
df.head()

now = datetime.now()
nowStr = now.strftime("%d%m%Y%H%M%S")
csvPath = path + 'invoice_'+nowStr+'.csv'
df.to_csv(csvPath) # Extracts csv to the directory with naming convention invoice_dmyhms.csv
