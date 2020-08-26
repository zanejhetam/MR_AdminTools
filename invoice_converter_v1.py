#############################################################
#
# INVOICE CONVERTER 
# 
# Written by Zane Jhetam for the Milton Rum Distillery, 2020
#
#############################################################
import os
import xml.etree.ElementTree as ET
import csv

# Get user input to determine which directory to pull
invoiceFolder = raw_input("Enter the path of the folder which contains your invoices in XML format...")
os.chdir(invoiceFolder)

# List of the rows of the eventual table, with header row pre-populated
rows = [["*ContactName", "EmailAddress", "POAddressLine1", "POAddressLine2", "POAddressLine3", "POAddressLine4", "POCity", "PORegion", \
        "POPostalCode", "POCountry", "*InvoiceNumber", "Reference", "*InvoiceDate", "*DueDate", "InventoryItemCode", "*Description", \
        "*Quantity", "*UnitAmount", "Discount", "*AccountCode", "*TaxType", "TrackingName1", "TrackingOption1", "TrackingName2", \
        "TrackingOption2", "Currency", "BrandingTheme"]]

# LOOP THROUGH EVERY FILE IN THE DIRECTORY
for filename in os.listdir(invoiceFolder):
    if not filename.endswith(".xml"):
        #Skip anything that isn't an XML File
        continue
    else:
        print("Trying to parse file: " + filename)
        
        # Parse the file
        tree = ET.parse(filename)
        root = tree.getroot()

        # SKIP IF NOT INVOICE
        # (1) Does it have a single 'BusinessDoc' element?
        if len(root.findall(".//BusinessDoc")) != 1:
            print("That doesn't appear to be an invoice - skipping")
            continue

        # (2) If yes to (1), the content of that element must also be 'Invoice'
        if root.findall(".//BusinessDoc")[0].text != "Invoice":
            print("That doesn't appear to be an invoice - skipping")
            continue

    #######################
    #     XML Parsing     #
    #######################

    # ContactName (the name of the shop, including location)
    contactName = root.findall("./ControlArea/Receiver/B2BParty/Organisation/Name")[0].text

    # Email Address - BLANK FOR NOW
    emailAddress = ""

    # POAddressLine1 - BLANK FOR NOW
    POAddressLine1 = ""

    # POAddressLine2 - BLANK FOR NOW
    POAddressLine2 = ""

    # POAddressLine3 - BLANK FOR NOW
    POAddressLine3 = ""

    # POAddressLine4 - BLANK FOR NOW
    POAddressLine4 = ""

    # POCity - BLANK FOR NOW
    POCity = ""

    # PORegion - BLANK FOR NOW
    PORegion = ""

    # POPostalCode - BLANK FOR NOW
    POPostalCode = ""

    # POCountry - BLANK FOR NOW
    POCountry = ""

    # InvoiceNumber (theirs)
    invoiceNumber = root.findall("./DataArea/InvoiceDetail/InvoiceHeader/InvoiceNumber")[0].text

    # Reference (MR's invoice number) - BLANK FOR NOW
    # N.B. if implementing in future, this is a concatenation in the following format: 
    # "DDMMYY - [Authorisor Name]"
    miltonRumReference = ""

    # InvoiceDate
    invoiceDate = root.findall("./DataArea/InvoiceDetail/InvoiceHeader/InvoiceDate/DateTime/Day")[0].text + "/" + \
                  root.findall("./DataArea/InvoiceDetail/InvoiceHeader/InvoiceDate/DateTime/Month")[0].text + "/" + \
                  root.findall("./DataArea/InvoiceDetail/InvoiceHeader/InvoiceDate/DateTime/Year")[0].text

    # PaymentDueDate
    dueDate = root.findall("./DataArea/InvoiceDetail/InvoiceHeader/ReferenceList/Reference")[1].findall("Identifier")[0].text 
    dueDate = dueDate.replace("-", "/")

    ########################################
    # Product handling section of invoice
    # WE MAY HAVE TO TURN THIS INTO A LOOP

    itemList = root.findall("./DataArea/InvoiceDetail/InvoiceItemList/InvoiceItem")

    itemIndex = 0 # Placeholder to save time if we do have to write the loop...

    # Description (N.B., this is actually after 'InventoryItemCode' in the invoice but we need it to infer the InventoryItemCode)
    description = itemList[itemIndex].findall("POItemDescription/ItemName")[0].text
    
    # InventoryItemCode - note this is inferred from the "Description"
    # N.B. DIFFERENT SHOPS ARE USING DIFFERENT DESCRIPTIONS
    # To be handled manually for now
    inventoryItemCode = ""

    # Quantity
    # Dan Murphy's invoices have two fields. One gives a unit 'magnitude', and one gives 'decimal places'
    # Additionally, they area counting cases in the door - not bottles in (so assume 6 btls/case)
    quantity = 6 * (int(itemList[itemIndex].findall("InvoicedQuantity/Quantity/Number/Value")[0].text) \
                / (10 ** int(itemList[itemIndex].findall("InvoicedQuantity/Quantity/Number/NumOfDec")[0].text)))

    # Unit Cost
    # We can calculate the per-bottle unit cost by dividing the quantity of bottles (calc'd above) by the Gross Value of a case
    unitAmount = ((float(itemList[itemIndex].findall("NetPrice/MonetaryValue/Number/Value")[0].text) \
                / (10 ** int(itemList[itemIndex].findall("NetPrice/MonetaryValue/Number/NumOfDec")[0].text))))/quantity

    # Discount - GENERALLY BLANK
    discount = ""

    # AccountCode (always 230)
    accountCode = "230"

    # TaxType - Always "GST On Income"
    taxType = "GST on Income"

    # TrackingName1 - BLANK FOR NOW
    TrackingName1 = ""

    # TrackingOption1 - BLANK FOR NOW
    TrackingOption1 = ""

    # TrackingName2 - BLANK FOR NOW
    TrackingName2 = ""

    # TrackingOption2 - BLANK FOR NOW
    TrackingOption2 = ""

    # Currency - BLANK FOR NOW
    Currency = ""

    # BrandingTheme - BLANK FOR NOW
    BrandingTheme = ""

    # Add the data to the list of rows    
    rows.append([contactName, emailAddress, POAddressLine1, POAddressLine2, POAddressLine3, POAddressLine4, POCity, PORegion, POPostalCode, POCountry,\
           invoiceNumber, miltonRumReference, invoiceDate, dueDate, inventoryItemCode, description, quantity, unitAmount, discount, accountCode, \
           taxType, TrackingName1, TrackingOption1, TrackingName2, TrackingOption2, Currency, BrandingTheme])

    #################################
    #       END LOOP                #
    #################################

# Open the file
output_file = open("collated_invoices_for_xero.csv", 'w')

with output_file:
    print ("Writing file....")
    # Set up the CSV Writer and output the file
    writer = csv.writer(output_file, csv.excel)
    writer.writerows(rows)

    print("File successfully written!")
