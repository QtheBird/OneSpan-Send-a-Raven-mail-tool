# ----------------------------------Doel van de tool ACHIEVED----------------------------------
# hoe minder ik moet ingeven of klikken hoe beter (dus als het een Picking list kan inlezen zodat ik het niet moet typen, perfect)
# ik zou vrije selectie geven over welke mails je wil, met checkboxes voor extra, misschien standard gewoon het meest voorkomende file key pswfile pswkey, en een checkbox "other" waar je zelf iets kan invullen moest er iets uitzonderlijks nodig zijn
# meerdere recipients tegelijk laten invullen
# idealiter salesforce integratie maar dat is een long shot
# vervaldatum standard een maand, maar overschrijfbaar indien nodig

# ----------------------------------stretch goals DONE----------------------------------

# EAN (en dus product type) kan via ingewerkte excel file, updates moeten dus gebeuren aan die excel file
#  Check even bij IT (quote Timmy: "stef of kletskop") of de ERP tool aangepast kan worden voor CRID en product type
# Check if the found 
#        als PO \d{4}F... of (RMA?)  
#        dan PO = temp, PO = ICO, ICO = temp
#        (bovenstaande om ICO en PO te wisselen in geval van FOC of RMA)
#  wat als er meerdere EAN codes zijn die files kunnen geven? -> in dat geval een multiple choice "waarvan wil je files leveren?"
#  En wat als er meerdere orders gecombineerd moeten worden? (meerdere opties kunnen selecteren en aantal optellen)
# support komma en puntseparation in amounts?
# meerdere mails niet altijd naar dezelfde recipients!

# ----------------------------------stretch goals To Do----------------------------------
# ESD uitlezen om accurater de vervaldatum in te geven automatisch
# als ESD overdatum,  dan vandaag als default nemen + 31 dagen voor vervaldag
# opmaak aanpassen van Aptos 12 naar Gib sans mt 11?
# CRID kan via OC in toekomst, dus best al implementeren
# checkbox where you can fill in what you need for a mailtype
# pin en puk (+zip psw)
# better recipient mail handling?
# WARNING als niet alle essentiele vakken ingevuld zijn

import re
import pypdfium2 as pdfium
import win32com.client as win32
import tkinter as tk
from tkinter import font
from tkinterdnd2 import DND_FILES, TkinterDnD
from datetime import datetime, timedelta
timeformat = '%d%b%Y'   #global variable om ineens alle instanties te kunnen aanpassen in toekomstige versies

def extract_pdf_data(path):
    """function that reads either a PL or an OC """
    doc = ""
    for page in pdfium.PdfDocument(path):
        doc += page.get_textpage().get_text_range()
    print(doc)
    if re.search("ORDER CONFIRMATION",doc):
        PO = re.findall(r"(?<=Purchase Order number : ).+",doc)[0]
        ICO = re.findall(r"(?<=Sales Order Number : )\d+",doc)[0]
        orderAmounts = re.findall(r"\d+(?= [\d,]+?.\d\d [\d,]+?.\d\d)",doc)
    elif re.search("PICKING LIST", doc):
        PO = re.findall(r"(?<=Your Purchase Order Number: ).+",doc)[0]
        ICO = re.findall(r"(?<=Our Order Number: )\d+",doc)[0]
        orderAmounts = re.findall(r"(?<= \d\d/\d\d/\d\d )[\d,]+",doc)
    elif re.search("PACKING SLIP",doc):
        PO = re.findall(r"(?<=Your Purchase Order Number: ).+",doc)[0]
        ICO = re.findall(r"(?<=Our Order Number: )\d+",doc)[0]
        orderAmounts = re.findall(r"[\d,]+(?=\s\d+\s\d+)",doc)
        print(orderAmounts)

    # check to see if FOC or RMA
    """if re.search(r"\dF\d", PO) or re.search(r"^R\d+",PO) or re.search(r"^RMA R\d+",PO):
        print("This is a FOC or RMA order")
        temp = PO
        PO = ICO
        ICO = temp"""
    
    tokenList = []
    if re.search(r"\d{13}",doc):
        tokenList=re.findall(r"(?<=\d{13} ).+",doc)
    bigList = []
    for i in range(0,len(tokenList)):
        bigList.append(orderAmounts[i]+": "+tokenList[i])
    if len(bigList)==1:
        read_amount_and_type(bigList)
    else:
        # default to the first item in the list
        temp = []
        temp.append(bigList[0])
        read_amount_and_type(temp)
    for item in bigList:
        listTokenType.insert(tk.END, item)

    # check for, and remove, trailing whitespaces
    PO = " ".join(PO.split())
    ICO = " ".join(ICO.split())
    return(PO,ICO)

def drop(event):
    """ definieer wat er moet gebeuren op het moment dat je een pdf dropt in het tekstvak """
    clear_all()
    entryPath.insert(tk.END, event.data)
    out = extract_pdf_data(event.data[1:-1])
    tkICO.set(out[1])
    tkPO.set(out[0])
    return event.action

def clear_path():
    """verwijder het pad naar het bestand """
    entryPath.delete(0,"end")

def clear_all():
    """ verwijder alle info ingevuld in het document """
    entryPath.delete(0,"end")
    tkICO.set("")
    tkPO.set("")
    tkAmount.set("")
    tkType.set("")
    tkCRID.set("")
    textMail.delete(0.0,"end")
    listTokenType.delete(0,"end")
    tkExpire.set((datetime.today() + timedelta(days=31)).strftime(timeformat).upper())
    check_default()
    tkDuplicate.set(1)
    for entry in boxRecList:
        entry.set("")
    for i in range(0, len(boxZipKeysRecList)):
        boxKeysRecList[i].set("")
        boxZipKeysRecList[i].set("")
    tkUrl.set("")
    tkCustomer.set("")

def get_data_to_mails():
    """ wrapper functie voor draw_mail (en de andere types) die alle data verzameld en dan de mail mail opmaakt """
    deliveryAmount = tkAmount.get()
    tokenType = tkType.get()
    ICO = tkICO.get()
    PO = tkPO.get()
    CRID = tkCRID.get()
    recipients = textMail.get(0.0,"end")
    expireDate = tkExpire.get()
    devOps = tkDevOps.get()
    duplicates = tkDuplicate.get()
    customerName = tkCustomer.get()
    Url = tkUrl.get()
    multipleKey = False

    if duplicates == 1:
        dupeTags = ["",""]
    else: 
        dupeTags = [""," A"," B"," C"," D"," E"," F"," G"," H"," I"]

    for dupe in range(1,duplicates+1):
        dupe = dupeTags[dupe]
        #mailList, ftpList, devOps
        if devOps:
            draw_devOps_mail(deliveryAmount,tokenType,ICO,PO,Url,customerName,dupe)
        else:
            if tkShared.get() !=0:
                for i in range(1, tkShared.get()+1):
                    multipleKey=True
                    newRecipients = boxKeysRecList[i-1].get() + ";" +boxRecList[1].get()+ ";"+ recipients
                    deliveryFile = "shared key " + str(i)
                    if boxftpList[1].get():
                        draw_ftp_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expireDate, newRecipients,dupe)
                    else:
                        draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID,newRecipients,dupe)
                    if boxBoxList[3].get(): #if zip psw of keys needed, then make them for every shared key
                        newRecipients = boxZipKeysRecList[i-1].get() + ";" +boxRecList[3].get()+ ";"+ recipients
                        deliveryFile = boxList[3].get() + " " + str(i)
                        if boxftpList[3].get():
                            draw_ftp_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expireDate, newRecipients,dupe)
                        else:
                            draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID,newRecipients,dupe)
            elif tkSplit.get() !=0:
                multipleKey=True
                for i in range(1, tkSplit.get()+1):
                    newRecipients = boxKeysRecList[i-1].get() + ";" +boxRecList[1].get()+ ";"+ recipients
                    deliveryFile = "split key " + str(i)
                    if boxftpList[1].get():
                        draw_ftp_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expireDate, newRecipients,dupe)
                    else:
                        draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID,newRecipients,dupe)
                    if boxBoxList[3].get(): #if zip psw of keys needed, then make them for every shared key
                        newRecipients = boxZipKeysRecList[i-1].get() + ";" +boxRecList[3].get()+ ";"+ recipients
                        deliveryFile = boxList[3].get() + " " + str(i)
                        if boxftpList[3].get():
                            draw_ftp_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expireDate, newRecipients,dupe)
                        else:
                            draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID,newRecipients,dupe)
            # now loop over all other possible mails 
            for i in range(0,len(boxList)):
                if boxBoxList[i].get():       #triggert enkel op een value 1 wat wil zeggen dat de box aangevinkt staat
                    newRecipients = boxRecList[i].get() + ";" + recipients
                    deliveryFile = boxList[i].get()
                    if multipleKey and i==3:
                        pass
                    elif boxftpList[i].get():      #als de overeenkomende ftp box ook aanstaat dan FTP mail in plaats van gewoon
                        draw_ftp_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expireDate, newRecipients,dupe)
                    else: 
                        draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID,newRecipients,dupe)

def read_amount_and_type(orderlist):
    """ splitst de EAN orderlijn naar type (dat verder verwerkt wordt) en naar hoeveelheid,
        zodat de hoeveelheden gesommeerd kunnen worden.
        """
    
    orderType = re.split(": ",orderlist[0])[1].lower()
    tkType.set(find_order(orderType))
    if len(orderlist) == 1:
        orderAmount = re.split(": ",orderlist[0])[0]
        tkAmount.set(orderAmount)
    elif len(orderlist) > 1:
        orderAmount = 0
        for order in orderlist:
            orderAmount += int(re.split(": ",order)[0])
        tkAmount.set(str(orderAmount))

def find_order(order):
    """filtert alle EAN descriptions (order) naar een output die wij in mails kunnen gebruiken (ordertype)
        wil je een nieuwe filter toevoegen, hou dan rekening met de volgorde: 
            eerst meer algemeen, daarna specifiek zodat je geen interferentie geeft met filters lager op de lijst, 
            en dat elk order in kleine letters staat, op "OAS" filteren doe je dus met "oas"
    """
    ordertype = "DP"
    if re.search("digipass go ",order):
        ordertype+= "GO" + re.findall(r"(?<=digipass go )\d+",order)[0]
    elif re.search("digipass go",order):
        ordertype+= "GO" + re.findall(r"(?<=digipass go)\d+",order)[0]
    elif re.search("digipass fx ",order):
        ordertype+= " FX" + re.findall(r"(?<=digipass fx )\d+",order)[0]
    elif re.search("digipass fx",order):
        ordertype+= " FX" + re.findall(r"(?<=digipass fx)\d+",order)[0]
    elif re.search("digipass ",order):
        ordertype+= re.findall(r"(?<=digipass )\d+",order)[0]
    elif re.search("digipass\d",order):
        ordertype+= re.findall(r"(?<=digipass)\d+",order)[0]
    elif re.search("fx ",order):
        ordertype+= " FX" + re.findall(r"(?<=fx )\d+",order)[0] 
    elif re.search("authentication suite",order):
        ordertype = "OAS"
    elif re.search("mobile authenticator es",order):
        ordertype = "MA ES"
    elif re.search("^mss",order):
        ordertype = "MSS"
    elif re.search("^mas",order):
        ordertype = "MAS"
    elif re.search("oas enterprise",order):
        ordertype = "OAS"
    elif re.search("^oas",order):
        ordertype = "OAS"
    elif re.search("onespan authentication",order):
        ordertype = "OAS"

    return ordertype

def call_find_order(event):
    """ functie die meerdere inputs verwerkt en telkens read_amount_and_type oproept
        output van curseselection is (0,) met telkens een comma en een cijfer toegevoegd 
        """
    templist = []
    bigList = listTokenType.get(0,"end")
    for k in listTokenType.curselection():
        templist.append(bigList[k])
    read_amount_and_type(templist)

def date_values():
    """functie die 'alle' mogelijke datums weergeeft waartussen de spinbox kan draaien
        dit door 45 dagen terug te gaan en 90 dagen verder van de voorkeursdatum
        de voorkeursdatum is datum vandaag + 1 maand (dit door gemiddelde 4.35 weken per maand op te tellen)
    """
    dates = []
    defaultDate = (datetime.today() + timedelta(days=31)).strftime(timeformat).upper()
    for i in range(0,91):
        dates.append((datetime.strptime(defaultDate,timeformat) + timedelta(days=i)).strftime(timeformat).upper())
    for i in range(-45,0):
        dates.append((datetime.strptime(defaultDate,timeformat) + timedelta(days=i)).strftime(timeformat).upper())
    return dates

def draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, recipients,dupe):
    """ maakt een draft en slaat die op in de inbox van de gebruiker
        alle inputs zijn strings, ook de getallen, 
        meerdere recipients in 1 string worden gescheiden door een puntkomma ;
    """
    mailSignature = "\n\n\nBest regards\nOnespan Logistics\n\n---------------------------------------------------------------------------\nOneSpan destroys the Digipass secrets of Digipass devices a certain period after delivery. This implies that redelivery of your DPX-, PSKC- or PIN-files is not possible anymore after this period.\n---------------------------------------------------------------------------\nIn case you have technical questions, please contact support@OneSpan.com\n---------------------------------------------------------------------------\nNotice of U.S. Export Controls for Restricted Products: Onespan and the Onespan products, technology software, deliverables, technical information, and related documents and materials and/or any other items deemed an export by the U.S. Government or EU jurisdictions are subject to the laws and regulations of the United States and the relevant EU jurisdictions. Diversion contrary to U.S., EU law is strictly prohibited.\nCONFIDENTIALITY NOTICE: The information contained in this transmittal, including any attachment, is privileged and confidential information and is intended only for the person or entity to which it is addressed.\n---------------------------------------------------------------------------"
    outlook = win32.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)  #  0: olMailItem
    mail.Subject = deliveryAmount+"pcs "+tokenType+" (ICO: "+ICO+" / PO: "+PO+" / CRID: "+CRID+")"+dupe
    mail.Body = "Dear Customer,\n\n\nPlease find attached the "+deliveryFile+" for your order for "+deliveryAmount+"pcs "+tokenType+" (ICO: "+ICO+" / PO: "+PO+" / CRID: "+CRID+")"+dupe+mailSignature
    mail.To = recipients
    mail.save()

def draw_ftp_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expiredate, recipients,dupe):
    """ maakt een draft en slaat die op in de inbox van de gebruiker
        alle inputs zijn strings, ook de getallen, 
        meerdere recipients in 1 string worden gescheiden door een puntkomma ;
    """
    mailSignature = "\n\n\nBest regards\nOnespan Logistics\n\n---------------------------------------------------------------------------\nOneSpan destroys the Digipass secrets of Digipass devices a certain period after delivery. This implies that redelivery of your DPX-, PSKC- or PIN-files is not possible anymore after this period.\n---------------------------------------------------------------------------\nIn case you have technical questions, please contact support@OneSpan.com\n---------------------------------------------------------------------------\nNotice of U.S. Export Controls for Restricted Products: Onespan and the Onespan products, technology software, deliverables, technical information, and related documents and materials and/or any other items deemed an export by the U.S. Government or EU jurisdictions are subject to the laws and regulations of the United States and the relevant EU jurisdictions. Diversion contrary to U.S., EU law is strictly prohibited.\nCONFIDENTIALITY NOTICE: The information contained in this transmittal, including any attachment, is privileged and confidential information and is intended only for the person or entity to which it is addressed.\n---------------------------------------------------------------------------"
    outlook = win32.Dispatch("Outlook.Application")

    date = str(datetime.today().strftime('%Y%m%d'))
    ftpPart = "\n\nhttps://ftp.onespan.com"+"\nusername: "+ICO+"_"+date+"_"+deliveryFile.upper() + "\npassword: \naccount expires: " + expiredate
    mail = outlook.CreateItem(0)  #  0: olMailItem
    mail.Subject = deliveryAmount+"pcs "+tokenType+" (ICO: "+ICO+" / PO: "+PO+" / CRID: "+CRID+")" + dupe
    mail.Body = "Dear Customer,\n\n\nPlease find herewith your credentials where you can download the " +deliveryFile+" for your order for "+deliveryAmount+"pcs "+tokenType+" (ICO: "+ICO+" / PO: "+PO+" / CRID: "+CRID+")"+dupe +ftpPart + mailSignature
    mail.To = recipients
    mail.save()

def draw_devOps_mail(deliveryAmount,tokenType,ICO,PO,Url,customerName,dupe):
    """ maakt een draft en slaat die op in de inbox van de gebruiker
        alle inputs zijn strings, ook de getallen, 
        meerdere recipients in 1 string worden gescheiden door een puntkomma ;
    """
    mailSignature = "\n\n\nBest regards\nOnespan Logistics\n\n---------------------------------------------------------------------------\nOneSpan destroys the Digipass secrets of Digipass devices a certain period after delivery. This implies that redelivery of your DPX-, PSKC- or PIN-files is not possible anymore after this period.\n---------------------------------------------------------------------------\nIn case you have technical questions, please contact support@OneSpan.com\n---------------------------------------------------------------------------\nNotice of U.S. Export Controls for Restricted Products: Onespan and the Onespan products, technology software, deliverables, technical information, and related documents and materials and/or any other items deemed an export by the U.S. Government or EU jurisdictions are subject to the laws and regulations of the United States and the relevant EU jurisdictions. Diversion contrary to U.S., EU law is strictly prohibited.\nCONFIDENTIALITY NOTICE: The information contained in this transmittal, including any attachment, is privileged and confidential information and is intended only for the person or entity to which it is addressed.\n---------------------------------------------------------------------------"
    outlook = win32.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)  #  0: olMailItem
    mail.Subject = deliveryAmount+"pcs "+tokenType+" (ICO: "+ICO+" / PO: "+PO+")"+dupe
    mail.Body = "Dear Customer,\n\n\nThe DPX-files and DPX-keys for order ICO: "+ICO+"/ PO: "+PO+dupe+"\n\n"+deliveryAmount+"pcs PRODUCTION "+tokenType+" for "+customerName+" URL: "+Url+"\nare available @: \\\\srv-colo1-fs\\tid-dpx\\"+ICO+"_"+customerName+mailSignature
    mail.To = "devops@onespan.com"
    mail.save()

def shared_keys_toggled():
    """set other key values to 0 to avoid duplicates"""
    tkSplit.set(0)
    tkKeyBox.set(0)
    tkKeyFTP.set(0)

def split_keys_toggled():
    """set other key values to 0 to avoid duplicates"""
    tkShared.set(0)
    tkKeyBox.set(0)
    tkKeyFTP.set(0)

def single_key_toggled():
    """set other key values to 0 to avoid duplicates"""
    tkShared.set(0)
    tkSplit.set(0)

def check_default():
    """ checks the default boxes whenever needed"""
    defaultbox = [1,1,1,1,0,0]
    defaultftp = [1,1,0,0,0,0]
    for i in range(0, len(boxList)):
        boxBoxList[i].set(defaultbox[i])
        boxftpList[i].set(defaultftp[i])
    tkSplit.set(0)
    tkShared.set(0)

def devops_check():
    """ when devops is clicked it will uncheck all other checkboxes by default
        when it is unclicked it sets other checkboxes back to the default mailing
    """
    if tkDevOps.get():              #triggert enkel als tkDevops AANGEZET wordt, wat wil zeggen dat al de rest uit moet
        for i in range(0,len(boxList)):     #nu alles 1 voor 1 uitzetten
            boxBoxList[i].set(0)
            boxftpList[i].set(0)
    else: 
        check_default()

def same_recipients_as_keys():
    """ when zip recipients are the same as key recipients, this will autofill the boxes upon toggle
    """
    if tkSameRecsMult.get():
        for i in range(0,len(boxZipKeysRecList)):
            boxZipKeysRecList[i].set(boxKeysRecList[i].get())
    else:
        for i in range(0,len(boxZipKeysRecList)):
            boxZipKeysRecList[i].set("")

# hoofdscherm opmaken, naam is root alle widgets die op het hoofdscherm komen moeten naar root verwijzen
# ineens ook koppelen aan drag en drop, dat wil zeggen dat je de file ANYWHERE op het hoofdscherm kan laten vallen
root = TkinterDnD.Tk()  
root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>",drop)

# tekstje voor drag en drop uit te leggen, grid zorgt voor plaatsing
lbDND = tk.Label(root,text="drag OC or PL anywhere:")
lbDND.grid(row=0, column=0)


# venster waar het pad naar de file in komt
entryPath = tk.Entry(root)
entryPath.grid(row=0,column=1,columnspan=3,sticky="W", ipadx=100)

# Label as spacer:
lbSpacer1 = tk.Label(root,text=" ")
lbSpacer1.grid(row=1, column=0)

# tweede spacer, deze is niet nodig als spacer maar dient puur zodat een drag en drop vak hier kan gemaakt worden
lbSpacer2 = tk.Label(root,text=" ")
lbSpacer2.grid(row=1, column=1,sticky="W",columnspan=3)

# label en tekstvak voor recipients, tekstvak moet groot genoeg zijn om meerdere mail-adressen op te vangen, gescheiden door ";" ","" or " " or "\n"
lbMail = tk.Label(root, text="global recipients (seperate with ;)",wraplength=90).grid(row=2, column=0)
textMail = tk.Text(root,width=40,height=5)
textMail.grid(row=2,column=1,columnspan=3,sticky="W")

# Label en tekstvak voor ICO
lbICO = tk.Label(root, text="ICO/Sales order: ").grid(row=3,column=0)
tkICO = tk.StringVar()
entryICO = tk.Entry(root, textvariable= tkICO).grid(row=3,column=1)

# label en tekstvak voor PO
lbPO = tk.Label(root,text="PO number: ").grid(row=3,column=2)
tkPO = tk.StringVar()
entryPO = tk.Entry(root, textvariable=tkPO).grid(row=3,column=3)

# label en tekstvak voor CRID
lbCRID = tk.Label(root,text="CRID: ").grid(row=4,column=0)
tkCRID = tk.StringVar()
entryCRID = tk.Entry(root, textvariable=tkCRID).grid(row=4,column=1)

# label en tekstvak voor expiration date
lbExpire = tk.Label (root, text="expire: ").grid(row=4,column=2)
tkExpire = tk.StringVar()
spinExpire = tk.Spinbox(root,textvariable=tkExpire,values=date_values(),state="normal",wrap=True).grid(row=4,column=3)


# label en tekstvak voor token amount
lbAmount = tk.Label(root,text="Token Amount: ").grid(row=5,column=0)
tkAmount = tk.StringVar()
entryAmount = tk.Entry(root, textvariable=tkAmount).grid(row=5,column=1)

#  label en tekstvak voor TokenType
lbType = tk.Label(root,text="Token Type: ").grid(row=5,column=2)
tkType = tk.StringVar()
entryType = tk.Entry(root, textvariable=tkType).grid(row=5,column=3)

# label en listbox voor alle order elementen onder deze ICO
lbTokenType = tk.Label(root,text="Select order(s)").grid(row=6,column=0)
listTokenType = tk.Listbox(root, selectmode="extended")
listTokenType.bind('<<ListboxSelect>>',call_find_order)
listTokenType.grid(row=6,column=1,sticky="W",columnspan=2, ipadx=40)

# apart venster "Frame", puur voor opmaak
mailFrame = tk.Frame(root)
mailFrame.grid(row=0,column=5,rowspan=11,sticky="NE")

# label voor keuze uit mails (in dat frame) en voor recipients
lbCheckBoxes = tk.Label(mailFrame,text="choose your mails here:").grid(row=0,column=0,sticky="W")
lbRecipients = tk.Label(mailFrame,text="add unique recipients:").grid(row=1,column=2,sticky="W")


# lijstje van alle opties voor mails (ook in dat frame)
boxList = []
boxBoxList = []
boxftpList = []
boxRecList = []

tkFile = tk.StringVar()
tkFile.set("file")
tkFileBox = tk.IntVar()
tkFileBox.set(1)
cbFile = tk.Checkbutton(mailFrame,textvariable=tkFile,variable=tkFileBox).grid(row=2,column=0,sticky="W")
tkFileFTP = tk.IntVar()
tkFileFTP.set(1)
cbFileFTP = tk.Checkbutton(mailFrame,text="FTP",variable=tkFileFTP).grid(row=2,column=1,sticky="W")
tkFileRec = tk.StringVar()
entryFileRec = tk.Entry(mailFrame, textvariable=tkFileRec).grid(row=2,column=2,sticky="W")
boxList.append(tkFile)
boxBoxList.append(tkFileBox)
boxftpList.append(tkFileFTP)
boxRecList.append(tkFileRec)

tkKey = tk.StringVar()
tkKey.set("key")
tkKeyBox= tk.IntVar()
tkKeyBox.set(1)
cbKey = tk.Checkbutton(mailFrame,textvariable=tkKey,variable=tkKeyBox, command=single_key_toggled).grid(row=3,column=0,sticky="W")
tkKeyFTP = tk.IntVar()
tkKeyFTP.set(1)
cbKeyFTP = tk.Checkbutton(mailFrame,text="FTP",variable=tkKeyFTP).grid(row=3,column=1,sticky="W")
tkKeyRec = tk.StringVar()
entryKeyRec = tk.Entry(mailFrame, textvariable=tkKeyRec).grid(row=3,column=2,sticky="W")
boxList.append(tkKey)
boxBoxList.append(tkKeyBox)
boxftpList.append(tkKeyFTP)
boxRecList.append(tkKeyRec)

tkFileZip = tk.StringVar()
tkFileZip.set("zip-psw of the file")
tkFileZipBox = tk.IntVar()
tkFileZipBox.set(1)
cbFileZip = tk.Checkbutton(mailFrame,textvariable=tkFileZip,variable=tkFileZipBox).grid(row=4,column=0,sticky="W")
tkFileZipFTP = tk.IntVar()
cbFileZipFTP = tk.Checkbutton(mailFrame,text="FTP",variable=tkFileZipFTP).grid(row=4,column=1,sticky="W")
tkFileZipRec = tk.StringVar()
entryFileZipRec = tk.Entry(mailFrame, textvariable=tkFileZipRec).grid(row=4,column=2,sticky="W")
boxList.append(tkFileZip)
boxBoxList.append(tkFileZipBox)
boxftpList.append(tkFileZipFTP)
boxRecList.append(tkFileZipRec)

tkKeyZip = tk.StringVar()
tkKeyZip.set("zip-psw of the key")
tkKeyZipBox = tk.IntVar()
tkKeyZipBox.set(1)
cbKeyZip = tk.Checkbutton(mailFrame,textvariable=tkKeyZip,variable=tkKeyZipBox).grid(row=5,column=0,sticky="W")
tkKeyZipFTP = tk.IntVar()
cbKeyZipFTP = tk.Checkbutton(mailFrame,text="FTP",variable=tkKeyZipFTP).grid(row=5,column=1,sticky="W")
tkKeyZipRec = tk.StringVar()
entryKeyZipRec = tk.Entry(mailFrame, textvariable=tkKeyZipRec).grid(row=5,column=2,sticky="W")
boxList.append(tkKeyZip)
boxBoxList.append(tkKeyZipBox)
boxftpList.append(tkKeyZipFTP)
boxRecList.append(tkKeyZipRec)

tkSerial = tk.StringVar()
tkSerial.set("serial list")
tkSerialBox = tk.IntVar()
tkSerialBox.set(0)
cbSerial = tk.Checkbutton(mailFrame,textvariable=tkSerial,variable=tkSerialBox).grid(row=6,column=0,sticky="W")
tkSerialFTP = tk.IntVar()
cbSerialFTP = tk.Checkbutton(mailFrame,text="FTP",variable=tkSerialFTP).grid(row=6,column=1,sticky="W")
tkSerialRec = tk.StringVar()
entrySerialRec = tk.Entry(mailFrame, textvariable=tkSerialRec).grid(row=6,column=2,sticky="W")
boxList.append(tkSerial)
boxBoxList.append(tkSerialBox)
boxftpList.append(tkSerialFTP)
boxRecList.append(tkSerialRec)

# knoppen voor shared en split keys
lbShared = tk.Label (mailFrame, text="shared keys: ").grid(row=7,column=0,sticky="W")
tkShared = tk.IntVar()
spinShared = tk.Spinbox(mailFrame,textvariable=tkShared,values=[0,2,3],wrap=True,command=shared_keys_toggled).grid(row=7,column=1)

lbSplit = tk.Label (mailFrame, text="split keys: ").grid(row=8,column=0,sticky="W")
tkSplit = tk.IntVar()
spinSplit = tk.Spinbox(mailFrame,textvariable=tkSplit,values=[0,2,3],wrap=True, command=split_keys_toggled).grid(row=8,column=1)

# ontvangers voor shared of split keys
lbmultipleRecipients = tk.Label(mailFrame,text="recipients for up to 3 keys: ").grid(row=7,column=2,sticky="W")
boxKeysRecList = []
tkKey1Rec = tk.StringVar()
entryKey1Rec = tk.Entry(mailFrame, textvariable=tkKey1Rec).grid(row=8,column=2)
boxKeysRecList.append(tkKey1Rec)
tkKey2Rec = tk.StringVar()
entryKey2Rec = tk.Entry(mailFrame, textvariable=tkKey2Rec).grid(row=9,column=2)
boxKeysRecList.append(tkKey2Rec)
tkKey3Rec = tk.StringVar()
entryKey3Rec = tk.Entry(mailFrame, textvariable=tkKey3Rec).grid(row=10,column=2)
boxKeysRecList.append(tkKey3Rec)

# ontvangers voor de zips van shared of split keys
lbZipMultipleRecipients = tk.Label(mailFrame,text="recipients key zip psw: ").grid(row=7,column=3,sticky="W")
boxZipKeysRecList = []
tkZipKey1Rec = tk.StringVar()
entryZipKey1Rec = tk.Entry(mailFrame, textvariable=tkZipKey1Rec).grid(row=8,column=3)
boxZipKeysRecList.append(tkZipKey1Rec)
tkZipKey2Rec = tk.StringVar()
entryZipKey2Rec = tk.Entry(mailFrame, textvariable=tkZipKey2Rec).grid(row=9,column=3)
boxZipKeysRecList.append(tkZipKey2Rec)
tkZipKey3Rec = tk.StringVar()
entryZipKey3Rec = tk.Entry(mailFrame, textvariable=tkZipKey3Rec).grid(row=10,column=3)
boxZipKeysRecList.append(tkZipKey3Rec)

# knop om 1 op 1 dezelfde ontvangers te nemen als shared of split keys
tkSameRecsMult = tk.IntVar()
cbSameRecsMult = tk.Checkbutton(mailFrame, variable=tkSameRecsMult, text="same as keys",command=same_recipients_as_keys).grid(row=11,column=3)

lbSpacer3 = tk.Label(mailFrame,text=" ").grid(row=9,column=0)

lbDuplicate = tk.Label(mailFrame,text="duplicate orders").grid(row=11,column=0,sticky="W")
lbDupeExplanation = tk.Label(mailFrame,text="duplicates use A, B, C, ... in mail to differentiate. 1 means no duplication",
                             wraplength=100,justify="left").grid(row=12,column=0,sticky="W")
tkDuplicate = tk.IntVar()
tkDuplicate.set(1)
spinDuplicate = tk.Spinbox(mailFrame,from_=1,to=9,wrap=True,textvariable=tkDuplicate).grid(row=11,column=1)

# bijkomend ALTERNATIEF voor devops EN/OF files naar een locatie in sharepoint
tkDevOps = tk.IntVar()
cbDevOps = tk.Checkbutton(mailFrame,text="devOps", variable=tkDevOps,command=devops_check).grid(row=1,column=1,sticky="W")

lbCustomer = tk.Label(root,text="Customer name (devops): ").grid(row=7,column=0)
tkCustomer = tk.StringVar()
entryCustomer = tk.Entry(root, textvariable=tkCustomer).grid(row=7,column=1)

lbUrl = tk.Label(root,text="URL (devops): ").grid(row=7, column=2)
tkUrl = tk.StringVar()
entryUrl = tk.Entry(root,textvariable=tkUrl).grid(row=7,column=3)
# url en klantnaam moet dus ook ingevuld kunnen worden voor devOps

# make a clear button for the file path
# btnClear = tk.Button(root, text="clear path", command=clear_path).grid(row=0,column=4)

# make a button that clears everything
btnClearALL = tk.Button(root, text="clear all", command=clear_all).grid(row=0,column=4)

# make a run button -> first it gets all data from all fields, then it drafts the mails
btnRun = tk.Button(root, text= "draft mails", command=get_data_to_mails).grid(row=8,column=4)
lbSpacer4 = tk.Label(root,text="        ").grid(row=9,column=8)

# customizing the window

#p1 = tk.PhotoImage(file = 'Onespan_Logo.png')
#root.iconphoto(False, p1)
root.title("Onespan: send-a-raven                                             -- version 0.3.2 (beta)")

root.mainloop()




