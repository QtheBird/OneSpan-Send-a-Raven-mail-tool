# ----------------------------------Doel van de tool----------------------------------
# hoe minder ik moet ingeven of klikken hoe beter (dus als het een Picking list kan inlezen zodat ik het niet moet typen, perfect)
# ik zou vrije selectie geven over welke mails je wil, met checkboxes voor extra, misschien standard gewoon het meest voorkomende file key pswfile pswkey, en een checkbox "other" waar je zelf iets kan invullen moest er iets uitzonderlijks nodig zijn
# meerdere recipients tegelijk laten invullen
# idealiter salesforce integratie maar dat is een long shot
# vervaldatum standard een maand, maar overschrijfbaar indien nodig

# ----------------------------------TO DO----------------------------------
# CRID kan via OC in toekomst
# EAN (en dus product type) kan via ingewerkte excel file, updates moeten dus gebeuren aan die excel file
#  Check even bij IT (stef of "kletskop") of de ERP tool aangepast kan worden voor CRID en product type
# Check if the found 
#        als PO \d{4}F... of (RMA?)  
#        dan PO = temp, PO = ICO, ICO = temp
#        (bovenstaande om ICO en PO te wisselen in geval van FOC of RMA)
# Can I use GIB sans MT 11 instead of APTOS 12? -> OPMAAK
#  wat als er meerdere EAN codes zijn die files kunnen geven? -> in dat geval een multiple choice "waarvan wil je files leveren?"
#  En wat als er meerdere orders gecombineerd moeten worden? (meerdere opties kunnen selecteren en aantal optellen)
#  stel ESD over tijd-> neem dan dag van vandaag als startuur



import re
import pypdfium2 as pdfium
import win32com.client as win32
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD


def extract_pdf_data(path):
    """function that reads either a PL or an OC """
    doc = ""
    for page in pdfium.PdfDocument(path):
        doc += page.get_textpage().get_text_range()
    
    if re.search("ORDER CONFIRMATION",doc):
        PO = re.findall(r"(?<=Purchase Order number : ).+",doc)[0]
        ICO = re.findall(r"(?<=Sales Order Number : )\d+",doc)[0]
        orderAmounts = re.findall(r"\d+(?= [\d,]+?.\d\d [\d,]+?.\d\d)",doc)
    elif re.search("PICKING LIST", doc):
        PO = re.findall(r"(?<=Your Purchase Order Number: ).+",doc)[0]
        ICO = re.findall(r"(?<=Our Order Number: )\d+",doc)[0]
        orderAmounts = re.findall(r"(?<= \d\d/\d\d/\d\d )\d+",doc)

    # check to see if FOC or RMA
    if re.search(r"\dF\d", PO) or re.search(r"^R\d+",PO):
        print("This is a FOC or RMA order")
        temp = PO
        PO = ICO
        ICO = temp
    
    tokenList = []
    if re.search(r"\d{13}",doc):
        tokenList=re.findall(r"(?<=\d{13} ).+",doc)

    bigList = []
    for i in range(0,len(tokenList)):
        bigList.append(orderAmounts[i]+": "+tokenList[i])
    if len(bigList)==1:
        read_amount_and_type(bigList)
    for item in bigList:
        listTokenType.insert(tk.END, item)
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
    listTokenType.delete(0,"end")

def get_data():
    """ wrapper functie voor drawmail die alle data verzameld en dan de mail mail opmaakt """
    deliveryAmount = tkAmount.get()
    tokenType = tkType.get()
    deliveryFile = ""
    ICO = tkICO.get()
    PO = tkPO.get()
    CRID = tkCRID.get()
    recipients = ""
    expiryDate = ""
    draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expiryDate,recipients)

def read_amount_and_type(orderlist):
    """ splitst de EAN orderlijn naar type (dat verder verwerkt wordt) en naar hoeveelheid,
        zodat de hoeveelheden gesommeerd kunnen worden.
        """
    print(orderlist)
    orderType = re.split(": ",orderlist[0])[1].lower()
    tkType.set(find_order(orderType))
    if len(orderlist) == 1:
        orderAmount = re.split(": ",orderlist[0])[0]
        tkAmount.set(orderAmount)
    elif len(orderlist) > 1:
        orderAmount = 0
        print("warning, multiple orders will sum but only the first order is displayed as token type")
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
    elif re.search("digipass",order):
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
    """ output van curseselection is (0,) met telkens een comma en een cijfer toegevoegd """
    templist = []
    bigList = listTokenType.get(0,"end")
    for k in listTokenType.curselection():
        templist.append(bigList[k])
    read_amount_and_type(templist)

def draw_mail(deliveryAmount,deliveryFile,tokenType,ICO,PO,CRID, expiryDate, recipients):
    """ maakt een draft en slaat die op in de inbox van de gebruiker
        alle inputs zijn strings ook de getallen, 
        meerdere recipients in 1 string worden gescheiden door een puntkomma ;
    """
    mailSignature = "\n\nBest regards\nOnespan Logistics\n---------------------------------------------------------------------------\nOneSpan destroys the Digipass secrets of Digipass devices a certain period after delivery. This implies that redelivery of your DPX-, PSKC- or PIN-files is not possible anymore after this period.\n---------------------------------------------------------------------------\nIn case you have technical questions, please contact support@OneSpan.com\n---------------------------------------------------------------------------\nNotice of U.S. Export Controls for Restricted Products: Onespan and the Onespan products, technology software, deliverables, technical information, and related documents and materials and/or any other items deemed an export by the U.S. Government or EU jurisdictions are subject to the laws and regulations of the United States and the relevant EU jurisdictions. Diversion contrary to U.S., EU law is strictly prohibited.\nCONFIDENTIALITY NOTICE: The information contained in this transmittal, including any attachment, is privileged and confidential information and is intended only for the person or entity to which it is addressed.\n---------------------------------------------------------------------------"
    outlook = win32.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)  #  0: olMailItem
    mail.Subject = deliveryAmount+"pcs "+tokenType+" (ICO: "+ICO+" / PO: "+PO+" / CRID: "+CRID+")"
    mail.Body = "Dear Customer,\n\nPlease find attached the "+deliveryFile+" for your order for "+deliveryAmount+"pcs "+tokenType+" (ICO: "+ICO+" / PO: "+PO+" / CRID: "+CRID+")"+mailSignature
    mail.To = recipients
    mail.save()


# hoofdscherm opmaken, naam is root alle widgets die op het hoofdscherm komen moeten naar root verwijzen
root = TkinterDnD.Tk()  


# tekstje voor drag en drop uit te leggen, grid zorgt voor plaatsing
lbDND = tk.Label(root,text="Drop OC or PL here:")

# voor user friendliness zorg ik ervoor dat de tekst links en lege ruimte onder het vak ook beschikbaar zijn als drag-and-drop
lbDND.drop_target_register(DND_FILES)
lbDND.dnd_bind('<<Drop>>', drop)
lbDND.grid(row=0, column=0)


# drag and drop functionaliteit zelf
entryPath = tk.Entry(root)
entryPath.drop_target_register(DND_FILES)
entryPath.dnd_bind('<<Drop>>', drop)
entryPath.grid(row=0,column=1,columnspan=3,sticky="W", ipadx=100)

# Label as spacer:
lbSpacer1 = tk.Label(root,text="---------------------")

# voor user friendliness zorg ik ervoor dat de tekst links en lege ruimte onder het vak ook beschikbaar zijn als drag-and-drop
lbSpacer1.drop_target_register(DND_FILES)
lbSpacer1.dnd_bind('<<Drop>>', drop)
lbSpacer1.grid(row=1, column=0)

# tweede spacer, deze is niet nodig als spacer maar dient puur zodat een drag en drop vak hier kan gemaakt worden
lbSpacer2 = tk.Label(root,text="----------------------------------------------------------------")

# voor user friendliness zorg ik ervoor dat de tekst links en lege ruimte onder het vak ook beschikbaar zijn als drag-and-drop
lbSpacer2.drop_target_register(DND_FILES)
lbSpacer2.dnd_bind('<<Drop>>', drop)
lbSpacer2.grid(row=1, column=1,sticky="W",columnspan=3)

# Label en tekstvak voor ICO
lbICO = tk.Label(root, text="ICO/Sales order: ").grid(row=2,column=0)
tkICO = tk.StringVar()
entryICO = tk.Entry(root, textvariable= tkICO).grid(row=2,column=1)

# label en tekstvak voor PO
lbPO = tk.Label(root,text="PO number: ").grid(row=2,column=2)
tkPO = tk.StringVar()
entryPO = tk.Entry(root, textvariable=tkPO).grid(row=2,column=3)

# label en tekstvak voor CRID
lbCRID = tk.Label(root,text="CRID: ").grid(row=3,column=0)
tkCRID = tk.StringVar()
entryCRID = tk.Entry(root, textvariable=tkCRID).grid(row=3,column=1)


# label en tekstvak voor token amount
lbAmount = tk.Label(root,text="Token Amount: ").grid(row=4,column=0)
tkAmount = tk.StringVar()
entryAmount = tk.Entry(root, textvariable=tkAmount).grid(row=4,column=1)

#  label en tekstvak voor TokenType
lbType = tk.Label(root,text="Token Type: ").grid(row=4,column=2)
tkType = tk.StringVar()
entryType = tk.Entry(root, textvariable=tkType).grid(row=4,column=3)

# label en listbox voor alle order elementen onder deze ICO
lbTokenType = tk.Label(root,text="Select order(s)").grid(row=5,column=0)
listTokenType = tk.Listbox(root, selectmode="extended")
listTokenType.bind('<<ListboxSelect>>',call_find_order)
listTokenType.grid(row=5,column=1,sticky="W",columnspan=2, ipadx=40)

# apart venster "Frame", puur voor opmaak
mailFrame = tk.Frame(root)
mailFrame.grid(row=0,column=5,rowspan=6,sticky="NE")

# label voor keuze uit mails (in dat frame)
lbCheckBoxes =tk.Label(mailFrame,text="choose your mails here:").grid(row=0,column=0,sticky="W")

# lijstje van alle opties voor mails (ook in dat frame)
tkFile = tk.StringVar()
tkFile.set("file")
tkFileBox = tk.IntVar()
tkFileBox.set(1)
cbFile = tk.Checkbutton(mailFrame,textvariable=tkFile,variable=tkFileBox).grid(row=1,column=0,sticky="W")


tkKey = tk.StringVar()
tkKey.set("key 1")
cbKey = tk.Checkbutton(mailFrame,textvariable=tkKey).grid(row=2,column=0,sticky="W")

tkFile = tk.StringVar()
tkFile.set("fileZipPSW")
cbFile = tk.Checkbutton(mailFrame,textvariable=tkFile).grid(row=3,column=0,sticky="W")

tkFile = tk.StringVar()
tkFile.set("KeyZipPSW")
cbFile = tk.Checkbutton(mailFrame,textvariable=tkFile).grid(row=4,column=0,sticky="W")

tkFile = tk.StringVar()
tkFile.set("CSV")
cbFile = tk.Checkbutton(mailFrame,textvariable=tkFile).grid(row=5,column=0,sticky="W")


# bijkomend lijstje voor FTP keuze
cbFileFTP = tk.Checkbutton(mailFrame,text="FTP").grid(row=1,column=1,sticky="W")
cbKeyFTP = tk.Checkbutton(mailFrame,text="FTP").grid(row=2,column=1,sticky="W")
cbFilePSWFTP = tk.Checkbutton(mailFrame,text="FTP").grid(row=3,column=1,sticky="W")
cbKeyPSWFTP = tk.Checkbutton(mailFrame,text="FTP").grid(row=4,column=1,sticky="W")
cbCSVFTP = tk.Checkbutton(mailFrame,text="FTP").grid(row=5,column=1,sticky="W")


# bijkomend ALTERNATIEF voor devops EN/OF files naar een locatie in sharepoint

# bijkomend lijstje met A,B,C... voor orders waarbij meerdere files verstuurd moeten worden

# bijkomend toggle voor shared of split keys

# make a clear button for the file path
# btnClear = tk.Button(root, text="clear path", command=clear_path).grid(row=0,column=4)

# make a button that clears everything
btnClearALL = tk.Button(root, text="clear all", command=clear_all).grid(row=0,column=4)

# make a run button -> first it gets all data from all fields, then it drafts the mails
btnRun = tk.Button(root, text= "draft mails", command=get_data).grid(row=6,column=1)

root.mainloop()




