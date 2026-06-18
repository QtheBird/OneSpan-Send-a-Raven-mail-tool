# mailDrafter

short explanation:

1. drag and drop supported files anywhere in the screen

2\. Tool supports OneSpan Picking List, Order confirmation, Packing Slips or Delivery Notes.
3. The first EAN will automatically be used to determine order type, click something else if you don't want the first. Do you need multiple EAN combinations? ctrl+click the ones you need, amounts get added up, type is determined by the first EAN in the list.

4\. Fill in your recipients in the boxes, there's a global box and there are specific boxes. When using split/shared keys, the recipient box of singular keys acts as a global box for the split/shared keys (ditto for key zip-psw, the same as keys button allows you to quick-copy from the split/shared keys if necessary)

5\. duplicate orders add A, B, C, D, ... to all mails to differentiate them

6\. any text box can be overwritten by the user and will be added to the mail

7\. expire dates are always 31 days in the future, sucks in a 30day month but if you don't mind, then the customer won't mind.

8\. dragging and dropping any file will clear ALL parameters, so don't customize anything without the file, you'll irrevocably lose your progress if you drop one anyway

9\. Copy subject creates a clipboard item you can paste as the subject for delivery portal mails (useful for QC only)



\# ----------------------------------Doel van de tool ACHIEVED----------------------------------

\# hoe minder ik moet ingeven of klikken hoe beter (dus als het een Picking list kan inlezen zodat ik het niet moet typen, perfect)

\# ik zou vrije selectie geven over welke mails je wil, met checkboxes voor extra, misschien standard gewoon het meest voorkomende file key pswfile pswkey, en een checkbox "other" waar je zelf iets kan invullen moest er iets uitzonderlijks nodig zijn

\# meerdere recipients tegelijk laten invullen

\# idealiter salesforce integratie maar dat is een long shot

\# vervaldatum standard een maand, maar overschrijfbaar indien nodig





\# ----------------------------------stretch goals DONE----------------------------------



\# Check if the found 

\#        als PO \\d{4}F... of (RMA?)  

\#        dan PO = temp, PO = ICO, ICO = temp

\#        (bovenstaande om ICO en PO te wisselen in geval van FOC of RMA)

\#  wat als er meerdere EAN codes zijn die files kunnen geven? -> in dat geval een multiple choice "waarvan wil je files leveren?"

\#  En wat als er meerdere orders gecombineerd moeten worden? (meerdere opties kunnen selecteren en aantal optellen)

\# support komma en puntseparation in amounts?

\# meerdere mails niet altijd naar dezelfde recipients!

\# devops met opties





\# ----------------------------------stretch goals To Dream of----------------------------------

\# ESD uitlezen om accurater de vervaldatum in te geven automatisch

\# als ESD overdatum,  dan vandaag als default nemen + 31 dagen voor vervaldag

\# opmaak aanpassen van Aptos 12 naar Gib sans mt 11?

\# CRID kan via OC in toekomst, dus best al implementeren

\# checkbox where you can fill in what you need for a mailtype

\# pin en puk (+zip psw)

\# better recipient mail handling?

\# WARNING als niet alle essentiele vakken ingevuld zijn

