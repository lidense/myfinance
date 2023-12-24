#Type this in the terminal to create a calc socket
# this explains how to use uno:
#https://blog.soutade.fr/post/2012/10/working-with-openofficelibreoffice-spreadsheets-with-python.html
#libreoffice --calc "--accept=socket,host=localhost,port=2002;urp;" --invisible

#PyOO - Pythonic interface to Apache OpenOffice API (UNO)
#PyOO allows you to control a running OpenOffice or LibreOffice program for reading and writing spreadsheet documents
import uno
import os
import unohelper

from pickletools import StackObject
import bs4 as bs 
import urllib.request

myList1Cardif=['FR0000295230','LU0340559557','LU0171307068','LU0823421689','LU0503631714','LU1951204046','FR0010288308',
'LU0217390656','LU2145461757','LU2240056015','LU2206556016'] #refers to list of isins
myList2Cardif=['MP-829523','MP-407710','0P0000VHO6','0P0000YSYV','0P0000PTZO','0P0001H0TH','MP-999283','MP-218358','0P0001KWJF','0P0001L4R7','0P0001KHJ0'] #refers to list of boursorama's tickers

myList1BourseDirect=['FR0000120271','FR0010601971','FR0013154002'] 
myList2BourseDirect=['1rPTTE','MP-546198','1rPDIM']
sheetList=[] #this is the list of all the sheets in the excel doc that is going to be connected to by uno
overallList1=[]
overallList2=[]
overallList1.append(myList1Cardif)
overallList1.append(myList1BourseDirect)
overallList2.append(myList2Cardif)
overallList2.append(myList2BourseDirect)

filename = 'cardif.txt.ods' #put here the name of the libreoffice file that you want to update

lol=[] #refers to list of lists 

 # Create an empty list to store the tickers from the current sheet
isinlist=[]

colCurrentPrice=6 #put here the number of the column of the Current Price (1st column is 0)
colDateOfUpdate=7 #put here the number of the column of the Date of Update
colWithIsin=3 #put here 
colWithTicker=1
colWithType=2
colSupport=0 #This is the column with the name of the support
#https://www.boursorama.com/bourse/opcvm/cours/MP-546198/
#https://www.boursorama.com/cours/1rPDIM/



#This is to obtain the exchange rate from USD to EUR and GBP
sauce=urllib.request.urlopen("https://www.xe.com/currencyconverter/convert/?Amount=1&From=USD&To=EUR").read()
soup=bs.BeautifulSoup(sauce,'lxml')
#exchUSD=soup.find('p', class_="result__BigRate-sc-1bsijpp-1 iGrAod").text
exchUSD=soup.find('p', class_="result__BigRate-sc-1bsijpp-1 dPdXSB").text
sauce=urllib.request.urlopen("https://www.xe.com/currencyconverter/convert/?Amount=1&From=GBP&To=EUR").read()
soup=bs.BeautifulSoup(sauce,'lxml')
#exchGBP=soup.find('p', class_="result__BigRate-sc-1bsijpp-1 iGrAod").text
exchGBP=soup.find('p', class_="result__BigRate-sc-1bsijpp-1 dPdXSB").text
print("exchange USD, GBP",float(exchUSD[0:6]),float(exchGBP[0:6]))

def build_lists(list2,list1):
    n=0
    for stock in list2:
        sauce=urllib.request.urlopen("https://www.boursorama.com/bourse/opcvm/cours/"+stock).read() 
        soup=bs.BeautifulSoup(sauce,'lxml')
        tags=soup.find('span', class_="c-instrument c-instrument--last")
        names=soup.find('a', class_='c-faceplate__company-link')
        date=soup.find('div',class_ ='c-faceplate__real-time')
        date=date.text.replace('\n                     OPCVM  dernier cours connu au ','')
        date=date.rstrip()
        names=names.text
        tags=tags.text
        tags=tags.replace(" ", "")
        l=[names.lstrip().rstrip(),list1[n],stock,tags,date]
        print(n,l)
        n+=1
        lol.append(l)
    return lol


def getStockRate(ticker,type):
    n=0
    stockrate=""
    dateofupdate=""
    names=""
    tags=''
    date=''


    if type=='opcvm':
        print("ticker",ticker)
        sauce=urllib.request.urlopen("https://www.boursorama.com/bourse/opcvm/cours/"+ticker).read()     
        soup=bs.BeautifulSoup(sauce,'lxml')
        tags=soup.find('span', class_="c-instrument c-instrument--last")
        names=soup.find('a', class_='c-faceplate__company-link')
        names=names.text
        
        date=soup.find('div',class_ ='c-faceplate__real-time')
        if date:
            date=date.text.replace('\n                     OPCVM  dernier cours connu au ','')
            date=date.rstrip()
        tags=tags.text
        tags=tags.replace(" ", "")
        currency=soup.find('span', class_="c-faceplate__price-currency").text.rstrip().lstrip()
        print(currency)
        if currency=='USD':
            tags=float(tags)*float(exchUSD[0:6])
        if currency=='GBX':
            tags=float(tags)*float(exchGBP[0:6])
            
        names=names.rstrip().lstrip()
        print(date,tags,names)
    if type=="action":
        url="http://boursorama.com/cours/"+ticker
        sauce=urllib.request.urlopen("http://boursorama.com/cours/"+ticker).read()
        soup=bs.BeautifulSoup(sauce,'lxml')
        tags=soup.find('span', class_="c-instrument c-instrument--last")
        names=soup.find('a', class_='c-faceplate__company-link')
        names=names.text
        names=names.rstrip().lstrip()
        tags=tags.text
        tags=tags.replace(' ','')

        currency=soup.find('span', class_="c-faceplate__price-currency").text.rstrip().lstrip()
        if currency=='USD':
            tags=float(tags)*float(exchUSD[0:6])
        if currency=='GBX':
            tags=(float(tags)*float(exchGBP[0:6]))/100
        date=soup.find('span',class_ ='c-instrument c-instrument--tradedate')
        date=date.text.rstrip().lstrip()
        print(date,tags,names)
    if type=="tracker":
        url="http://boursorama.com/cours/"+ticker
        sauce=urllib.request.urlopen("https://bourse.boursorama.com/bourse/trackers/cours/"+ticker).read()
        soup=bs.BeautifulSoup(sauce,'lxml')
        tags=soup.find('span', class_="c-instrument c-instrument--last")
        names=soup.find('a', class_='c-faceplate__company-link')
        names=names.text
        names=names.rstrip().lstrip()
        tags=tags.text
        currency=soup.find('span', class_="c-faceplate__price-currency").text.rstrip().lstrip()
        if currency=='USD':
            tags=float(tags)*float(exchUSD[0:6])
        if currency=='GBX':
            tags=float(tags)*float(exchGBP[0:6])
        date=soup.find('span',class_ ='c-instrument c-instrument--tradedate')
        date=date.text.rstrip().lstrip()
  

        print(date,tags,names)

    if  type=='bond':
        #sauce=urllib.request.urlopen("https://www.teleborsa.it/obbligazioni-titoli-di-stato/btpi-15st23-2-6-b2y8-it0004243512-SVQwMDA0MjQzNTEy").read()
        sauce=urllib.request.urlopen("https://www.borsaitaliana.it/borsa/obbligazioni/mot/btp/scheda/IT0004243512.html").read()
        soup=bs.BeautifulSoup(sauce,'lxml')
        #tags=soup.find('span',class_='t-text -black-warm-60 -formatPrice')
        tags=soup.find('span',class_= 't-text -right')
        #tags=tags.text
        tags=tags.text
        print(tags)
        tags=tags.replace(",",".")
        tags.replace(' ','')

        date=date[-18:]
        date=date.rstrip().lstrip()
    
        #for heading in soup.find_all(["h1", "h2", "h3"]):
        #    print(heading.name + ' ' + heading.text.strip())
        #names=soup.find('h1')

        names=soup.find('h1',class_='t-text -flola-bold -size-xlg -inherit')
        names=names.text
        names=names.rstrip().lstrip()
        #date=soup.find('span',class_ ='t-text -block -size-xs | -xs')
        date=soup.find_all('span',class_= 't-text -right')
        #date=date.text
        date=date[1].text
        print(date)
        date=date[-19:].rstrip().lstrip()
        tags=tags.rstrip().lstrip()
      
        print(tags)
        #for ch in tags:
        #    print (ch)
        print("price", tags, "names: ",names,"date: ", date)
   
        
    
    print("tags", tags)
    stockrate=float(tags)
    dateofupdate=date




    
    return stockrate,dateofupdate,names

def connect(port, filename):
    global url
    # get the uno component context from the PyUNO runtime
    #LibreOffice will be listening for localhost connection on port 2002
    localContext = uno.getComponentContext()

    # create the UnoUrlResolver
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)

    # connect to the running office
    ctx = resolver.resolve("uno:socket,host=localhost,port=" + str(port) + ";urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager

    # get the central desktop object
    DESKTOP =smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    url = unohelper.systemPathToFileUrl(os.path.abspath(filename))

    doc = DESKTOP.loadComponentFromURL(url, '_blank', 0, ()) 
    

    return doc

def getUsedArea(sheet):
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)

    cursor.gotoStartOfUsedArea(True)

    return cursor

#The used_range object represents the area of the sheet that contains actual content, excluding empty rows and columns.
# applies f to every used row in the spreadsheet
def iterRows(sheet, f):
    used_range = getUsedArea(sheet)

    for row in used_range.Rows:
        f(row)


## Establish a connection to the LibreOffice instance using the socket and port specified
doc = connect(2002, 'cardif.txt.ods')

# Get a list of all sheets in the document
sheets = doc.getSheets()

# Create an enumerator to iterate through the sheets
sheet_enum = sheets.createEnumeration()

# Iterate through each sheet in the document
while sheet_enum.hasMoreElements():
    # Get the current sheet
    sheet = sheet_enum.nextElement()
    print (sheet.getName())
    
    # Add the name of the sheet to a list
    sheetList.append(sheet.getName())


# Iterate through the list of sheets
for i in sheetList:

    if i!='Global':

        #lol=build_lists('myList2'+i,'myList1'+i)

        #print("lol",lol)
        
         # Get the current sheet using its name
        sheet = doc.getSheets().getByName(i)
 
        iterRows(sheet,lambda row: isinlist.append((row.getCellByPosition(colWithTicker, 0).String)))
        #iterRows(sheet,lambda row: isinlist.append((row.getCellByPosition(colWithIsin, 0).String)))
        #print("isinlist",isinlist)

        #print the list of tickers in the current sheet
        print("isinlist:", isinlist)
        n=0

        for l in isinlist:
            if l!=''and l!='Ticker':
                print(n)
                ticker=sheet.getCellByPosition(colWithTicker, n).String
                print(ticker)
                type=sheet.getCellByPosition(colWithType, n).String
                print(type)
                stockrate,dateofupdate,names=getStockRate(ticker,type)
                print(stockrate, dateofupdate)
                sheet.getCellByPosition(colCurrentPrice, n).setValue(stockrate)
                sheet.getCellByPosition(colDateOfUpdate, n).setString(dateofupdate)
                sheet.getCellByPosition(colSupport, n).setString(names)
                n+=1
            else:
                n+=1
                
        isinlist=[]

        









#print(value)
#iterRows(sheet,lambda row: print(row.getCellByPosition(0, 0).String))

#iterRows(sheet,lambda row: isinlist.append((row.getCellByPosition(2, 0).String)))


#print(isinlist)


#doc.storeToURL(unohelper.systemPathToFileUrl(os.path.abspath(filename2)), ())
#doc.close(True)
