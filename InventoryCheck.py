#Import packages
import pyodbc 
import codecs
#Query DBs for computers.
def get_ad():
    import subprocess
    import codecs
    p = subprocess.Popen(["powershell.exe", "script/that/queries/AD.ps1"])
    p.communicate()
    i = codecs.open('output/of/ps1/script.txt', encoding='utf-16')
    format1 = []
    for row in i.readlines():
        format1.append(str(row))
    paren = []
    for x in format1:
        paren.append(x.replace('(', ''))
    quote = []
    for x in paren:
        quote.append(x.replace('"', ''))
    comma = []
    for x in quote:
        comma.append(x.replace(',', ''))
    oparen = []
    for x in comma:
        oparen.append(x.replace(')', ''))
    squote = []
    for x in oparen:
        squote.append(x.replace("'", ""))
    final = []
    for x in squote:
        final.append(x.strip())
    f = open('destination/of/output0', 'w+')
    for x in final:
        f.write(x + '\n')
    f.close()
def get_trend():
    cnxn = pyodbc.connect("Driver={SQL Server}; Server=SERVER; Database=DATABASE; Trusted_Connection=yes;")
    cursor = cnxn.cursor()
    cursor.execute("SQL QUERY HERE")
    format1 = []
    for row in cursor:
        format1.append(str(row))
    paren = []
    for x in format1:
        paren.append(x.replace('(', ''))
    quote = []
    for x in paren:
        quote.append(x.replace('"', ''))
    comma = []
    for x in quote:
        comma.append(x.replace(',', ''))
    oparen = []
    for x in comma:
        oparen.append(x.replace(')', ''))
    squote = []
    for x in oparen:
        squote.append(x.replace("'", ""))
    final = []
    for x in squote:
        final.append(x.strip())
    f = open('destination/of/output1', 'w+')
    for x in final:
        f.write(x + '\n')
    f.close()
def get_trendee():
    cnxn = pyodbc.connect("Driver={SQL Server}; Server=SERVER; Database=DATABASE; Trusted_Connection=yes;")
    cursor = cnxn.cursor()
    cursor.execute("SQL QUERY HERE")
    format1 = []
    for row in cursor:
        format1.append(str(row))
    paren = []
    for x in format1:
        paren.append(x.replace('(', ''))
    quote = []
    for x in paren:
        quote.append(x.replace('"', ''))
    comma = []
    for x in quote:
        comma.append(x.replace(',', ''))
    oparen = []
    for x in comma:
        oparen.append(x.replace(')', ''))
    squote = []
    for x in oparen:
        squote.append(x.replace("'", ""))
    final = []
    for x in squote:
        final.append(x.strip())
    f = open('destination/of/output2', 'w+')
    for x in final:
        f.write(x + '\n')
    f.close()
def get_helpDesk():
    cnxn = pyodbc.connect("Driver={SQL Server}; Server=SERVER; Database=DATABASE; Trusted_Connection=yes;")
    cursor = cnxn.cursor()
    cursor.execute("SQL QUERY HERE")
    format1 = []
    for row in cursor:
        format1.append(str(row))
    paren = []
    for x in format1:
        paren.append(x.replace('(', ''))
    quote = []
    for x in paren:
        quote.append(x.replace('"', ''))
    comma = []
    for x in quote:
        comma.append(x.replace(',', ''))
    oparen = []
    for x in comma:
        oparen.append(x.replace(')', ''))
    squote = []
    for x in oparen:
        squote.append(x.replace("'", ""))
    final = []
    for x in squote:
        final.append(x.strip())
    f = open('destination/of/output3', 'w+')
    for x in final:
        f.write(x + '\n')
    f.close()
get_trend()
get_helpDesk()
get_trendee()
get_ad()
#Open files where DB output is kept.
f = open('destination/of/output2', 'r')
g = open('destination/of/output3', 'r')
h = open('destination/of/output1', 'r')
i = open('destination/of/output0', 'r')
#Read files created in functions and split to lists of desktops and laptops for the duplicate checks.
officeScanD = []
officeScanL = []
helpDeskD = []
helpDeskL = []
trendD = []
trendL = []
ADd = []
ADl = []
for x in i.readlines():
    x = x.rstrip('\n')
    if x.startswith('D'):
        ADd.append(x)
    else:
        ADl.append(x)
for x in f.readlines():
    x = x.rstrip('\n')
    if x.startswith('L'):
        officeScanL.append(x)
    else:
        officeScanD.append(x)
for x in g.readlines():
    x = x.rstrip('\n')
    if x.startswith('D'):
        helpDeskD.append(x)
    else:
        helpDeskL.append(x)
for x in h.readlines():
    x = x.rstrip('\n')
    if x.startswith('D'):
        trendD.append(x)
    else:
        trendL.append(x)
#combine lists for DB checks.
OfficeScan_Len = []
for x in officeScanD:
    OfficeScan_Len.append(x)
for x in officeScanL:
    OfficeScan_Len.append(x)
AD_Len = []
for x in ADd:
    AD_Len.append(x)
for x in ADl:
    AD_Len.append(x)
Trend_Len = []
for x in trendD:
    Trend_Len.append(x)
for x in trendL:
    Trend_Len.append(x)
helpdesk_Len = []
for x in helpDeskD:
    helpdesk_Len.append(x)
for x in helpDeskL:
    helpdesk_Len.append(x)
#check DBs against each other for missing assets
def AD_to_TrendAV():#AD to TrendAV
    D = set(AD_Len).difference(Trend_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def AD_to_TrendEE():#AD to TrendEE
    D = set(AD_Len).difference(OfficeScan_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def AD_to_helpDesk():#AD to Helpdesk
    D = set(AD_Len).difference(helpdesk_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def TrendAV_to_AD():#TrendAV to AD
    D = set(Trend_Len).difference(AD_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def TrendAV_to_TrendEE():#TrendAV to TrendEE
    D = set(Trend_Len).difference(OfficeScan_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def TrendAV_to_Helpdesk():#TrendAV to Helpdesk
    D = set(Trend_Len).difference(helpdesk_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def TrendEE_to_AD():#TrendEE to AD
    D = set(OfficeScan_Len).difference(AD_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D 
def TrendEE_to_TrendAV():#TrendEE to TrendAV
    D = set(OfficeScan_Len).difference(Trend_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D 
def TrendEE_to_Helpdesk():#TrendEE to Helpdesk
    D = set(OfficeScan_Len).difference(helpdesk_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D 
def Helpdesk_to_AD():#Helpdesk to AD
    D = set(helpdesk_Len).difference(AD_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def Helpdesk_to_TrendAV():#Helpdesk to TrendAV
    D = set(helpdesk_Len).difference(Trend_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
def Helpdesk_to_TrendEE():#Helpdesk to TrendEE
    D = set(helpdesk_Len).difference(OfficeScan_Len)
    if str(D) == 'set()':
        D = 'Nothing'
    return D
a = AD_to_TrendAV()
b = TrendAV_to_AD()
c = AD_to_helpDesk()
d = Helpdesk_to_AD()
e = AD_to_TrendEE()
j = TrendEE_to_AD()
A = TrendAV_to_TrendEE()
B = TrendAV_to_Helpdesk()
C = TrendEE_to_TrendAV()
I = TrendEE_to_Helpdesk()
E = Helpdesk_to_TrendAV()
F = Helpdesk_to_TrendEE()
#Duplicates check.
trendDdupes = []
trendDnew = sorted(set(trendD))
for x in range(len(trendDnew)):
    if trendD.count(trendDnew[x]) > 1:
        trendDdupes.append(trendDnew[x])
trendLdupes = []
trendLnew = sorted(set(trendL))
for x in range(len(trendLnew)):
    if trendL.count(trendLnew[x]) > 1:
        trendLdupes.append(trendLnew[x])
officeScanDduplicates = []
OSdnew = sorted(set(officeScanD))
for x in range(len(OSdnew)):
    if officeScanD.count(OSdnew[x]) > 1:
        officeScanDduplicates.append(OSdnew[x])
officeScanLduplicates = []
OSlnew = sorted(set(officeScanL))
for x in range(len(OSlnew)):
    if officeScanL.count(OSlnew[x]) > 1:
        officeScanLduplicates.append(OSlnew[x])
helpDeskDduplicates = []
HelpL = sorted(set(helpDeskL))
for x in range(len(HelpL)):
    if helpDeskL.count(HelpL[x]) > 1:
        helpDeskDduplicates.append(HelpL[x])
helpDeskLduplicates = []
HelpD = sorted(set(helpDeskD))
for x in range(len(HelpD)):
    if helpDeskD.count(HelpD[x]) > 1:
        helpDeskLduplicates.append(HelpD[x])
ADDduplicates = []
ADD = sorted(set(ADd))
for x in range(len(ADD)):
    if ADd.count(ADD[x]) > 1:
        ADDduplicates.append(ADD[x])
ADLduplicates = []
ADL = sorted(set(ADl))
for x in range(len(HelpD)):
    if ADl.count(HelpD[x]) > 1:
        ADLduplicates.append(HelpD[x])
#Change set() to No Desktops or No Laptops
if str(trendDdupes) == '[]':
    trendDdupes = ''
if str(trendLdupes) == '[]':
    trendLdupes = ''
if str(officeScanDduplicates) == '[]':
    officeScanDduplicates = ''  
if str(officeScanLduplicates) == '[]':
    officeScanLduplicates = ''  
if str(helpDeskDduplicates) == '[]':
    helpDeskDduplicates = ''  
if str(helpDeskLduplicates) == '[]':
    helpDeskLduplicates = ''  
if str(ADDduplicates) == '[]':
    ADDduplicates = ''  
if str(ADLduplicates) == '[]':
    ADLduplicates = ''
if str(ADLduplicates) == '' and str(ADDduplicates) == '':
    ADDduplicates = 'No duplicates.'
if str(trendDdupes) == '' and str(trendLdupes) == '':
    trendDdupes = 'No duplicates.'
if str(officeScanLduplicates) == '' and str(officeScanDduplicates) == '':
    officeScanDduplicates = 'No duplicates.'
if str(helpDeskLduplicates) == '' and str(helpDeskDduplicates) == '':
    helpDeskDduplicates = 'No duplicates.'
#write output to file.
z = open('path/for/final/output.txt', 'w+')
z.write(
    'Below are the counts from each DB.'
    '\n'
    '\n'
    '' + str(len(AD_Len)) + ' Assets in AD'
    '\n'
    '' + str(len(Trend_Len)) + ' Assets in TrendAV'
    '\n'
    '' + str(len(OfficeScan_Len)) + ' Assets in Trendee'
    '\n'
    '' + str(len(helpdesk_Len)) + ' Assets in Helpdesk'
    '\n'
    '\n'
    'Below are the duplicates in each DB.'
    '\n'
    '\n'
    'AD: ' + str(ADDduplicates) + ' ' + str(ADLduplicates) + ''
    '\n'
    'TrendAV: ' + str(trendDdupes) + ' ' + str(trendLdupes) + ''
    '\n'
    'Trendee: ' + str(officeScanDduplicates) + ' ' + str(officeScanLduplicates) + ''
    '\n'
    'Helpdesk: ' + str(helpDeskDduplicates) + ' ' + str(helpDeskLduplicates) + ''
    '\n'
    '\n'
    'AD COMPARED TO TRENDAV, HELPDESK, AND TRENDEE'
    '\n'
    '\n'
    'TrendAV is missing ' + str(a) + ' from AD.'
    '\n'
    '\n'
    'Helpdesk is missing ' + str(c) + ' from AD.'
    '\n'
    '\n'
    'TrendEE is missing ' + str(e) + ' from AD.'
    '\n'
    '\n'
    '-------------------------------------------'
    '\n'
    '\n'
    'TRENDAV COMPARED TO AD, TRENDEE, AND HELPDESK'
    '\n'
    '\n'
    'AD is missing ' + str(b) + ' from TrendAV.'
    '\n'
    '\n'
    'TrendEE is missing ' + str(A) + ' from TrendAV.'
    '\n'
    '\n'
    'Helpdesk is missing ' + str(B) + ' from TrendAV.'
    '\n'
    '\n'
    '-------------------------------------------'
    '\n'
    '\n'
    'TRENDEE COMPARED TO AD, TRENDAV, AND HELPDESK'
    '\n'
    '\n'
    'AD is missing ' + str(j) + ' From TrendEE.'    
    '\n'
    '\n'
    'TrendAV is missing ' + str(C) + ' from TrendEE.'  
    '\n'
    '\n'
    'Helpdesk is missing ' + str(I) + ' from TrendEE.'  
    '\n'
    '\n'
    '-------------------------------------------'
    '\n'
    '\n'
    'HELPDESK COMPARED TO AD, TRENDAV, AND TRENDEE'
    '\n'
    '\n'
    'AD is missing ' + str(d) + ' from Helpdesk.'
    '\n'
    '\n'
    'TrendAV is missing ' + str(E) + '  from Helpdesk.'
    '\n'
    '\n'
    'TrendEE is missing ' + str(F) + '  from Helpdesk.'
    '\n'
    '\n'
    '-------------------------------------------')
#Close all open files.
f.close()
g.close()
h.close()
i.close()
z.close()
