from jira.client import JIRA
import datetime
import xlsxwriter
import time

#Start Time
startu_time = time.time()

#Authentication
'''Connect to JIRA server'''
options = {'server': 'https://tkts.sys.<ServerName>.net'}
'''Provide username and password'''
jira = JIRA(options, basic_auth=('UserName', 'Pwd***'))
print('Auth Successful')
#Dates
today = str(datetime.datetime.today())[0:10]
todayminusone = str(datetime.datetime.now() - datetime.timedelta(days = 1))[0:10]
todayminustwo = str(datetime.datetime.now() - datetime.timedelta(days = 2))[0:10]
todayminusthree = str(datetime.datetime.now() - datetime.timedelta(days = 3))[0:10]
todayminusfour = str(datetime.datetime.now() - datetime.timedelta(days = 4))[0:10]
todayminusfive = str(datetime.datetime.now() - datetime.timedelta(days = 5))[0:10]
todayminussix = str(datetime.datetime.now() - datetime.timedelta(days = 6))[0:10]
todayminusseven = str(datetime.datetime.now() - datetime.timedelta(days = 7))[0:10]

Array1 = []
Array2 = []
Array3 = []
Array4 =[]
print('entering projects')
Project = ['project=PROJ1','project=PROJ2','project=PROJ3','project=PROJ4']
for proj in Project:
    issues = jira.search_issues(proj,0,500)
    TargetIssues = []
    Aging1 = []
    Aging2 = []
    Aging3 = []
    Aging4 = []

    # Count Variables

    Critical = 0
    Blocker = 0
    Major = 0
    Minor = 0

    Critical1 = 0
    Blocker1 = 0
    Major1 = 0
    Minor1 = 0

    A1C = 0
    A1B = 0
    A1MA = 0
    A1MI = 0

    A1ETC = 0
    A1ETB = 0
    A1ETMA = 0
    A1ETMI = 0

    A2C = 0
    A2B = 0
    A2MA = 0
    A2MI = 0

    A2ETC = 0
    A2ETB = 0
    A2ETMA = 0
    A2ETMI = 0

    A3C = 0
    A3B = 0
    A3MA = 0
    A3MI = 0
    A3ET = 0

    A3ETC = 0
    A3ETB = 0
    A3ETMA = 0
    A3ETMI = 0

    A4C = 0
    A4B = 0
    A4MA = 0
    A4MI = 0
    A4ET = 0

    A4ETC = 0
    A4ETB = 0
    A4ETMA = 0
    A4ETMI = 0

    for issue in issues:
        CreatedDate = issue.fields.created[0:10]
        UpdatedDate = issue.fields.updated[0:10]
        ResolvedDate = str(issue.fields.resolutiondate)[0:10]
        IssueElement = []
        IssueElement.append(issue.key)
        IssueElement.append(issue.fields.status.name)
        IssueElement.append(issue.fields.priority.name)
        IssueElement.append(CreatedDate)
        IssueElement.append(issue.fields.resolution)
        IssueElement.append(ResolvedDate)
        TargetIssues.append(IssueElement)

        if (CreatedDate == today) or (CreatedDate == todayminusone):
            Aging0to1 = []
            Aging0to1.append(issue.key)
            Aging0to1.append(issue.fields.status.name)
            Aging0to1.append(issue.fields.priority.name)
            Aging0to1.append(CreatedDate)
            Aging0to1.append(issue.fields.resolution)
            Aging1.append(Aging0to1)
        elif (CreatedDate == todayminustwo) or (CreatedDate == todayminusthree):
            Aging1to3 = []
            Aging1to3.append(issue.key)
            Aging1to3.append(issue.fields.status.name)
            Aging1to3.append(issue.fields.priority.name)
            Aging1to3.append(CreatedDate)
            Aging1to3.append(issue.fields.resolution)
            Aging2.append(Aging1to3)
        elif (CreatedDate == todayminusfour) or (CreatedDate == todayminusfive):
            Aging3to5 = []
            Aging3to5.append(issue.key)
            Aging3to5.append(issue.fields.status.name)
            Aging3to5.append(issue.fields.priority.name)
            Aging3to5.append(CreatedDate)
            Aging3to5.append(issue.fields.resolution)
            Aging3.append(Aging3to5)
        else:
            Aging5orMore = []
            Aging5orMore.append(issue.key)
            Aging5orMore.append(issue.fields.status.name)
            Aging5orMore.append(issue.fields.priority.name)
            Aging5orMore.append(CreatedDate)
            Aging5orMore.append(issue.fields.resolution)
            Aging4.append(Aging5orMore)
    print('calculating')
    for EE in Aging1:
        if str(EE[4]) == 'None' and EE[1] != 'External Team Engaged' and EE[1] != 'Stalled' and EE[1] != 'Queued' and EE[1] != 'Awaiting More Information' and EE[2] == 'Critical':
            A1C = A1C + 1
        elif str(EE[4]) == 'None' and EE[1] != 'External Team Engaged' and EE[1] != 'Stalled' and EE[1] != 'Queued' and EE[1] != 'Awaiting More Information' and EE[2] == 'Blocker':
            A1B = A1B + 1
        elif str(EE[4]) == 'None' and EE[1] != 'External Team Engaged' and EE[1] != 'Stalled' and EE[1] != 'Queued' and EE[1] != 'Awaiting More Information' and EE[2] == 'Major':
            A1MA = A1MA + 1
        elif str(EE[4]) == 'None' and EE[1] != 'External Team Engaged' and EE[1] != 'Stalled' and EE[1] != 'Queued' and EE[1] != 'Awaiting More Information' and EE[2] == 'Minor':
            A1MI = A1MI + 1
        elif EE[1] == 'External Team Engaged' or EE[1] == 'Stalled' or EE[1] == 'Queued' or EE[1] == 'Awaiting More Information':
            if EE[2] == 'Critical':
                A1ETC = A1ETC+1
            elif EE[2] == 'Blocker':
                A1ETB = A1ETB+1
            elif EE[2] == 'Major':
                A1ETMA = A1ETMA+1
            elif EE[2] == 'Minor':
                A1ETMI = A1ETMI+1
    A1TOT = A1C+A1B+A1MA+A1MI



    for EE2 in Aging2:
        if str(EE2[4]) == 'None' and EE2[1] != 'External Team Engaged' and EE2[1] != 'Stalled' and EE2[1] != 'Queued' and EE2[1] != 'Awaiting More Information'  and EE2[2] == 'Critical':
            A2C = A2C + 1
        elif str(EE2[4]) == 'None' and EE2[1] != 'External Team Engaged' and EE2[1] != 'Stalled' and EE2[1] != 'Queued' and EE2[1] != 'Awaiting More Information' and EE2[2] == 'Blocker':
            A2B = A2B + 1
        elif str(EE2[4]) == 'None' and EE2[1] != 'External Team Engaged' and EE2[1] != 'Stalled' and EE2[1] != 'Queued' and EE2[1] != 'Awaiting More Information' and EE2[2] == 'Major':
            A2MA = A2MA + 1
        elif str(EE2[4]) == 'None' and EE2[1] != 'External Team Engaged' and EE2[1] != 'Stalled' and EE2[1] != 'Queued' and EE2[1] != 'Awaiting More Information' and EE2[2] == 'Minor':
            A2MI = A2MI + 1
        elif EE2[1] == 'External Team Engaged' or EE2[1] == 'Stalled' or EE2[1] == 'Queued' or EE2[1] == 'Awaiting More Information':
            if EE2[2] == 'Critical':
                A2ETC = A2ETC+1
            elif EE2[2] == 'Blocker':
                A2ETB = A2ETB+1
            elif EE2[2] == 'Major':
                A2ETMA = A2ETMA+1
            elif EE2[2] == 'Minor':
                A2ETMI = A2ETMI+1
    A2TOT = A2C+A2B+A2MA+A2MI


    for EE3 in Aging3:
        if str(EE3[4]) == 'None' and EE3[1] != 'External Team Engaged' and EE3[1] != 'Stalled' and EE3[1] != 'Queued' and EE3[1] != 'Awaiting More Information' and EE3[2] == 'Critical':
            A3C = A3C + 1
        elif str(EE3[4]) == 'None' and EE3[1] != 'External Team Engaged' and EE3[1] != 'Stalled' and EE3[1] != 'Queued' and EE3[1] != 'Awaiting More Information' and EE3[2] == 'Blocker':
            A3B = A3B + 1
        elif str(EE3[4]) == 'None' and EE3[1] != 'External Team Engaged' and EE3[1] != 'Stalled' and EE3[1] != 'Queued' and EE3[1] != 'Awaiting More Information' and EE3[2] == 'Major':
            A3MA = A3MA + 1
        elif str(EE3[4]) == 'None' and EE3[1] != 'External Team Engaged' and EE3[1] != 'Stalled' and EE3[1] != 'Queued' and EE3[1] != 'Awaiting More Information' and EE3[2] == 'Minor':
            A3MI = A3MI + 1
        elif EE3[1] == 'External Team Engaged' or EE3[1] == 'Stalled' or EE3[1] == 'Queued' or EE3[1] == 'Awaiting More Information':
            if EE3[2] == 'Critical':
                A3ETC = A3ETC+1
            elif EE3[2] == 'Blocker':
                A3ETB = A3ETB+1
            elif EE3[2] == 'Major':
                A3ETMA = A3ETMA+1
            elif EE3[2] == 'Minor':
                A3ETMI = A3ETMI+1
    A3TOT = A3C+A3B+A3MA+A3MI


    for EE4 in Aging4:
        if str(EE4[4]) == 'None' and EE4[1] != 'External Team Engaged' and EE4[1] != 'Stalled' and EE4[1] != 'Queued' and EE4[1] != 'Awaiting More Information' and EE4[2] == 'Critical':
            A4C = A4C + 1
        elif str(EE4[4]) == 'None' and EE4[1] != 'External Team Engaged' and EE4[1] != 'Stalled' and EE4[1] != 'Queued' and EE4[1] != 'Awaiting More Information' and EE4[2] == 'Blocker':
            A4B = A4B + 1
        elif str(EE4[4]) == 'None' and EE4[1] != 'External Team Engaged' and EE4[1] != 'Stalled' and EE4[1] != 'Queued' and EE4[1] != 'Awaiting More Information'  and EE4[2] == 'Major':
            A4MA = A4MA + 1
        elif str(EE4[4]) == 'None' and EE4[1] != 'External Team Engaged' and EE4[1] != 'Stalled' and EE4[1] != 'Queued' and EE4[1] != 'Awaiting More Information'  and EE4[2] == 'Minor':
            A4MI = A4MI + 1
        elif EE4[1] == 'External Team Engaged' or EE4[1] == 'Stalled' or EE4[1] == 'Queued' or EE4[1] == 'Awaiting More Information':
            if EE4[2] == 'Critical':
                A4ETC = A4ETC+1
            elif EE4[2] == 'Blocker':
                A4ETB = A4ETB+1
            elif EE4[2] == 'Major':
                A4ETMA = A4ETMA+1
            elif EE4[2] == 'Minor':
                A4ETMI = A4ETMI+1
    A4TOT = A4C+A4B+A4MA+A4MI


    for Element in TargetIssues:

        if (Element[5] == today) or (Element[5] == todayminusone) or (Element[5] == todayminustwo) or (Element[5] == todayminusthree) or (Element[5] == todayminusfour) or (Element[5] == todayminusfive) or (Element[5] == todayminussix):
            if Element[1] == 'Closed' or Element[1] == 'Resolved':
                if Element[2] == 'Critical':
                    Critical1 = Critical1 + 1
                    #print((Element[0])+","+(Element[1])+","+Element[5])
                elif Element[2] == 'Blocker':
                    Blocker1 = Blocker1 + 1
                    #print((Element[0]) + "," + (Element[1]) + "," + Element[5])
                elif Element[2] == 'Major':
                    Major1 = Major1 + 1
                    #print((Element[0])+","+(Element[1])+","+Element[5])
                elif Element[2] == 'Minor':
                    Minor1 = Minor1 + 1
                    #print((Element[0])+","+(Element[1])+","+Element[5])
                else:
                    print("Invalid Status")
            else:
                pass
        else:
            pass

        if str(Element[4]) == 'None' and Element[2] == 'Critical':
            Critical = Critical+1
        elif str(Element[4]) == 'None' and Element[2] == 'Blocker':
            Blocker = Blocker+1
        elif str(Element[4]) == 'None' and Element[2] == 'Major':
            Major = Major+1
        elif str(Element[4]) == 'None' and Element[2] == 'Minor':
            Minor = Minor+1
        else:
            pass



    #Grand Total Calculation

    ETC = A1ETC+A2ETC+A3ETC+A4ETC
    ETB = A1ETB+A2ETB+A3ETB+A4ETB
    ETMA = A1ETMA+A2ETMA+A3ETMA+A4ETMA
    ETMI = A1ETMI+A2ETMI+A3ETMI+A4ETMI
    ET = ETC+ETB+ETMA+ETMI

    GTC = A1C+A2C+A3C+A4C+ETC
    GTB = A1B+A2B+A3B+A4B+ETB
    GTMA = A1MA+A2MA+A3MA+A4MA+ETMA
    GTMI = A1MI+A2MI+A3MI+A4MI+ETMI
    GT = GTC+GTB+GTMA+GTMI

#Store all variables in an Array
    print('storing variable in array')

    if proj == 'project=PROJ1':
        Array1.append(Critical)
        Array1.append(Critical1)
        Array1.append(Blocker)
        Array1.append(Blocker1)
        Array1.append(Major)
        Array1.append(Major1)
        Array1.append(Minor)
        Array1.append(Minor1)
        Array1.append(A1TOT)
        Array1.append(A1C)
        Array1.append(A1B)
        Array1.append(A1MA)
        Array1.append(A1MI)
        Array1.append(A2TOT)
        Array1.append(A2C)
        Array1.append(A2B)
        Array1.append(A2MA)
        Array1.append(A2MI)
        Array1.append(A3TOT)
        Array1.append(A3C)
        Array1.append(A3B)
        Array1.append(A3MA)
        Array1.append(A3MI)
        Array1.append(A4TOT)
        Array1.append(A4C)
        Array1.append(A4B)
        Array1.append(A4MA)
        Array1.append(A4MI)
        Array1.append(ET)
        Array1.append(ETC)
        Array1.append(ETB)
        Array1.append(ETMA)
        Array1.append(ETMI)
        Array1.append(GT)
        Array1.append(GTC)
        Array1.append(GTB)
        Array1.append(GTMA)
        Array1.append(GTMI)
    elif proj == 'project=PROJ2':
        Array2.append(Critical)
        Array2.append(Critical1)
        Array2.append(Blocker)
        Array2.append(Blocker1)
        Array2.append(Major)
        Array2.append(Major1)
        Array2.append(Minor)
        Array2.append(Minor1)
        Array2.append(A1TOT)
        Array2.append(A1C)
        Array2.append(A1B)
        Array2.append(A1MA)
        Array2.append(A1MI)
        Array2.append(A2TOT)
        Array2.append(A2C)
        Array2.append(A2B)
        Array2.append(A2MA)
        Array2.append(A2MI)
        Array2.append(A3TOT)
        Array2.append(A3C)
        Array2.append(A3B)
        Array2.append(A3MA)
        Array2.append(A3MI)
        Array2.append(A4TOT)
        Array2.append(A4C)
        Array2.append(A4B)
        Array2.append(A4MA)
        Array2.append(A4MI)
        Array2.append(ET)
        Array2.append(ETC)
        Array2.append(ETB)
        Array2.append(ETMA)
        Array2.append(ETMI)
        Array2.append(GT)
        Array2.append(GTC)
        Array2.append(GTB)
        Array2.append(GTMA)
        Array2.append(GTMI)
    elif proj == 'project=PROJ3':
        Array3.append(Critical)
        Array3.append(Critical1)
        Array3.append(Blocker)
        Array3.append(Blocker1)
        Array3.append(Major)
        Array3.append(Major1)
        Array3.append(Minor)
        Array3.append(Minor1)
        Array3.append(A1TOT)
        Array3.append(A1C)
        Array3.append(A1B)
        Array3.append(A1MA)
        Array3.append(A1MI)
        Array3.append(A2TOT)
        Array3.append(A2C)
        Array3.append(A2B)
        Array3.append(A2MA)
        Array3.append(A2MI)
        Array3.append(A3TOT)
        Array3.append(A3C)
        Array3.append(A3B)
        Array3.append(A3MA)
        Array3.append(A3MI)
        Array3.append(A4TOT)
        Array3.append(A4C)
        Array3.append(A4B)
        Array3.append(A4MA)
        Array3.append(A4MI)
        Array3.append(ET)
        Array3.append(ETC)
        Array3.append(ETB)
        Array3.append(ETMA)
        Array3.append(ETMI)
        Array3.append(GT)
        Array3.append(GTC)
        Array3.append(GTB)
        Array3.append(GTMA)
        Array3.append(GTMI)
    elif proj == 'project=PROJ4':
        Array4.append(Critical)
        Array4.append(Critical1)
        Array4.append(Blocker)
        Array4.append(Blocker1)
        Array4.append(Major)
        Array4.append(Major1)
        Array4.append(Minor)
        Array4.append(Minor1)
        Array4.append(A1TOT)
        Array4.append(A1C)
        Array4.append(A1B)
        Array4.append(A1MA)
        Array4.append(A1MI)
        Array4.append(A2TOT)
        Array4.append(A2C)
        Array4.append(A2B)
        Array4.append(A2MA)
        Array4.append(A2MI)
        Array4.append(A3TOT)
        Array4.append(A3C)
        Array4.append(A3B)
        Array4.append(A3MA)
        Array4.append(A3MI)
        Array4.append(A4TOT)
        Array4.append(A4C)
        Array4.append(A4B)
        Array4.append(A4MA)
        Array4.append(A4MI)
        Array4.append(ET)
        Array4.append(ETC)
        Array4.append(ETB)
        Array4.append(ETMA)
        Array4.append(ETMI)
        Array4.append(GT)
        Array4.append(GTC)
        Array4.append(GTB)
        Array4.append(GTMA)
        Array4.append(GTMI)
    else:
        print("Unidentified Project. Please Check")

# Store in Grand Total Array

HGTArray = []
HGTA1 = Array1[8]+Array3[8]+Array4[8]+Array2[8]
HGTA2 = Array1[13]+Array3[13]+Array4[13]+Array2[13]
HGTA3 = Array1[18]+Array3[18]+Array4[18]+Array2[18]
HGTA4 = Array1[23]+Array3[23]+Array4[23]+Array2[23]
HGTET = Array1[28]+Array3[28]+Array4[28]+Array2[28]
HGT = Array1[33]+Array3[33]+Array4[33]+Array2[33]
HGTArray.append(HGT)
HGTArray.append(HGTA1)
HGTArray.append(HGTA2)
HGTArray.append(HGTA3)
HGTArray.append(HGTA4)
HGTArray.append(HGTET)


#Excel Data-Population
print('generating excel')

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('JiraReport.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.

date=time.strftime("%m/%d/%Y")
cols = ["Team", "Priority", "No:of Tickets \nopen as on "+date, "No:of Tickets \nclosed this week"]

# Start from the first cell. Rows and columns are zero indexed.
row = 2
col = 0
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'yellow'})
worksheet.merge_range('A1:D1', 'Ticket status as on '+date,merge_format)
format = workbook.add_format()
format.set_bold()
format.set_text_wrap()
format.set_bg_color('cyan')
format.set_border()
worksheet.set_column(2, 3, 20)

#Normal Format

format3 = workbook.add_format()
format3.set_text_wrap()
format3.set_border()

#Grass Green Format

format2 = workbook.add_format()
format2.set_bold()
format2.set_border()
format2.set_bg_color('silver')


# Iterate over the data and write it out row by row.
for index, col in enumerate(cols):
    worksheet.write(row, index,col,format)

worksheet.set_column(0, 0, 20)

#Proj 1 Data
worksheet.merge_range('A4:A8', 'Team B - AMS',format2)
worksheet.write(3, 1,"Critical",format3)
worksheet.write(4, 1,"Blocker",format3)
worksheet.write(5, 1,"Major",format3)
worksheet.write(6, 1,"Minor",format3)
worksheet.write(7, 1,"Total",format2)
worksheet.write(3, 2,Array3[0],format3)
worksheet.write(3, 3,Array3[1],format3)
worksheet.write(4, 2,Array3[2],format3)
worksheet.write(4, 3,Array3[3],format3)
worksheet.write(5, 2,Array3[4],format3)
worksheet.write(5, 3,Array3[5],format3)
worksheet.write(6, 2,Array3[6],format3)
worksheet.write(6, 3,Array3[7],format3)
worksheet.write(7, 2,(Array3[0]+Array3[2]+Array3[4]+Array3[6]),format2)
worksheet.write(7, 3,(Array3[1]+Array3[3]+Array3[5]+Array3[7]),format2)

#Proj 2 Data

worksheet.merge_range('A9:A13', 'Team B - Inventory',format2)
worksheet.write(8, 1,"Critical",format3)
worksheet.write(9, 1,"Blocker",format3)
worksheet.write(10, 1,"Major",format3)
worksheet.write(11, 1,"Minor",format3)
worksheet.write(12, 1,"Total",format2)
worksheet.write(8, 2,Array1[0],format3)
worksheet.write(8, 3,Array1[1],format3)
worksheet.write(9, 2,Array1[2],format3)
worksheet.write(9, 3,Array1[3],format3)
worksheet.write(10, 2,Array1[4],format3)
worksheet.write(10, 3,Array1[5],format3)
worksheet.write(11, 2,Array1[6],format3)
worksheet.write(11, 3,Array1[7],format3)
worksheet.write(12, 2,(Array1[0]+Array1[2]+Array1[4]+Array1[6]),format2)
worksheet.write(12, 3,(Array1[1]+Array1[3]+Array1[5]+Array1[7]),format2)

#Proj 3 Data

worksheet.merge_range('A14:A18', 'Team B - Payment',format2)
worksheet.write(13, 1,"Critical",format3)
worksheet.write(14, 1,"Blocker",format3)
worksheet.write(15, 1,"Major",format3)
worksheet.write(16, 1,"Minor",format3)
worksheet.write(17, 1,"Total",format2)
worksheet.write(13, 2,Array4[0],format3)
worksheet.write(13, 3,Array4[1],format3)
worksheet.write(14, 2,Array4[2],format3)
worksheet.write(14, 3,Array4[3],format3)
worksheet.write(15, 2,Array4[4],format3)
worksheet.write(15, 3,Array4[5],format3)
worksheet.write(16, 2,Array4[6],format3)
worksheet.write(16, 3,Array4[7],format3)
worksheet.write(17, 2,(Array4[0]+Array4[2]+Array4[4]+Array4[6]),format2)
worksheet.write(17, 3,(Array4[1]+Array4[3]+Array4[5]+Array4[7]),format2)

#Proj 4 Data

worksheet.merge_range('A19:A23', 'Unified Notes',format2)
worksheet.write(18, 1,"Critical",format3)
worksheet.write(19, 1,"Blocker",format3)
worksheet.write(20, 1,"Major",format3)
worksheet.write(21, 1,"Minor",format3)
worksheet.write(22, 1,"Total",format2)
worksheet.write(18, 2,Array2[0],format3)
worksheet.write(18, 3,Array2[1],format3)
worksheet.write(19, 2,Array2[2],format3)
worksheet.write(19, 3,Array2[3],format3)
worksheet.write(20, 2,Array2[4],format3)
worksheet.write(20, 3,Array2[5],format3)
worksheet.write(21, 2,Array2[6],format3)
worksheet.write(21, 3,Array2[7],format3)
worksheet.write(22, 2,(Array2[0]+Array2[2]+Array2[4]+Array2[6]),format2)
worksheet.write(22, 3,(Array2[1]+Array2[3]+Array2[5]+Array2[7]),format2)

#Total
GrandOpen = (Array4[0]+Array4[2]+Array4[4]+Array4[6]+Array1[0]+Array1[2]+Array1[4]+Array1[6]+Array3[0]+Array3[2]+Array3[4]+Array3[6]+Array2[0]+Array2[2]+Array2[4]+Array2[6])
GrandClosed = (Array4[1]+Array4[3]+Array4[5]+Array4[7]+Array1[1]+Array1[3]+Array1[5]+Array1[7]+Array3[1]+Array3[3]+Array3[5]+Array3[7]+Array2[1]+Array2[3]+Array2[5]+Array2[7])
worksheet.write(23, 0,"Grand Total",format2)
worksheet.write(23, 1,"",format2)
worksheet.write(23, 2,GrandOpen,format2)
worksheet.write(23, 3,GrandClosed,format2)

#Aging Report

#Proj 1 Data

worksheet.write(8, 5,"Team B - AMS",format2)
worksheet.write(4, 5,"Critical",format3)
worksheet.write(5, 5,"Blocker",format3)
worksheet.write(6, 5,"Major",format3)
worksheet.write(7, 5,"Minor",format3)
worksheet.write(8, 6,Array3[8],format2)
worksheet.write(4, 6,Array3[9],format3)
worksheet.write(5, 6,Array3[10],format3)
worksheet.write(6, 6,Array3[11],format3)
worksheet.write(7, 6,Array3[12],format3)
worksheet.write(8, 7,Array3[13],format2)
worksheet.write(4, 7,Array3[14],format3)
worksheet.write(5, 7,Array3[15],format3)
worksheet.write(6, 7,Array3[16],format3)
worksheet.write(7, 7,Array3[17],format3)
worksheet.write(8, 8,Array3[18],format2)
worksheet.write(4, 8,Array3[19],format3)
worksheet.write(5, 8,Array3[20],format3)
worksheet.write(6, 8,Array3[21],format3)
worksheet.write(7, 8,Array3[22],format3)
worksheet.write(8, 9,Array3[23],format2)
worksheet.write(4, 9,Array3[24],format3)
worksheet.write(5, 9,Array3[25],format3)
worksheet.write(6, 9,Array3[26],format3)
worksheet.write(7, 9,Array3[27],format3)
worksheet.write(8, 10,Array3[28],format2)
worksheet.write(4, 10,Array3[29],format3)
worksheet.write(5, 10,Array3[30],format3)
worksheet.write(6, 10,Array3[31],format3)
worksheet.write(7, 10,Array3[32],format3)
worksheet.write(8, 11,Array3[33],format2)
worksheet.write(4, 11,Array3[34],format3)
worksheet.write(5, 11,Array3[35],format3)
worksheet.write(6, 11,Array3[36],format3)
worksheet.write(7, 11,Array3[37],format3)

#Proj 2 Data

worksheet.write(13, 5,"Team B - Inventory",format2)
worksheet.write(9, 5,"Critical",format3)
worksheet.write(10, 5,"Blocker",format3)
worksheet.write(11, 5,"Major",format3)
worksheet.write(12, 5,"Minor",format3)
worksheet.write(13, 6,Array1[8],format2)
worksheet.write(9, 6,Array1[9],format3)
worksheet.write(10, 6,Array1[10],format3)
worksheet.write(11, 6,Array1[11],format3)
worksheet.write(12, 6,Array1[12],format3)
worksheet.write(13, 7,Array1[13],format2)
worksheet.write(9, 7,Array1[14],format3)
worksheet.write(10, 7,Array1[15],format3)
worksheet.write(11, 7,Array1[16],format3)
worksheet.write(12, 7,Array1[17],format3)
worksheet.write(13, 8,Array1[18],format2)
worksheet.write(9, 8,Array1[19],format3)
worksheet.write(10, 8,Array1[20],format3)
worksheet.write(11, 8,Array1[21],format3)
worksheet.write(12, 8,Array1[22],format3)
worksheet.write(13, 9,Array1[23],format2)
worksheet.write(9, 9,Array1[24],format3)
worksheet.write(10, 9,Array1[25],format3)
worksheet.write(11, 9,Array1[26],format3)
worksheet.write(12, 9,Array1[27],format3)
worksheet.write(13, 10,Array1[28],format2)
worksheet.write(9, 10,Array1[29],format3)
worksheet.write(10, 10,Array1[30],format3)
worksheet.write(11, 10,Array1[31],format3)
worksheet.write(12, 10,Array1[32],format3)
worksheet.write(13, 11,Array1[33],format2)
worksheet.write(9, 11,Array1[34],format3)
worksheet.write(10, 11,Array1[35],format3)
worksheet.write(11, 11,Array1[36],format3)
worksheet.write(12, 11,Array1[37],format3)

#Proj 3 Data

worksheet.write(18, 5,"Team B - Payment",format2)
worksheet.write(14, 5,"Critical",format3)
worksheet.write(15, 5,"Blocker",format3)
worksheet.write(16, 5,"Major",format3)
worksheet.write(17, 5,"Minor",format3)
worksheet.write(18, 6,Array4[8],format2)
worksheet.write(14, 6,Array4[9],format3)
worksheet.write(15, 6,Array4[10],format3)
worksheet.write(16, 6,Array4[11],format3)
worksheet.write(17, 6,Array4[12],format3)
worksheet.write(18, 7,Array4[13],format2)
worksheet.write(14, 7,Array4[14],format3)
worksheet.write(15, 7,Array4[15],format3)
worksheet.write(16, 7,Array4[16],format3)
worksheet.write(17, 7,Array4[17],format3)
worksheet.write(18, 8,Array4[18],format2)
worksheet.write(14, 8,Array4[19],format3)
worksheet.write(15, 8,Array4[20],format3)
worksheet.write(16, 8,Array4[21],format3)
worksheet.write(17, 8,Array4[22],format3)
worksheet.write(18, 9,Array4[23],format2)
worksheet.write(14, 9,Array4[24],format3)
worksheet.write(15, 9,Array4[25],format3)
worksheet.write(16, 9,Array4[26],format3)
worksheet.write(17, 9,Array4[27],format3)
worksheet.write(18, 10,Array4[28],format2)
worksheet.write(14, 10,Array4[29],format3)
worksheet.write(15, 10,Array4[30],format3)
worksheet.write(16, 10,Array4[31],format3)
worksheet.write(17, 10,Array4[32],format3)
worksheet.write(18, 11,Array4[33],format2)
worksheet.write(14, 11,Array4[34],format3)
worksheet.write(15, 11,Array4[35],format3)
worksheet.write(16, 11,Array4[36],format3)
worksheet.write(17, 11,Array4[37],format3)

#Proj 4 Data

worksheet.write(23, 5,"Unified Notes",format2)
worksheet.write(19, 5,"Critical",format3)
worksheet.write(20, 5,"Blocker",format3)
worksheet.write(21, 5,"Major",format3)
worksheet.write(22, 5,"Minor",format3)
worksheet.write(23, 6,Array2[8],format2)
worksheet.write(19, 6,Array2[9],format3)
worksheet.write(20, 6,Array2[10],format3)
worksheet.write(21, 6,Array2[11],format3)
worksheet.write(22, 6,Array2[12],format3)
worksheet.write(23, 7,Array2[13],format2)
worksheet.write(19, 7,Array2[14],format3)
worksheet.write(20, 7,Array2[15],format3)
worksheet.write(21, 7,Array2[16],format3)
worksheet.write(22, 7,Array2[17],format3)
worksheet.write(23, 8,Array2[18],format2)
worksheet.write(19, 8,Array2[19],format3)
worksheet.write(20, 8,Array2[20],format3)
worksheet.write(21, 8,Array2[21],format3)
worksheet.write(22, 8,Array2[22],format3)
worksheet.write(23, 9,Array2[23],format2)
worksheet.write(19, 9,Array2[24],format3)
worksheet.write(20, 9,Array2[25],format3)
worksheet.write(21, 9,Array2[26],format3)
worksheet.write(22, 9,Array2[27],format3)
worksheet.write(23, 10,Array2[28],format2)
worksheet.write(19, 10,Array2[29],format3)
worksheet.write(20, 10,Array2[30],format3)
worksheet.write(21, 10,Array2[31],format3)
worksheet.write(22, 10,Array2[32],format3)
worksheet.write(23, 11,Array2[33],format2)
worksheet.write(19, 11,Array2[34],format3)
worksheet.write(20, 11,Array2[35],format3)
worksheet.write(21, 11,Array2[36],format3)
worksheet.write(22, 11,Array2[37],format3)

#Horizontal Grand Total

worksheet.write(24, 5,"Grand Total",format2)
worksheet.write(24, 6,HGTA1,format2)
worksheet.write(24, 7,HGTA2,format2)
worksheet.write(24, 8,HGTA3,format2)
worksheet.write(24, 9,HGTA4,format2)
worksheet.write(24, 10,HGTET,format2)
worksheet.write(24, 11,HGT,format2)

format1 = workbook.add_format()
format1.set_bold()
format1.set_text_wrap()
format1.set_border()
format1.set_bg_color('cyan')
worksheet.write('F3',"",format1)
worksheet.write('F4',"Team",format1)
worksheet.write('G4',"0-1 Day",format1)
worksheet.write('H4',"1-3 Days",format1)
worksheet.write('I4',"3-5 Days",format1)
worksheet.write('J4',">5 days",format1)
worksheet.merge_range('F1:L1', 'Aging Report as on '+date,merge_format)
worksheet.merge_range('G3:J3', 'Aging of tickets pending within team ',format1)
worksheet.set_column(10, 10, 20)
worksheet.merge_range('K3:K4', 'Pending with other \nteam/closure \nconfirmation',format1)
worksheet.write('L3',"",format1)
worksheet.write('L4',"Grand Total",format1)

#Close Workbook

workbook.close()

#Print Total Time Taken

print("\nProcess Completed in %s seconds" % (time.time() - startu_time))