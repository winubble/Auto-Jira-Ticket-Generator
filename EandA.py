import pandas as pd
import numpy as np

# JIRA libraries
from jira.resources import IssueLink
import jira.client
from jira.client import JIRA
from jira import JIRA
import re

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Here is extraction part
# which extract the deploy rows from the oringial workshhets
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# import the excel file into a dataframe
df_whole = pd.read_excel('Cruise Configuration - Nautilus - test 2020.xlsx', sheet_name = None)

# count how many sheets in the whole excel sheet
sheetCnt = len(df_whole.keys())

# create a list and initialize the list item in the list
# for storing the subsheets
df_sheet = []
for i in range(0,sheetCnt):
    df_sheet.append("df_{}".format(i))

# create a list and append in the node name for filling use
node_fill = []
for sheet in df_whole.keys():
    node_fill.append(sheet)

# open the sheet seperately and store them seperately in the created list
for i in range(0,sheetCnt):
    df_sheet[i] = pd.read_excel('Cruise Configuration - Nautilus - test 2020.xlsx', sheet_name = i)

# create a dictionary for storing the site ticket EN number
site_dict = {}
for site in node_fill:    
    site_dict[site] = ''

# set a variable for storing the current position
pos = 0

# create a dataframe for storing the deploy row
df_out = pd.DataFrame()

# loop through all sheets in the excel sheets
for sheet in df_sheet:
    # Drop unused columns
    # only keep column from column 1 to column 12
    # which is NODE to OPERATION
    sheet.drop(sheet.iloc[:, 13::], inplace = True, axis = 1)
    
    # grab the site EN number and store to dictionary
    site_dict[node_fill[pos]] = sheet.columns[0]
        
    # Assign the column names
    # since the first row initially is empty
    sheet.columns = ['Node', 'Junction Box', 'Port','Communications','Cable to Connector Panel', 'Cable', 'IP Address',
                 'Instrument Category', 'Instrument', 'Serial Number', 'Device ID', 'Work Ticket', 'Operation']
    
    # Drop the first row
    sheet = sheet.drop(index=0)
    
    # ~~~~~~~~~~~~~~~~~~~deal with the sub ticket
    # Set up the intial fill-in value
    currPort = ''
    currCom = ''
    currCtCP = ''
    currCEI = ''
    currIP = ''
    currIC = ''
    currIns = ''
    currSN = ''
    currDI = ''
    currWT = ''
    currOp = ''
    
    # check in the port column, whether already see the port in strin already
    # initially set to false
    meetString = False
    
    # iterate the rows in the subsheet
    # and record the values when encouter to the first valid record
    for index, row in sheet.iterrows():
        if(isinstance(row['Port'], str)):
            currPort = row['Port']
            currCom = row['Communications']
            currCtCP = row['Cable to Connector Panel']
            currCEI = row['Cable']
            currIP = row['IP Address']
            currIC = row['Instrument Category']
            currIns = row['Instrument']
            currSN = row['Serial Number']
            currDI = row['Device ID']
            currWT = row['Work Ticket']
            currOp = row['Operation']
            break
            
    # iterate rhw rows in the subsheet and do the subtickts fill        
    for index, row in sheet.iterrows():
        # if meet a empty port cell and already seen a string port before
        # check whether this row is empty or not
        # if is not, fill in with current fill-in values
        if(meetString == True and (not isinstance(row['Port'], str))):
            
            isCom = isinstance(row['Communications'], str)
            isCtCP = isinstance(row['Cable to Connector Panel'], str)
            isCEI = isinstance(row['Cable'], str)
            isIP = isinstance(row['IP Address'], str)
            isIC = isinstance(row['Instrument Category'], str)
            isIns = isinstance(row['Instrument'], str)
            isSN = isinstance(row['Serial Number'], str)
            isDI = isinstance(row['Device ID'], str)
            isWT = isinstance(row['Work Ticket'], str)
            isOp = isinstance(row['Operation'], str)
            
            # check whether the current is empty or not
            # in other words, the possibility of being a subtickets
            isNotEmpty = isCom or isCtCP or isCEI or isIP or isIC or isIns or isSN or isDI or isWT or isOp
               
#        isNotEmpty = row['Communications'] != np.nan or row['Cable to Connector Panel'] != np.nan 
#        or row['Cable EXT ID'] != np.nan or row['IP Address'] != np.nan or row['IP Address'] != np.nan 
#        or row['Instrument Category'] != np.nan or row['Instrument'] != np.nan or row['Serial Number'] != np.nan 
#        or row['Device ID'] != np.nan or row['Work Ticket'] != np.nan or row['Operation'] != np.nan

            # if is not empty -> is a subtickets -> fill the empty cells using current fill-in values
            # note: the fill-in values is inherited from the last port-non-empty row
            if(isNotEmpty):
            
                # fill out the empty cells of children tickets
                row['Port'] = currPort
                        
                if(not isCom):
                    row['Communications'] = currCom
                
                #if row['Communications'] == np.nan:
                 #   row['Communications'] = currCom
                
                if (not isCtCP):
                    row['Cable to Connector Panel'] = currCtCP
                
                if (not isCEI):
                    row['Cable'] = currCEI
                
                if (not isIP):
                    row['IP Address'] = currIP
                
                if (not isIC):
                    row['Instrument Category'] = currIC
                
                if (not isIns):
                    row['Instrument'] = currIns
                
                if (not isSN):
                    row['Serial Number'] = currSN
                
                if (not isDI):
                    row['Device ID'] = currDI
                    
                # do not inherit the work ticket value
                #if (not isWT):
                    #row['Work Ticket'] = currWT
                
                if (not isOp):
                    row['Operation'] = currOp
                    
        # if the current row is a new ticket, assign the current row values to the fill-in values
        if(isinstance(row['Port'], str) and row['Port'] != currPort):
            meetString = True
            currPort = row['Port']
            currCom = row['Communications']
            currCtCP = row['Cable to Connector Panel']
            currCEI = row['Cable']
            currIP = row['IP Address']
            currIC = row['Instrument Category']
            currIns = row['Instrument']
            currSN = row['Serial Number']
            currDI = row['Device ID']
            currWT = row['Work Ticket']
            currOp = row['Operation']
                   
    junction_box_fill = ""
    # find the juction box fill
    for item in sheet['Junction Box']:
        thisType = isinstance(item, str) 
        if thisType:
            junction_box_fill = item
        
    # fill the Node column and Junction Box column            
    sheet['Node'] = sheet['Node'].fillna(node_fill[pos])
    sheet['Junction Box'] = sheet['Junction Box'].fillna(junction_box_fill)
    
    # Extract the deploy row from the current sheet and append the row into the output dataframe
    for index, row in sheet.iterrows():
        if(row["Operation"] == "Deploy"):
            df_out = df_out.append(row, ignore_index = True)[list(sheet)]
    
    #jump to the next sheet
    pos += 1

# insert the column needed in the jira import tool
df_out.insert(11, "Component", "Test and Development")
df_out.insert(13, "Linked To", np.nan)

# print(df_out)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Here is automation part
# which generate the tickets automatically
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# suppress the warning
pd.set_option('mode.chained_assignment', None)

# initialzie the variable
__instrument = ''
__serialNumber = ''
__deviceID = ''

# Descriptions
__location = ''
__communications = ''
__cableID = ''
__instrumentCategory = ''
__ipAddress = ''
__junctionBox = ''
__port = ''


__components = ''
__outwardIssue = ''
__summaryTitle = 'Instrument Qualification'

def create_ticket(row):
    # Assign the values
    __instrument = row['Instrument']
    __serialNumber = row['Serial Number']
    __deviceID = row['Device ID']

    __location = row['Node']
    __communications = row['Communications']
    __cableID = row['Cable']
    __instrumentCategory = row['Instrument Category']
    __ipAddress = row['IP Address']

    __components = row['Component']
    __junctionBox = row['Junction Box']
    __port = row['Port']
    __outwardIssue = row['Linked To'] 


    # Connect to jira
    # Authentication done by using username and password
    username = 'mtcelec2'
    password = '1q2w3e4R!'

    jira = JIRA(
        basic_auth = (username, password),
        options = {'server': 'http://142.104.193.65:8080'}
        #options = {'server': 'https://jira.oceannetworks.ca/'}
    )

    #create a ticket for current row
    new_issue = jira.create_issue(
        project = {'key': 'EN'}, 
        summary = "%s: %s SI: %s DI: %s" % (__summaryTitle, __instrument, __serialNumber, __deviceID),
        description = "Site Location: %s\n Communications: %s\n Cable EXT ID: %s\n Instrument Category: %s\n IP Address: %s\n Junction Box: %s\n Port: %s\n" % (__location, __communications, __cableID, __instrumentCategory, __ipAddress, __junctionBox, __port), 
        issuetype = {'name': 'Task'}, 
        components = [{'name' : __components}],
        customfield_10794 = {'id': "10453"},            # Bill of work to Customers (Default: ONC Internal)
        customfield_10592 = "%s" % __serialNumber,      # Serial # field
        customfield_10070 = __deviceID)                 # Device ID field

    # add the linkedto
    if(isinstance(row['Linked To'], str)):
        jira.create_issue_link("Related", new_issue.key, __outwardIssue, None)

    return new_issue.key

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# initialize the current parent EN number
currEN = ''

# assign the first value of the port column to currEN
# for comparing if it is a parent or a child
currPort = df_out.iloc[0]['Port']
meetInitialPort = False

for index, row in df_out.iterrows():
            
    if(row['Port'] == currPort and meetInitialPort == True):
        # This is a child row
        # check whether has the ticket already
        # if yes, update the ticket with linkedto, remain the currEN and currPort the same 
        # if no, fill the linkedto(both row and df) with currEN and create a new ticket and write into dataframe
  
        # fill the linkedTo with currEN
        df_out['Linked To'][index] = currEN
        row['Linked To'] = currEN
            
        # create a new jira ticket
        myKey = create_ticket(row)
            
        # write the new ticket URL into dataframe
        df_out['Work Ticket'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
        # This IS WRONG: row['Work Ticket'] = "http://142.104.193.65:8080/browse/%s" % myKey. because it won't change the df
            
        # do not need to update the EN number or currPort because this is a child ticket
            
        
    # This is row one
    if(row['Port'] == currPort and meetInitialPort == False):
        # Mark as already seen the first row
        meetInitialPort = True
        
        # link to the site, create a new ticket, write into the dataframe and record the EN number

        # loop through the dictionary, and write the site EN number to both row and df
        for mysite, siteEN in site_dict.items(): 
            if (mysite == row['Node']):
                df_out['Linked To'][index] = siteEN
                row['Linked To'] = siteEN
                 
        myKey = create_ticket(row)
            
        # write the new ticket URL into dataframe
        df_out['Work Ticket'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
        # THIS IS WRONG: row['Work Ticket'] = "http://142.104.193.65:8080/browse/%s" % myKey
            
        # record EN number
        currEN = myKey

        # no need to record the current port since this is the first and I already record it

    if(row['Port'] != currPort):
        # this means this row is a new parent row
        # link to the site, create a new ticket, write the URL to the dataframe, update the currEN, update the currPort
         
        # loop through the dictionary, and write the site EN number to both row and df
        for mysite, siteEN in site_dict.items(): 
            if (mysite == row['Node']):
                df_out['Linked To'][index] = siteEN
                row['Linked To'] = siteEN
                   
        # create a ticket, write the URL to dataframe and update the currEN, update the currPort
        myKey = create_ticket(row)
            
        # write the new ticket URL into dataframe
        df_out['Work Ticket'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
       
        # update EN number
        currEN = myKey

        # record the current port
        currPort = row['Port']        


print(df_out)
#df_out.to_excel(r'testSimple.xlsx', index = False)