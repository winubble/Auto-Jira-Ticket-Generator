import pandas as pd
import numpy as np

# JIRA libraries
from jira.resources import IssueLink
import jira.client
from jira.client import JIRA
from jira import JIRA
import re
"""
Assume this case the work ticket columns are all empty
"""

# import the deploy excel file into a dataframe
df_deploy = pd.read_excel('testoutput.xlsx', sheet_name = 0)

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

def create_ticket(row, currEN):
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
        jira.create_issue_link("Related", new_issue.key, currEN, None)

    return new_issue.key


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# initialize the current parent EN number
currEN = ''

# assign the first value of the port column to currEN
# for comparing if it is a parent or a child
currPort = df_deploy.iloc[0]['Port']
meetInitialPort = False

for index, row in df_deploy.iterrows():
            
    if(row['Port'] == currPort and meetInitialPort == True):
        # This is a child row
        # check whether has the ticket already
        # if yes, update the ticket with linkedto, remain the currEN and currPort the same 
        # if no, fill the linkedto(both row and df) with currEN and create a new ticket and write into dataframe

            
        # fill the linkedTo with currEN
        df_deploy['Linked To'][index] = currEN

        row['Linked To'] = currEN
            
        # create a new jira ticket
        myKey = create_ticket(row, currEN)
            
        # write the new ticket URL into dataframe
        df_deploy['Work Ticket'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
        # This IS WRONG: row['Work Ticket'] = "http://142.104.193.65:8080/browse/%s" % myKey. because it won't change the df
            
        # do not need to update the EN number or currPort because this is a child ticket
            
        
    # This is row one
    if(row['Port'] == currPort and meetInitialPort == False):
        # Mark as already seen the first row
        meetInitialPort = True
        
        # Check whether it has the ticket already
        # if yes, record the EN number as currEN
        # if not, create the ticket and record the EN number as currEN
        if(isinstance(row['Work Ticket'], str)):
            # Then it has the ticket
            # Use regular expression to record the EN number
            findEN = 'EN'
            currEN = row['Work Ticket'][row['Work Ticket'].find(findEN):]
        else:
            # Then it has no ticket
            # create a new ticket, write into the dataframe and record the EN number
            
            myKey = create_ticket(row, currEN)
            
            # write the new ticket URL into dataframe
            df_deploy['Work Ticket'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
            # THIS IS WRONG: row['Work Ticket'] = "http://142.104.193.65:8080/browse/%s" % myKey
            
            # record EN number
            currEN = myKey

            # no need to record the current port since this is the first and I already record it

    if(row['Port'] != currPort):
        #print("This is row[port]:" + row['Port'])
        #print("this is currPort"+ currPort)
        
        # this means this row is a new parent row
        # check whether it has a ticket already
        # if yes, then leave it alone, and update the currEN, update the currPort
        # if no, then create a new ticket, write the URL to the dataframe, update the currEN, update the currPort
        if(isinstance(row['Work Ticket'], str)):
            # Then it has a ticket
            # Use regular expression to record the EN number
            findEN = 'EN'
            currEN = row['Work Ticket'][row['Work Ticket'].find(findEN):]
            currPort = row['Port']   
        else:
            # Then it does not have a ticket
            # create a ticket, write the URL to dataframe and update the currEN, update the currPort
            myKey = create_ticket(row, currEN)
            
            # write the new ticket URL into dataframe
            df_deploy['Work Ticket'][index] = "http://142.104.193.65:8080/browse/%s" % myKey
       
            # update EN number
            currEN = myKey

            # record the current port
            currPort = row['Port']        


print(df_deploy)
#df_deploy.to_excel(r'testSimple.xlsx', index = False)