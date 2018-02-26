from jira.client import JIRA
import cgi
import xlrd
import sys
from xlrd import open_workbook

wb = open_workbook ("JIRA Issue Creation Tracker.xlsx")
sheet = wb.sheet_by_index(0)
jira_user = 'aravind.v'
jira_password = 'Tel@123'
jira_server = 'https://tejira.tataelxsi.co.in'
jira_project_key = 'TES'


options = {
    'server': jira_server
}

   
jira = JIRA(options, basic_auth=(jira_user, jira_password))
print "Started Creating User Stories in JIRA"
#for rowItrator in range(1, sheet.nrows):
for rowItrator in range(1, 3):
  print rowItrator 
  jira_assignee = sheet.cell(rowItrator, 3).value
  jira_label = sheet.cell(rowItrator, 2).value
  jira_type = sheet.cell(rowItrator, 5).value

  story_dct = {
    'project' : { 'key': jira_project_key },
    'summary' : sheet.cell(rowItrator, 1).value,
    'issuetype' : { 'name' : jira_type },
    'description' : sheet.cell(rowItrator, 4).value, 
    'assignee' : { "name" : jira_assignee },
    'labels' : [jira_label,' '],

    }
  
  child = jira.create_issue(fields=story_dct)
  print("created child: " + child.key)
  #child.update(reporter={'name': 'Aravind V'})
 
  #jira.add_issues_to_epic(sheet.cell(rowItrator, 1).value, [child.key], ignore_epics=True)
  print sheet.cell(rowItrator, 2).value
  #jira.add_issues_to_epic(str(sheet.cell(rowItrator, 1).value), [str(child.key)], ignore_epics=True)
  print "===================================="
