#!/usr/bin/python3

import requests
import json
import sys
import datetime
import xlsxwriter
import argparse



###
# LOGIN
###

def login(api_key, api_secret):
  url = "https://api.timeular.com/api/v3/developer/sign-in"

  payload = json.dumps({
    "apiKey": api_key,
    "apiSecret": api_secret
  })
  headers = {
    'Content-Type': 'application/json'
  }

  response = requests.request("POST", url, headers=headers, data=payload)

  if response.status_code != 200:
    print("[-] Login unsuccessful.")
    print(" ---- ")
    print(response.text)
    print(" ---- ")
    sys.exit(1)
  else:
    print("[+] Login successful.")
    session = json.loads(response.text)["token"]

  return session



###
# GET ACTIVITIES
###

def get_activities(session):
  url = "https://api.timeular.com/api/v3/activities"

  payload = {}
  headers = {
    'Authorization': 'Bearer ' + str(session)
  }

  response = requests.request("GET", url, headers=headers, data=payload)

  activities = json.loads(response.text)

  return activities



###
# TRANSLATE ACTIVITY
###

def translate_activity(activities, activityname=False, activityid=False):

  if activityname == False and activityid == False:
    print("[-] Internal error while translating activity.")
    sys.exit(1)

  if activityid:
    for a in activities['activities']:
      if a['id'] == activityid:
        return a['name']

  if activityname:
    for a in activities['activities']:
      if a['name'] == activityname:
        return a['id']

  print('[-] Activity could not be found.')
  print('    Following activities are available:')
  for a in activities['activities']:
    print('    - ' + str(a['name']))

  sys.exit(1)
  return False



###
# GET ENTRIES
###

def get_entries(session, start_time, end_time):
  url = "https://api.timeular.com/api/v3/time-entries/" + str(start_time) + "/" + str(end_time)

  payload = {}
  headers = {
    'Authorization': 'Bearer ' + str(session)
  }

  response = requests.request("GET", url, headers=headers, data=payload)

  entries = json.loads(response.text)

  if len(entries) == 0:
    print("[-] No entries found for selected dates.")
    sys.exit(1)

  return entries



###
# FILTER ENTRIES
###

def parse_entries(entries, activityid=False):
  entries_in_scope = []

  if activityid:
    for te in entries["timeEntries"]:
      if te["activityId"] == activityid:
        entries_in_scope.append(te)
      else:
        continue
  else:
    for te in entries["timeEntries"]:
      entries_in_scope.append(te)

  if len(entries_in_scope) == 0:
    print("[-] No entries left for selected project and dates.")
    sys.exit(1)

  return entries_in_scope



###
# CALCULATE FIRST AND LAST DAY OF MONTH
###

def calculate_first_and_last_day(month):

  result = {
    "start_time": "",
    "end_time": ""
  }

  today = datetime.datetime.now()

  if month == "current":
    result["start_time"] = today.replace(day=1).strftime('%Y-%m-%dT00:00:00.000')

    next_month = today.replace(day=28) + datetime.timedelta(days=4)
    result["end_time"] = (next_month - datetime.timedelta(days=next_month.day)).strftime('%Y-%m-%dT23:59:59.999')

    return result

  elif month == "last":

    last_day_of_prev_month = today.replace(day=1) - datetime.timedelta(days=1)
    start_day_of_prev_month = today.replace(day=1) - datetime.timedelta(days=last_day_of_prev_month.day)

    result["start_time"] = start_day_of_prev_month.strftime('%Y-%m-%dT00:00:00.000')    
    result["end_time"] = last_day_of_prev_month.strftime("%Y-%m-%dT23:59:59.999")

    return result



###
# EXPORT TO CSV
###

def export_times(entries, filename, activities):

  workbook = xlsxwriter.Workbook(filename)
  worksheet = workbook.add_worksheet()

  entries_to_export = [["Day", "Start", "End", "Duration", "Description", "Project"]]

  for e in entries:
    tmp_start_to = datetime.datetime.strptime(e["duration"]["startedAt"][:-4], '%Y-%m-%dT%H:%M:%S')
    tmp_end_to = datetime.datetime.strptime(e["duration"]["stoppedAt"][:-4], '%Y-%m-%dT%H:%M:%S')

    if tmp_start_to.strftime("%m/%d") != tmp_end_to.strftime("%m/%d"):
      print("[!] Start and endtime on different days. Double-check the entry!")
    
    tmp_day = tmp_start_to.strftime("%m/%d")
    tmp_start = tmp_start_to.strftime("%H:%M")
    tmp_end = tmp_end_to.strftime("%H:%M")
    tmp_duration = str(tmp_end_to - tmp_start_to)
    tmp_description = e["note"]["text"]
    tmp_project = translate_activity(activities, activityid=e["activityId"])
    entries_to_export.append([tmp_day, tmp_start, tmp_end, tmp_duration, tmp_description, tmp_project])

  row = 0

  start_end_format = workbook.add_format({'num_format': 'hh:mm'})
  duration_format = workbook.add_format({'num_format': 'hh:mm:ss'})

  for r in entries_to_export:
     worksheet.write(row, 0, r[0])
     worksheet.write(row, 1, r[1], start_end_format)
     worksheet.write(row, 2, r[2], start_end_format)
     worksheet.write(row, 3, r[3], duration_format)
     worksheet.write(row, 4, r[4])
     worksheet.write(row, 5, r[5])
     row += 1

  worksheet.write(row, 2, "Sum:")
  worksheet.write(row, 3, "=SUM(D2:D" + str(row) + ")", duration_format)
  worksheet.write(row+1, 2, "Sum Rounded:")
  worksheet.write(row+1, 3, "=ROUND(D" + str(row+1) + "*(24*60/15),0)/(24*60/15)", duration_format)

  workbook.close()



if __name__ == '__main__':

  parser = argparse.ArgumentParser(description="Timeular export...")

  parser.add_argument('-p', '--project', help="Define the project name, as it is in timeular. Use \"all\" to extract all projects.", required=True)
  parser.add_argument('-lm', '--lastmonth', help="Extract all times from last month.", default=False, action="store_true")
  parser.add_argument('-cm', '--currentmonth', help="Extract all times from current month.", default=False, action="store_true")
  parser.add_argument('-sd', '--startday', help="First day of the period to capture.")
  parser.add_argument('-ld', '--lastday', help="Last day of the period to capture.")

  args = parser.parse_args()

  activityname = args.project

  if args.lastmonth == True:
    period = calculate_first_and_last_day("last")
    start_time = period["start_time"]
    end_time = period["end_time"]
  elif args.currentmonth == True:
    period = calculate_first_and_last_day("current")
    start_time = period["start_time"]
    end_time = period["end_time"]
  else:
    if args.startday == False and args.lastday == False:
      print("[-] No dates given.")
      sys.exit(1)
    else:
      try:
        start_time = datetime.datetime.strptime(args.startday, '%Y-%m-%d').strftime('%Y-%m-%dT00:00:00.000')
        end_time = datetime.datetime.strptime(args.lastday, '%Y-%m-%d').strftime("%Y-%m-%dT23:59:59.999")
      except ValueError:
        print("[-] Date given in wrong format - required: %Y-%m-%d")
        sys.exit(1)

  try:
    with open("api-key.json", "r") as f:
      apicreds = json.load(f)
  except:
    print("[-] API credentials not found.")
    sys.exit(1)

  session = login(apicreds['apiKey'], apicreds['apiSecret'])

  filename = datetime.datetime.now().strftime('%Y%m%d') + "_" + str(activityname) + '_' + start_time[:10] + "_to_" + end_time[:10] + '_TimeExport.xlsx'

  activities = get_activities(session)
  entries = get_entries(session, start_time, end_time)

  if activityname != "all":
    activityid = translate_activity(activities, activityname=activityname)
    entries = parse_entries(entries, activityid)
  else:
    entries = parse_entries(entries)
  

  export_times(entries, filename, activities)
  
  
