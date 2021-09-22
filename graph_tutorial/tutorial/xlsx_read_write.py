import requests
import json

graph_url = 'https://graph.microsoft.com/v1.0'
# #### Layout.html
# <li class="nav-item" data-turbolinks="false">
# <a class="nav-link{% if request.resolver_match.view_name == 'xlsxread' %} active{% endif %}" href="{% url 'xlsxread' %}">Calendar</a>

#### 
def xlsx_read(request):
  context = initialize_context(request)
  user = context['user']

  print(user)

  token = get_token(request)

  filecode = "01UDQNFF77ECOFWLH43VG3MJIHT2T3GJRK"
  sheetname =  "Sheet1"
  startrange =  "a1"
  endrange = "d15"

  events = get_xlsx_contents(
    token,
    filecode,
    sheetname, 
    startrange, 
    endrange)

  if events:
    # Convert the ISO 8601 date times to a datetime object
    # This allows the Django template to format the value nicely
    # for event in events['value']:
    #   event['start']['dateTime'] = parser.parse(event['start']['dateTime'])
    #   event['end']['dateTime'] = parser.parse(event['end']['dateTime'])

    # context['events'] = events['value']

    print(events)

  return render(request, 'tutorial/calendar.html', context)
  
### Code in graph_helper.py or xlsx_read_write.py

def get_xlsx_contents(token, filecode,sheetname):
  # Set headers
  headers = {
    'Authorization': 'Bearer {0}'.format(token)
  }

  # Configure query parameters to
  # modify the results
  # query_params = {
  #     '$select': 'values'
  # }

  # Send GET to /me/events

  ## For the All records without range
  # events = requests.get("{0}/me/drive/items/{1}/workbook/worksheets('{2}')/UsedRange(valuesOnly=true)?$select=values".format(graph_url, filecode, sheetname),
  #                       headers=headers)

  ### For the All records with range
  # events = requests.get("{0}/me/drive/items/{1}/workbook/worksheets('{2}')/range(address='a1:v500')/UsedRange(valuesOnly=true)?$select=values".format(graph_url, filecode, sheetname),
  #                       headers=headers)

  ### For range and expected format

  events = requests.get("{0}/me/drive/items/{1}/workbook/worksheets('{2}')/range(address='a1:v500')/UsedRange(valuesOnly=true)?$format=atom,$select=values".format(graph_url, filecode, sheetname),
                        headers=headers)

  # # Return the JSON result
  # print("EVENT")
  # print(events.json())
  # return events.json()

## for front end 
  values = events.json()
  # values = values["values"]
  values = values["text"]
  print(type(values))
  print(values)

  import pandas as pd

  df = pd.DataFrame(values[1:], columns=values[0])

  print(type(df["Project Start Date"]))


  df.rename(columns = {'Project Name':'Project_Name',
                        'Team Members':'Team_Members',
                        'In Progress':'In_Progress',
                        'Development Completion date':'Development_Completion_date',
                        'System Testing Completion date':'System_Testing_Completion_date',
                        'UAT Completion date':'UAT_Completion_date',
                        'Production Deployment date':'Production_Deployment_date',
                        'Key Risk':'Key_Risk',
                        'Project completion in percentage':'Project_completion_in_percentage',
                        'Project Start Date':'Project_Start_Date',
                        'Project End date':'Project_End_Date',
                        'Actual Cost':'Actual_Cost',
                        'Planned Cost':'Planned_Cost',
                        'Estimated Hours':'Estimated_Hours',
                        'Hours Spent':'Hours_Spent',
                        'PMS Report':'PMS_Report'
                        }, inplace = True)
  # df = pd.DataFrame()
  # print(df)

  df.replace({'\n': '<br>'}, regex=True, inplace= True)

  # return df.reset_index().to_json(orient = 'records')
  return df.reset_index().to_dict('records')