# -*- coding: utf-8 -*-
"""
Created on Wed Mar 05 10:40:16 2014

@author: ayip
"""

import facebook
import pyodbc

user = 'yip'
pw = 'Password5'
#conn_str = 'DRIVER={SQL Server};SERVER=madb;DATABASE=AnalyticsTestDB;UID='+user+';PWD='+pw
conn_str = 'DRIVER={SQL Server};SERVER=madb;DATABASE=AnalyticsTestDB;Trusted_Connection=yes'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

import win32com.client
w = win32com.client.Dispatch('imacros')

print 'Trying to get access_token'
form_id = ["u_0_e","u_0_a","u_0_b","u_0_c","u_0_d"]
n = 0
i = 0
w.iimInit('-fx',1)
while n < 1:
    code = "URL GOTO=https://developers.facebook.com/tools/explorer/145634995501895/?method=GET&path=me%3Ffields%3Did%2Cname&version="+"\n"
    code += "WAIT SECONDS=5"+'\n'
    code += "TAG POS=1 TYPE=A ATTR=ID:get_access_token"+'\n'
    code += "WAIT SECONDS=5"+'\n'
    code += "TAG POS=1 TYPE=BUTTON ATTR=TXT:Get<SP>Access<SP>Token"+'\n'
    code += "WAIT SECONDS=5"+'\n'
    code += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:"+form_id[i]+" ATTR=NAME:access_token EXTRACT=TXT"
    n = w.iimPlayCode(code)
    i+=1
access_token = w.iimGetLastExtract()
w.iimClose()

print 'Got a token!'
#access_token = 'CAACEdEose0cBAA1GjCT0q5gvNBA41S3BeAoUefQVU9xYPhdvmHlhsla6deSRC6OZBUJOyAu17O2j85ZAgQWaIWNTNPe5ocZCVZArJw2F8Kom1qjTZAvwTdlmFASQ7YZBL1IaGHfufUyu4dKJqNuEdBeq8EEK5U4fjT4zZChrRcRkzO4HbYbqCbq97YAovcZC8dwZD'

graph = facebook.GraphAPI(access_token)
FB_PAGE_DB = 'USCellular_FB_Page'

object_id = '165716504171'  # US Cellular
page_metrics = ['page_impressions','page_impressions_unique','page_impressions_paid',
          'page_impressions_paid_unique','page_impressions_organic_unique','page_impressions_by_age_gender_unique',
          'page_fans','page_fans_gender_age','page_fans_by_like_source','page_fan_adds','page_fan_removes',
          'page_storytellers','page_storytellers_by_age_gender','page_positive_feedback_by_type']

type_key = {11:'Group created',12:'Event created',46:'Status update',
            56:'Post on wall from another user',66:'Note created',
            80:'Link posted',128:'Video posted',247:'Photos posted',
            237:'App story',257:'Comment created',272:'App story',
            285:'Check in to a place',308:'Post in Group'}

period = ['86400','0']  # day, lifetime

#find days with missing demo fields
date_rows = cursor.execute('SELECT [date] FROM ['+FB_PAGE_DB+'] where [page_fans_gender_age_F.25-34] is null').fetchall()
dates = []
for d in date_rows:
    dates.append(d[0])

#update days with missing demos
for d in dates:
    for p in period:
        fql_line = 'SELECT metric, value FROM insights WHERE object_id="'+object_id+'"'
        fql_line += ' and period="'+p+'" and end_time=end_time_date("'+d+'")'
        fql_line += ' and metric in ("'+'","'.join(page_metrics)+'")'
        result = graph.fql(fql_line)
        for r in result:
            sql_line = 'UPDATE [dbo].['+FB_PAGE_DB+'] SET ['+r.values()[0]+']='
            if r.values()[0] not in ('page_positive_feedback_by_type','page_fans_by_like_source',
            'page_storytellers_by_age_gender','page_impressions_by_age_gender_unique','page_fans_gender_age'):
                sql_line = 'UPDATE [dbo].['+FB_PAGE_DB+'] SET ['+r.values()[0]+']='
                if isinstance(r.values()[1],str):
                    sql_line+="'"+r.values()[1]+"' WHERE [date]='"+d+"'"
                else:
                    sql_line+=str(r.values()[1])+" WHERE [date]='"+d+"'"
                cursor.execute(sql_line)
                cursor.commit()
            else:
                # need to iterate over feedback types
                key = r['metric']
                try:
                    for k in r['value'].keys():
                        sql_line = 'UPDATE [dbo].['+FB_PAGE_DB+'] SET ['+key+'_'+k+']='
                        if isinstance(r.values()[1],str):
                            sql_line+="'"+r['value'][k]+"' WHERE [date]='"+d+"'"
                        else:
                            sql_line+=str(r['value'][k])+" WHERE [date]='"+d+"'"
                        try:
                            print 'Metric: '+key+'_'+k
                            cursor.execute(sql_line)
                            cursor.commit()
                        except:
                            print 'New Metric: '+key+'_'+k
                            cursor.execute('ALTER TABLE [dbo].['+FB_PAGE_DB+'] ADD ['+key+'_'+k+'] int NULL')
                            cursor.execute('ALTER TABLE dbo.USCellular_FB_Page SET (LOCK_ESCALATION = TABLE)')
                            cursor.commit()
                            cursor.execute(sql_line)
                            cursor.commit()
                except:
                    print 'No values for '+key