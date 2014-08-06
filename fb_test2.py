# -*- coding: utf-8 -*-
"""
Created on Thu Feb 06 10:15:37 2014

@author: ayip
"""

import facebook
from datetime import date, timedelta

access_token = 'CAACEdEose0cBAHSLUMmc4Ud49vZBZAti5ez4Oji3FjBYMWrlkj2ZCauyb4ZAVIGKZBtmBXnCjdWMZBD6oCM9OckywnvWnRLCDHyx5uLv0KIaaKeIcpZBSsrfH3RF1gdjDeOjiO6eTVvFTcVQJE7JAtuZBkmFKgtu6ZA9ZBN1tykLz6wggUr26mTlbXsGDQv4E5s5QZD'
graph = facebook.GraphAPI(access_token)
today = date(2013,2,2)
source = []

while today > date(2013,1,1):
    fql_line = 'SELECT metric, value FROM insights WHERE object_id="165716504171" and period="86400" and end_time=end_time_date("'+today.isoformat()+'") and metric in ("page_fans_by_like_source")'
    result = graph.fql(fql_line)
    if result != []:
        for key in result[0]['value'].keys():
            if key not in source:
                source.append(key)
                print key
    today = today-timedelta(1)