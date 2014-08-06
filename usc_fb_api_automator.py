import facebook
import pyodbc
import codecs
import os
from datetime import date,timedelta,datetime

user = 'yip'
pw = 'Password5'
#conn_str = 'DRIVER={SQL Server};SERVER=madb;DATABASE=AnalyticsTestDB;UID='+user+';PWD='+pw
conn_str = 'DRIVER={SQL Server};SERVER=madb;DATABASE=AnalyticsTestDB;Trusted_Connection=yes'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

os.chdir('C:\\Users\\ayip\\Documents\\python scripts')
#app_id = '466514266809468'
#app_secret = '7a4e0158e9be931491723a7e3c2858e7'
#
#app_token = facebook.get_app_access_token(app_id,app_secret)
#print app_token
#
#r = requests.get('https://graph.facebook.com/oauth/access_token?grant_type=client_credentials&client_id='+app_id+'&client_secret='+app_secret)
#access_token = r.text.split('=')[1]
#print access_token
form_id = ["u_0_e","u_0_a","u_0_b","u_0_c","u_0_d"]
log = open('log.txt','w')
import win32com.client
w = win32com.client.Dispatch('imacros')
n = 0
i = 0
w.iimInit('-fx',1)
log.write('Trying to get access_token\n')
print 'Trying to get access_token'
while n < 1:
    code = "URL GOTO=https://developers.facebook.com/tools/explorer/145634995501895/?method=GET&path=me%3Ffields%3Did%2Cname&version="+"\n"
    code += "WAIT SECONDS=5"+'\n'
    code += "TAG POS=1 TYPE=A ATTR=ID:get_access_token"+'\n'
    code += "WAIT SECONDS=5"+'\n'
    code += "TAG POS=1 TYPE=BUTTON ATTR=TXT:Get<SP>Access<SP>Token"+'\n'
    code += "WAIT SECONDS=5"+'\n'
    code += "TAG POS=1 TYPE=INPUT:TEXT FORM=ID:"+form_id[i]+" ATTR=NAME:access_token EXTRACT=TXT"
    n = w.iimPlayCode(code)
    i += 1
access_token = w.iimGetLastExtract()
w.iimClose()

log.write('Got a token!\n')
print 'Got a token!'
#if we need to hard code a token
#access_token = 'CAACEdEose0cBAFy1JgBaYlyQrLBIBZAyZAdZBJupuTasg3TV86v5VLZCouN0lXZAZB4TiHku5r2yvUgbOCyVB7BAEm182UbE7NMQBWFMTTMpcWXVePlnpplJ9kZAPaqjYEn2n1y0ZAwK9djVOreOu1MT6LRSJcVFWbrc3xUCyiKFYV3WZCB5dRrq9DSq5UuVLMkcZD'
graph = facebook.GraphAPI(access_token)
FB_PAGE_DB = 'USCellular_FB_Page'
FB_POST_INFO_DB = 'USCellular_FB_Post_Info'
FB_POST_DATA_DB = 'USCellular_FB_Post_Data'
#object_id = '100164536718543'  # massmutual
object_id = '165716504171'  # US Cellular
page_metrics = ['page_impressions','page_impressions_unique','page_impressions_paid',
          'page_impressions_paid_unique','page_impressions_organic_unique','page_impressions_by_age_gender_unique',
          'page_fans','page_fans_gender_age','page_fans_by_like_source','page_fan_adds','page_fan_removes',
          'page_storytellers','page_storytellers_by_age_gender','page_positive_feedback_by_type']
post_metrics = ['post_impressions','post_impressions_unique','post_impressions_paid',
                'post_impressions_paid_unique','post_impressions_organic',
                'post_impressions_organic_unique']
type_key = {11:'Group created',12:'Event created',46:'Status update',
            56:'Post on wall from another user',66:'Note created',
            80:'Link posted',128:'Video posted',247:'Photos posted',
            237:'App story',257:'Comment created',272:'App story',
            285:'Check in to a place',308:'Post in Group'}

# get post ids currently in db
db_posts_row = cursor.execute('SELECT DISTINCT [created_time],[post_id] FROM ['+FB_POST_INFO_DB+'] ORDER BY [created_time]').fetchall()
db_posts = []
for d in db_posts_row:
    db_posts.append(d[1])
# get post ids from page
# save post ids to file for reference


if os.path.exists(object_id+'_posts.txt'):
    log.write('Post file detected: using it'+'\n')
    print 'Post file detected: using it'
    posts = []
    p_file = codecs.open(object_id+'_posts.txt',encoding='utf-8')
    for line in p_file:
        posts.append(line.replace('\r\n',''))
    p_file.close()

    p_file = codecs.open(object_id+'_posts.txt','a',encoding='utf-8')
    unix_time = datetime(1970,1,1)
    join_time = datetime(2012,12,31)
    current_time = datetime.today()-timedelta(1)
    while current_time > join_time:
        log.write( 'Working on post_ids: '+current_time.isoformat()+'\n')
        print 'Working on post_ids: '+current_time.isoformat()
        result = graph.fql('SELECT post_id FROM stream WHERE source_id='+object_id+
        ' and actor_id='+object_id+' and created_time>'+
        str(int((current_time-unix_time).total_seconds()))+' LIMIT 100')
        for r in result:
            if r.values()[0] not in posts:
                posts.append(r.values()[0])
                p_file.write(r.values()[0]+'\r\n')
                log.write( 'Writing post_id: '+r.values()[0]+'\n')
                print 'Writing post_id: '+r.values()[0]
            else:
                # if we hit repeating post_ids, break out of loop
                current_time = join_time
                break
        current_time = current_time - timedelta(1)
else:
    log.write( 'No post file detected: making a new one'+'\n')
    print 'No post file detected: making a new one'
    p_file = codecs.open(object_id+'_posts.txt','w',encoding='utf-8')
    posts = []
    unix_time = datetime(1970,1,1)
    join_time = datetime(2012,12,31)
    current_time = datetime.today()-timedelta(1)
    while current_time > join_time:
        log.write( 'Working on post_ids: '+current_time.isoformat()+'\n')
        print 'Working on post_ids: '+current_time.isoformat()
        result = graph.fql('SELECT post_id FROM stream WHERE source_id='+object_id+
        ' and actor_id='+object_id+' and created_time>'+
        str(int((current_time-unix_time).total_seconds()))+' LIMIT 100')
        for r in result:
            if r.values()[0] not in posts:
                posts.append(r.values()[0])
                p_file.write(r.values()[0]+'\r\n')
                log.write( 'Writing post_id: '+r.values()[0]+'\n')
                print 'Writing post_id: '+r.values()[0]
        current_time = current_time - timedelta(1)
p_file.close()

period = ['86400','0']  # day, lifetime
today = date.today()-timedelta(1)
end_date = date(2012,12,31)   # only change if we need to remake db
date_rows = cursor.execute('SELECT DISTINCT [date] FROM ['+FB_PAGE_DB+']').fetchall()
dates = []
for d in date_rows:
    dates.append(d[0])

# page metrics
while today > end_date:
    log.write( 'Working on FB Page: '+today.isoformat()+'\n')
    print 'Working on FB Page: '+today.isoformat()
    if today.isoformat() not in dates:
        cursor.execute("insert into [dbo].["+FB_PAGE_DB+"] ([date]) values ('"+today.isoformat()+"')")
        cursor.commit()
        for p in period:
            fql_line = 'SELECT metric, value FROM insights WHERE object_id="'+object_id+'"'
            fql_line += ' and period="'+p+'" and end_time=end_time_date("'+today.isoformat()+'")'
            fql_line += ' and metric in ("'+'","'.join(page_metrics)+'")'
            result = graph.fql(fql_line)
            for r in result:
                log.write( 'Metric: '+r.values()[0]+'\n')
                if r.values()[0] not in ('page_positive_feedback_by_type','page_fans_by_like_source',
                'page_storytellers_by_age_gender','page_impressions_by_age_gender_unique','page_fans_gender_age'):
                    sql_line = 'UPDATE [dbo].['+FB_PAGE_DB+'] SET ['+r.values()[0]+']='
                    if isinstance(r.values()[1],str):
                        sql_line+="'"+r.values()[1]+"' WHERE [date]='"+today.isoformat()+"'"
                    else:
                        sql_line+=str(r.values()[1])+" WHERE [date]='"+today.isoformat()+"'"
                    cursor.execute(sql_line)
                    cursor.commit()
                else:
                    # need to iterate over feedback types
                    key = r['metric']
                    try:
                        for k in r['value'].keys():
                            sql_line = 'UPDATE [dbo].['+FB_PAGE_DB+'] SET ['+key+'_'+k+']='
                            if isinstance(r.values()[1],str):
                                sql_line+="'"+r['value'][k]+"' WHERE [date]='"+today.isoformat()+"'"
                            else:
                                sql_line+=str(r['value'][k])+" WHERE [date]='"+today.isoformat()+"'"
                            try:
                                log.write( 'Metric: '+key+'_'+k+'\n')
                                cursor.execute(sql_line)
                                cursor.commit()
                            except:
                                log.write( 'New Metric: '+key+'_'+k+'\n')
                                cursor.execute('ALTER TABLE [dbo].['+FB_PAGE_DB+'] ADD ['+key+'_'+k+'] int NULL')
                                cursor.execute('ALTER TABLE dbo.USCellular_FB_Page SET (LOCK_ESCALATION = TABLE)')
                                cursor.commit()
                                cursor.execute(sql_line)
                                cursor.commit()
                    except:
                        log.write( 'No values for '+key+'\n')
    else:
        log.write( today.isoformat()+' already in database!'+'\n')
        print today.isoformat()+' already in database!'
    today = today - timedelta(1)

today = date.today()-timedelta(1)
# post metrics
for post_id in posts:
    # update post_info if new post
    # else, add new post data
    log.write( 'Working on FB Posts: '+post_id+' '+today.isoformat()+'\n')
    print 'Working on FB Posts: '+post_id+' '+today.isoformat()
    if post_id not in db_posts:
        log.write( post_id+' not in database: adding info'+'\'n')
        print post_id+' not in database: adding info'
        sql_line = 'INSERT INTO [dbo].['+FB_POST_INFO_DB+'] '
        sql_cols = ['post_id']
        sql_vals = [post_id]
        fql_line = 'SELECT message,created_time,promotion_status,type FROM stream'
        fql_line += " WHERE post_id='"+post_id+"'"
        result = graph.fql(fql_line)
        for key in result[0].keys():
            sql_cols.append(key)
            # match type code to text
            if key == "type":
                if result[0][key] == None:
                    sql_vals.append('None')
                else:
                    sql_vals.append(type_key[result[0][key]])
            # need to convert created_time into isoformat
            elif key == "created_time":
                sql_vals.append(datetime.fromtimestamp(result[0][key]).isoformat())
            else:
                try:
                    if result[0][key].values() == None:
                        sql_vals.append('None')
                    else:
                        sql_vals.append(result[0][key].values().replace("'","''"))
                except:
                    if result[0][key] == None:
                        sql_vals.append('None')
                    else:
                        sql_vals.append(result[0][key].replace("'","''"))
        sql_line += '(['+'],['.join(sql_cols)+"]) values ('"+"','".join(sql_vals)+"')"
        cursor.execute(sql_line)
        cursor.commit()
    # check if line has been inputted in post_data
    check = cursor.execute("select post_id from [dbo].["+FB_POST_DATA_DB+"] where [post_id]='"+post_id+"' and [date]='"+today.isoformat()+"'").fetchall()
    if check == []:
        sql_line = 'INSERT INTO [dbo].['+FB_POST_DATA_DB+'] ([date],[post_id]) VALUES ('
        sql_line += "'"+today.isoformat()+"','"+post_id+"')"
        cursor.execute(sql_line)
        cursor.commit()

        fql_line = 'SELECT metric, value FROM insights WHERE object_id="'+post_id+'"'
        fql_line += ' and period="0" and metric in ("'+'","'.join(post_metrics)+'")'
        result = graph.fql(fql_line)
        for r in result:
            log.write( 'Metric: '+r.values()[0]+'\n')
            sql_line = 'UPDATE [dbo].['+FB_POST_DATA_DB+'] SET ['+r.values()[0]+']='
            if isinstance(r.values()[1],str):
                sql_line+="'"+r.values()[1]+"' WHERE [date]='"+today.isoformat()+"' and post_id='"+post_id+"'"
            else:
                sql_line+=str(r.values()[1])+" WHERE [date]='"+today.isoformat()+"' and post_id='"+post_id+"'"
            cursor.execute(sql_line)
            cursor.commit()
        fql_line = 'SELECT like_info.like_count,share_count,comment_info.comment_count FROM stream'
        fql_line += ' WHERE post_id="'+post_id+'"'
        result = graph.fql(fql_line)
        if result != []:
            for key in result[0].keys():
                log.write( 'Metric: '+key+'\n')
                if key=='share_count':
                    sql_line = 'UPDATE [dbo].['+FB_POST_DATA_DB+'] SET ['+key+']='+str(result[0][key])+" WHERE [date]='"+today.isoformat()+"' and post_id='"+post_id+"'"
                else:
                    sql_line = 'UPDATE [dbo].['+FB_POST_DATA_DB+'] SET ['+result[0][key].keys()[0]+']='+str(result[0][key].values()[0])+" WHERE [date]='"+today.isoformat()+"' and post_id='"+post_id+"'"
                cursor.execute(sql_line)
                cursor.commit()
    else:
        log.write( post_id+' '+today.isoformat()+' already in database!'+'\n')
        print post_id+' '+today.isoformat()+' already in database!'

conn.close()
log.close()
print 'Finished!'