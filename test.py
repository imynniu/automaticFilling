import openpyxl
import requests
from time import sleep
import time
import json

# 请求头
my_header = {
    "User-Agent":"User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36 MicroMessenger/7.0.9.501 NetType/WIFI MiniProgramEnv/Windows WindowsWechat",
    "imprint":"oWRkU0X0y2TYFRDqFFdGW153oLM0",
    "Referer":"https://servicewechat.com/wx23d8d7ea22039466/1674/page-frame.html",
}
# "Accept-Encoding": "gzip, deflate, br",
# "Connection": " keep-alive",
# "Host": "a.welife001.com"
##url 地址
host = 'https://a.welife001.com'


##参数定义
# 提交地址
url_submit ='/applet/notify/feedbackWithOss'

# 数据源
filename = 'data.xlsx'
sheet_name = 'Sheet1'
dict_data = {"extra":1,"id":"632886752186540a0fb177df","daka_day":"2022-10-09","submit_type":"submit","networkType":"wifi","member_id":"63292a61a6ea452aaf90f859","op":"add","invest":{"is_tmp":False,"is_private":False,"_id":"6328867577181f99d2d26fc9","subject":[{"seq":0,"cate":1,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fcb","name":"是","checked":True},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fcc","name":"否","checked":False}],"_id":"6328867577181f99d2d26fca","title":"身体是否健康","required":True,"valid":True},{"seq":1,"cate":2,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fce","name":""},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fcf","name":""}],"_id":"6328867577181f99d2d26fcd","title":"体温（早 中 晚）","required":True,"input":{"content":"36.5 36.6 36.6","file":[]},"valid":True},{"seq":2,"cate":4,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd1","name":""},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd2","name":""}],"_id":"6328867577181f99d2d26fd0","title":"手机定位","required":True,"input":{"content":"{\"title\":\"大连市甘井子区人民政府\",\"address\":\"辽宁省大连市甘井子区明珠广场1号\",\"location\":[121.525529,38.953054]}"},"valid":True},{"seq":3,"cate":1,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd4","name":"是","checked":False},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd5","name":"否","checked":True}],"_id":"6328867577181f99d2d26fd3","title":"是否处于隔离状态","required":True,"valid":True},{"seq":4,"cate":1,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd7","name":"常态化","checked":True},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd8","name":"低风险","checked":False},{"seq":2,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd9","name":"中风险","checked":False},{"seq":3,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fda","name":"高风险","checked":False}],"_id":"6328867577181f99d2d26fd6","title":"目前所在地区风险等级（国务院小程序查询）","required":True,"valid":True},{"seq":5,"cate":2,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fdc","name":""},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fdd","name":""}],"_id":"6328867577181f99d2d26fdb","title":"是否有需要指导员协助事宜","required":False,"valid":True,"input":{"content":"否","file":[]}}],"create_at":"2022-09-19T15:10:45.068Z","update_at":"2022-09-19T15:10:45.068Z","__v":0,"time":1665293821680,"day":"2022-10-09"},"feedback_text":""}

dict_data_or = {"extra":1,"id":"632886752186540a0fb177df","daka_day":"2022-10-09","submit_type":"submit","networkType":"wifi","member_id":"63292a61a6ea452aaf90f859","op":"add","invest":{"is_tmp":False,"is_private":False,"_id":"6328867577181f99d2d26fc9","subject":[{"seq":0,"cate":1,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fcb","name":"是","checked":True},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fcc","name":"否","checked":False}],"_id":"6328867577181f99d2d26fca","title":"身体是否健康","required":True,"valid":True},{"seq":1,"cate":2,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fce","name":""},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fcf","name":""}],"_id":"6328867577181f99d2d26fcd","title":"体温（早 中 晚）","required":True,"input":{"content":"36.5 36.6 36.6","file":[]},"valid":True},{"seq":2,"cate":4,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd1","name":""},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd2","name":""}],"_id":"6328867577181f99d2d26fd0","title":"手机定位","required":True,"input":{"content":"{\"title\":\"大连市甘井子区人民政府\",\"address\":\"辽宁省大连市甘井子区明珠广场1号\",\"location\":[121.525529,38.953054]}"},"valid":True},{"seq":3,"cate":1,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd4","name":"是","checked":False},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd5","name":"否","checked":True}],"_id":"6328867577181f99d2d26fd3","title":"是否处于隔离状态","required":True,"valid":True},{"seq":4,"cate":1,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd7","name":"常态化","checked":True},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd8","name":"低风险","checked":False},{"seq":2,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fd9","name":"中风险","checked":False},{"seq":3,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fda","name":"高风险","checked":False}],"_id":"6328867577181f99d2d26fd6","title":"目前所在地区风险等级（国务院小程序查询）","required":True,"valid":True},{"seq":5,"cate":2,"inputs_count":0,"inputs":[],"item_details":[{"seq":0,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fdc","name":""},{"seq":1,"checks_count":0,"rate":0,"file":[],"checkedlist":[],"_id":"6328867577181f99d2d26fdd","name":""}],"_id":"6328867577181f99d2d26fdb","title":"是否有需要指导员协助事宜","required":False,"valid":True,"input":{"content":"否","file":[]}}],"create_at":"2022-09-19T15:10:45.068Z","update_at":"2022-09-19T15:10:45.068Z","__v":0,"time":1665293821680,"day":"2022-10-09"},"feedback_text":""}

# 日志输出方式
# with open("text.txt","a") as file:
#     file.write("What I want to add on goes here")

# 读取数据
data = []
# 读取数据
def read_data(filename,sheet_name):
    wb = openpyxl.load_workbook(filename)
    sheet = wb[sheet_name]
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=sheet.max_column, max_row=sheet.max_row):
        temp_data = []
        for d in row :
            temp_data.append(d.value)
        data.append(temp_data)
    return 1

# 修改信息
def modify_data(data_person):
    # 问卷id
    id_person = data_person[0]
    dict_data['id'] = id_person
    # 打卡日期
    time_today = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    dict_data['daka_day'] = time_today
    # 成员id
    member_id_person = data_person[1]
    dict_data['member_id'] = member_id_person
    temperature_person =data_person[2]
    # 问卷日期
    dict_data['invest']['day'] = time_today
    # 打卡时间戳
    dict_data['invest']['time'] = int(time.time()*1000)
    # 体温
    dict_data['invest']['subject'][1]['input']['content']=temperature_person
    # 地理位置
    city = data_person[5]
    buildings = data_person[6]
    lat = data_person[7]
    lng = data_person[8]
    loc = {
        "title":buildings,
        "address":city,
        "location":[float(lat), float(lng)]
    }
    dict_data['invest']['subject'][2]['input']['content'] = json.dumps(loc)
    # 联系赵洪吉
    content_person = data_person[9]
    dict_data['invest']['subject'][5]['input']['content'] = content_person




###提交表单
def send_request(url,request_data,log_name):
    r= requests.post(url,json=request_data, headers= my_header)
    if(r.json()['msg'] =='ok'):
        with open("log.txt","a") as file:
            file.write(time.strftime('\n'+'%Y-%m-%d %H:%M:%S', time.localtime(time.time()))+' '+log_name+' 的信息上传成功！！')
    return 1


if __name__ == '__main__':
    while(1):
        data.clear()
        re_read = read_data(filename, sheet_name)
        for data_temp in data:
            dict_data = dict_data_or.copy()
            modify_data(data_temp)
            url =host+url_submit
            # print(dict_data)
            send_request(url,dict_data,data_temp[10])
        print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        print('填报成功，等待八小时后再次填报......')
        sleep(28800)