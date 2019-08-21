#encoding:utf-8
import requests
import re
import time
import openpyxl
import xlrd
import json
#下载音乐
def download_music(mname):
    excel_dict=read_excel(mname+'.xlsx')
    # 读取字典中的键和值（歌曲名，songmid）
    for songname,songmid in excel_dict.items():
        # 用songmid拼接链接
        data =json.dumps({"req": {"module": "CDN.SrfCdnDispatchServer","method": "GetCdnDispatch","param": {"guid": "5300073565","calltype": 0,"userip": ""}},"req_0": {"module": "vkey.GetVkeyServer","method": "CgiGetVkey","param": {"guid": "5300073565","songmid": [songmid],"songtype": [0],"uin": "0","loginflag": 1,"platform": "20"}},"comm": {"uin": 0,"format": "json","ct": 20,"cv": 0}})
        url = 'https://u.y.qq.com/cgi-bin/musicu.fcg?callback=getplaysongvkey2904725697971273&g_tk=5381&jsonpCallback=getplaysongvkey2904725697971273&loginUin=0&hostUin=0&format=jsonp&inCharset=utf8&outCharset=utf-8&notice=0&platform=yqq&needNewCode=0&data={}'.format(data)
        response = requests.get(url).content.decode('utf-8')
        purl_data = re.findall(r'\w+\((.*)\)',response)[0]
        data_json = json.loads(purl_data)
        purl = data_json['req_0']['data']['midurlinfo'][0]['purl']
        # 匹配下载链接，拼接链接
        download_url = 'http://124.232.144.154/amobile.music.tc.qq.com/{}'.format(purl)
        print("=======正在下载:{}=======".format(songname))
        music = requests.get(download_url)
        if music.content==None:
            continue
        # 保存音乐
        with open("{}.m4a".format(re.sub(r'[\s+|@<>:\\"/]','',songname)),"wb") as m:
             m.write(music.content)
        time.sleep(2)
# 读取excel中歌曲名及songmid，取出一个字典：｛歌曲名1：songmid1，歌曲名2：songmid2，......｝
def read_excel(excel_name):
    book=xlrd.open_workbook(excel_name,encoding_override='utf-8')
    sheet=book.sheet_by_name('sheet1')
    c_values=sheet.col_values(5) #读取第六列
    c_values_1=sheet.col_values(0)  #读取第一列
    excel_dict={}
    # 封装字典｛歌曲名：songmid｝
    for i in range(1,len(c_values)):
        excel_dict[c_values_1[i]]=c_values[i]
    return excel_dict
# 保存歌手歌曲名单。excel  表头为['歌曲名','专辑','时长','播放链接','歌词','songmid']
def write_excel(mname):
    url = 'https://c.y.qq.com/soso/fcgi-bin/client_search_cp'
    url_lrc='https://c.y.qq.com/lyric/fcgi-bin/fcg_query_lyric_yqq.fcg'
    # 请求头
    headers={
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-encoding':'gzip, deflate, br',
            'Accept-language':'zh-CN,zh;q=0.9',
            'Origin':'https://y.qq.com',
            'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36',
            'Referer':'https://y.qq.com/portal/search.html'
                        }
    dd=openpyxl.Workbook()
    sheet=dd.active
    sheet.title='sheet1'
    sheet['A1']='歌曲名'
    sheet['B1']='专辑'
    sheet['C1']='时长'
    sheet['D1']='播放链接'
    sheet['E1']='歌词'
    sheet['F1']='songmid'
    for x in range(300):
        # 请求参数
        params = {
        'ct':'24',
        'qqmusic_ver': '1298',
        'new_json':'1',
        'remoteplace':'sizer.yqq.song_next',
        'searchid':'64405487069162918',
        't':'0',
        'aggr':'1',
        'cr':'1',
        'catZhida':'1',
        'lossless':'0',
        'flag_qc':'0',
        'p':str(x+1),
        'n':'20',
        'w':mname,
        'g_tk':'5381',
        'loginUin':'0',
        'hostUin':'0',
        'format':'json',
        'inCharset':'utf8',
        'outCharset':'utf-8',
        'notice':'0',
        'platform':'yqq.json',
        'needNewCode':'0'
        }
        # 将参数封装为字典
        res_music = requests.get(url,headers=headers,params=params)
        # 调用get方法，下载这个字典
        json_music = res_music.json()
        # 使用json()方法，将response对象，转为列表/字典
        list_music = json_music['data']['song']
        if 'list' not in list_music.keys():
            break
        else:
            list_music=list_music['list']
        # 一层一层地取字典，获取歌单列表
        for music in list_music:
        # list_music是一个列表，music是它里面的元素
            music_name=music['name']
            music_album=music['album']['name']
            music_time=str(music['interval'])+'秒'
            music_url='https://y.qq.com/n/yqq/song/'+music['mid']+'.html'
            music_id=music['id']
            songmid=music['mid']
            params_music = {
            'nobase64': '1',
            'musicid': music_id,
            '-': 'jsonp1',
            'g_tk': '5381',
            'loginUin': '0',
            'hostUin': '0',
            'format': 'json',
            'inCharset': 'utf8',
            'outCharset': 'utf-8',
            'notice': '0',
            'platform': 'yqq.json',
            'needNewCode': '0',
            }
            res_lrc=requests.get(url_lrc,headers=headers,params=params_music)
            json_lrc=res_lrc.json()
            try:
                lrc_list=json_lrc['lyric']
            except KeyError:
                lrc_list='无'
            lrc_str=re.sub('\[*\]*\&*\#*[a-zA-Z]*[0-9]*\;*', '', lrc_list)
            str_music = music_name + '\n所属专辑：' + music_album + '\n播放时长：' + music_time + '\n播放链接：'+music_url + '\n歌词：  '+lrc_str+'\n\n'
            print(str_music)
        # 封装列表写到表格
            music_list=[music_name,music_album,music_time,music_url,lrc_str,songmid]
            sheet.append(music_list)
        time.sleep(0.5)
    dd.save(mname+'.xlsx')
# 选择功能，返回数字
def choice_gongneng():
    choice_num = input('------请选择功能（输入数字即可）:------\n1.下载歌手所有歌曲信息\n2.下载歌手所有歌曲\n3.退出\n')
    while True:
        if choice_num in ['1', '2', '3']:
            return choice_num
        else:
            choice_num = input('------输入错误------\n请选择功能（输入数字即可）：\n1.下载歌手所有歌曲信息\n2.下载歌手所有歌曲\n')

if __name__ == '__main__':
    mname = input('请输入歌手名：\n')#也可以写到下面的循环里
    while True:
        choice_num=choice_gongneng()
        if choice_num =='1':
            write_excel(mname)
            print('------保存歌曲信息成功！------')
        elif choice_num=='2':
            try:
                download_music(mname)
                print('----下载完毕----')
            except FileNotFoundError:#如果没有歌手的歌曲信息表，则先下载歌曲信息表到xlsx，然后再下载歌曲
                write_excel(mname)
                download_music(mname)
                print('----下载完毕----')
        else:
            exit()