#!/usr/bin/env python
#encoding: utf-8
#Author: guoxudong
import xlwt
import paramiko
import time

# global infoList
infoList = []
def get_ssh(ip, user, pwd):
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ip, 22, user, pwd, timeout=15)
        return ssh
    except Exception,e:
        print e
        return "False"
def init_chan(ssh,command):
    ssh_t = ssh.get_transport()
    chan = ssh_t.open_session()
    chan.setblocking(0)                                                               # 设置非阻塞
    chan.exec_command(command)
    return chan

def dojob(appName,ip, commandIn="ll -h /home/admin/taobao-tomcat-production-7.0.59.3/deploy|grep "):
    ssh = get_ssh(ip, 'deploy', 'OyeelDS1007')                                     # 连接远程服务器
    chan=init_chan(ssh,commandIn+appName+"*.war")
    # chan=init_chan(ssh,commandIn+appName+"*")
    flag=True
    while flag:
        # import time
        while chan.recv_ready():
            log_msg = chan.recv(10000).strip()
            log = log_msg.split(' ')                                                 #处理ls数据
            now = time.localtime().tm_year                                           #记录当前年份，方便生成时间戳
            if log[-4] == '':
                month = log[-5]
                bar = log[-6]
            else:
                month = log[-4]
                bar = log[-5]
            creatTime = month+' '+log[-3]+' '+log[-2]+' '+str(now)                  #拼接时间戳需要的格式
            timestamp = time.mktime(time.strptime(creatTime, "%b %d %H:%M %Y"))    #生成时间戳
            timeArray = time.localtime(timestamp)
            otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)         #时间戳转换为指定格式日期
            # changeTime = str(now)+' ' + log[-4] + ' ' + log[-3] + ' ' + log[-2]    #拼接创建时间（在结果中显示）
            log_tuple = (ip,log[-1],bar,otherStyleTime,int(timestamp))               #[-1]：war包名；[-5]：包大小；changeTime：创建时间；int(timestamp)：时间戳，用于比较发布的先后时间
            infoList.append(log_tuple)
            print "\t"+ip+"\t"+log_msg
            flag=False;
            break;

def job(server):
    print server.getAlias()                                                         # 打印别名
    for ip in server.getServers():                                                  # 循环遍历IP
        dojob(server.getAlias(),ip,server.getPath())
    print "\n"

from enum import Enum
wl_path="ls -l /apps/tomcat/webapps|grep "
alitomcat_path="ls -l /home/admin/taobao-tomcat-production-7.0.59.3/deploy|grep "
class SERVERS(Enum):
    jg={"alias":"jg","server":["10.70.14.163","10.70.14.164"]}
    jg_service={"alias":"jg-service","server":["10.70.14.170","10.70.14.189"]}
    shgt_service_yy={"alias":"shgt-service-yy","server":["10.70.14.81","10.70.14.82"]}
    shgt_service_yyoc={"alias":"shgt-service-yyoc","server":["10.70.14.194","10.70.14.196"]}
    oc={"alias":"oc","server":["10.70.14.79","10.70.14.80"]}
    shgt_service_zy={"alias":"shgt-service-zy","server":["10.70.14.41","10.70.14.42","10.70.14.45","10.70.14.46"]}
    shgt_service_search={"alias":"shgt-service-search","server":["10.70.14.47","10.70.14.48"]}
    search_ng={"alias":"search-ng","server":["10.70.14.63","10.70.14.64"]}
    shgt_service_dd={"alias":"shgt-service-dd","server":["10.70.14.49","10.70.14.50"]}
    buyer_ng={"alias":"buyer-ng","server":["10.70.14.71","10.70.14.72"]}
    seller_ng={"alias":"seller-ng","server":["10.70.14.43","10.70.14.44"]}
    account_ng={"alias":"account-ng","server":["10.70.14.35","10.70.14.36"]}
    bid_ng={"alias":"bid-ng","server":["10.70.14.67","10.70.14.68"]}
    shgt_business_track={"alias":"shgt-business-track","server":["10.70.14.61","10.70.14.62"]}
    sso={"alias":"sso","server":["10.70.14.69","10.70.14.70"]}
    shgt_service_pay={"alias":"shgt-service-pay","server":["10.70.14.53","10.70.14.54"]}
    pay_ng={"alias":"pay-ng","server":["10.70.14.51","10.70.14.52"]}
    shgt_service_wl={"alias":"shgt-service-wl","server":["10.70.14.57","10.70.14.58"]}
    trace_ng={"alias":"trace-ng","server":["10.70.14.59","10.70.14.60"]}
    jk_misc={"alias":"jk-misc","server":["10.70.14.179","10.70.14.180"]}
    jk_finance={"alias":"jk-finance","server":["10.70.14.175","10.70.14.176"]}
    home_ng={"alias":"home-ng","server":["10.70.14.185","10.70.14.186","10.70.14.187","10.70.14.188"]}
    statics={"alias":"statics","server":["10.70.14.112"]}
    wl={"alias":"wl","server":["10.70.14.239","10.70.14.229"],"path":wl_path}
    service_wl={"alias":"service-wl","server":["10.70.14.240","10.70.14.241"],"path":wl_path}

    def getAlias(self):
        return self.value["alias"]
    def getServers(self):
        return self.value["server"]
    def getPath(self):
        return self.value.get("path",alitomcat_path)

# 执行方式一：单独执行一条
#job(SERVERS.jg)

# 执行方式二：批量执行
# lists = [SERVERS.jg,SERVERS.shgt_service_yy,SERVERS.shgt_service_yyoc,SERVERS.jg_service,SERVERS.oc,SERVERS.shgt_service_zy
#          ,SERVERS.shgt_service_search,SERVERS.search_ng,SERVERS.shgt_service_dd,SERVERS.buyer_ng,SERVERS.seller_ng,SERVERS.account_ng
#          ,SERVERS.bid_ng,SERVERS.shgt_business_track,SERVERS.sso,SERVERS.shgt_service_pay,SERVERS.pay_ng,SERVERS.shgt_service_wl
#          ,SERVERS.trace_ng,SERVERS.jk_finance,SERVERS.jk_misc,SERVERS.home_ng,SERVERS.statics,SERVERS.wl,SERVERS.service_wl]
# lists = [SERVERS.oc,SERVERS.shgt_service_yy,SERVERS.shgt_service_yyoc]
lists = [SERVERS.jg_service]
for item in lists:
    job(item)

#设置Excel表格样式
def set_style(name, height, bold=False, back=False):
    style = xlwt.XFStyle()                                                      # 初始化样式
    font = xlwt.Font()                                                          # 为样式创建字体
    font.name = name                                                            # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    borders = xlwt.Borders()                                                    # 设置边框
    borders.left = xlwt.Borders.THIN  # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    if back:
        patterni = xlwt.Pattern()                                               # 为样式创建图案
        patterni.pattern = 2                                                    # 设置底纹的图案索引，1为实心，2为50%灰色，对应为excel文件单元格格式中填充中的图案样式
        patterni.pattern_fore_colour = 0x16                                     # 设置底纹的前景色，对应为excel文件单元格格式中填充中的背景色
        patterni.pattern_back_colour = 0x16                                     # 设置底纹的背景色，对应为excel文件单元格格式中填充中的图案颜色
        style.pattern = patterni                                                # 为样式设置图案
    return style
#生成Excel表格
wb = xlwt.Workbook(encoding = 'utf-8')
ws = wb.add_sheet('正式发布列表')                                              #设置工作表名称
# styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: bold on;'); # 80% like
ws.col(0).width = 5000                                                          #设置行宽
ws.col(1).width = 6500
ws.col(2).width = 4000
ws.col(3).width = 6000
ws.write(0,0,'ip',set_style('Times New Roman',280,True,True))                #设置首行名称以及样式
ws.write(0,1,'包名',set_style('Times New Roman',280,True,True))
ws.write(0,2,'发布包大小',set_style('Times New Roman',280,True,True))
ws.write(0,3,'最后更新时间',set_style('Times New Roman',280,True,True))
#对数据按照时间戳进行降序排列
sortList =  sorted(infoList, key=lambda time: time[4], reverse=True)           # sort by timestamp
#遍历数据插入Excel表格
for index,information in enumerate(sortList):
    for num,info in enumerate(information):
        if num > 3:                                                             #不显示时间戳
            break
        ws.write(index + 1,num,information[num],set_style('Times New Roman',260,False))

# filetime = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime())
# fileName = filetime + '_checkList.xls'
try:
    wb.save('checkList.xls')                                                  #切记不可xls文件打开的时候跑脚本，否则将会报错
except IOError as e:
    print '保存文件时出错：',e
    print '请关闭Excel文件，然后重启程序'

#附录：背景色对照
_colour_map_text = """\
aqua 0x31
black 0x08
blue 0x0C
blue_gray 0x36
bright_green 0x0B
brown 0x3C
coral 0x1D
cyan_ega 0x0F
dark_blue 0x12
dark_blue_ega 0x12
dark_green 0x3A
dark_green_ega 0x11
dark_purple 0x1C
dark_red 0x10
dark_red_ega 0x10
dark_teal 0x38
dark_yellow 0x13
gold 0x33
gray_ega 0x17
gray25 0x16
gray40 0x37
gray50 0x17
gray80 0x3F
green 0x11
ice_blue 0x1F
indigo 0x3E
ivory 0x1A
lavender 0x2E
light_blue 0x30
light_green 0x2A
light_orange 0x34
light_turquoise 0x29
light_yellow 0x2B
lime 0x32
magenta_ega 0x0E
ocean_blue 0x1E
olive_ega 0x13
olive_green 0x3B
orange 0x35
pale_blue 0x2C
periwinkle 0x18
pink 0x0E
plum 0x3D
purple_ega 0x14
red 0x0A
rose 0x2D
sea_green 0x39
silver_ega 0x16
sky_blue 0x28
tan 0x2F
teal 0x15
teal_ega 0x15
turquoise 0x0F
violet 0x14
white 0x09
yellow 0x0D"""