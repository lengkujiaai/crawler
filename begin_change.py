import pdfplumber
from fpdf import FPDF
import win32api
import win32print
#import webbrowser
import time
import os
import re
import shutil

import warnings
warnings.filterwarnings("ignore")
#Python的第三方包往往依赖其它的包进行开发。一旦依赖的包发生较大的版本升级，那么往往会出现兼容性问题， 引起编译器警告或报错

#author:shihailong
#date:20220906
 
import sys

def parse_pdf(FILE_PATH):
    #FILE_PATH= 'C:\\Users\\lengk\\beibaoshibie\\pdf_change\\test.pdf'
    #FILE_PATH= 'test.pdf'
    with pdfplumber.open(FILE_PATH) as pdf:
        content = ''
        for i in range(len(pdf.pages)):
            page = pdf.pages[i]
            new_txt = page.extract_text()
            page_content = '\n'.join(page.extract_text().split('\n')[:-1])
            content = content + page_content
        c = content.split('\n')
    return content


def save(content, file_dir, SAVE_FILE_NAME, OPEN_PATH):
    #SAVE_FILE_PATH= 'C:\\Users\\lengk\\beibaoshibie\\pdf_change\\parse_and_save.pdf'
    SAVE_FILE_PATH = file_dir + '中文版' + SAVE_FILE_NAME
    DELETE_FILE_PATH = file_dir + SAVE_FILE_NAME
    count_header = 0 #用来控制头部显示
    save_pdf = FPDF()
    #save_pdf.add_font('SIMYOU','','SIMYOU.ttf',True)
    save_pdf.add_font('youyuan','','youyuan2.ttf',True)
    save_pdf.add_page()
    save_pdf.set_font("youyuan", size=8)
    c = content.split('\n')
    #print(c)
    #length = len(c) - 2
    #c = c[:length]
    count_comment = 1#由于中英文替换中有两个Comment,值为1代表跑道代码，值为2代表其它。并且1在前面出现，替换后赋值2
    count_average = 0#用来计算显示000-1800后面的Average这行的
    for item in c:
        #print(item)
        count_header+=1
        if(count_header < 3):
            if(count_header == 1):
                save_pdf.set_font("youyuan", size=15)
                item = "摩擦系数测试数据总结"
            if(count_header == 2):
                save_pdf.set_font("youyuan", size=10)
                item = '版权： 中文测试版'
            save_pdf.cell(200, 5, txt=item, ln=2, align='C')
        else:
            save_pdf.set_font("youyuan", size=8)
            if(count_average == 19):#这个只能放在这里，不能放后面
                count_average += 1#跳出去
                save_pdf.cell(200, 5, txt='\n', ln=2)
                #f.write('\n')
            if(count_average == 18):#这个只能放在这里，不能放后面
                count_average += 1#跳出去
                item = re.sub(' +', ' ', item)
                item = item[1:]
                item = item.replace(' ','               ')
            #new_list = item.split(' ')
            #print('new_list: ')
            #print(new_list)
            if(item.find('Date') != -1):
                if(item.find('January') != -1):
                    item = item.replace('January', '一月')
                if(item.find('February') != -1):
                    item = item.replace('February', '二月')
                if(item.find('March') != -1):
                    item = item.replace('March', '三月')
                if(item.find('April') != -1):
                    item = item.replace('April', '四月')
                if(item.find('May') != -1):
                    item = item.replace('May', '五月')
                if(item.find('June') != -1):
                    item = item.replace('June', '六月')
                if(item.find('July') != -1):
                    item = item.replace('July', '七月')
                if(item.find('August') != -1):
                    item = item.replace('August', '八月')
                if(item.find('September') != -1):
                    item = item.replace('September', '九月')
                if(item.find('October') != -1):
                    item = item.replace('October', '十月')
                if(item.find('November') != -1):
                    item = item.replace('November', '十一月')
                if(item.find('December') != -1):
                    item = item.replace('December', '十二月')
            if(item.find('00-') != -1):
                item = re.sub('00-', '00- ', item)
                item = re.sub(' +', ' ', item)
                new_list = item.split(' ')
                #print(item)
                #print(new_list)
                #print(len(new_list))
                #print(count_header)
                num1000 = int(new_list[2])
                if(len(new_list) == 5):
                    if(num1000 < 1000):
                        item = new_list[1] + new_list[2] + '             '+new_list[3] +'            '+new_list[4]
                    elif(num1000 == 1000):
                        item = new_list[1] + new_list[2] + '            '+new_list[3] +'            '+new_list[4]
                    elif(num1000 == 10000):
                        item = new_list[1] + new_list[2] + '          '+new_list[3] +'            '+new_list[4]
                    elif(num1000 < 10000):
                        item = new_list[1] + new_list[2] + '           '+new_list[3] +'            '+new_list[4]
                    else:
                        item = new_list[1] + new_list[2] + '         '+new_list[3] +'            '+new_list[4]
                if(len(new_list) == 9):#在长度为3100到3500时读取有问题，特殊处理，但是读取还是少了3300-3400的数据
                    item = new_list[1] + new_list[2] + '           '+new_list[3] +'            '+new_list[4]
                    save_pdf.cell(200, 5, txt=item, ln=2)
                    item = new_list[5] + new_list[6] + '           '+new_list[7] +'            '+new_list[8]
                if(len(new_list) == 7):
                    if(num1000 < 1000):
                        item = new_list[1] + new_list[2] + '             '+new_list[3] +'            '+new_list[4]+'                 '+new_list[5]+'             '+new_list[6]
                    elif(num1000 == 1000):
                        item = new_list[1] + new_list[2] + '            '+new_list[3] +'            '+new_list[4]+'                 '+new_list[5]+'             '+new_list[6]
                    else:
                        item = new_list[1] + new_list[2] + '           '+new_list[3] +'            '+new_list[4]+'                 '+new_list[5]+'             '+new_list[6]
            if(item.find('Section')!=-1 and item.find('Avg')==-1):
                list_avg = item.split(' ')
                empty1 = '                                  '
                empty2 = '                       '
                empty3 = '                     '
                if(len(list_avg) == 6):#显示3列数据的情况
                    item = list_avg[1] + ' '+list_avg[2] + empty1 + list_avg[3]+ empty2 + list_avg[4]+ empty3 + list_avg[5]
                elif(len(list_avg) == 4):#显示1列数据的情况
                    item = list_avg[1] + ' '+list_avg[2] + empty1 + list_avg[3]
                item = item.replace('Section A','A段')
                item = item.replace('Section B','B段')
                item = item.replace('Section C','C段')
            if(item.find('Run Average')!=-1):
                list_avg = item.split(' ')
                empty1 = '                               '
                empty2 = '                       '
                empty3 = '                     '
                if(len(list_avg) == 6):#显示3列数据的情况
                    item = list_avg[1] + ' '+list_avg[2] + empty1 + list_avg[3]+ empty2 + list_avg[4]+ empty3 + list_avg[5]
                if(len(list_avg) == 4):#显示1列数据的情况
                    item = list_avg[1] + ' '+list_avg[2] + empty1 + list_avg[3]
            if(item.find('Run 1 Data')!=-1):
                #save_pdf.cell(200, 5, txt='\n', ln=2)
                item = "                         跑道1                               跑道2"
            if(item.find('FILE HEADER')!=-1):
                item = item.replace('FILE HEADER', '基本信息')
            if(item.find('km/hr')!=-1):
                item = re.sub(' +', ' ', item)
                item = item[1:]
                item = item.replace(' ','                      ')
            if(item.find('Runway:')!=-1):
                item = item.replace('Runway:', '跑道编号:')
            if(item.find('Comment:')!=-1 and count_comment == 1):
                count_comment = 2
                item = item.replace('Comment:', '跑道代码:')
            if(item.find('Date:')!=-1):
                item = item.replace('Date:', '测试日期:')
            if(item.find('Time:')!=-1):
                item = item.replace('Time:', '时间:')
            if(item.find('Version:')!=-1):
                item = item.replace('Version:', '软件版本:')
            if(item.find('Mode')!=-1):
                item = item.replace('Mode', '测试标准')
            if(item.find('Water')!=-1):
                item = item.replace('Water', '喷水与否')
                if(item.find('Was Applied')!=-1):
                    item = item.replace('Was Applied', '水')
                else:
                    item = ' 喷水与否: 无水'
            if(item.find('Unit')!=-1):
                item = item.replace('Unit:', '单位:')
            if(item.find('Scale')!=-1):
                item = ' 换算比：1016摩擦系数测试距离读取成1000'
            if(item.find('Set Length')!=-1):
                item = item.replace('Set Length', '测试距离')
                item = item.replace('(Set to the nearest 100 meters)', '(设置成100米为最小单位)')
            if(item.find('Run1 Length')!=-1):
                item = item.replace('Run1 Length', '跑道1 长度')
            if(item.find('Run2 Length')!=-1):
                item = item.replace('Run2 Length', '跑道2 长度')
            if(item.find('Run Length')!=-1):
                item = item.replace('Run Length', '跑道长度')
            if(item.find('OPERATOR MESSAGES')!=-1):
                item = item.replace('OPERATOR MESSAGES', '测试人员信息')
            if(item.find('No Messages')!=-1):
                item = item.replace('No Messages', '无')
            if(item.find('ADDITIONAL DATA')!=-1):
                item = item.replace('ADDITIONAL DATA', '附加数据')
            if(item.find('Operator')!=-1):
                item = item.replace('Operator', '操作员')
            if(item.find('Tire Type')!=-1):
                item = item.replace('Tire Type', '测试轮胎类型')
            if(item.find('Comment')!=-1 and count_comment == 2):
                item = item.replace('Comment', '其它')
            if(item.find('AVERAGE VALUES')!=-1):
                item = item.replace('AVERAGE VALUES', '平均值')
            
            if(item.find('Distance')!=-1):
                item = item.replace('Distance', '跑道位置')
            if(item.find('Avg Speed')!=-1):
                item = item.replace('Avg Speed', '平均速度')
            if(item.find('Avg Friction')!=-1):
                item = item.replace('Avg Friction', '平均摩擦系数')
            if(item.find('Run Average')!=-1):
                item = item.replace('Run Average', '平均值')
            if(item.find('Overall Friction Average')!=-1):
                item = item.replace('Overall Friction Average', '摩擦系数总体平均值')
            if(item.find('Overall Avg')!=-1):#针对第3版的更改
                new_list = item.split(' ')
                #print(new_list)
                #item = item.replace('Overall Avg', '平均值')
                empty1 = '              '
                empty2 = '            '
                if(len(new_list) == 5):
                    item = '平均值' + empty1 + new_list[3] + empty2 + new_list[4]
            if(item.find('Average')!=-1):#这个平均值要放在最后，否则把其它的部分给干扰了
                item = item.replace('Average', '平均值  ')
                item = re.sub(' +', ' ', item)
                new_list = item.split(' ')
                item = new_list[1] +'              '+ new_list[2] + '            '+new_list[3] +'                 '+new_list[4] + '             '+new_list[5]
            if(item.find('Fric: Run 1')!=-1):
                new_list = item.split('Fric: Run 1')
                save_pdf.cell(200, 5, txt=new_list[0], ln=2)
                #save_pdf.cell(200, 5, txt='', ln=2)
                #save_pdf.cell(200, 5, txt='分段摩擦系数平均值', ln=2)
                item = re.sub(' +', ' ', item)
                list0 = item.split(' ')
                item = '分段摩擦系数平均值            摩擦系数：跑道1            摩擦系数：跑道2                平均值'
            if(item.find('SECTION FRICTION AVERAGES')!=-1):
                item = '分段摩擦系数平均值'
            if(item.find('Fric:')!=-1):#针对第3版的更改，不要挪位置
                continue
            if(item.find('RUN DATA')!=-1):#针对第3版的更改，不要挪位置
                continue
            if(item.find('Meters')!=-1):
                item = item.replace('Meters', '米')
            if(item.find('EVENTS')!=-1):
                #因为有的EVENTS和上一行的数字在一行，有的不在一行，无奈之下才这样处理的
                item = item.replace('EVENTS', '')
                save_pdf.cell(200, 5, txt=item, ln=2)
                item = '事件：'
            if(item.find('No Events')!=-1):
                item = item.replace('No Events', '没有事件')
            save_pdf.cell(200, 5, txt=item, ln=2)
    save_pdf.output(SAVE_FILE_PATH)
    #print(SAVE_FILE_PATH)
    os.remove(DELETE_FILE_PATH)
    #webbrowser.open(SAVE_FILE_PATH, new=2)

def getFlist(file_dir):
    root = ''
    dirs = ''
    files = ''
    for root,dirs,files in os.walk(file_dir):
        a = 0
        #print('root_dir: ', root)
        #print('sub_dirs: ', dirs)
        #print('files: ', files)
    return files


def translate_and_print():
    print('开始运行英文pdf转换中文pdf程序')
    print('----ok----')
    file_dir = 'work_dir\\'
    #file_dir是存放英文版pdf的目录，例如: file_dir'C:\\Users\\pdf_dir\\'
    read_file_dir = file_dir  #读取的英文版目录
    save_file_dir = file_dir  #要保存的中文版目录，现在是在一个位置，如果中文版想和英文版不一样，就改这里
    
    #--------我的常量--------------------------
    SOURCE_FILE = 'E:\\pdf_change\\pdf_change_v1\\'
    default_HPRT_MT800 = 'HP LaserJet Pro MFP M128fw[f311e6]'
    #--------客户的常量------------------------
    #SOURCE_FILE = 'C:\\Tradewind\\FTMS\\pdf_change_v1\\'
    #default_HPRT_MT800 = 'HPRT MT800'
    
    OPEN_PATH = SOURCE_FILE + 'work_dir\\'
    #filenames = getFlist(file_dir)
    #time.sleep(5)
    while(True):
        try:
            #time.sleep(4)
            try:
                filenames2 = getFlist(SOURCE_FILE+'english\\') 
                for name in filenames2:
                    print(filenames2)
                    print(len(filenames2))
                    #print(name)
                    if(name.find('.pdf') != -1):
                        print(name)
                        print('开始移动---')
                        time.sleep(1)
                        shutil.move(SOURCE_FILE + 'english\\' + name, OPEN_PATH + name)
                        print('已经把pdf从目录english移动到目录work_dir')
            except  Exception as e:
                pass
            
            filenames = getFlist(file_dir)
            
            if(len(filenames) == 1):
                if(filenames[0].find('中文版') == -1):
                    SAVE_FILE_NAME = filenames[0]
                    content = parse_pdf(read_file_dir + SAVE_FILE_NAME)
                    save(content, save_file_dir, SAVE_FILE_NAME, OPEN_PATH)
                    print('保存新的中文版')
                    time.sleep(0.01)
                    
                    #--------------------------这里是打印设置----begin-------------------------------------
                    printers = win32print.EnumPrinters(3)
 
                    # 获取默认打印机
                    default_printers = win32print.GetDefaultPrinter()
                    print(default_printers)
                    # 指定另一个打印机名作为默认打印机，如果在其它地方已经设置了默认打印机，则可以不进行下面的设置
                    #win32print.SetDefaultPrinter('HP Color MFP E87640-50-60 PCL-6 (V4) (网络)')# 这里可以换成其他打印机名称
                    #win32print.SetDefaultPrinter('HP LaserJet Pro MFP M128fw[f311e6]')# 这里可以换成其他打印机名称
                    win32print.SetDefaultPrinter(default_HPRT_MT800)
                    default_printers = win32print.GetDefaultPrinter()
                    #print(default_printers)
                    
 
                    # 设置权限作为获得句柄语句的参数，有时也可不用
                    printaccess = {"DesiredAccess":win32print.PRINTER_ACCESS_USE}# 较低的权限
                    print_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}# 较高的权限
                    # 获取指定打印机句柄
                    pHandle = win32print.OpenPrinter(default_printers,print_DEFAULTS)# 这里使用默认打印机，第2个权限参数是可选选项，但如果不设置足够高的权限可能无法成功更改打印参数设置
                    # 根据指定打印机句柄获取指定打印机信息
                    properties = win32print.GetPrinter(pHandle,2)#传入1返回1个元祖，传入2返回1个字典
                    # 获取打印机打印参数设置——pDevMode类
                    devmode = properties['pDevMode']
                    devmode.Orientation = 1
                    properties['pDevMode'] = devmode
                    win32print.SetPrinter(pHandle,2,properties,0)
                    #--------------------------这里是打印设置----end----------------------------------------
                    
                    #--------------------------下面是调用打印机打印-----------------------------------------
                    GHOSTSCRIPT_PATH = SOURCE_FILE + "GHOSTSCRIPT\\bin\\gswin32.exe"
                    GSPRINT_PATH =  SOURCE_FILE + "GSPRINT\\gsprint.exe"
                    currentprinter = win32print.GetDefaultPrinter()
                    auto_print_pdf = '" "' + OPEN_PATH + '中文版' + filenames[0]+ '"'
                    #print('auto_print_pdf: ')
                    #print('                ' + auto_print_pdf)
                    win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH+'" -printer "'+currentprinter+ auto_print_pdf, '.', 0)
                    #time.sleep(5)
                    print('打印命令发送完成')
            else:
                for item in filenames:
                    time.sleep(0.01)
                    if(item.find('中文版') != -1):
                        print('删除旧的中文版:' + item)
                        #print(item)
                        os.remove(file_dir + item)
                    #else:
                    #    print('保留英文版')
            
        except  Exception as e:
            print('----try 中出现错误----')
            print(e)
            #pass
translate_and_print()