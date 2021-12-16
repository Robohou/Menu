# Menu
Use python to generate useful menu
import os, os.path, xlwt
 
BRANCH = '├─'
LAST_BRANCH = '└─'
TAB = '│  '
EMPTY_TAB = '   '
 
global line
 
def setStyle(name, height,color, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    # 字体类型：比如宋体、仿宋也可以是汉仪瘦金书繁
    font.name = name
    # 是否为粗体
    font.bold = bold
    # 设置字体颜色
    font.colour_index = color
    # 字体大小
    font.height = height
    # 字体是否斜体
    font.italic = False
    # 字体下划,当值为11时。填充颜色就是蓝色
    font.underline = 10
    # 字体中是否有横线struck_out
    font.struck_out =False
    # 定义格式
    style.font = font

    return style

# 遍历当前文件夹下的子文件夹和文件
def get_dir_list(path, placeholder=''):
    global line
    folder_list = [folder for folder in os.listdir(path) if os.path.isdir(os.path.join(path, folder))] # 全部子文件夹名称
    file_list = [file for file in os.listdir(path) if os.path.isfile(os.path.join(path, file))] # 全部子文件名称
    for folder in folder_list:
        folder_path = os.path.join(path, folder)
        if list(os.listdir(folder_path)):
            
            # 文件夹名 写入文本文档，并添加缩进
            if folder != folder_list[-1]:
                txt_index.write(placeholder + BRANCH + folder +':'+ '\n')  
            else:
                txt_index.write(placeholder + (BRANCH if file_list else LAST_BRANCH) + folder + '\n')
            # 写入Excel并增加超链接
            link = 'HYPERLINK("%s";"%s")'%(folder_path.replace(main_path+"\\",""),folder+':')  # 相对路径
            sheet_index.write(line, placeholder.count(TAB), xlwt.Formula(link),setStyle(u'微软雅黑', 210, 2, True))#2为红色，1为白色
            # link = 'HYPERLINK("%s";"%s")'%(folder_path,folder)  # 绝对路径
            # sheet_index2.write(line, placeholder.count(TAB), xlwt.Formula(link))
            line = line + 1
            get_dir_list(folder_path, placeholder + TAB)  # 内层文件夹进深一层，并搜索
        else:
            os.rmdir(folder_path) # 删除空文件夹
 
    for file in file_list:
        if str.upper(os.path.splitext(file)[-1]) not in [".TIF", ".PNG",".JPG",".DB",".PY"]:#排除Thrumb.db文件的干扰
            file = file.replace("‑","")
            # 文件名 写入文本文档，并添加缩进
            try:
                if file != file_list[-1]:
                    txt_index.write(placeholder + BRANCH + file + '\n')
                else:
                    txt_index.write(placeholder + LAST_BRANCH + file_list[-1] + '\n') 
            except UnicodeEncodeError:
                    txt_index.write(placeholder + BRANCH + "特殊字符，无法写入txt文件！！！" + '\n')
                    print("这个文件名里有特殊字符："+ file)
            link ='HYPERLINK("%s";"%s")'%(os.path.join(path, file).replace(main_path+"\\",""),file) 
            sheet_index.write(line, placeholder.count(TAB), xlwt.Formula(link),setStyle(u'微软雅黑', 210, 4, False))
            # link = 'HYPERLINK("%s";"%s")'%(os.path.join(path, file),file)
            # sheet_index2.write(line, placeholder.count(TAB), xlwt.Formula(link))
            line = line + 1
 
if __name__ == '__main__':
 
    line = 1
    book = xlwt.Workbook()
    main_path = "./" #  main_path 目录就是当前目录

    sheet_index = book.add_sheet('目录')  # 目录必须放置于 main_path 下
    sheet_index.write(0, 0, "超级链接目录",setStyle(u'仿宋', 250, 2, True))
    sheet_index.write(0, 3, "由H朝中自动编辑，也可以手动修改",setStyle(u'仿宋', 200, 2, False))
    sheet_index2 = book.add_sheet('关于本表格')
    sheet_index2.write(0, 0, "作者：侯朝中 18010092511")
    sheet_index2.write(1, 0, "微信号:Robohou")
    sheet_index2.write(2, 0, "博客：https://www.cnblogs.com/robohou/")
    with open(main_path+"\\"+'目录.txt','w') as txt_index:
        txt_index.write(main_path+"\n")
        get_dir_list(main_path)
    txt_index.close()
    book.save(main_path+"\\"+'目录(带链接).xls')
    print("完成了！")
