---
layout: post
title: 使用python-docx 根据图片生成word报告
date: 2024-05-29
Author: 李然
categories: 
tags: [code,python]
comments: true
--- 

根据图片生成word报告

以下是代码

```Python
from copy import deepcopy
from datetime import datetime
import os

from docx import Document
from PIL import Image
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from concurrent.futures import ThreadPoolExecutor, wait, ALL_COMPLETED
from docx.enum.text import WD_BREAK
from docx.text.paragraph import Paragraph
from docx.oxml.xmlchemy import OxmlElement

bugMap = {
    "杆塔树障": "基础",
    "杆塔未封顶": "基础",
    "杆塔异物": "基础",
    "施工遗留": "基础",
    "杆塔鸟巢": "基础",
    "杆塔倾斜": "基础",
    "塔基植被覆盖": "基础",
    "塔基杂物堆积": "基础",
    "杆塔倾斜": "基础",
    "塔基树障": "基础",
    "杆塔异物": "基础",
    "杆塔裂纹": "基础",
    "杆塔未封顶": "基础",
    "杆塔损伤": "基础",
    "塔头破损": "基础",
    "杆塔破损": "基础",
    "拉线松弛": "基础",
    "横担锈蚀": "基础",
    "绝缘子脱落": "绝缘子",
    "绝缘子破损": "绝缘子",
    "绝缘子老化": "绝缘子",
    "绝缘子倾斜": "绝缘子",
    "绝缘子污秽": "绝缘子",
    "绝缘子灼伤": "绝缘子",
    "绝缘子雷击": "绝缘子",
    "釉面剥落": "绝缘子",
    "绑带松脱": "绝缘子",
    "绝缘子绑带安装不规范": "绝缘子",
    "金具锈蚀": "金具",
    "销钉缺失": "金具",
    "销钉退出": "金具",
    "销钉安装不规范": "金具",
    "螺母松动": "金具",
    "螺母缺失": "金具",
    "防震锤锈蚀": "金具",
    "防震锤脱落": "金具",
    "导线缠绕": "导地线",
    "导线脱落": "导地线",
    "导线悬挂异物": "导地线",
    "导线断股": "导地线",
    "导线松股": "导地线",
    "导线固定不牢": "导地线",
    "地线悬挂异物": "导地线",
    "绝缘保护壳破损": "附属设施",
    "绝缘保护壳缺失": "附属设施",
    "标识牌脱落": "附属设施",
    "通道树障": "通道",
    "通道施工": "通道",
    "变压器漏油": "变压器",
    "变压器渗油": "变压器",
    "避雷器雷击": "避雷器",
    "避雷器破损": "避雷器",
    "线耳脱落": "避雷器",
    "避雷器连接线脱落": "避雷器",
}

bugTypeCountMap = {}
total_statis_map = {}
image_index = {}
bug_type_map = {1: "危急", 2: "严重", 3: "一般"}


# statis_add_row 汇总行数小于图片数量时，添加行
def statis_add_row(table, images):
    tr = len(table.rows)
    while tr - 1 < len(images):
        new_row = deepcopy(table.rows[-1])
        table.rows[-1]._tr.addnext(new_row._element)  # 在最后一行后面添加
        table.rows[-1]._tr.addprevious(new_row._element)
        tr += 1


# 缺陷描述和缺陷类别不匹配时，模糊匹配
def fuzzy_match(bug_detail):
    if "绝缘子" in bug_detail:
        return "绝缘子"
    if any(keyword in bug_detail for keyword in ["杆塔", "塔基", "塔头", "塔顶"]):
        return "基础"
    if any(keyword in bug_detail for keyword in ["金具", "销钉", "螺母"]):
        return "金具"
    if any(keyword in bug_detail for keyword in ["保护壳", "标识牌"]):
        return "附属设施"
    if any(keyword in bug_detail for keyword in ["地线", "导线"]):
        return "导地线"
    if "避雷器" in bug_detail:
        return "避雷器"
    if "变压器" in bug_detail:
        return "变压器"
    if "通道" in bug_detail:
        return "通道"
    debug_log(f"{bug_detail} 未匹配到缺陷类别", 2)
    return ""


def get_bug_type(bug_reason):
    bugType = bugMap.get(bug_reason, "")
    if len(bugType) == 0:
        bugType = fuzzy_match(bug_reason)
        if len(bugType) > 0:
            debug_log(
                f"缺陷描述:\033[32m[{bug_reason}]\033[m 未匹配到缺陷类别,已模糊匹配为 >>> \033[32m{bugType}\033[m",
                1,
            )
    return bugType


# statis_add_table 汇总数据写入
def set_detail_statis(table, images, bug_type):
    debug_log(f"开始处理 {bug_type_map.get(bug_type,'')}缺陷汇总表")
    c = 1
    for i in images:

        picName, picType = i.split(".")
        bugLevel = imageBugLevelMap.get(i, "")

        bugType = get_bug_type(imageBugReasonMap.get(i, ""))
        # 汇总数据
        if len(bugType) > 0:
            update_bug_type_count(bugType, bugLevel)

        update_cell(table, c, 0, str(c))
        update_cell(table, c, 1, picName[:-3])
        update_cell(table, c, 2, bugType)
        update_cell(table, c, 3, bugLevel)
        c += 1
    debug_log("一般缺陷汇总表 写入完成")


# 更新单元格
def update_cell(table, row_idx, col_idx, text):
    cell = table.cell(row_idx, col_idx)
    cell.text = text
    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


# 设置单元格居中
def cell_set_center(cell):
    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


# 复制单元格字体格式
def copy_cell_font_name(cell1, cell2):
    cell1.paragraphs[0].runs[0].font.name = cell2.paragraphs[0].runs[0].font.name


# 复制单元格字体大小
def copy_cell_font_size(cell1, cell2):
    cell1.paragraphs[0].runs[0].font.size = cell2.paragraphs[0].runs[0].font.size


def update_bug_type_count(bugType, bugLevel):
    # bugLevelCountMap = bugTypeCountMap.get(bugType, {})
    # bugLevelCountMap[bugLevel] = bugLevelCountMap.get(bugLevel, 0) + 1
    # bugLevelCountMap["合计"] = bugLevelCountMap.get("合计", 0) + 1
    # bugTypeCountMap[bugType] = bugLevelCountMap

    # bugLevelCountMap2 = bugTypeCountMap.get("合计", {})
    # bugLevelCountMap2[bugLevel] = bugLevelCountMap2.get(bugLevel, 0) + 1
    # bugLevelCountMap2["合计"] = bugLevelCountMap2.get("合计", 0) + 1
    # bugTypeCountMap["合计"] = bugLevelCountMap2

    bugTypeCountMap.setdefault(bugType, {}).update(
        {
            bugLevel: bugTypeCountMap.get(bugType, {}).get(bugLevel, 0) + 1,
            "合计": bugTypeCountMap.get(bugType, {}).get("合计", 0) + 1,
        }
    )
    bugTypeCountMap.setdefault("合计", {}).update(
        {
            bugLevel: bugTypeCountMap.get("合计", {}).get(bugLevel, 0) + 1,
            "合计": bugTypeCountMap.get("合计", {}).get("合计", 0) + 1,
        }
    )


# 每个图片插入数据到一个表格
def deal_table(table, pic):
    picName, picType = pic.split(".")
    p1, p2, p3, p4 = picName.split("_")
    update_cell(table, 1, 0, p1)
    update_cell(table, 1, 1, p2)
    update_cell(table, 1, 2, p4)
    update_cell(table, 2, 1, p3)

    table_insert_image(table, 9, pic)
    close_up_pic = closeUpMap.get(picName, "")
    if len(close_up_pic) > 0:
        table_insert_image(table, 12, close_up_pic)
    debug_log(f"明细表 {picName} 处理完成")


# 插入图片
def table_insert_image(table, cell_idx, pic):
    image_path = "./pic/" + pic
    cell = table._cells[cell_idx]

    cell.paragraphs[0].add_run().add_picture(
        image_path, width=cell.width * 0.9, height=table.rows[3].height * 0.9
    )
    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


# 缺陷数量统计表
def bug_num_statis(table):
    bugLevelCountMap = bugTypeCountMap.get("合计", {})
    if len(bugLevelCountMap) == 0:
        return
    col = 0
    for key in table.rows[0].cells:
        count = bugLevelCountMap.get(key.text, 0)
        if count > 0:
            table.cell(1, col).text = str(count)
            table.cell(1, col).paragraphs[0].runs[0].font.name = statis_number_font
            if is_set_statis_number_size:
                copy_cell_font_size(table.cell(1, col), table.cell(0, col))
            cell_set_center(table.cell(1, col))
            total_statis_map[key.text] = count
        col += 1


# 缺陷类别统计表
def bug_type_statis(table):
    row = 0
    for rows in table.rows:
        bugLevelCountMap = bugTypeCountMap.get(rows.cells[0].text, {})
        if len(bugLevelCountMap) > 0 or row >= 2:
            col = 0
            for key in table.rows[1].cells:
                count = bugLevelCountMap.get(key.text, 0)
                if count > 0 or col >= 1:
                    table.cell(row, col).text = str(count)
                    table.cell(row, col).paragraphs[0].runs[
                        0
                    ].font.name = statis_number_font
                    if is_set_statis_number_size:
                        copy_cell_font_size(table.cell(row, col), table.cell(1, col))

                    cell_set_center(table.cell(row, col))
                    if key.text == "合计":
                        total_statis_map[rows.cells[0].text] = count
                col += 1
        row += 1


# 缺陷情况总览
def set_total_description(doc):
    total_description_paragraph = match_text_paragraph(doc, "本次现场巡检")
    if total_description_paragraph == 0:
        debug_log(" 定位缺陷情况总览模块失败......", 2)
        return False
    para = doc.paragraphs[total_description_paragraph]
    tpl = para.text
    font_size = para.runs[0].font.size
    content = tpl.format(
        total_bug=total_statis_map.get("合计", 0),
        weiji_bug=total_statis_map.get("危急", 0),
        yanzhong_bug=total_statis_map.get("严重", 0),
        yiban_bug=total_statis_map.get("一般", 0),
        bileiqi_bug=total_statis_map.get("避雷器", 0),
        bianyaqi_bug=total_statis_map.get("变压器", 0),
        daodixian_bug=total_statis_map.get("导地线", 0),
        fushu_bug=total_statis_map.get("附属设施", 0),
        jichu_bug=total_statis_map.get("基础", 0),
        jinjv_bug=total_statis_map.get("金具", 0),
        jueyuanzi_bug=total_statis_map.get("绝缘子", 0),
        tongdao_bug=total_statis_map.get("通道", 0),
    )
    debug_log(f"缺陷情况总览文字: {content}")
    doc.paragraphs[total_description_paragraph].text = content
    doc.paragraphs[total_description_paragraph].runs[0].font.size = font_size
    return True


# 匹配文字所在段落
def match_text_paragraph(doc, text):
    pi = 0
    for p in doc.paragraphs:
        if text in p.text:
            return pi
        pi += 1
    return 0


closeUpMap = {}
imageBugLevelMap = {}
imageBugReasonMap = {}
imageTowerMap = {}
imageRouteNameMap = {}
imageTypeMap = {}


def deal_close_up_image(pic):
    picName, picType = pic.split(".")
    if len(picName.split("_")) != 5:
        debug_log(
            f"图片名称不规范,不规范的图片为：\033[35m{picName}.{picType}\033[m ", 2
        )
        return False
    if len(picName.split("_特写")) != 2:
        debug_log(
            f"图片名称不规范,不规范的图片为：\033[35m{picName}.{picType}\033[m ", 2
        )
        return False
    closeUpName, _ = picName.split("_特写")
    closeUpMap[closeUpName] = pic
    return True


# 获取待处理的图片
def get_images():
    commonList = []
    criticalList = []
    emergencyList = []
    for root, dirs, pics in os.walk("./pic"):
        for pic in pics:
            picName, picType = pic.split(".")
            if len(picName.split("_")) != 4:
                if not deal_close_up_image(pic):
                    return
                continue
            route_name, tower_num, bug_reason, bug_level = picName.split("_")
            imageBugLevelMap[pic] = bug_level
            imageBugReasonMap[pic] = bug_reason
            imageTowerMap[pic] = tower_num
            imageRouteNameMap[pic] = route_name
            imageTypeMap[pic] = picType
            match bug_level:
                case "危急":
                    emergencyList.append(pic)
                case "严重":
                    criticalList.append(pic)
                case "一般":
                    commonList.append(pic)
    return emergencyList, criticalList, commonList


def get_statis_table(doc, index=1):
    i = 0
    sort = 1
    for t in doc.tables:
        if t._cells[1].text == "缺陷描述":
            if index == sort:
                return i
            sort += 1
        i += 1


def get_detail_table(doc, index=1):
    i = 0
    sort = 1
    for t in doc.tables:
        if t._cells[0].text == "线路名称":
            if index == sort:
                return i
            sort += 1
        i += 1


def match_detail_paragraph(doc, title):
    pi = 0
    for p in doc.paragraphs:
        if title in p.text:
            return pi
        pi += 1
    return 0


# def par_index(paragraph):
#     "Get the index of the paragraph in the document"
#     doc = paragraph._parent
#     # the paragraphs elements are being generated on the fly,
#     # they change all the time
#     # so in order to index, we must use the elements
#     l_elements = [p._element for p in doc.paragraphs]
#     return l_elements.index(paragraph._element)


def par_index(doc, para):
    i = 0
    for p in doc.paragraphs:
        debug_log(p.text, 2)
        if p == para:
            return i
        i += 1


def insert_paragraph_after(paragraph, text=None, style=None):
    """Insert a new paragraph after the given paragraph."""
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    if style is not None:
        new_para.style = style
    return new_para


def missing_table_num(doc, table_index, image_list):
    for i in range(table_index, table_index + len(image_list)):
        if i >= len(doc.tables):
            return table_index + len(image_list) - i
        table = doc.tables[i]
        if i > table_index and (table._cells[0].text != "线路名称"):
            return table_index + len(image_list) - i
    return 0


def add_missing_table(doc, tbl, paragraph_index, add_num):
    if add_num > 0:
        debug_log(f"危急部分缺少{add_num}个表格")
        for i in range(add_num):
            new_tbl = deepcopy(tbl)
            insert_paragraph_after(
                doc.paragraphs[paragraph_index + i],
            ).add_run().add_break(WD_BREAK.PAGE)
            doc.paragraphs[paragraph_index + i]._p.addnext(new_tbl)
        debug_log(f"<危急>部分插入{add_num}个表格成功")


def add_missing_rows(doc, table_index, image_list, bug_type):
    statis_table_rows = len(doc.tables[table_index].rows) - 1
    if statis_table_rows < len(image_list):
        statis_add_row(doc.tables[table_index], image_list)
        debug_log(f"{bug_type}缺陷汇总表插入{len(image_list)-statis_table_rows-1}行")


# 生成模板
def get_template(emergencyList, criticalList, commonList, fileName):
    # 实例化一个Document对象，相当于打开word软件，新建一个空白文件
    doc = Document(fileName)
    # tables = doc.tables  # 获取文档中所有表格对象的列表

    emergency_detail_table_index = get_detail_table(doc, 1)
    critical_detail_table_index = get_detail_table(doc, 2)
    common_detail_table_index = get_detail_table(doc, 3)
    if not (
        emergency_detail_table_index
        * critical_detail_table_index
        * common_detail_table_index
    ):
        debug_log("定位危急缺陷明细表失败", 2)
        return False
    emergency_statis_table_index = get_statis_table(doc, 1)
    critical_statis_table_index = get_statis_table(doc, 2)
    common_statis_table_index = get_statis_table(doc, 3)

    emergency_detail_paragraph = match_detail_paragraph(doc, "危急缺陷明细表")
    if emergency_detail_paragraph == 0:
        debug_log("定位危急缺陷明细表失败", 2)
        return False

    # 判断  汇总表行数是否足够
    add_missing_rows(doc, emergency_statis_table_index, emergencyList, "危急")
    add_missing_rows(doc, critical_statis_table_index, criticalList, "严重")
    add_missing_rows(doc, common_statis_table_index, commonList, "一般")

    tbl = doc.tables[emergency_detail_table_index]._tbl

    # 判断 <危急> 详情表数量是否足够
    addNum = missing_table_num(doc, emergency_detail_table_index, emergencyList)
    # <危急> 添加表格
    add_missing_table(doc, tbl, emergency_detail_paragraph, addNum)
    critical_detail_table_index += addNum
    common_detail_table_index += addNum

    # 判断 <严重> 详情表数量是否足够
    critical_detail_paragraph = match_detail_paragraph(doc, "严重缺陷明细表")
    if critical_detail_paragraph == 0:
        debug_log("严重缺陷明细表失败", 2)
        return False
    addNum = missing_table_num(doc, critical_detail_table_index, criticalList)
    # <严重> 添加表格
    add_missing_table(doc, tbl, critical_detail_paragraph, addNum)
    common_detail_table_index += addNum

    # 判断 <一般> 详情表数量是否足够
    common_detail_paragraph = match_detail_paragraph(doc, "一般缺陷明细表")
    if critical_detail_paragraph == 0:
        debug_log("一般缺陷明细表失败", 2)
        return False
    addNum = missing_table_num(doc, common_detail_table_index, commonList)
    add_missing_table(doc, tbl, common_detail_paragraph, addNum)

    doc.save("tpl.docx")
    debug_log("生成模板成功")
    return True


def deal_one_type_table(doc, table_index, iamge_list, bug_type):
    debug_log(f"开始处理 {bug_type_map.get(bug_type,'')}明细表")

    picIndex = 0
    for i in range(
        table_index + 1,
        table_index + 1 + len(iamge_list),
    ):
        table = doc.tables[i]
        pic = iamge_list[picIndex]
        deal_table(table, pic)
        picIndex += 1

    debug_log(f"{bug_type_map.get(bug_type,'')}明细表 处理完成")


# 处理数据
def deal(emergencyList, criticalList, commonList, fileName):
    # 处理数据
    doc = Document("tpl.docx")
    tables = doc.tables  # 获取文档中所有表格对象的列表
    emergency_statis_table_index = get_statis_table(doc, 1)
    critical_statis_table_index = get_statis_table(doc, 2)
    common_statis_table_index = get_statis_table(doc, 3)

    set_detail_statis(tables[emergency_statis_table_index], emergencyList, EMERGENCY)

    set_detail_statis(tables[critical_statis_table_index], criticalList, CRITICAL)

    set_detail_statis(tables[common_statis_table_index], commonList, COMMON)

    deal_one_type_table(doc, emergency_statis_table_index, emergencyList, EMERGENCY)
    deal_one_type_table(doc, critical_statis_table_index, criticalList, CRITICAL)
    deal_one_type_table(doc, common_statis_table_index, commonList, COMMON)

    bug_num_statis(tables[bug_num_table_index])
    debug_log("缺陷数量统计表 写入完成")
    bug_type_statis(tables[bug_type_table_index])
    debug_log("缺陷类别统计表 写入完成")

    if set_total_description(doc):
        debug_log("缺陷情况总览 写入完成")
    debug_log("处理结束，正在保存文件...")
    doc.save(fileName)
    debug_log("文件保存文件成功")


def debug_log(message, log_level=0):
    level_tips = ""
    match log_level:
        case 0:
            if not debug:
                return
            level_tips = "[INFO]   "
        case 1:
            level_tips = "\033[33m[WARNING]\033[m"
        case 2:
            level_tips = "\033[31m[ERROR]\033[m  "
    print(f"{level_tips}{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")


# 处理exif信息
def clearexif(imageList):

    def clear(image):
        f = Image.open(image)  # 你的图片文件
        f.save(image)  # 替换掉你的图片文件
        f.close()

    executor = ThreadPoolExecutor(ThreadPoolNum)
    all_tasks = [executor.submit(clear, imageList[i]) for i in range(len(imageList))]
    wait(all_tasks, return_when=ALL_COMPLETED)


templateFileName = "template.docx"  # 模板文件名称
bug_num_table_index = 4  # 缺陷数量表位置
bug_type_table_index = bug_num_table_index + 1  # 缺陷类别表位置
statis_number_font = "Times New Roman"  # 统计表数字字体
is_set_statis_number_size = True
debug = True  # 是否开启提示
warn = True  # 是否开启警告信息
ThreadPoolNum = 10
EMERGENCY = 1
CRITICAL = 2
COMMON = 3


if __name__ == "__main__":
    debug_log("程序开始运行...")
    tmpName = input("\033[32m请输入待生成的文件名称(回车确认):\033[m")
    if tmpName == "":
        tmpName = "res"
    fileName = f"{tmpName}.docx"
    emergencyList, criticalList, commonList = get_images()
    if len(commonList) + len(criticalList) > +len(emergencyList) > 0:
        if get_template(emergencyList, criticalList, commonList, templateFileName):
            clearexif(emergencyList)
            clearexif(criticalList)
            clearexif(commonList)
            deal(emergencyList, criticalList, commonList, fileName)
            debug_log(f"请查看 \033[32m{fileName}\033[m 文件")
    debug_log(f"程序运行结束！")


```
