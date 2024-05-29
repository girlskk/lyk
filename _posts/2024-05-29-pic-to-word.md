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


# statis_add_table 汇总行数小于图片数量时，添加行
def statis_add_table(table, images):
    tr = len(table.rows)
    while tr - 1 < len(images):
        new_row = deepcopy(table.rows[-1])
        table.rows[-1]._tr.addnext(new_row._element)  # 在最后一行后面添加
        table.rows[-1]._tr.addprevious(new_row._element)
        tr += 1


# 缺陷描述和缺陷类别不匹配时，模糊匹配
def fuzzy_mattch(bug_detail):
    if "绝缘子" in bug_detail:
        return "绝缘子"
    elif "杆塔" in bug_detail or "塔基" in bug_detail or "塔头" in bug_detail:
        return "基础"
    elif "金具" in bug_detail or "销钉" in bug_detail or "螺母" in bug_detail:
        return "金具"
    elif "保护壳" in bug_detail or "标识牌" in bug_detail:
        return "附属设施"
    elif "地线" in bug_detail or "导线" in bug_detail:
        return "导地线"
    elif "避雷器" in bug_detail:
        return "避雷器"
    elif "变压器" in bug_detail:
        return "变压器"
    elif "通道" in bug_detail:
        return "通道"
    if warn:
        print(f"{bug_detail} 未匹配到缺陷类别")
    return ""


# statis_add_table 汇总数据写入
def set_detail_statis(table, images):
    c = 1
    for i in images:
        picName, picType = i.split(".")
        try:
            table.cell(c, 1).text = picName[:-3]
            copy_cell_font_size(table.cell(c, 1), table.cell(c, 0))
            copy_cell_font_size(table.cell(c, 1), table.cell(c, 0))
            cell_set_center(table.cell(c, 1))
        except:
            if warn:
                print(f"{picName}:获取缺陷描述失败")
        try:
            k = picName[:-3].split("_")[-1]
            bugType = bugMap.get(k, "")
            bugLevel = picName[-2:]
            if len(bugType) == 0:
                bugType = fuzzy_mattch(k)
                if len(bugType) > 0:
                    debugLog(f"{k}未匹配到缺陷类别,已模糊匹配为{bugType}")
            table.cell(c, 2).text = bugType
            copy_cell_font_size(table.cell(c, 2), table.cell(c, 0))
            copy_cell_font_size(table.cell(c, 2), table.cell(c, 0))
            cell_set_center(table.cell(c, 2))

            # 汇总数据
            if len(bugType) > 0:
                bugLevelCountMap = bugTypeCountMap.get(bugType, {})
                bugLevelCountMap[bugLevel] = bugLevelCountMap.get(bugLevel, 0) + 1
                bugLevelCountMap["合计"] = bugLevelCountMap.get("合计", 0) + 1
                bugTypeCountMap[bugType] = bugLevelCountMap

                bugLevelCountMap2 = bugTypeCountMap.get("合计", {})
                bugLevelCountMap2[bugLevel] = bugLevelCountMap2.get(bugLevel, 0) + 1
                bugLevelCountMap2["合计"] = bugLevelCountMap2.get("合计", 0) + 1
                bugTypeCountMap["合计"] = bugLevelCountMap2
        except:
            if warn:
                print(f"{picName}:获取缺陷类别失败")
        try:
            table.cell(c, 3).text = picName[-2:]
            copy_cell_font_size(table.cell(c, 3), table.cell(c, 0))
            copy_cell_font_size(table.cell(c, 3), table.cell(c, 0))
            cell_set_center(table.cell(c, 3))
        except:
            if warn:
                print(f"{picName}:获取缺陷等级失败")

        table.cell(c, 0).text = str(c)
        copy_cell_font_size(table.cell(c, 0), table.cell(c, 3))
        copy_cell_font_size(table.cell(c, 0), table.cell(c, 3))
        cell_set_center(table.cell(c, 0))
        c += 1


# 设置表格中单元格文字和图片
def set_cell_text(cell0, cell, text):
    cell.text = text
    copy_cell_font_name(cell, cell0)
    copy_cell_font_size(cell, cell0)
    cell_set_center(cell)
    return cell


# 每个图片插入数据到一个表格
def deal_table(table, pic):
    cells = table._cells
    picName, picType = pic.split(".")
    p1, p2, p3, p4 = picName.split("_")
    set_cell_text(cells[0], cells[4], p1)
    set_cell_text(cells[0], cells[5], p2)
    # set_cell_text(cells[0],cells[6],p2)
    set_cell_text(cells[0], cells[7], p4)
    set_cell_text(cells[0], cells[9], p3)
    image_path = "./pic/" + pic

    cells[12].paragraphs[0].add_run().add_picture(
        image_path, width=cells[12].width * 0.9, height=table.rows[3].height * 0.9
    )
    cells[12].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 左右居中
    debugLog(f"明细表 {picName} 处理完成")


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


# 复制单元格字体格式
def copy_cell_font_name(cell1, cell2):
    cell1.paragraphs[0].runs[0].font.name = cell2.paragraphs[0].runs[0].font.name


# 复制单元格字体大小
def copy_cell_font_size(cell1, cell2):
    cell1.paragraphs[0].runs[0].font.size = cell2.paragraphs[0].runs[0].font.size


# 设置单元格左右居中
def cell_set_center(cell1):
    # 左右居中
    cell1.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


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
        print("<worning> 定位缺陷情况总览模块失败......")
        return
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
    debugLog(f"缺陷情况总览文字: {content}")
    doc.paragraphs[total_description_paragraph].text = content
    doc.paragraphs[total_description_paragraph].runs[0].font.size = font_size


# 匹配文字所在段落
def match_text_paragraph(doc, text):
    pi = 0
    for p in doc.paragraphs:
        if text in p.text:
            return pi
        pi += 1
    return 0


# 获取待处理的图片
def get_images():
    commonList = []
    otherList = []
    for root, dirs, pics in os.walk("./pic"):
        for pic in pics:
            picName, picType = pic.split(".")
            if len(picName.split("_")) != 4:
                print(picName, "名称不规范")
                return
            level = picName.split("_")[-1]
            if level == "一般":
                commonList.append(pic)
            else:
                otherList.append(pic)
    return commonList, otherList


# 生成模板
def get_template(commonList, otherList, fileName):
    # 实例化一个Document对象，相当于打开word软件，新建一个空白文件
    doc = Document(fileName)
    tables = doc.tables  # 获取文档中所有表格对象的列表
    level1_table_index = first_table_index + 3

    level1_paragraph_index = match_text_paragraph(doc, "危急、严重缺陷明细表")
    if level1_paragraph_index == 0:
        print("定位危急、严重缺陷明细表失败")
        return
    tbl = doc.tables[level1_table_index]._tbl

    # 判断 <危急，严重> 详情表数量是否足够
    addNum = 0
    level2_table_index = level1_table_index + 2  # 第一个一般缺陷明细表位置
    for i in range(level1_table_index, level1_table_index + len(otherList)):
        table = tables[i]
        if i > level1_table_index and (table._cells[0].text != "线路名称"):
            level2_table_index = i + 1
            addNum = level1_table_index + len(otherList) - i
            break
    # 判断 <危急，严重> 汇总表行数是否足够
    statis_table_rows = len(tables[level1_table_index - 1].rows) - 1
    if statis_table_rows < len(otherList):
        statis_add_table(tables[level1_table_index - 1], otherList)
        debugLog(f"危急、严重缺陷汇总表插入{len(otherList)-statis_table_rows-1}行")
    # 判断<一般>汇总表行数是否足够
    statis_table_rows = len(tables[level2_table_index - 1].rows) - 1
    if statis_table_rows < len(commonList):
        statis_add_table(tables[level2_table_index - 1], commonList)
        debugLog(f"一般缺陷汇总表插入{len(commonList)-statis_table_rows}行")
    # <危急，严重> 添加表格
    if addNum > 0:
        debugLog("危急，严重部分缺少", addNum, "个表格")
        for i in range(addNum):
            new_tbl = deepcopy(tbl)
            doc.paragraphs[level1_paragraph_index]._p.addnext(new_tbl)
        debugLog("<危急，严重>部分插入", addNum, "个表格成功")

    # 判断 <一般> 详情表数量是否足够
    addNum = 0
    for i in range(level2_table_index, level2_table_index + len(commonList)):
        if i > len(tables) - 1:
            addNum = level2_table_index + len(commonList) - i
            break
        table = tables[i]
        if i > level2_table_index and (table._cells[0].text != "线路名称"):
            addNum = level2_table_index + len(commonList) - i
            break
    if addNum > 0:
        debugLog("<一般>部分缺少", addNum, "个表格")
        for i in range(addNum):
            new_tbl = deepcopy(tbl)
            doc.paragraphs[-1]._p.addnext(new_tbl)
        debugLog("<一般>部分插入", addNum, "个表格成功")

    doc.save("tpl.docx")
    debugLog("生成模板成功")


# 处理数据
def deal(commonList, otherList, fileName):
    # 处理数据
    doc = Document("tpl.docx")
    tables = doc.tables  # 获取文档中所有表格对象的列表
    level1_table_index = first_table_index + 3

    debugLog("开始处理 危急、严重缺陷汇总表")
    set_detail_statis(tables[first_table_index + 2], otherList)
    debugLog("危急、严重缺陷汇总表 写入完成")

    debugLog("开始处理 危急、严重缺陷明细表")
    picIndex = 0
    level2_table_index = level1_table_index + 2
    for i in range(level1_table_index, level1_table_index + len(otherList)):
        table = tables[i]
        level2_table_index = i + 2
        pic = otherList[picIndex]
        deal_table(table, pic)
        picIndex += 1
    debugLog("危急、严重缺陷明细表 处理完成")

    debugLog("开始处理 一般缺陷汇总表")
    set_detail_statis(tables[level2_table_index - 1], commonList)
    debugLog("一般缺陷汇总表 写入完成")

    debugLog("开始处理 一般缺陷明细表")
    picIndex = 0
    for i in range(level2_table_index, level2_table_index + len(commonList)):
        table = tables[i]
        pic = commonList[picIndex]
        deal_table(table, pic)
        picIndex += 1
    debugLog("一般缺陷明细表处理完成")

    bug_num_statis(tables[first_table_index])
    debugLog("缺陷数量统计表 写入完成")
    bug_type_statis(tables[first_table_index + 1])
    debugLog("缺陷类别统计表 写入完成")

    set_total_description(doc)
    debugLog("缺陷情况总览 写入完成")
    doc.save(fileName)


def debugLog(log):
    if debug:
        print(log)


# 处理exif信息
def clearexif(imageList):

    def clear(image):
        f = Image.open(image)  # 你的图片文件
        f.save(image)  # 替换掉你的图片文件
        f.close()

    executor = ThreadPoolExecutor(ThreadPoolNum)
    all_tasks = [executor.submit(clear, i) for i in range(imageList)]
    wait(all_tasks, return_when=ALL_COMPLETED)


templateFileName = "test.docx"  # 模板文件名称
first_table_index = 4  # 第一个表位置
statis_number_font = "Times New Roman"  # 统计表数字字体
is_set_statis_number_size = True
debug = True  # 是否开启提示
warn = True  # 是否开启警告信息
ThreadPoolNum = 10

if __name__ == "__main__":
    print("程序开始运行...")
    tmpName = input("请输入待生成的文件名称(回车确认):")
    if tmpName == "":
        tmpName = "res"
    fileName = f"{tmpName}.docx"
    commonList, otherList = get_images()
    get_template(commonList, otherList, templateFileName)
    clearexif(commonList)
    clearexif(otherList)
    deal(commonList, otherList, fileName)
    print(f"程序运行结束！请查看<{fileName}>文件")

```
