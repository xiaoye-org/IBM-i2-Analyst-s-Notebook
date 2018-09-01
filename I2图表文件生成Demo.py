# coding: utf-8

# Func：调用COM接口生成I2的图表文件
# Author：chenggh
# His：2018-8-30 Created

import win32com.client
import numpy as np
import os

# 创建I2 Notebook对象
objApp = win32com.client.Dispatch('LinkNotebook.Application.7')
# 是否显示I2软件窗口
objApp.Visible = True 

# 添加一个空白图表
objApp.Charts.Add('')

# 获取当前图表对象
objChart = objApp.Charts.CurrentChart
# 获取当前图表的实体（节点）集合
objEntityTypeColl = objChart.EntityTypes

# 添加一个图表中节点类型 并指定图标为Person
objEntityType = objEntityTypeColl.Add(u'客户','Person',0)

# 添加节点
## 获取当前图表Icon类型节点的样式
objIconStyle = objChart.CurrentIconStyle
objIconStyle.Type = objEntityType
## 添加两个节点 函数原型为LNChart.CreateIcon (Style, X, Y, Label, Identity)
objIcon1 = objChart.CreateIcon(objIconStyle,150,150,u'客户1','001')
## 节点2 加框突出显示
objIconStyle.SetSubItemVisible(8,True)
objIcon2 = objChart.CreateIcon(objIconStyle,350,150,u'客户2','002')

# 添加链接
## 获取当前图表链接的样式
objLinkStyle = objChart.CurrentLinkStyle
## 设置链接的箭头
objLinkStyle.ArrowStyle = 1 # 0-ArrowStyle.ArrowNone  1-ArrowStyle.ArrowOnHead
## 创建链接 函数原型为LNChart.CreateLink (Style, EndFrom, EndTo, Label)
objLink1 = objChart.CreateLink(objLinkStyle,objIcon1,objIcon2,u'1000.02')

# 添加图表属性 给当前链接加一个属性：交易笔数
objAttributeClassColl = objChart.AttributeClasses
objAttributeClass = objAttributeClassColl.Find(u"交易笔数")
if objAttributeClass is None:
   objAttributeClass = objChart.CreateAttributeClass(u"交易笔数",1, "", True, True, True)
## 将当前链接的交易笔数 如设置为4笔 
objLink1.SetAttributeValue(objAttributeClass,4.0)

# 判断文件是否存在后 保存图表文件
## 需要注意的是win32com只支持绝对路径文件名
lfilename = u'f:\I2测试图表.anb'
if os.path.isfile(lfilename):
    os.remove(lfilename)
## 保存图表文件
objChart.SaveChart(lfilename)

# 关闭当前图表 退出I2 释放COM对象
objApp.Charts.CloseChart(objChart,0)
objApp.Quit()
objApp = None

