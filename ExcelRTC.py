#!/bin/env python
# -*- encoding: utf-8 -*-

##
#   @file ExcelRTC.py
#   @brief PortPointControl Component

import win32com
import pythoncom
import pdb
from win32com.client import *
import pprint
import datetime
import msvcrt

import thread


import optparse
import sys,os,platform
import re
import time
import random
import commands
import math



import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming

from PyQt4 import QtCore, QtGui
from MainWindow import MainWindow

from CalcControl import *

excel_comp = None



excelcontrol_spec = ["implementation_id", "ExcelControl",
                  "type_name",         "ExcelControl",
                  "description",       "Excel Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.file_path", "NewFile",
                  "conf.default.SlideNumberInRelative", "1",
                  "conf.default.SlideFileInitialNumber", "0",
                  "conf.__widget__.file_path", "text",
                  "conf.__widget__.SlideNumberInRelative", "radio",
                  "conf.__widget__.SlideFileInitialNumber", "spin",
                  "conf.__constraints__.SlideNumberInRelative", "(0,1)",
                  "conf.__constraints__.SlideFileInitialNumber", "0<=x<=1000",
                  ""]



##
# @brief 文字を行番号に変換
# @patam m_str 変換前の文字列
# @return 行番号
#
def convertStrToVal(m_str):
  if len(m_str) > 0:
    m_c = ord(m_str[0]) - 64

    if len(m_str) == 1:
      return m_c
    else:
      if ord(m_str[1]) < 91 and ord(m_str[1]) > 64:
        m_c2 = ord(m_str[1]) - 64
        return m_c2 + m_c*26
      else:
        return m_c


##
# @class ExcelPortObject
# @brief 追加するポートのクラス
#


class ExcelPortObject(CalcDataPort.CalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)
        
    
    ##
    # @brief 
    # @param self 
    # @param m_cal ExcelRTC
    def update_cellName(self, m_cal):

        cell, sheet, m_len = m_cal.m_excel.getCell(self._col, self._row, self._sn, self._length, False)
        if cell:
          self.update_cellNameSub(cell, m_len)

        

    
        
    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSingle(self, cell, m_len):

        cell.Value2 = self._name
        

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSeq(self, cell, m_len):
        v = []
            
        for i in range(0, m_len):
          v.append([self._name + ":" + str(i)])

        cell.Value2 = v
        

    

    ##
    # @brief 
    # @param self 
    # @param b データ名
    # @param count カウンター
    # @param m_len 行の範囲
    # @param cell セルオブジェクト
    # @return 
    def input_cellNameEx(self, b, count, m_len, cell):

        v = []
        for i in range(0,m_len):
          if i == count[0]:
            v.append(b)
          else:
            v.append(cell.Value2[0][i])
        cell.Value2 = v
        
                    
        count[0] += 1
        if count[0] >= m_len:
            return False
        return True

        
            

    

    ##
    # @brief 
    # @param self 
    # @param m_cal ExcelRTC
    def getCell(self, m_cal):
        return m_cal.m_excel.getCell(self._num, self._row, self._sn, self._length)
        

    

    
                    
    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param sheet シートオブジェクト
    # @param m_cal ExcelRTC
    def putOut(self, cell, sheet, m_cal):
        
        m_string = CalcDataPort.DataType.String
        m_value = CalcDataPort.DataType.Value
        
        cell.Interior.ColorIndex = 6
        
        if  self._length == "":
          val = cell.Value2
        else:
          val = []
          for i in cell.Value2[0]:
            val.append(i)

        

        if self._num > 1 and self.state == True:
          
          
          cell2, sheet2, m_len2 = m_cal.m_excel.getCell(self._num-1, self._row, self._sn, self._length)
          
          cell2.Interior.ColorIndex = 0
          
        


        return val

##
# @class ExcelInPort
# @brief
#

class ExcelInPort(CalcDataPort.CalcInPort, ExcelPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcInPort.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        ExcelPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        ExcelPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        ExcelPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return ExcelPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return ExcelPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return ExcelPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcInPort.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcInPort.update_cellNameSub(self, cell, m_len)
        
    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        cell, sheet, m_len = self.getCell(m_cal)
        if cell != None:
              
          cell.Value2 = b
          if self.state:
            self._num = self._num + 1


        
                    
##
# @class ExcelInPortSeq
# @brief 
class ExcelInPortSeq(CalcDataPort.CalcInPortSeq, ExcelPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcInPortSeq.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        ExcelPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        ExcelPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        ExcelPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return ExcelPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return ExcelPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return ExcelPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcInPortSeq.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcInPortSeq.update_cellNameSub(self, cell, m_len)

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        cell, sheet, m_len = self.getCell(m_cal)
        if cell != None:
          v = []
          for j in range(0, len(b)):
            if m_len > j:
              v.append(b[j])
              
          cell.Value2 = v
          if self.state:
              self._num = self._num + 1
        
##
# @class ExcelInPortEx
# @brief 
class ExcelInPortEx(CalcDataPort.CalcInPortEx, ExcelPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcInPortEx.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)
        self.v = []
        
    def update_cellName(self, m_cal):
        ExcelPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        ExcelPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        ExcelPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return ExcelPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return ExcelPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return ExcelPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcInPortEx.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcInPortEx.update_cellNameSub(self, cell, m_len)
    

    ##
    # @brief
    # @param self 
    # @param b データ
    # @param count カウンター
    # @param m_len 行の範囲
    # @param cell セルオブジェクト
    # @param d_type データタイプ
    def putDataEx(self, b, count, m_len, cell, d_type):
        self.v.append(b)
                            
        count[0] += 1
        
        if count[0] >= m_len:
            
            cell.Value2 = self.v
            self.v = []
            return False
        return True

        
##
# @class ExcelOutPort
# @brief 
class ExcelOutPort(CalcDataPort.CalcOutPort, ExcelPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcOutPort.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        ExcelPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        ExcelPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        ExcelPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return ExcelPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return ExcelPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return ExcelPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcOutPort.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcOutPort.update_cellNameSub(self, cell, m_len)
    

##
# @class ExcelOutPortSeq
# @brief 
class ExcelOutPortSeq(CalcDataPort.CalcOutPortSeq, ExcelPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcOutPortSeq.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        ExcelPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        ExcelPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        ExcelPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return ExcelPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return ExcelPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return ExcelPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcOutPortSeq.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcOutPortSeq.update_cellNameSub(self, cell, m_len)


##
# @class ExcelOutPortEx
# @brief 
#
class ExcelOutPortEx(CalcDataPort.CalcOutPortEx, ExcelPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcOutPortEx.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        ExcelPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        ExcelPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        ExcelPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return ExcelPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return ExcelPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return ExcelPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcOutPortEx.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcOutPortEx.update_cellNameSub(self, cell, m_len)
    def putDataEx(self, count, val, d_type):
        return CalcDataPort.CalcOutPortEx.putDataEx(self, count, val, d_type)



##
# @class ExcelControl
# @brief Excelを操作するためのRTCのクラス
#

class ExcelControl(CalcControl):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    CalcControl.__init__(self, manager)

    global excel_comp
    excel_comp = self

    
    
    prop = OpenRTM_aist.Manager.instance().getConfig()
    fn = self.getProperty(prop, "excel.filename", "")
    self.m_excel = ExcelObject()
    if fn != "":
      str1 = [fn]
      OpenRTM_aist.replaceString(str1,"/","\\")
      fn = os.path.abspath(str1[0])
    self.m_excel.Open(fn)

    self.conf_filename = ["NewFile"]

    self.m_CalcInPort = ExcelInPort
    self.m_CalcInPortSeq = ExcelInPortSeq
    self.m_CalcInPortEx = ExcelInPortEx

    self.m_CalcOutPort = ExcelOutPort
    self.m_CalcOutPortSeq = ExcelOutPortSeq
    self.m_CalcOutPortEx = ExcelOutPortEx
    
    
    
    
    
    return

  ##
  # @brief rtc.confの設定を取得する関数
  #
  def getProperty(self, prop, key, value):
        
        if  prop.findNode(key) != None:
            
            value = prop.getProperty(key)
        return value

  ##
  # @brief コンフィギュレーションパラメータが変更されたときに呼び出される関数
  # @param self 
  #
  def configUpdate(self):
      CalcControl.configUpdate(self)
      return
      """self._configsets.update("default","file_path")
      str1 = [self.conf_filename[0]]
      OpenRTM_aist.replaceString(str1,"/","\\")
      sfn = str1[0]
      tfn = os.path.abspath(sfn)
      if sfn == "NewFile":
        self.m_excel.Open("")
      else:
        
        self.m_excel.initCom()
        self.m_excel.Open(tfn)"""
        #self.m_excel.closeCom()


  def addActionLock(self):
    return
    """tid = str(thread.get_ident())
    self.m_excel.comObjects[tid].xlApplication.ScreenUpdating = False"""

  def removeActionLock(self):
    return
    """tid = str(thread.get_ident())
    self.m_excel.comObjects[tid].xlApplication.ScreenUpdating = True"""

  def setCellColor(self, op):
    pass


  def onInitialize(self):
    CalcControl.onInitialize(self)
    
    
    
    return RTC.RTC_OK
    
  
  def onActivated(self, ec_id):
    
    #self.m_excel.initCom()
    
    CalcControl.onActivated(self, ec_id)
    
    

    #self.file = open('text3.txt', 'w')

    
    
    
    return RTC.RTC_OK

  def onDeactivated(self, ec_id):
    
    CalcControl.onDeactivated(self, ec_id)
    
    #self.file.close()
    return RTC.RTC_OK


  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    
    CalcControl.onExecute(self, ec_id)
    #cell,sheet,m_len = self.m_excel.getCell(1, "A", "Sheet1", "C")
    #cell.Value2 = [10, 10, 11]

    return RTC.RTC_OK

  
  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def on_shutdown(self, ec_id):
      CalcControl.on_shutdown(self, ec_id)
      global excel_comp
      excel_comp = None
      return RTC.RTC_OK


  
  
class ExcelComObject:
   def __init__(self,xlApplication,xlWorkbooks,xlWorkbook,xlWorksheets,xlWorksheet):  
      self.xlApplication = xlApplication
      self.xlWorkbooks = xlWorkbooks
      self.xlWorkbook = xlWorkbook
      self.xlWorksheets = xlWorksheets
      self.xlWorksheet = xlWorksheet

  
##
# @class ExcelObject
# @brief Excelを操作するクラス
#
class ExcelObject:

    ##
    # @brief コンストラクタ
    # @param self 
    #
    def __init__(self):
        self.filename = " "
        
        """self.xlApplication = None
        self.xlWorkbooks = None
        self.xlWorkbook = None
        self.xlWorksheets = None
        self.xlWorksheet = {}"""
        self.comObjects = {}
        

        self.thread_xlApplication = None
        self.thread_xlWorkbooks = None
        self.thread_xlWorkbook = None
        self.thread_xlWorksheets = None

        self.t_xlApplication = None
        self.t_xlWorkbooks = None
        self.t_xlWorkbook = None
        self.t_xlWorksheets = None
        self.t_xlWorksheet = {}
        

        self.red = 255
        self.green = 255
        self.blue = 0


    ##
    # @brief 
    # @param self
    # @param r
    # @param g
    # @param b 
    #
    def setColor(self, r, g, b):
      self.red = r
      self.green = g
      self.blue = b

    def resetCellColor(c, l, sn, elen):
      if sn in self.xlWorksheet:
        ws = self.xlWorksheet[sn]
        t_l = convertStrToVal(l)
        t_leng = convertStrToVal(elen)

        if c > 0 and t_l > 0 and t_leng >= t_l:
          c1 = ws.Cells(c,t_l)
          c2 = ws.Cells(c,t_leng)

          mr = ws.Range(c1,c2)

          mr.Interior.ColorIndex = 0

    def saveRTC(self, sf):
      if "保存用" in self.xlWorksheet:
        ws = self.xlWorksheet["保存用"]
        for i in range(0, len(sf)):
          c1 = ws.Cells(1+i,1)
          c1.Value2 = sf[i]

    def loadRTC(self):
      if "保存用" in self.xlWorksheet:
        ws = self.xlWorksheet["保存用"]
        v = []
        for i in range(0, 100):
          c1 = ws.Cells(1+i,1)
          try:
            tmp = c1.Text.encode("utf-8")
            if tmp == "":
              return v
            else:
              v.append(tmp)
          except:
            return v


    def getCell(self, c, l, sn, leng, mt = True):

      xlWorksheet = self.t_xlWorksheet
      
      
      
      if mt == True:
        self.initCom()
        tid = str(thread.get_ident())
        xlWorksheet = self.comObjects[tid].xlWorksheet

      
        
        
      
      if sn in xlWorksheet:
        
        ws = xlWorksheet[sn]
        v = []
        
        t_l = convertStrToVal(l)
        
        t_leng = t_l
        
        try:
          if leng == "":
            t_leng = t_l
          else:
            t_leng = convertStrToVal(leng)
        except:
          t_leng = t_l
        
        if t_l > t_leng:
          t_leng = t_l
        
        if c > 0 and t_l > 0 and t_leng >= t_l:

          c1 = ws.Cells(c,t_l)
          c2 = ws.Cells(c,t_leng)

          mr = ws.Range(c1,c2)
          
          
          return mr,ws,t_leng-t_l+1
      return None,None,None



    def setCellValue(self, c, l, sn, state, v):
      if sn in self.t_xlWorksheet:
        ws = self.t_xlWorksheet[sn]
        t_l = convertStrToVal(l)
        mnum = len(v)
        if not state:
          mnum = 1
        for i in range(0, mnum):
          if c+i > 0 and t_l > 0:
            c1 = ws.Cells(c+i,t_l)
            c2 = ws.Cells(c+i,t_l+len(v[i])-1)

            mr = ws.Range(c1,c2)

            mr.Value2 = v[i]


    def getCellValue(self, c, l, sn, leng):
      if sn in self.t_xlWorksheet:
        ws = self.t_xlWorksheet[sn]
        v = []
        t_l = convertStrToVal(l)
        t_leng = t_l
        try:
          t_leng = convertStrToVal(leng)
        except:
          t_leng = t_l
          
        if t_l > t_leng:
          t_leng = t_l

        if c > 0 and t_l > 0 and t_leng >= t_l:

          if c > 1:
            c1 = ws.Cells(c-1,t_l)
            c2 = ws.Cells(c-1,t_leng)

            mr = ws.Range(c1,c2)

            mr.Interior.ColorIndex = 0

          c1 = ws.Cells(c,t_l)
          c2 = ws.Cells(c,t_leng)

          mr = ws.Range(c1,c2)

          mr.Interior.ColorIndex = 6

        for i in range(0, t_leng-t_l+1):
          if c > 0 and t_l+i > 0:
            c1 = ws.Cells(c,t_l+i)
            try:
              v.append(c1.Value2)
            except:
              v.append(0)

        return v

        
    

    ##
    # @brief 
    # @param self 
    # @param xlWorksheets 
    #
    def setSheet(self, xlWorksheets, xlWorksheet):
      count = xlWorksheets.Count

      

      for i in range(1, count+1):
          item = xlWorksheets.Item(i)
          xlWorksheet[item.Name.encode("utf-8")] = item

      
          
    
    ##
    # @brief 
    # @param self 
    #
    def preInitCom(self):
        self.thread_xlApplication = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_xlApplication)
        self.thread_xlWorkbooks = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_xlWorkbooks)
        self.thread_xlWorkbook = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_xlWorkbook)
        self.thread_xlWorksheets = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_xlWorksheets)

    ##
    # @brief 
    # @param self
    #
    def initCom(self):
        tid = str(thread.get_ident())
        
        #if self.xlApplication == None:
        if tid in self.comObjects:
          pass
        else:
          pythoncom.CoInitialize()
          
          xlApplication = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_xlApplication, pythoncom.IID_IDispatch))
          
          xlWorkbooks = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_xlWorkbooks, pythoncom.IID_IDispatch))
          xlWorkbook = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_xlWorkbook, pythoncom.IID_IDispatch))
          xlWorksheets = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_xlWorksheets, pythoncom.IID_IDispatch))
          xlWorksheet = {}
          
          self.setSheet(xlWorksheets, xlWorksheet)
          
          self.comObjects[tid] = ExcelComObject(xlApplication,xlWorkbooks,xlWorkbook,xlWorksheets,xlWorksheet)
          
    ##
    # @brief 
    # @param self
    #
    def closeCom(self):
        pythoncom.CoUninitialize()

    ##
    # @brief Excelファイルを開く関数
    # @param self
    # @param fn ファイルパス
    #
    def Open(self, fn):
        if self.filename == fn:
            return
        self.filename = fn

        

        try:
            
            if self.t_xlApplication == None:
              t_xlApplication = win32com.client.Dispatch("Excel.Application")
            else:
              t_xlApplication = self.t_xlApplication
            
            
            t_xlApplication.Visible = True
            try:
                t_xlWorkbooks = t_xlApplication.Workbooks
                

                try:
                    
                    t_xlWorkbook = None
                    if self.filename == "":
                        t_xlWorkbook = t_xlWorkbooks.Add()
                        
                    else:
                        t_xlWorkbook = t_xlWorkbooks.Open(self.filename)

                    
                    
                    self.t_xlApplication = t_xlApplication
                    self.t_xlWorkbooks = t_xlWorkbooks
                    self.t_xlWorkbook = t_xlWorkbook

                    self.t_xlWorksheets = self.t_xlWorkbook.Worksheets

                    
                    

                    self.setSheet(self.t_xlWorksheets, self.t_xlWorksheet)
                    
                    

                    if "保存用" in self.t_xlWorksheet:
                      pass
                    else:
                      self.t_xlWorksheets.Add(None, self.t_xlWorksheets.Item(self.t_xlWorksheets.Count))
                      wsp = self.t_xlWorksheets.Item(self.t_xlWorksheets.Count)
                      wsp.Name = u"保存用"
                      
                      self.t_xlWorksheet["保存用"] = wsp
                      self.t_xlWorksheets.Select()
                      

                    self.preInitCom()

                    
                except:
                    return
            except:
                return
        except:
            return


##
# @brief
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
    profile = OpenRTM_aist.Properties(defaults_str=excelcontrol_spec)
    manager.registerFactory(profile,
                            ExcelControl,
                            OpenRTM_aist.Delete)
    comp = manager.createComponent("ExcelControl")

def main():
    """m_excel = ExcelObject()
    m_excel.Open("")
    cell,sheet,m_len = m_excel.getCell(1, "A", "Sheet1", "C", False)
    print type(cell.Value2)
    cell.Interior.ColorIndex = 6
    #m_excel.setCellValue(1, "A", "Sheet1", True, [[0,1,2],[0,1,2]])
    #print m_excel.getCellValue(1,"A","Sheet1","B")
    cell,sheet,m_len = m_excel.getCell(1, "A", "Sheet1", "C")
    v = []
    for i in range(0,len(cell.Value2[0])):
      if i == 1:
        v.append(10)
      else:
        v.append(cell.Value2[0][i])
    cell.Value2 = v#"""

    #print thread.get_ident()
    

    mgr = OpenRTM_aist.Manager.init(sys.argv)
    mgr.setModuleInitProc(MyModuleInit)
    mgr.activateManager()
    mgr.runManager(True)

    global excel_comp
    
    app = QtGui.QApplication([""])
    mainWin = MainWindow(excel_comp, mgr)
    #mainWin = MainWindow(None, None)
    mainWin.show()
    app.exec_()
    
    
if __name__ == "__main__":
    main()
