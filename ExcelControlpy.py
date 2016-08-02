#!/bin/env python
# -*- encoding: utf-8 -*-

##
#   @file ExcelControlpy.py
#   @brief ExcelControlpy Component



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
from ExcelObject import *

excel_comp = None



excelcontrolpy_spec = ["implementation_id", "ExcelControlpy",
                  "type_name",         "ExcelControlpy",
                  "description",       "Excel Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "Office",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.file_path", "NewFile",
                  "conf.default.actionLock", "1",
                  "conf.default.red", "255",
                  "conf.default.green", "255",
                  "conf.default.blue", "0",
                  "conf.default.stime", "0.05",
                  "conf.default.stCell_row", "A",
                  "conf.default.stCell_col", "1",
                  "conf.default.stCell_sheetname", "Sheet1",
                  "conf.dataport0.port_type", "DataInPort",
                  "conf.dataport0.data_type", "TimedFloat",
                  "conf.dataport0.column", "1",
                  "conf.dataport0.start_row", "A",
                  "conf.dataport0.end_row", "A",
                  "conf.dataport0.sheetname", "sheet1",
                  "conf.dataport0.c_move", "1",
                  "conf.dataport0.Attach_Port", "None",
                  "conf.__widget__.actionLock", "radio",
                  "conf.__widget__.red", "spin",
                  "conf.__widget__.green", "spin",
                  "conf.__widget__.blue", "spin",
                  "conf.__widget__.file_path", "text",
                  "conf.__widget__.port_type", "radio",
                  "conf.__widget__.column", "spin",
                  "conf.__widget__.start_row", "text",
                  "conf.__widget__.end_row", "text",
                  "conf.__widget__.sheetname", "text",
                  "conf.__widget__.data_type", "radio",
                  "conf.__widget__.c_move", "radio",
                  "conf.__widget__.Attach_Port", "text",
                  "conf.__widget__.stCell_col", "spin",
                  "conf.__constraints__.actionLock", "(0,1,2)",
                  "conf.__constraints__.red", "0<=x<=255",
                  "conf.__constraints__.green", "0<=x<=255",
                  "conf.__constraints__.blue", "0<=x<=255",
                  "conf.__constraints__.column", "1<=x<=1000",
                  "conf.__constraints__.stCell_col", "1<=x<=1000",
                  "conf.__constraints__.port_type", "(DataInPort,DataOutPort)",
                  "conf.__constraints__.data_type", """(TimedDouble,TimedLong,TimedFloat,TimedShort,TimedULong,TimedUShort,TimedChar,TimedWChar,
                                                    TimedBoolean,TimedOctet,TimedString,TimedWString,TimedDoubleSeq,TimedLongSeq,TimedFloatSeq,
                                                    TimedShortSeq,TimedULongSeq,TimedUShortSeq,TimedCharSeq,TimedWCharSeq,TimedOctetSeq,TimedStringSeq,
                                                    TimedWStringSeq,TimedRGBColour,TimedPoint2D,TimedVector2D,TimedPose2D,TimedVelocity2D,TimedAcceleration2D,
                                                    TimedPoseVel2D,TimedSize2D,TimedGeometry2D,TimedCovariance2D,TimedPointCovariance2D,TimedCarlike,TimedSpeedHeading2D,
                                                    TimedPoint3D,TimedVector3D,TimedOrientation3D,TimedPose3D,TimedVelocity3D,TimedAngularVelocity3D,TimedAcceleration3D,
                                                    TimedAngularAcceleration3D,TimedPoseVel3D,TimedSize3D,TimedGeometry3D,TimedCovariance3D,TimedSpeedHeading3D,TimedOAP)""",
                  "conf.__constraints__.c_move", "(0,1)",
                  ""]






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
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        m_string = CalcDataPort.DataType.String
        m_value = CalcDataPort.DataType.Value
        
        cell, sheet, m_len = self.getCell(m_cal)
        if cell != None:
          cell.Value2 = b
          """if self._dataType[2] == m_string:
            cell.Value2 = b
          elif self._dataType[2] == m_value:
            if self.state:
              pass
            else:
              db = 0
              if self.count != 0:
                if  self._length == "":
                  db = (b - cell.Value2)/m_cal.stime[0]
                else:
                  db = (b - cell.Value2[0][0])/m_cal.stime[0]
                
              celld, sheetd, m_lend = m_cal.m_excel.getCell(self._num+1, self._row, self._sn, "")
              celld.Value2 = db
              
              cellf, sheetf, m_lenf = m_cal.m_excel.getCell(self._num+2, self._row, self._sn, "")
              
              fb = b*m_cal.stime[0]
              
              if self.count != 0:
                fb = cellf.Value2 + b*m_cal.stime[0]
                
                
              
              
              cellf.Value2 = fb
              
            cell.Value2 = b
            self.count += 1"""

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
        m_string = CalcDataPort.DataType.String
        m_value = CalcDataPort.DataType.Value
        
        cell, sheet, m_len = self.getCell(m_cal)
        if cell != None:
          v = []
          for j in range(0, len(b)):
            if m_len > j:
              v.append(b[j])
          cell.Value2 = v
          """if self._dataType[2] == m_string:
              cell.Value2 = v

              
          elif self._dataType[2] == m_value:
            
            if self.state:
              pass
            else:
            
              if  self._length == "":
                db = 0
                
                if self.count != 0:
                  db = (v[0] - cell.Value2)/m_cal.stime[0]
                  
                celld, sheetd, m_lend = m_cal.m_excel.getCell(self._num+1, self._row, self._sn, "")
                celld.Value2 = db
                
                cellf, sheetf, m_lenf = m_cal.m_excel.getCell(self._num+2, self._row, self._sn, "")
                
                fb = v[0]*m_cal.stime[0]
                
                if self.count != 0:
                  fb = cellf.Value2 + v[0]*m_cal.stime[0]
 
                cellf.Value2 = fb
              else:
                db = []
                tv2 = cell.Value2[0]
                for i in range(0, len(v)):
                  if self.count != 0:
                    db.append((v[i] - tv2[i])/m_cal.stime[0])
                  else:
                    db.append(0)

                celld, sheetd, m_lend = m_cal.m_excel.getCell(self._num+1, self._row, self._sn, self._length)
                celld.Value2 = db

                cellf, sheetf, m_lenf = m_cal.m_excel.getCell(self._num+2, self._row, self._sn, self._length)

                fb = []
                tv3 = cellf.Value2[0]
                for i in range(0, len(v)):
                  if self.count != 0:
                    fb.append(tv3[i] + v[i]*m_cal.stime[0])
                  else:
                    fb.append(v[i]*m_cal.stime[0])
                    
                cellf.Value2 = fb
                  
            cell.Value2 = v
            self.count += 1"""
            
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

class ExcelControlpy(CalcControl):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    CalcControl.__init__(self, manager)

    global excel_comp
    excel_comp = self

    
    
    """prop = OpenRTM_aist.Manager.instance().getConfig()
    fn = self.getProperty(prop, "excel.filename", "")
    self.m_excel = ExcelObject()
    if fn != "":
      str1 = [fn]
      OpenRTM_aist.replaceString(str1,"/","\\")
      fn = os.path.abspath(str1[0])
    self.m_excel.Open(fn)"""

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


    ##
    # @brief セルの文字列を取得
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @return セルの文字列
    #
  def get_string(self, l, c, sn):
    cell,sheet,m_len = self.m_excel.getCell(c, l, sn, "")
    if cell and m_len == 1:
      return str(cell.Value2)
    elif cell:
      return str(cell.Value2)
    else:
      return ""
    
    
    ##
    # @brief セルの値を設定
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @param v 設定する値
    #
  def set_value(self, l, c, sn, v):
    cell,sheet,m_len = self.m_excel.getCell(c, l, sn, "")
    if cell:
      cell.Value2 = v


    ##
    # @brief セルの文字列を設定
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @param v 設定する文字列
    #
  def set_string(self, l, c, sn, v):
    cell,sheet,m_len = self.m_excel.getCell(c, l, sn, "")
    if cell:
      cell.Value2 = v

    ##
    # @brief 画面の更新停止
    # @param self 
    # 
  def addActionLock(self):
    return
    """tid = str(thread.get_ident())
    self.m_excel.comObjects[tid].xlApplication.ScreenUpdating = False"""

    ##
    # @brief 画面の更新再開
    # @param self 
    # 
  def removeActionLock(self):
    return
    """tid = str(thread.get_ident())
    self.m_excel.comObjects[tid].xlApplication.ScreenUpdating = True"""

    ##
    # @brief データポートと関連付けしてあるセルの色を設定
    # @param self 
    # @param op データポートオブジェクト
    #  
  def setCellColor(self, op):
    pass

  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  #
  def onInitialize(self):
    CalcControl.onInitialize(self)

    self.bindParameter("file_path", self.conf_filename, "NewFile")
    
    self._configsets.update("default","file_path")
    self.m_excel = ExcelObject()
    fn = self.conf_filename[0]
    if fn == "NewFile":
        fn = ""

    if fn != "":
      str1 = [fn]
      OpenRTM_aist.replaceString(str1,"/","\\")
      fn = os.path.abspath(str1[0])
    
    self.m_excel.Open(fn)
    
    
    return RTC.RTC_OK
    
  ##
  # @brief 活性化時のコールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  def onActivated(self, ec_id):
    
    #self.m_excel.initCom()
    
    #CalcControl.onActivated(self, ec_id)
    
    

    #self.file = open('text3.txt', 'w')

    
    
    
    return RTC.RTC_OK

  ##
  # @brief 不活性化時のコールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  def onDeactivated(self, ec_id):
    
    CalcControl.onDeactivated(self, ec_id)
    
    #self.file.close()
    return RTC.RTC_OK

  def setTime(self):
    try:
      self.set_value(self.stCell_row[0], self.stCell_col[0], self.stCell_sheetname[0], self.m_time)
    except:
      pass

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


  



##
# @brief
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
    profile = OpenRTM_aist.Properties(defaults_str=excelcontrolpy_spec)
    manager.registerFactory(profile,
                            ExcelControlpy,
                            OpenRTM_aist.Delete)
    comp = manager.createComponent("ExcelControlpy")

##
# @brief 
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
