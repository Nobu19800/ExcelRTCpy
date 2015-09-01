#!/bin/env python
# -*- encoding: utf-8 -*-

##
#   @file ExcelObject.py
# @brief Excelを操作するためのクラス、関数

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
# @class ExcelComObject
# @brief 
# 
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



    

