#!/bin/env python
# -*- encoding: utf-8 -*-

##
#
# @file CalcControl.py

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math



from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', home+'/OOoRTC/CalcIDL', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', '.\\OOoRTC\\CalcIDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL']




import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming


import threading



import SpreadSheet_idl
from omniORB import PortableServer


import SpreadSheet, SpreadSheet__POA

import CalcDataPort
import OOoRTC


##
# @class CalcConfigUpdateParam
# @brief コンフィギュレーションパラメータが更新されたときのコールバック
#

class CalcConfigUpdateParam(OpenRTM_aist.ConfigurationSetListener):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param e_rtc OOoCalcRTC
    #
   def __init__(self,e_rtc):
        self.m_rtc =  e_rtc

    ##
    # @brief 
    # @param self 
    # @param config_param_name 
    #
   def __call__(self, config_param_name):
        self.m_rtc.configUpdate()


##
# @class mSpreadSheet_i
# @brief サービスポートSpreadSheet
#
class mSpreadSheet_i (SpreadSheet__POA.mSpreadSheet):


    ##
    # @brief コンストラクタ
    # @param self 
    # @param m_comp OOoCalcRTC
    #
    def __init__(self, m_comp):

        self.m_comp = m_comp

    



    ##
    # @brief セルの文字列を取得
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @return セルの文字列
    #
    def get_string(self, l, c, sn):
        guard = OpenRTM_aist.ScopedLock(self.m_comp._mutex)
        try:
            tmp = self.m_comp.get_string(l, c, sn)
        except:
            pass
        del guard

        return tmp
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief セルの値を設定
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @param v 設定する値
    #
    def set_value(self, l, c, sn, v):
        guard = OpenRTM_aist.ScopedLock(self.m_comp._mutex)
        try:
            self.m_comp.set_value(l, c, sn, v)
        except:
            pass
        del guard
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 未実装
    #
    def get_string_range(self, l1, c1, l2, c2, sn):
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 未実装
    #
    def set_value_range(self, l, c, sn, v):
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None

    ##
    # @brief セルの文字列を設定
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @param v 設定する文字列
    #
    def set_string(self, l, c, sn, v):
        guard = OpenRTM_aist.ScopedLock(self.m_comp._mutex)
        try:
            self.m_comp.set_string(l, c, sn, v)
        except:
            pass
        del guard
            
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None

    ##
    # @brief 未実装
    #
    def set_string_range(self, l, c, sn, v):
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None





##
# @class CalcControl
# @brief 表計算ソフトを操作するためのRTCのクラス
#

class CalcControl(OpenRTM_aist.DataFlowComponentBase):

    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    self.OutPorts = {}
    self.InPorts = {}
    self.ConfOutPorts = {}
    self.ConfInPorts = {}

    

    self._SpreadSheetPort = OpenRTM_aist.CorbaPort("SpreadSheet")
    self._spreadsheet = mSpreadSheet_i(self)

    
    

    self.conf_data_type = ["TimedFloat"]
    self.conf_port_type = ["DataInPort"]
    self.conf_column = [1]
    self.conf_start_row = ["A"]
    self.conf_end_row = ["A"]
    self.conf_sheetname = ["sheet1"]
    self.actionLock = [1]
    self.red = [255]
    self.green = [255]
    self.blue = [0]
    self.c_move = [1]
    self.Attach_Port = ["None"]

    
    self.stime = [0.05]
    self.stCell_row = ["A"]
    self.stCell_col = [1]
    self.stCell_sheetname = ["sheet1"]

    self._mutex = threading.RLock()
    self.guard = None

    self.m_CalcInPort = CalcDataPort.CalcInPort
    self.m_CalcInPortSeq = CalcDataPort.CalcInPortSeq
    self.m_CalcInPortEx = CalcDataPort.CalcInPortEx

    self.m_CalcOutPort = CalcDataPort.CalcOutPort
    self.m_CalcOutPortSeq = CalcDataPort.CalcOutPortSeq
    self.m_CalcOutPortEx = CalcDataPort.CalcOutPortEx

    self.m_time = 0

    
    
    
    return

    ##
    # @brief セルの文字列を取得
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @return セルの文字列
    #
  def get_string(self, l, c, sn):
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
    pass

    ##
    # @brief セルの文字列を設定
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @param v 設定する文字列
    #
  def set_string(self, l, c, sn, v):
    pass
    
  ##
  # @brief 実行周期を設定する関数
  # @param self 
  # @param rate：実行周期
  #

  def mSetRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # @brief 活性化するための関数
  # @param self 
  #    

  def mActivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # @brief 不活性化するための関数
  # @param self 
  #

  def mDeactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  ##
  # @brief コンフィギュレーションパラメータによりアウトポートを追加する関数
  # @param self 
  # @param name データポート名
  # @param data_type データ型
  # @param row 行番号
  # @param col 列番号
  # @param mlen 行番号の範囲
  # @param sn シート名
  # @param mstate 列を移動するか
  # @param t_attachports 関連付けしたインポート

  def addConfOutPort(self, name, data_type, row, col, mlen, sn, mstate, t_attachports):

    sig = CalcDataPort.DataType.Single
    sec = CalcDataPort.DataType.Sequence
    ext = CalcDataPort.DataType.Extend

    
    
    m_data_o, m_data_type =  CalcDataPort.GetDataSType(data_type)
    

    if m_data_o:
        
        
        m_outport = OpenRTM_aist.OutPort(name, m_data_o)
        self.addOutPort(name, m_outport)
        

        if m_data_type[1] == sig:
            self.ConfOutPorts[name] = self.m_CalcOutPort(m_outport, m_data_o, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)
        elif m_data_type[1] == sec:
            self.ConfOutPorts[name] = self.m_CalcOutPortSeq(m_outport, m_data_o, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)
        elif m_data_type[1] == ext:
            self.ConfOutPorts[name] = self.m_CalcOutPortEx(m_outport, m_data_o, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)

  ##
  # @brief コンフィギュレーションパラメータによりインポートを追加する関数
  # @param self 
  # @param name データポート名
  # @param data_type データ型
  # @param row 行番号
  # @param col 列番号
  # @param mlen 行番号の範囲
  # @param sn シート名
  # @param mstate 列を移動するか
  # @param t_attachports 関連付けしたアウトポート
  
  def addConfInPort(self, name, data_type, row, col, mlen, sn, mstate, t_attachports):
    sig = CalcDataPort.DataType.Single
    sec = CalcDataPort.DataType.Sequence
    ext = CalcDataPort.DataType.Extend

    
    
    m_data_i, m_data_type =  CalcDataPort.GetDataSType(data_type)
    
    if m_data_i:
        
        
        m_inport = OpenRTM_aist.InPort(name, m_data_i)
        self.addInPort(name, m_inport)

        
        
        
        #self.InPorts[name] = self.m_CalcPortObject(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)
        if m_data_type[1] == sig:
            self.ConfInPorts[name] = self.m_CalcInPort(m_inport, m_data_i, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)
        elif m_data_type[1] == sec:
            self.ConfInPorts[name] = self.m_CalcInPortSeq(m_inport, m_data_i, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)
        elif m_data_type[1] == ext:
            self.ConfInPorts[name] = self.m_CalcInPortEx(m_inport, m_data_i, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)
        
        m_inport.addConnectorDataListener(OpenRTM_aist.ConnectorDataListenerType.ON_BUFFER_WRITE,
                                          DataListener(self.ConfInPorts[name],self))

  ##
  # @brief アウトポート追加の関数
  # @param self
  # @param name アウトポートの名前
  # @param m_inport 接続するインポート
  # @param row データを書き込む行番号
  # @param sn 接続するインポートのパス
  # @param mstate 列を移動するか
  # @param t_attachports 関連付けするデータポート
  # @return 追加したアウトポート
  def mAddOutPort(self, name, m_inport, row, col, mlen, sn, mstate, t_attachports, autoCon = True):

    sig = CalcDataPort.DataType.Single
    sec = CalcDataPort.DataType.Sequence
    ext = CalcDataPort.DataType.Extend

    m_data_o = None
    m_data_type = None
    
    if autoCon:
        m_data_o, m_data_type =  CalcDataPort.GetDataType(m_inport[1])
    else:
        m_data_o, m_data_type =  CalcDataPort.GetDataSType(m_inport[1])
    

    if m_data_o:
        
        m_outport = OpenRTM_aist.OutPort(name, m_data_o)
        self.addOutPort(name, m_outport)
        
        if autoCon:
            OOoRTC.ConnectPort(m_inport[1], m_outport._objref, name)

        if m_data_type[1] == sig:
            self.OutPorts[name] = self.m_CalcOutPort(m_outport, m_data_o, name, row, col, mlen, sn, mstate, m_inport, m_data_type, t_attachports)
        elif m_data_type[1] == sec:
            self.OutPorts[name] = self.m_CalcOutPortSeq(m_outport, m_data_o, name, row, col, mlen, sn, mstate, m_inport, m_data_type, t_attachports)
        elif m_data_type[1] == ext:
            self.OutPorts[name] = self.m_CalcOutPortEx(m_outport, m_data_o, name, row, col, mlen, sn, mstate, m_inport, m_data_type, t_attachports)
        
        return self.OutPorts[name]

    return None
                
    

            
        
  ##
  # @brief インポート追加の関数
  # @param self
  # @param name インポートの名前
  # @param m_outport 接続するアウトポート
  # @param row データを書き込む行番号
  # @param sn 書き込むシート
  # @param mstate 列を移動するか
  # @param t_attachports 関連付けするデータポート
  # @return 追加したインポート
        
  def mAddInPort(self, name, m_outport, row, col, mlen, sn, mstate, t_attachports, autoCon = True):
    sig = CalcDataPort.DataType.Single
    sec = CalcDataPort.DataType.Sequence
    ext = CalcDataPort.DataType.Extend


    m_data_i = None
    m_data_type = None
    
    if autoCon:
        m_data_i, m_data_type =  CalcDataPort.GetDataType(m_outport[1])
    else:
        m_data_i, m_data_type =  CalcDataPort.GetDataSType(m_outport[1])
    
    if m_data_i:
        m_inport = OpenRTM_aist.InPort(name, m_data_i)
        self.addInPort(name, m_inport)

        if autoCon:
            OOoRTC.ConnectPort(m_inport._objref, m_outport[1], name)

        
        
        #self.InPorts[name] = self.m_CalcPortObject(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)
        if m_data_type[1] == sig:
            self.InPorts[name] = self.m_CalcInPort(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)
        elif m_data_type[1] == sec:
            self.InPorts[name] = self.m_CalcInPortSeq(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)
        elif m_data_type[1] == ext:
            self.InPorts[name] = self.m_CalcInPortEx(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)


        
        m_inport.addConnectorDataListener(OpenRTM_aist.ConnectorDataListenerType.ON_BUFFER_WRITE,
                                          DataListener(self.InPorts[name], self))

        return self.InPorts[name]

    return None

  ##
  # @brief データポート全削除の関数
  # @param self 
  #
  def mRemoveAllPort(self):
      for n,op in self.OutPorts.items():
          op._port.disconnect_all()
          self.removePort(op._port)
      self.OutPorts = {}

      for n,ip in self.InPorts.items():
          ip._port.disconnect_all()
          self.removePort(ip._port)
      self.InPorts = {}
  
  ##
  # @brief アウトポート削除の関数
  # @param self 
  # @param outport 削除するアウトポート
  #
  
  def mRemoveOutPort(self, outport):
      outport._port.disconnect_all()
      self.removePort(outport._port)
      del self.OutPorts[outport._name]

  ##
  # @brief インポート削除の関数
  # @param self 
  # @param outport 削除するインポート
  #

  def mRemoveInPort(self, inport):
      inport._port.disconnect_all()
      self.removePort(inport._port)
      del self.InPorts[inport._name]


  ##
  # @brief コンフィギュレーションパラメータが変更されたときに呼び出される関数
  # @param self 
  #
  
  def configUpdate(self):
      
      for i in range(0, 100):
          dn = "dataport" + str(i+1)
          
          
          if self._configsets.haveConfig(dn):
              
              self._configsets.activateConfigurationSet(dn)
              self._configsets.update(dn)

              
              tdt = ""
              tmp = None
              if self.ConfInPorts.has_key(dn):
                  if self.conf_port_type[0] != "DataInPort":
                      del self.ConfInPorts[dn]
                  else:
                      tmp = self.ConfInPorts[dn]
                      tdt = "DataInPort"
              if self.ConfOutPorts.has_key(dn):
                  if self.conf_port_type[0] != "DataOutPort":
                      del self.ConfOutPorts[dn]
                  else:
                      tmp = self.ConfOutPorts[dn]
                      tdt = "DataOutPort"

              data_type = ""
              if tmp != None:
                  profile = tmp._port.get_port_profile()
                  #props = nvlist_to_dict(profile.properties)
                  #data_type =  props['dataport.data_type']
                  data_type =  OOoRTC.nvlist_getValue(profile.properties,'dataport.data_type')
                  if data_type.startswith('IDL:'):
                    data_type = data_type[4:]
                    colon = data_type.rfind(':')
                  if colon != -1:
                    data_type = data_type[:colon]

                    data_type = data_type.replace('RTC/','')
              

              if int(self.conf_column[0]) > 0 and len(self.conf_start_row[0]) > 0:
                  c_move = True
                  if int(self.c_move[0]) == 0:
                      c_move = False
                  Attach_Port = {}
                  tA = re.split(",",self.Attach_Port[0])
                  for k in tA:
                       if k != "" and k != "None":
                           Attach_Port[k] = k

                  if tdt != None and data_type == self.conf_data_type[0]:# and self.conf_port_type[0] == tdt:
                      tmp._row = self.conf_start_row[0]
                      tmp._sn = self.conf_sheetname[0]
                      tmp._col = self.conf_column[0]
                      tmp._length = self.conf_end_row[0]
                      tmp.attachports = Attach_Port 
                      tmp.state = c_move
                      

                  else:
                      
                      if tmp != None:
                          tmp._port.disconnect_all()
                          self.removePort(tmp._port)

                      
                      
                      if self.conf_port_type[0] == "DataInPort":
                        self.addConfInPort(dn, self.conf_data_type[0], self.conf_start_row[0], int(self.conf_column[0]), self.conf_end_row[0], self.conf_sheetname[0], c_move, Attach_Port)
                      elif self.conf_port_type[0] == "DataOutPort":
                        self.addConfOutPort(dn, self.conf_data_type[0], self.conf_start_row[0], int(self.conf_column[0]), self.conf_end_row[0], self.conf_sheetname[0], c_move, Attach_Port)
                      

  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  #
  
  def onInitialize(self):
    

    self._SpreadSheetPort.registerProvider("spreadsheet", "SpreadSheet::mSpreadSheet", self._spreadsheet)
    self.addPort(self._SpreadSheetPort)

    
    self.addConfigurationSetListener(OpenRTM_aist.ConfigurationSetListenerType.ON_SET_CONFIG_SET, CalcConfigUpdateParam(self))

    self.bindParameter("data_type", self.conf_data_type, "TimedFloat")
    self.bindParameter("port_type", self.conf_port_type, "DataInPort")
    self.bindParameter("column", self.conf_column, "1")
    self.bindParameter("start_row", self.conf_start_row, "A")
    self.bindParameter("end_row", self.conf_end_row, "A")
    self.bindParameter("sheetname", self.conf_sheetname, "sheet1")
    self.bindParameter("actionLock", self.actionLock, "1")
    self.bindParameter("Red", self.red, "255")
    self.bindParameter("Green", self.green, "255")
    self.bindParameter("Blue", self.blue, "0")
    self.bindParameter("c_move", self.c_move, "1")
    self.bindParameter("Attach_Port", self.Attach_Port, "None")
    self.bindParameter("stime", self.stime, "0.05")
    self.bindParameter("stCell_row", self.stCell_row, "A")
    self.bindParameter("stCell_col", self.stCell_col, "1")
    self.bindParameter("stCell_sheetname", self.stCell_sheetname, "sheet1")
    
    
    
    
    
    
    
    return RTC.RTC_OK

    ##
    # @brief 画面の更新停止
    # @param self 
    # 
  def addActionLock(self):
    pass

    ##
    # @brief 画面の更新再開
    # @param self 
    # 
  def removeActionLock(self):
    pass

    ##
    # @brief データポートと関連付けしてあるセルの色を設定
    # @param self 
    # @param op データポートオブジェクト
    #  
  def setCellColor(self, op):
    pass
  
  ##
  # @brief 不活性化時のコールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onDeactivated(self, ec_id):

    
        
    self.guard = OpenRTM_aist.ScopedLock(self._mutex)

    try:
        self.addActionLock()
        
        
        for n,op in self.OutPorts.items():
            self.setCellColor(op)
            

        for n,op in self.ConfOutPorts.items():
            self.setCellColor(op)
            

        self.removeActionLock()
    except:
        pass
    
    del self.guard

    for n,op in self.ConfOutPorts.items():
        op._num = int(op._col)
        op.count = 0

    for n,ip in self.ConfInPorts.items():
        ip._num = int(ip._col)
        ip.count = 0

    for n,op in self.OutPorts.items():
        op._num = int(op._col)
        op.count = 0

    for n,ip in self.InPorts.items():
        ip._num = int(ip._col)
        ip.count = 0

    self.m_time = 0

    
    
    return RTC.RTC_OK

  ##
  # @brief 関連付けたインポート、アウトポートの処理
  # @param self 
  # @param ip インポート
  # @param OutPorts アウトポートのリスト
  # @param InPorts インポートのリスト
  def udAPort(self, ip, OutPorts, InPorts):
      
      for n,p in ip.attachports.items():
        if OutPorts.has_key(p) == True:
            
            op = OutPorts[p]
            if len(op.attachports) != 0:
                
                Flag = True
                for i,j in op.attachports.items():
                    if InPorts.has_key(j) == True:
                        #if len(self.InPorts[j].buffdata) == 0:
                        if InPorts[j]._port.isNew() != True:
                            Flag = False
                    else:
                        Flag = False
                        
                if Flag:
                    self.guard = OpenRTM_aist.ScopedLock(self._mutex)

                    try:
                        for i,j in op.attachports.items():
                            InPorts[j].putData(self)
                        
                        op.putData(self)
                    except:
                        pass
                    
                    del self.guard

  ##
  # @brief インポートと関連付けしたアウトポートのデータ入力後、インポートのデータ出力
  # @param self 
  # @param ip インポート
  def updateAPort(self, ip):
      self.udAPort(ip, self.OutPorts, self.InPorts)
      self.udAPort(ip, self.ConfOutPorts, self.ConfInPorts)
     
      
  def setTime(self):
    pass

  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    
    
    
    self.guard = OpenRTM_aist.ScopedLock(self._mutex)

    try:
        if int(self.actionLock[0]) != 0:
            
            self.addActionLock()
            

        

        self.setTime()

        
                
        for n,ip in self.ConfInPorts.items():
            if len(ip.attachports) == 0:
                ip.putData(self)

        for n,ip in self.InPorts.items():
            if len(ip.attachports) == 0:
                ip.putData(self)
                

        if int(self.actionLock[0]) == 2:
            self.removeActionLock()
                
        
        for n,op in self.OutPorts.items():
            
            if len(op.attachports) == 0:
                op.putData(self)
                
        

        for n,op in self.ConfOutPorts.items():
            if len(op.attachports) == 0:
                op.putData(self)


                
                
                
            
                
        if int(self.actionLock[0]) == 1:
            self.removeActionLock()

    except:
        pass
        
    del self.guard

    self.m_time += self.stime[0]

    
    
    return RTC.RTC_OK

  

    
  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param 
  # @return RTC::ReturnCode_t
  def onFinalize(self):
      
      return RTC.RTC_OK


  ##
  # @brief データポートと関連付けしたセルに名前を入力
  # @param self 
  #
  def update_cellName(self):
      for n,op in self.ConfOutPorts.items():
        op.update_cellName(self)
            
      for n,ip in self.ConfInPorts.items():
        ip.update_cellName(self)


            
    
      for n,op in self.OutPorts.items():
        op.update_cellName(self)
            
      for n,ip in self.InPorts.items():
        ip.update_cellName(self)


##
# @class DataListener
# @brief データが書き込まれたときに呼び出されるコールバック関数
#


class DataListener(OpenRTM_aist.ConnectorDataListenerT):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param m_port データポートオブジェクト
    # @param m_rtc OOoCalcRTC
    #
  def __init__(self, m_port, m_rtc):
    self.m_port = m_port
    self.m_rtc = m_rtc

    ##
    # @brief デストラクタ
    # @param self 
    #
  def __del__(self):
    pass

    ##
    # @brief
    # @param self 
    # @param info コネクタの情報
    # @param cdrdata データ
  def __call__(self, info, cdrdata):
    
    data = OpenRTM_aist.ConnectorDataListenerT.__call__(self, info, cdrdata, self.m_port._data)
    
    guard = OpenRTM_aist.ScopedLock(self.m_port._mutex)
    try:
        self.m_port.buffdata.append(data.data)
    except:
        pass
    del guard
    

    
    self.m_rtc.updateAPort(self.m_port)
