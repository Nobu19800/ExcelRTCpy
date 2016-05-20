# -*- coding: utf-8 -*-

##
#
# @file MainWindow.py
# @brief ExcelRTCの操作GUIのメインウインドウ

from PyQt4 import QtCore, QtGui
import OOoRTC
import CalcDataPort

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


import threading

##
# @class TreeNode
# @brief ツリーノード
#
class TreeNode:
    ##
    # @brief コンストラクタ
    # @param self 
    # @param node ツリーアイテム
    # @param mw メインウインドウオブジェクト
    def __init__(self, node, mw):
        self.node = node
        self.window = mw
    ##
    # @brief 子ノード追加
    # @param self
    # @param child 子ノード
    def appendChild(self, child):
        self.node.addChild(child.node)
    ##
    # @brief 親ノード取得
    # @param self
    # @return 親ノード
    def getParent(self):
        parent = self.node.parent()
        if parent == None:
            return None
        return self.window.getTreeNode(parent)
    ##
    # @brief 子ノード数取得
    # @param self
    # @return 子ノード数
    def getChildCount(self):
        return int(self.node.childCount())
    ##
    # @brief 表示名取得
    # @param self
    # @return 表示名
    def getDisplayValue(self):
        return str(self.node.text(0).toLocal8Bit())

##
# @class MainWindow
# @brief メインウインドウのウィジェット
#
class MainWindow(QtGui.QMainWindow):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param rtc ExcelRTC
    # @param mgr マネージャーオブジェクト
    def __init__(self, rtc, mgr):
        super(MainWindow, self).__init__()
        self.rtc = rtc
        self.mgr = mgr
        self.treeNodeList = []
        self._paths = []
        self.selItem = None

        

        
        self.layout = QtGui.QHBoxLayout()


        #self.UB.clicked.connect(self.UpdateComp)        
        

        self.l_layout = QtGui.QVBoxLayout()
        
        self.treeWidget = QtGui.QTreeWidget(self)
        self.treeWidget.itemClicked.connect(self.treeWidgetSlot)
        self.l_layout.addWidget(self.treeWidget)
        
        
        self.portListcomboBox = QtGui.QComboBox()
        self.portListcomboBox.setLineEdit(QtGui.QLineEdit())
        self.portListcomboBox.activated[str].connect(self.portListSlot)
        self.l_layout.addWidget(self.portListcomboBox)
        
        self.sub_l_layout = QtGui.QHBoxLayout()
        
        self.addButton = QtGui.QPushButton(u"作成")
        self.sub_l_layout.addWidget(self.addButton)
        self.addButton.clicked.connect(self.addSlot)  
        
        self.deleteButton = QtGui.QPushButton(u"削除")
        self.sub_l_layout.addWidget(self.deleteButton)
        self.deleteButton.clicked.connect(self.deleteSlot)
        
        
        self.l_layout.addLayout(self.sub_l_layout)
        self.layout.addLayout(self.l_layout)
        

        self.m_layout = QtGui.QVBoxLayout()

        self.infotextBox = QtGui.QLineEdit()
        self.m_layout.addWidget(self.infotextBox)
        
        self.collabel = QtGui.QLabel(u"行番号")
        self.m_layout.addWidget(self.collabel)

        self.coltextBox = QtGui.QLineEdit()
        self.m_layout.addWidget(self.coltextBox)
        self.coltextBox.setText("2")

        self.rawlabel = QtGui.QLabel(u"列番号")
        self.m_layout.addWidget(self.rawlabel)

        self.sub_m_layout = QtGui.QHBoxLayout()

        self.rawtextBox = QtGui.QLineEdit()
        self.sub_m_layout.addWidget(self.rawtextBox)
        self.rawtextBox.setText("A")

        self.lentextBox = QtGui.QLineEdit()
        self.sub_m_layout.addWidget(self.lentextBox)
        

        self.m_layout.addLayout(self.sub_m_layout)

        self.sheetlabel = QtGui.QLabel(u"シート名")
        self.m_layout.addWidget(self.sheetlabel)

        self.sheetcomboBox = QtGui.QComboBox()
        self.m_layout.addWidget(self.sheetcomboBox)   
        self.naneServerlabel = QtGui.QLabel(u"ネームサーバーのアドレス")
        self.m_layout.addWidget(self.naneServerlabel)

        self.namingServertextBox = QtGui.QLineEdit()
        self.m_layout.addWidget(self.namingServertextBox)
        self.namingServertextBox.setText("localhost")
        
        self.namingServerButton = QtGui.QPushButton(u"ツリー表示")
        self.m_layout.addWidget(self.namingServerButton)
        self.namingServerButton.clicked.connect(self.namingServerSlot)

        self.resetButton = QtGui.QPushButton(u"行番号初期化")
        self.m_layout.addWidget(self.resetButton)
        self.resetButton.clicked.connect(self.resetSlot)

        self.allResetButton = QtGui.QPushButton(u"全ての行番号を初期化")
        self.m_layout.addWidget(self.allResetButton)
        self.allResetButton.clicked.connect(self.allResetSlot)

        self.loadButton = QtGui.QPushButton(u"読み込み")
        self.m_layout.addWidget(self.loadButton)
        self.loadButton.clicked.connect(self.loadSlot)

        self.layout.addLayout(self.m_layout)
        
        self.r_layout = QtGui.QVBoxLayout()

        self.r_layout.addStretch()
        
        self.inPortlabel = QtGui.QLabel(u"InPort")
        self.r_layout.addWidget(self.inPortlabel)
        
        self.inPortcomboBox = QtGui.QComboBox()
        self.r_layout.addWidget(self.inPortcomboBox)
        
        self.portlabel = QtGui.QLabel(u"関連付けしたInPort")
        self.r_layout.addWidget(self.portlabel)
        
        self.attachcomboBox = QtGui.QComboBox()
        self.r_layout.addWidget(self.attachcomboBox)
        
        self.attachButton = QtGui.QPushButton(u"関連付け")
        self.r_layout.addWidget(self.attachButton)
        self.attachButton.clicked.connect(self.attachBSlot)
        
        self.detachButton = QtGui.QPushButton(u"関連付け解除")
        self.r_layout.addWidget(self.detachButton)
        self.detachButton.clicked.connect(self.detachBSlot)
        
        self.rowcheckBox = QtGui.QCheckBox(u"列を移動させるか")
        self.r_layout.addWidget(self.rowcheckBox)
        self.rowcheckBox.setCheckState(QtCore.Qt.Checked)

        self.datatypelabel = QtGui.QLabel(u"データ型")
        self.r_layout.addWidget(self.datatypelabel)

        self.datatypecomboBox = QtGui.QComboBox()
        self.r_layout.addWidget(self.datatypecomboBox)
        
        self.datatypecomboBox.addItem(u"")
        for d in CalcDataPort.DataType.DataTypeList:
            self.datatypecomboBox.addItem(d.decode("utf-8"))

        
        self.porttypecomboBox = QtGui.QComboBox()
        self.r_layout.addWidget(self.porttypecomboBox)
        self.porttypecomboBox.addItem(u"DataInPort")
        self.porttypecomboBox.addItem(u"DataOutPort")

        self.layout.addLayout(self.r_layout)
        
        

        #Button1.clicked.connect(SetButton)
        
        """item1 = QtGui.QTreeWidgetItem(["test2"])
        self.treeWidget.addTopLevelItem(item1)
        item1_1 = QtGui.QTreeWidgetItem(["child"])
        item1.addChild(item1_1)
        item1 = QtGui.QTreeWidgetItem(self.treeWidget)
        item1.setText(0, u"アイテム1")
        item1_1 = QtGui.QTreeWidgetItem(["child"])
        item1.addChild(item1_1)"""
        #self.layout.addStretch()
	

	
	#self.setRTCTree()

	#self.treeWidget.clear()
        
        self.widget = QtGui.QWidget()
        self.widget.setLayout(self.layout)
        
        self.setCentralWidget(self.widget)
        self.setWindowTitle("ExcelRTC")
        self.setUnifiedTitleAndToolBarOnMac(True)
        
    ##
    # @brief ツリーのマウスでの操作に対するコールバック
    # @param self 
    #
    def treeWidgetSlot(self, obj):
        self.datatypecomboBox.setCurrentIndex(0)
        self.selItem = self.getTreeNode(obj)
        
        t_comp, nd = OOoRTC.JudgePort(self, self._paths)
        
        self.portListcomboBox.lineEdit().setText("")

        if t_comp:
            for n,o in self.rtc.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    self.updateTree(o)
                    return
            for n,i in self.rtc.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    self.updateTree(i)
                    return
        else:
            return

        
        self.infotextBox.setText(u"未作成")

        """tmp = self.getTreeNode(obj)
        print tmp.getParent()
        print tmp.getChildCount()
        print tmp.getDisplayValue()
        self.selItem = tmp
        print OOoRTC.JudgePort(self, self._paths)"""

    ##
    # @brief ツリーの選択位置が変わったときに各テキストボックスの内容を変更する関数
    # @param self 
    # @param m_port データポートオブジェクト
    #
    def updateTree(self, m_port):
    
    
        self.sheetcomboBox.setCurrentIndex(self.sheetcomboBox.findText(m_port._sn))

        
        if len(m_port._port_a[0]) == 2:
            self.datatypecomboBox.setCurrentIndex(self.datatypecomboBox.findText(m_port._port_a[1]))
            self.porttypecomboBox.setCurrentIndex(self.porttypecomboBox.findText(m_port._port_a[0][0]))
        else:
            self.datatypecomboBox.setCurrentIndex(0)
        
    
        self.rawtextBox.setText(m_port._row)

    
        self.lentextBox.setText(m_port._length)
    

    
        self.infotextBox.setText(u'作成済み')

    
        self.coltextBox.setText(str(m_port._col))

    
        if m_port.state:
            self.rowcheckBox.setCheckState(QtCore.Qt.Checked)
        else:
            self.rowcheckBox.setCheckState(QtCore.Qt.Unchecked)

        self.updateInPortList()
        self.updateAttachPort(m_port)
        
        
    ##
    # @brief データポートリストのコールバック
    # @param self
    # @param string テキスト
    #
    def portListSlot(self, string):
        self.updateInPortList()
        pn = str(self.portListcomboBox.currentText().toLocal8Bit())
        if self.rtc.InPorts.has_key(pn) == True:
            self.updateTree(self.rtc.InPorts[pn])
        elif self.rtc.OutPorts.has_key(pn) == True:
            self.updateTree(self.rtc.OutPorts[pn])

    ##
    # @brief インポートを追加する関数
    # @param self 
    # @param name データポート名
    # @param i_port 接続するデータポート
    def compAddOutPort(self, name, i_port, autoCon = True):
        row = str(self.rawtextBox.text().toLocal8Bit())
        sn = str(self.sheetcomboBox.currentText().toLocal8Bit())
        if sn == "":
            return False
        col = str(self.coltextBox.text().toLocal8Bit())
        
        mlen = str(self.lentextBox.text().toLocal8Bit())
        
        mstate = int(self.rowcheckBox.checkState())
        mst = True
        if mstate == QtCore.Qt.Unchecked:
            mst = False
        tcomp = self.rtc.mAddOutPort(name, i_port, row, col, mlen, sn, mst, {}, autoCon)
        if tcomp:
            tcomp.update_cellName(self.rtc)
        return True

    ##
    # @brief インポートを追加する関数
    # @param self 
    # @param name データポート名
    # @param o_port 接続するデータポート
    def compAddInPort(self, name, o_port, autoCon = True):
        row = str(self.rawtextBox.text().toLocal8Bit())
        
        sn = str(self.sheetcomboBox.currentText().toLocal8Bit())
        if sn == "":
            return False
        col = str(self.coltextBox.text().toLocal8Bit())
        
        mlen = str(self.lentextBox.text().toLocal8Bit())
        
        mstate = int(self.rowcheckBox.checkState())
        mst = True
        if mstate == QtCore.Qt.Unchecked:
            mst = False
        tcomp = self.rtc.mAddInPort(name, o_port, row, col, mlen, sn, mst, {}, autoCon)
        if tcomp:
            tcomp.update_cellName(self.rtc)
        return True

    ##
    # @brief 作成したポートの設定を保存する関数
    # self 
    #
    def updateSaveSheet(self):
        
        for i in range(1, 30):
                cell,sheet,m_len = self.rtc.m_excel.getCell(i,"A",'保存用',"",False)
                
                cell.Value2 = ''
                

                cell,sheet,m_len = self.rtc.m_excel.getCell(i,"B",'保存用',"",False)
                cell.Value2 = ''

                cell,sheet,m_len = self.rtc.m_excel.getCell(i,"C",'保存用',"",False)
                cell.Value2 = ''

                cell,sheet,m_len = self.rtc.m_excel.getCell(i,"D",'保存用',"",False)
                cell.Value2 = ''

                cell,sheet,m_len = self.rtc.m_excel.getCell(i,"E",'保存用',"",False)
                cell.Value2 = ''

                cell,sheet,m_len = self.rtc.m_excel.getCell(i,"F",'保存用',"",False)
                cell.Value2 = ''

                cell,sheet,m_len = self.rtc.m_excel.getCell(i,"G",'保存用',"",False)
                cell.Value2 = ''

           
        count = 1
        PortList = []
        
        for n,o in self.rtc.OutPorts.items():
            PortList.append(o)
        for n,i in self.rtc.InPorts.items():
            
            PortList.append(i)
        

        

        for p in PortList:

                    cell,sheet,m_len = self.rtc.m_excel.getCell(count,"A",'保存用',"",False)
                    pn = ''
                    for j in range(0, len(p._port_a[0])):
                        if j == 0:
                            pn = p._port_a[0][j]
                        else:
                            pn = pn + ':' + p._port_a[0][j]
                    cell.Value2 = str(pn)

                    cell,sheet,m_len = self.rtc.m_excel.getCell(count,"B",'保存用',"",False)
                    cell.Value2 = p._row

                    cell,sheet,m_len = self.rtc.m_excel.getCell(count,"C",'保存用',"",False)
                    cell.Value2 = p._col

                    cell,sheet,m_len = self.rtc.m_excel.getCell(count,"D",'保存用',"",False)
                    cell.Value2 = p._length

                    cell,sheet,m_len = self.rtc.m_excel.getCell(count,"E",'保存用',"",False)
                    cell.Value2 = p._sn

                    cell,sheet,m_len = self.rtc.m_excel.getCell(count,"F",'保存用',"",False)
                    cell.Value2 = str(p.state)

                    cell,sheet,m_len = self.rtc.m_excel.getCell(count,"G",'保存用',"",False)
                    pn = ''
                    tmp = 0
                    
                    for k,j in p.attachports.items():
                        if tmp == 0:
                            pn = j
                        else:
                            pn = pn + ':' + j
                        tmp += 1
                            
                        
                    cell.Value2 = str(pn)


                    count = count + 1
            
        else:
            return

    
    def loadParam(self, count):
        cell,sheet,m_len = self.rtc.m_excel.getCell(count,"B",'保存用',"",False)
        if cell.Value2 == None:
            row = "A"
        row = cell.Value2
                    
                    
        cell,sheet,m_len = self.rtc.m_excel.getCell(count,"C",'保存用',"",False)
        if cell.Value2 == None:
            col = 1
        col = int(cell.Value2)

        cell,sheet,m_len = self.rtc.m_excel.getCell(count,"D",'保存用',"",False)
        if cell.Value2 == None:
            mlen = ""
        else:
            mlen = cell.Value2

        cell,sheet,m_len = self.rtc.m_excel.getCell(count,"E",'保存用',"",False)
        if cell.Value2 == None:
            sn = "Sheet1"
        sn = cell.Value2

        cell,sheet,m_len = self.rtc.m_excel.getCell(count,"F",'保存用',"",False)
        
                    
        if str(cell.Value2) == "True":
            mstate = True
        else:
            mstate = False

        cell,sheet,m_len = self.rtc.m_excel.getCell(count,"G",'保存用',"",False)


        t_attachports = {}
        if cell.Value2:
            tmp = re.split(':',cell.Value2)
                        
            for pp in tmp:
                if pp != "":
                    t_attachports[pp] = pp

        return row,col,mlen,sn,mstate,t_attachports

    ##
    # @brief 読み込んだ保存用シートからポートを作成する関数
    # self 
    #
    def loadSheet(self):
        
        self.rtc.mRemoveAllPort()
        count = 1
        m_hostname = ''
        _path = []
        

        while True:
            
            
            cell,sheet,m_len = self.rtc.m_excel.getCell(count,"A",'保存用',"",False)

            if cell.Value2 == None:
                return
            
            m_name = re.split(':',cell.Value2)
            
            
            if len(m_name) < 2:
                return

            if len(m_name) == 2:
                for dn in CalcDataPort.DataType.DataTypeList:
                    if m_name[1] == dn:
                        m_name[1] = dn
                #m_name[1] = m_name[1].decode("cp932").encode("utf-8")
                F_Name = m_name[1] + str(OpenRTM_aist.uuid1())
                row,col,mlen,sn,mstate,t_attachports = self.loadParam(count)
                if m_name[0] == 'DataInPort':
                    self.rtc.mAddInPort(F_Name, [[m_name[0],m_name[1]],m_name[1]], row, col, mlen, sn, mstate, t_attachports, False)
                elif m_name[0] == 'DataOutPort':
                    self.rtc.mAddOutPort(F_Name, [[m_name[0],m_name[1]],m_name[1]], row, col, mlen, sn, mstate, t_attachports, False)
            else:
                if m_hostname == m_name[1]:
                    pass
                else:
                    _paths = OOoRTC.GetPathList(m_name[1], self.mgr, None)
                    m_hostname = m_name[1]

                if _paths == None:
                    return

                for p in _paths:
                    if p[0] == m_name:
                        F_Name = p[0][-2] + p[0][-1]
                        profile = p[1].get_port_profile()
                        #props = nvlist_to_dict(profile.properties)

                        

                        row,col,mlen,sn,mstate,t_attachports = self.loadParam(count)
                        
                        if OOoRTC.nvlist_getValue(profile.properties,'port.port_type') == 'DataInPort':
                            self.rtc.mAddOutPort(F_Name, p, row, col, mlen, sn, mstate, t_attachports)
                        elif OOoRTC.nvlist_getValue(profile.properties,'port.port_type') == 'DataOutPort':
                            self.rtc.mAddInPort(F_Name, p, row, col, mlen, sn, mstate, t_attachports)            
    
            count = count + 1
                                

                


    ##
    # @brief 関連付けしたポートのリストを更新する関数
    # @param self 
    # @param m_port データポートオブジェクト
    def updateAttachPort(self, m_port):
        
        self.attachcomboBox.clear()
    
        for n,i in m_port.attachports.items():
            self.attachcomboBox.addItem (i)


    ##
    # @brief ポートのパラメータを設定する関数
    # @param self 
    # @param m_port データポートオブジェクト
    #
    def setPortParam(self, m_port):
        m_port._row = str(self.rawtextBox.text().toLocal8Bit())
        m_port._sn = str(self.sheetcomboBox.currentText().toLocal8Bit())
        m_port._col = int(self.coltextBox.text().toLocal8Bit())
        m_port._length = str(self.lentextBox.text().toLocal8Bit())
        m_port.update_cellName(self.rtc)
        mstate = int(self.rowcheckBox.checkState())
        if mstate == QtCore.Qt.Unchecked:
            m_port.state = False
        else:
            m_port.state = True
        self.updateSaveSheet()

    ##
    # @brief インポートのリストを更新する関数
    # @param self
    #
    def updateInPortList(self):

        self.inPortcomboBox.clear()
        
        
        for n,i in self.rtc.InPorts.items():
            self.inPortcomboBox.addItem (i._name)

    ##
    # @brief データポートのリストを更新する関数
    # @param self 
    #
    def updateDataPortList(self):

        self.portListcomboBox.clear()
        #dpcb_control.Text = ""
        
        for n,i in self.rtc.InPorts.items():
            self.portListcomboBox.addItem(i._name)

        for n,i in self.rtc.OutPorts.items():
            self.portListcomboBox.addItem(i._name)

        self.portListcomboBox.lineEdit().setText("")
            
    ##
    # @brief ポート作成ボタンのコールバック
    # @param self 
    #
    def addSlot(self):
        
        pn = str(self.portListcomboBox.currentText().toLocal8Bit())
        if self.rtc.InPorts.has_key(pn) == True:
            self.setPortParam(self.rtc.InPorts[pn])
            self.updateSaveSheet()
            return
        elif self.rtc.OutPorts.has_key(pn) == True:
            self.setPortParam(self.rtc.OutPorts[pn])
            self.updateSaveSheet()
            return

        dt = str(self.datatypecomboBox.currentText().toLocal8Bit())
        if dt != "":
            F_Name = dt + str(OpenRTM_aist.uuid1())
            pt = str(self.porttypecomboBox.currentText().toLocal8Bit())
            if pt == "DataOutPort":
                if self.compAddOutPort(F_Name, [[pt,dt],dt], False):
                    self.massageBoxFunc('',dt.decode("utf-8")+u"型のOutPortを作成しました。")
            else:
                if self.compAddInPort(F_Name, [[pt,dt],dt], False):
                    self.massageBoxFunc('',dt.decode("utf-8")+u"型のInPortを作成しました。")

            self.updateInPortList()
            self.updateDataPortList()
            self.updateSaveSheet()
            return
            
        t_comp, nd = OOoRTC.JudgePort(self, self._paths)
        if t_comp:
            
            for n,o in self.rtc.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    self.setPortParam(o)
                    
                    return
            for n,i in self.rtc.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    self.setPortParam(i)
                    
                    return
                
            F_Name = t_comp[0][-2] + t_comp[0][-1]
            profile = t_comp[1].get_port_profile()
            #props = nvlist_to_dict(profile.properties)

            if OOoRTC.nvlist_getValue(profile.properties,'port.port_type') == 'DataInPort':
                self.compAddOutPort(F_Name, t_comp)
            elif OOoRTC.nvlist_getValue(profile.properties,'port.port_type') == 'DataOutPort':
                self.compAddInPort(F_Name, t_comp)

            

            self.updateSaveSheet()
            self.infotextBox.setText(u"作成済み")

            self.updateInPortList()
            self.updateDataPortList()

            self.massageBoxFunc('',t_comp[0][-2]+u"の"+t_comp[0][-1]+u"と通信するデータポートを作成しました。")

        
        """print self.coltextBox.text()
        print self.rawtextBox.text()
        print self.lentextBox.text()
        print self.namingServertextBox.text()
        print self.rowcheckBox.checkState()
        print QtCore.Qt.Checked"""

    ##
    # @brief ツリーノード取得
    # @param self 
    # @param obj ツリーアイテム
    # @return ツリーノード取得
    #
    def getTreeNode(self, obj):
        for i in self.treeNodeList:
            if i.node == obj:
                return i
        return None

    ##
    # @brief 選択中のツリーアイテム取得
    # @param self 
    # @return 選択中のツリーアイテム
    #
    def getSelection(self):
        return self.selItem

    ##
    # @brief メッセージボックス表示
    # @param self 
    # @param title タイトル
    # @param mes メッセージ
    #

    def massageBoxFunc(self, title, mes):
        msgbox = QtGui.QMessageBox(self)
        msgbox.setText(mes)
        msgbox.setModal(True)
        ret = msgbox.exec_()

    ##
    # @brief ポートの削除の関数
    # @param self 
    # @param m_port データポートオブジェクト
    #
    def delPortTC(self, m_port):
        self.clearInfo()
        
        self.updateSaveSheet()
        self.massageBoxFunc("", u'削除しました')

        self.portListcomboBox.lineEdit().setText("")

    ##
    # @brief ポート削除ボタンのコールバック
    # @param self 
    #
    def deleteSlot(self):
        pn = str(self.portListcomboBox.currentText().toLocal8Bit())
        if self.rtc.InPorts.has_key(pn) == True:
            i = self.rtc.InPorts[pn]
            self.rtc.mRemoveInPort(i)
            self.delPortTC(i)
            return
        elif self.rtc.OutPorts.has_key(pn) == True:
            o = self.rtc.OutPorts[pn]
            self.rtc.mRemoveOutPort(o)
            self.delPortTC(o)
            return
        
        t_comp, nd = OOoRTC.JudgePort(self, self._paths)

        if t_comp:
            for n,o in self.rtc.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    self.rtc.mRemoveOutPort(o)
                    self.delPortTC(o)
                    return
            for n,i in self.rtc.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    self.rtc.mRemoveInPort(i)
                    self.delPortTC(i)
                    return
        else:
            self.massageBoxFunc("", u'データポートを選択してください')
            return
        
        self.massageBoxFunc("", u'削除済みです')

        
    ##
    # @brief ポートを削除したときに各テキストボックスを変更する関数
    # @param self 
    #
    def clearInfo(self):
        self.infotextBox.setText(u'未作成')
        self.coltextBox.setText("2")

        self.updateInPortList()
        self.updateDataPortList()


    ##
    # @brief ツリー作成ボタンのコールバック
    # @param self 
    #

    def namingServerSlot(self):
        self.setRTCTree()

    ##
    # @brief データを書き込む列の初期化ボタンのコールバック
    # @param self 
    #
    def resetSlot(self):
        t_comp, nd = OOoRTC.JudgePort(self, self._paths)
        if t_comp:
            for n,o in self.rtc.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    o._num = int(o._col)
                    return
            for n,i in self.rtc.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    i._num = int(i._col)
                    return
        else:
            self.massageBoxFunc("", u'データポートを選択してください')
            return
        
        self.massageBoxFunc("", u'削除済みです')

    ##
    # @brief データを書き込む列を全て初期化するボタンのコールバック
    # @param self 
    #
    def allResetSlot(self):
        for n,o in self.rtc.OutPorts.items():
            o._num = int(o._col)
        for n,i in self.rtc.InPorts.items():
            i._num = int(i._col)

    ##
    # @brief ファイル読み込みボタンのコールバック
    # @param self 
    #
    def loadSlot(self):
        self.loadSheet()
        self.updateDataPortList()
        #fileName = QtGui.QFileDialog.getOpenFileName(self,u"開く","","Config File (*.conf);;Python File (*.py);;All Files (*)")
	

	#ba = str(fileName.toLocal8Bit())
	#self.SC.open(ba)

    ##
    # @brief ポート関連付けの関数
    # @param self
    # @param m_port データポートオブジェクト
    #
    def attachTC(self, m_port):
        iname = str(self.inPortcomboBox.currentText().toLocal8Bit())
        if self.rtc.InPorts.has_key(iname) == True:
                        
            m_port.attachports[iname] = iname
            self.rtc.InPorts[iname].attachports[m_port._name] = m_port._name

            self.updateSaveSheet()
            self.updateAttachPort(m_port)

            self.massageBoxFunc("", m_port._name+u"と"+iname+u"を関連付けしました")
            

            
            self.inPortcomboBox.setCurrentIndex(self.inPortcomboBox.findText(iname))
                    
        else:
            self.massageBoxFunc(u'エラー'u'インポートの名前が正しくありません')
            return

    ##
    # @brief ポート関連付けボタンのコールバック
    # @param self
    #
    def attachBSlot(self):
        pn = str(self.portListcomboBox.currentText().toLocal8Bit())
        
        if self.rtc.OutPorts.has_key(pn) == True:
            o = self.rtc.OutPorts[pn]
            self.attachTC(o)
            return
        
        t_comp, nd = OOoRTC.JudgePort(self, self._paths)

        if t_comp:
            for n,o in self.rtc.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    self.attachTC(o)
                    return
            
        else:
            self.massageBoxFunc("", u'データポートを選択してください')
            return
        
        self.massageBoxFunc("", u'削除済みです')


    ##
    # @brief ポート関連付け解除の関数
    # @param self
    # @param m_port データポートオブジェクト
    #
    def detachTC(self, m_port):
        iname = str(self.attachcomboBox.currentText().toLocal8Bit())

        if m_port.attachports.has_key(iname) == True:
            del m_port.attachports[iname]
            if self.rtc.InPorts[iname].attachports.has_key(m_port._name) == True:
                del self.rtc.InPorts[iname].attachports[m_port._name]
                UpdateSaveSheet()  
                UpdateAttachPort(dlg_control, m_port)

                self.massageBoxFunc("", m_port._name+u"と"+iname+u"の関連付けを解除しました")

            else:
                self.massageBoxFunc(u'エラー'u'インポートの名前が正しくありません')

    ##
    # @brief ポート関連付け解除ボタンのコールバック
    # @param self
    #
    def detachBSlot(self):
        pn = str(self.portListcomboBox.currentText().toLocal8Bit())
        
        if self.rtc.OutPorts.has_key(pn) == True:
            o = self.rtc.OutPorts[pn]
            self.detachTC(o)
            return
        
        t_comp, nd = OOoRTC.JudgePort(self, self._paths)

        if t_comp:
            for n,o in self.rtc.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    self.detachTC(o)
                    return
            
        else:
            self.massageBoxFunc("", u'データポートを選択してください')
            return
        
        self.massageBoxFunc("", u'削除済みです')

    ##
    # @brief ツリーノード作成
    # @param self
    # @param name ノード名
    # @param sel 未使用
    #
    def createNode(self, name, sel):
        tmp = TreeNode(QtGui.QTreeWidgetItem([name]), self)
        self.treeNodeList.append(tmp)
        return tmp


    ##
    # @brief ツリーにネーミングサーバーのオブジェクトを登録する関数
    # @param self 
    #
    def setRTCTree(self):
        if self.mgr != None:
            self.treeWidget.clear()
            self.treeNodeList = []
            
            self.sheetcomboBox.clear()
            for i,j in self.rtc.m_excel.t_xlWorksheet.items():
                self.sheetcomboBox.addItem(i.decode("utf-8"))
            if str(self.sheetcomboBox.currentText().toLocal8Bit()) == u"保存用".encode("cp932"):
                self.sheetcomboBox.setCurrentIndex(1)
            address = str(self.namingServertextBox.text())
            orb = self.mgr._orb
            namingserver = OOoRTC.SetNamingServer(address, orb, self.massageBoxFunc, "Unicode")
            if namingserver == None:
                return
            #print namingserver
            tmp = QtGui.QTreeWidgetItem(["/"])
            self.treeWidget.addTopLevelItem(tmp)
            root = TreeNode(tmp, self)
            self.treeNodeList.append(root)
            
            oChild = self.createNode(address, True)
            root.appendChild(oChild)

            _path = ['/', address]
            self._paths = []

            OOoRTC.rtc_get_rtclist(namingserver, self._paths, _path, oChild, self)
            

            
            
            

            
            
