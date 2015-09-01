# -*- coding: utf-8 -*-

##
#
# @file CalcDataPort.py
# @brief 表計算ソフトのRTCのデータポート操作関連のクラス、関数

import optparse
import sys,os,platform
import re


from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL']

    
    
    



import time
import random
import commands
import RTC
import OpenRTM_aist


from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming

import threading

import OOoRTC


##
# @class DataType
# @brief データのタイプ
#

class DataType:
    Single = 0
    Sequence = 1
    Extend = 2

    String = 0
    Value = 1

    DataTypeList = ["TimedDouble","TimedLong","TimedFloat","TimedShort","TimedULong","TimedUShort","TimedChar","TimedWChar",
                        "TimedBoolean","TimedOctet","TimedString","TimedWString","TimedDoubleSeq","TimedLongSeq","TimedFloatSeq",
                        "TimedShortSeq","TimedULongSeq","TimedUShortSeq","TimedCharSeq","TimedWCharSeq","TimedOctetSeq","TimedStringSeq",
                        "TimedWStringSeq","TimedRGBColour","TimedPoint2D","TimedVector2D","TimedPose2D","TimedVelocity2D","TimedAcceleration2D",
                        "TimedPoseVel2D","TimedSize2D","TimedGeometry2D","TimedCovariance2D","TimedPointCovariance2D","TimedCarlike","TimedSpeedHeading2D",
                        "TimedPoint3D","TimedVector3D","TimedOrientation3D","TimedPose3D","TimedVelocity3D","TimedAngularVelocity3D","TimedAcceleration3D",
                        "TimedAngularAcceleration3D","TimedPoseVel3D","TimedSize3D","TimedGeometry3D","TimedCovariance3D","TimedSpeedHeading3D","TimedOAP"]
    def __init__(self):
        pass


##
# @brief
# @param data_type データ型
# @return データオブジェクト、[データ型、データのタイプ、データ型の名前]
def GetDataSType(data_type):
    sig = DataType.Single
    sec = DataType.Sequence
    ext = DataType.Extend

    m_string = DataType.String
    m_value = DataType.Value
    
    if data_type == 'TimedDouble':
        dt = RTC.TimedDouble(RTC.Time(0,0),0)
        return dt, [float, sig, m_value, data_type]
    elif data_type == 'TimedLong':
        dt = RTC.TimedLong(RTC.Time(0,0),0)
        return dt, [long, sig, m_value, data_type]
    elif data_type == 'TimedFloat':
        dt = RTC.TimedFloat(RTC.Time(0,0),0)
        return dt, [float, sig, m_value, data_type]
    elif data_type == 'TimedInt':
        dt = RTC.TimedInt(RTC.Time(0,0),0)
        return dt, [int, sig, m_value, data_type]
    elif data_type == 'TimedShort':
        dt = RTC.TimedShort(RTC.Time(0,0),0)
        return dt, [int, sig, m_value, data_type]
    elif data_type == 'TimedUDouble':
        dt = RTC.TimedUDouble(RTC.Time(0,0),0)
        return dt, [float, sig, m_value, data_type]
    elif data_type == 'TimedULong':
        dt = RTC.TimedULong(RTC.Time(0,0),0)
        return dt, [long, sig, m_value, data_type]
    elif data_type == 'TimedUFloat':
        dt = RTC.TimedUFloat(RTC.Time(0,0),0)
        return dt, [float, sig, m_value, data_type]
    elif data_type == 'TimedUInt':
        dt = RTC.TimedUInt(RTC.Time(0,0),0)
        return dt, [int, sig, m_value, data_type]
    elif data_type == 'TimedUShort':
        dt = RTC.TimedUShort(RTC.Time(0,0),0)
        return dt, [int, sig, m_value, data_type]
    elif data_type == 'TimedChar':
        dt = RTC.TimedChar(RTC.Time(0,0),0)
        return dt, [str, sig, m_string, data_type]
    elif data_type == 'TimedWChar':
        dt = RTC.TimedWChar(RTC.Time(0,0),0)
        return dt, [str, sig, m_string, data_type]
    elif data_type == 'TimedBoolean':
        dt = RTC.TimedBoolean(RTC.Time(0,0),0)
        return dt, [bool, sig, m_value, data_type]
    elif data_type == 'TimedOctet':
        dt = RTC.TimedOctet(RTC.Time(0,0),0)
        return dt, [chr, sig, m_value, data_type]
    elif data_type == 'TimedString':
        dt = RTC.TimedString(RTC.Time(0,0),0)
        return dt, [str, sig, m_string, data_type]
    elif data_type == 'TimedWString':
        dt = RTC.TimedWString(RTC.Time(0,0),0)
        return dt, [str, sig, m_string, data_type]
    elif data_type == 'TimedDoubleSeq':
        dt = RTC.TimedDoubleSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value, data_type]
    elif data_type == 'TimedLongSeq':
        dt = RTC.TimedLongSeq(RTC.Time(0,0),[])
        return dt, [long, sec, m_value, data_type]
    elif data_type == 'TimedFloatSeq':
        dt = RTC.TimedFloatSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value, data_type]
    elif data_type == 'TimedIntSeq':
        dt = RTC.TimedIntSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value, data_type]
    elif data_type == 'TimedShortSeq':
        dt = RTC.TimedShortSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value, data_type]
    elif data_type == 'TimedUDoubleSeq':
        dt = RTC.TimedUDoubleSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value, data_type]
    elif data_type == 'TimedULongSeq':
        dt = RTC.TimedULongSeq(RTC.Time(0,0),[])
        return dt, [long, sec, m_value, data_type]
    elif data_type == 'TimedUFloatSeq':
        dt = RTC.TimedUFloatSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value, data_type]
    elif data_type == 'TimedUIntSeq':
        dt = RTC.TimedUIntSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value, data_type]
    elif data_type == 'TimedUShortSeq':
        dt = RTC.TimedUShortSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value, data_type]
    elif data_type == 'TimedCharSeq':
        dt = RTC.TimedCharSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string, data_type]
    elif data_type == 'TimedWCharSeq':
        dt = RTC.TimedWCharSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string, data_type]
    elif data_type == 'TimedBooleanSeq':
        dt = RTC.TimedBooleanSeq(RTC.Time(0,0),[])
        return dt, [bool, sec, m_value, data_type]
    elif data_type == 'TimedOctetSeq':
        dt = RTC.TimedOctetSeq(RTC.Time(0,0),[])
        return dt, [chr, sec, m_value, data_type]
    elif data_type == 'TimedStringSeq':
        dt = RTC.TimedStringSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string, data_type]
    elif data_type == 'TimedWStringSeq':
        dt = RTC.TimedWStringSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string, data_type]
    elif data_type == "TimedRGBColour":
        dt = RTC.TimedRGBColour(RTC.Time(0,0),RTC.RGBColour(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedPoint2D":
        dt = RTC.TimedPoint2D(RTC.Time(0,0),RTC.Point2D(0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedVector2D":
        dt = RTC.TimedVector2D(RTC.Time(0,0),RTC.Vector2D(0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedPose2D":
        dt = RTC.TimedPose2D(RTC.Time(0,0),RTC.Pose2D(RTC.Point2D(0,0),0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedVelocity2D":
        dt = RTC.TimedVelocity2D(RTC.Time(0,0),RTC.Velocity2D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedAcceleration2D":
        dt = RTC.TimedAcceleration2D(RTC.Time(0,0),RTC.Acceleration2D(0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedPoseVel2D":
        dt = RTC.TimedPoseVel2D(RTC.Time(0,0),RTC.PoseVel2D(RTC.Pose2D(RTC.Point2D(0,0),0),RTC.Velocity2D(0,0,0)))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedSize2D":
        dt = RTC.TimedSize2D(RTC.Time(0,0),RTC.Size2D(0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedGeometry2D":
        dt = RTC.TimedGeometry2D(RTC.Time(0,0),RTC.Geometry2D(RTC.Point2D(0,0),RTC.Size2D(0,0)))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedCovariance2D":
        dt = RTC.TimedCovariance2D(RTC.Time(0,0),RTC.Covariance2D(0,0,0,0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedPointCovariance2D":
        dt = RTC.TimedPointCovariance2D(RTC.Time(0,0),RTC.PointCovariance2D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedCarlike":
        dt = RTC.TimedCarlike(RTC.Time(0,0),RTC.Carlike(0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedSpeedHeading2D":
        dt = RTC.TimedSpeedHeading2D(RTC.Time(0,0),RTC.SpeedHeading2D(0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedPoint3D":
        dt = RTC.TimedPoint3D(RTC.Time(0,0),RTC.Point3D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedVector3D":
        dt = RTC.TimedVector3D(RTC.Time(0,0),RTC.Vector3D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedOrientation3D":
        dt = RTC.TimedOrientation3D(RTC.Time(0,0),RTC.Orientation3D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedPose3D":
        dt = RTC.TimedPose3D(RTC.Time(0,0),RTC.Pose3D(RTC.Point3D(0,0,0), RTC.Orientation3D(0,0,0)))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedVelocity3D":
        dt = RTC.TimedVelocity3D(RTC.Time(0,0),RTC.Velocity3D(0,0,0,0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedAngularVelocity3D":
        dt = RTC.TimedAngularVelocity3D(RTC.Time(0,0),RTC.AngularVelocity3D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedAcceleration3D":
        dt = RTC.TimedAcceleration3D(RTC.Time(0,0),RTC.Acceleration3D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedAngularAcceleration3D":
        dt = RTC.TimedAngularAcceleration3D(RTC.Time(0,0),RTC.AngularAcceleration3D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedPoseVel3D":
        dt = RTC.TimedPoseVel3D(RTC.Time(0,0),RTC.PoseVel3D(RTC.Pose3D(RTC.Point3D(0,0,0), RTC.Orientation3D(0,0,0)),RTC.Velocity3D(0,0,0,0,0,0)))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedSize3D":
        dt = RTC.TimedSize3D(RTC.Time(0,0),RTC.Size3D(0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedGeometry3D":
        dt = RTC.TimedGeometry3D(RTC.Time(0,0),RTC.Geometry3D(RTC.Pose3D(RTC.Point3D(0,0,0), RTC.Orientation3D(0,0,0)),RTC.Size3D(0,0,0)))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedCovariance3D":
        dt = RTC.TimedCovariance3D(RTC.Time(0,0),RTC.Covariance3D(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedSpeedHeading3D":
        dt = RTC.TimedSpeedHeading3D(RTC.Time(0,0),RTC.SpeedHeading3D(0,RTC.Orientation3D(0,0,0)))
        return dt, [str, ext, m_value, data_type]
    elif data_type == "TimedOAP":
        dt = RTC.TimedOAP(RTC.Time(0,0),RTC.OAP(RTC.Velocity3D(0,0,0,0,0,0),RTC.Velocity3D(0,0,0,0,0,0),RTC.Velocity3D(0,0,0,0,0,0)))
        return dt, [str, ext, m_value, data_type]
    else:
        return None, None
    

##
# @brief データポートのデータ型を返す関数
# @param m_port データポート
# @return データオブジェクト、[データ型、データのタイプ、データ型の名前]
#
        
def GetDataType(m_port):
    

    m_string = DataType.String
    m_value = DataType.Value
    
    profile = m_port.get_port_profile()
    
    #props = nvlist_to_dict(profile.properties)
    #data_type =  props['dataport.data_type']
    data_type = OOoRTC.nvlist_getValue(profile.properties,'dataport.data_type')
    if data_type.startswith('IDL:'):
        data_type = data_type[4:]
    colon = data_type.rfind(':')
    if colon != -1:
        data_type = data_type[:colon]

    data_type = data_type.replace('RTC/','')

    return GetDataSType(data_type)

##
# @class CalcPortObject
# @brief 追加するポートのクラス
#


class CalcPortObject:
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
        
        self._port = port
        self._data = data
        self._name = name
        
        
        self._num = int(col)

        
        
        self._row = row
        self._length = mlen
        
        self._sn = sn
        self._port_a = port_a
        self._dataType = m_dataType
        self.buffdata = []
        self.attachports = t_attachports 
        self.state = mstate
        self._col = col

        self.count = 0

        self._mutex = threading.RLock()

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def update_cellName(self, m_cal):
        pass

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSub(self, cell, m_len):
        pass

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSingle(self, cell, m_len):
        pass

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSeq(self, cell, m_len):
        pass

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameEx(self, cell, m_len):
        b = self._name + ":"
        count = [0]
        if self._dataType[3] == "TimedRGBColour":
            if self.input_cellNameEx(b + str("r"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("g"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("b"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedPoint2D":
            if self.input_cellNameEx(b + str("x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("y"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedVector2D":
            if self.input_cellNameEx(b + str("x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("y"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedPose2D":
            if self.input_cellNameEx(b + str("position.x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("heading"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedVelocity2D":
            if self.input_cellNameEx(b + str("vx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("vy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("va"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedAcceleration2D":
            if self.input_cellNameEx(b + str("ax"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("ay"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedPoseVel2D":
            if self.input_cellNameEx(b + str("pose.position.x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.position.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.heading"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.vx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.vy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.va"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedSize2D":
            if self.input_cellNameEx(b + str("l"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("w"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedGeometry2D":
            if self.input_cellNameEx(b + str("pose.position.x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.position.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.heading"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("size.l"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("size.w"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedCovariance2D":
            if self.input_cellNameEx(b + str("xx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xt"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("yy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("yt"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("tt"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedPointCovariance2D":
            if self.input_cellNameEx(b + str("xx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("yy"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedCarlike":
            if self.input_cellNameEx(b + str("speed"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("steeringAngle"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedSpeedHeading2D":
            if self.input_cellNameEx(b + str("speed"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("heading"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedPoint3D":
            if self.input_cellNameEx(b + str("x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("z"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedVector3D":
            if self.input_cellNameEx(b + str("x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("z"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedOrientation3D":
            if self.input_cellNameEx(b + str("r"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("p"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("y"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedPose3D":
            if self.input_cellNameEx(b + str("position.x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.z"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.r"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.p"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.y"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedVelocity3D":
            if self.input_cellNameEx(b + str("vx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("vy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("vz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("vr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("vp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("va"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedAngularVelocity3D":
            if self.input_cellNameEx(b + str("avx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("avy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("avz"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedAcceleration3D":
            if self.input_cellNameEx(b + str("ax"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("ay"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("az"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedAngularAcceleration3D":
            if self.input_cellNameEx(b + str("aax"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("aay"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("aaz"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedPoseVel3D":
            if self.input_cellNameEx(b + str("pose.position.x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.position.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.position.z"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.orientation.r"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.orientation.p"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.orientation.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.vx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.vy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.vz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.vr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.vp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("velocities.va"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedSize3D":
            if self.input_cellNameEx(b + str("l"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("w"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("h"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedGeometry3D":
            if self.input_cellNameEx(b + str("pose.position.x"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.position.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.position.z"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.orientation.r"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.orientation.p"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pose.orientation.y"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("size.l"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("size.w"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("size.h"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedCovariance3D":
            if self.input_cellNameEx(b + str("xx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("xa"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("yy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("yz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("yr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("ya"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("zz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("za"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("rr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("rp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("ra"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("pa"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("aa"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedSpeedHeading3D":
            if self.input_cellNameEx(b + str("speed"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("direction.r"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("direction.p"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("direction.y"), count, m_len, cell) == False:
                return
        if self._dataType[3] == "TimedOAP":
            if self.input_cellNameEx(b + str("orientation.vx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.vy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.vz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.vr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.vp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("orientation.va"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("approach.vx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("approach.vy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("approach.vz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("approach.vr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("approach.vp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("approach.va"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.vx"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.vy"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.vz"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.vr"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.vp"), count, m_len, cell) == False:
                return
            if self.input_cellNameEx(b + str("position.va"), count, m_len, cell) == False:
                return

    ##
    # @brief 
    # @param self 
    # @param b データ名
    # @param count カウンター
    # @param m_len 行の範囲
    # @param cell セルオブジェクト
    # @return 
    def input_cellNameEx(self, b, count, m_len, cell):
        pass
            

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def putData(self, m_cal):
        pass

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def getCell(self, m_cal):
        pass

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        pass

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def putIn(self, m_cal):
        m_string = DataType.String
        m_value = DataType.Value

        tmbd = []
        

        if len(self.attachports) != 0:
            self.buffdata = []
            if self._port.isNew():
                data = self._port.read()
                self.buffdata = [data.data]

        guard = OpenRTM_aist.ScopedLock(self._mutex)
        try:
            tmbd = self.buffdata[:]
            self.buffdata = []
        except:
            pass
        del guard
    
        #tms = len(tmbd)-1
        if self.state:
            pass
        else:
            if len(tmbd) != 0:
                tmbd = [tmbd[-1]]
        for b in tmbd:
            self.updateIn(b, m_cal)

        
                    
    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param sheet シートオブジェクト
    # @param m_cal OOoCalcRTC
    def putOut(self, cell, sheet, m_cal):
        pass

##
# @class CalcInPort
# @brief
#

class CalcInPort(CalcPortObject):
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
        CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def putData(self, m_cal):
        self.putIn(m_cal)

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        pass

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSub(self, cell, m_len):
        self.update_cellNameSingle(cell, m_len)

        
                    
##
# @class CalcInPortSeq
# @brief 
class CalcInPortSeq(CalcPortObject):
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
        CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def putData(self, m_cal):
        self.putIn(m_cal)

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        pass

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSub(self, cell, m_len):
        self.update_cellNameSeq(cell, m_len)

##
# @class CalcInPortEx
# @brief 
class CalcInPortEx(CalcPortObject):
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
        CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def putData(self, m_cal):
        self.putIn(m_cal)

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def update_cellNameSub(self, cell, m_len):
        self.update_cellNameEx(cell, m_len)

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        m_string = DataType.String
        m_value = DataType.Value

        cell, sheet, m_len = self.getCell(m_cal)
        
        
        if cell != None:
            

            count = [0]

            if self.state:
                self._num = self._num + 1

            if self._dataType[3] == "TimedRGBColour":
	        if self.putDataEx(b.r, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.g, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.b, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedPoint2D":
	        if self.putDataEx(b.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.y, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedVector2D":
	        if self.putDataEx(b.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.y, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedPose2D":
	        if self.putDataEx(b.position.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.heading, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedVelocity2D":
	        if self.putDataEx(b.vx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.vy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.va, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedAcceleration2D":
	        if self.putDataEx(b.ax, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.ay, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedPoseVel2D":
	        if self.putDataEx(b.pose.position.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.position.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.heading, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.vx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.vy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.va, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedSize2D":
	        if self.putDataEx(b.l, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.w, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedGeometry2D":
	        if self.putDataEx(b.pose.position.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.position.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.heading, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.size.l, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.size.w, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedCovariance2D":
	        if self.putDataEx(b.xx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xt, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.yy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.yt, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.tt, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedPointCovariance2D":
	        if self.putDataEx(b.xx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.yy, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedCarlike":
	        if self.putDataEx(b.speed, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.steeringAngle, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedSpeedHeading2D":
	        if self.putDataEx(b.speed, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.heading, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedPoint3D":
	        if self.putDataEx(b.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.z, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedVector3D":
	        if self.putDataEx(b.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.z, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedOrientation3D":
	        if self.putDataEx(b.r, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.p, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.y, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedPose3D":
	        if self.putDataEx(b.position.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.z, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.r, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.p, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.y, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedVelocity3D":
	        if self.putDataEx(b.vx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.vy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.vz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.vr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.vp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.va, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedAngularVelocity3D":
	        if self.putDataEx(b.avx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.avy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.avz, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedAcceleration3D":
	        if self.putDataEx(b.ax, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.ay, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.az, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedAngularAcceleration3D":
	        if self.putDataEx(b.aax, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.aay, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.aaz, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedPoseVel3D":
	        if self.putDataEx(b.pose.position.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.position.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.position.z, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.orientation.r, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.orientation.p, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.orientation.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.vx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.vy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.vz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.vr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.vp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.velocities.va, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedSize3D":
	        if self.putDataEx(b.l, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.w, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.h, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedGeometry3D":
	        if self.putDataEx(b.pose.position.x, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.position.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.position.z, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.orientation.r, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.orientation.p, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pose.orientation.y, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.size.l, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.size.w, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.size.h, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedCovariance3D":
	        if self.putDataEx(b.xx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.xa, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.yy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.yz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.yr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.ya, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.zz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.za, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.rr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.rp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.ra, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.pa, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.aa, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedSpeedHeading3D":
	        if self.putDataEx(b.speed, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.direction.r, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.direction.p, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.direction.y, count, m_len, cell, m_value) == False:
	            return
	    if self._dataType[3] == "TimedOAP":
	        if self.putDataEx(b.orientation.vx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.vy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.vz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.vr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.vp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.orientation.va, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.approach.vx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.approach.vy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.approach.vz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.approach.vr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.approach.vp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.approach.va, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.vx, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.vy, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.vz, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.vr, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.vp, count, m_len, cell, m_value) == False:
	            return
	        if self.putDataEx(b.position.va, count, m_len, cell, m_value) == False:
	                return
        
        
        
        

    ##
    # @brief
    # @param self 
    # @param b データ
    # @param count カウンター
    # @param m_len 行の範囲
    # @param cell セルオブジェクト
    # @param d_type データタイプ
    def putDataEx(self, b, count, m_len, cell, d_type):
        return False

        
##
# @class CalcOutPort
# @brief 
class CalcOutPort(CalcPortObject):
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
        CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    ##
    # @brief
    # @param self 
    # @param m_cal OOoCalcRTC
    def putData(self, m_cal):
        cell, sheet, m_len = self.getCell(m_cal)
        
        if cell != None:
            val = self.putOut(cell, sheet, m_cal)
            if self._length == "":
                if val != "":
                    
                    self._data.data = self._dataType[0](val)
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1
            else:
                flag = True
                if val[0] != "":
                    self._data.data = self._dataType[0](val[0])
                else:
                    flag = False

                if flag:
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1

    ##
    # @brief
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSub(self, cell, m_len):
        self.update_cellNameSingle(cell, m_len)

##
# @class CalcOutPortSeq
# @brief 
class CalcOutPortSeq(CalcPortObject):
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
        CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def putData(self, m_cal):
        cell, sheet, m_len = self.getCell(m_cal)

        if cell != None:
            val = self.putOut(cell, sheet, m_cal)
            if self._length == "":
                if val != "":
                    self._data.data = [self._dataType[0](val)]
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1
            else:
                flag = True
                self._data.data = []
                for v in val:
                    if v != "":
                        self._data.data.append(self._dataType[0](v))
                    else:
                        flag = False

                if flag:
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1

    ##
    # @brief 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSub(self, cell, m_len):
        self.update_cellNameSeq(cell, m_len)

##
# @class CalcOutPortEx
# @brief 
#
class CalcOutPortEx(CalcPortObject):
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
        CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    ##
    # @brief
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSub(self, cell, m_len):
        self.update_cellNameEx(cell, m_len)

    ##
    # @brief
    # @param self 
    # @param m_cal OOoCalcRTC
    def putData(self, m_cal):
        
        m_string = DataType.String
        m_value = DataType.Value
        
        

        count = [0]

        

        cell, sheet, m_len = self.getCell(m_cal)
        
        

        if cell != None:
            val = self.putOut(cell, sheet, m_cal)
            
            
            if self._length == "":
                val = [val]


            if self._dataType[3] == "TimedRGBColour":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.r = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.g = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.b = ans
                else:
                    return
            if self._dataType[3] == "TimedPoint2D":
                ans = self.putDataEx(count, val, m_value)
                
                if ans != None:
                    self._data.data.x = ans
                else:
                    return
                
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.y = ans
                else:
                    return
            if self._dataType[3] == "TimedVector2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.y = ans
                else:
                    return
                
            if self._dataType[3] == "TimedPose2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.heading = ans
                else:
                    return
                     
            if self._dataType[3] == "TimedVelocity2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.va = ans
                else:
                    return
                    
            if self._dataType[3] == "TimedAcceleration2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.ax = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.ay = ans
                else:
                    return
                
            if self._dataType[3] == "TimedPoseVel2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.heading = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.vx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.vy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.va = ans
                else:
                    return
                    
            if self._dataType[3] == "TimedSize2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.l = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.w = ans
                else:
                    return

            if self._dataType[3] == "TimedGeometry2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.heading = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.size.l = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.size.w = ans
                else:
                    return
            if self._dataType[3] == "TimedCovariance2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xt = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.yy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.yt = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.tt = ans
                else:
                    return
                
            if self._dataType[3] == "TimedPointCovariance2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.yy = ans
                else:
                    return
                
            if self._dataType[3] == "TimedCarlike":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.speed = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.steeringAngle = ans
                else:
                    return
                
            if self._dataType[3] == "TimedSpeedHeading2D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.speed = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.heading = ans
                else:
                    return
                
            if self._dataType[3] == "TimedPoint3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.z = ans
                else:
                    return
                
            if self._dataType[3] == "TimedVector3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.z = ans
                else:
                    return
                
            if self._dataType[3] == "TimedOrientation3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.r = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.p = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.y = ans
                else:
                    return
                
            if self._dataType[3] == "TimedPose3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.z = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.r = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.p = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.y = ans
                else:
                    return
                    
                
            if self._dataType[3] == "TimedVelocity3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.vp = ans
                else:
                    return
            if self._dataType[3] == "TimedAngularVelocity3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.avx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.avy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.avz = ans
                else:
                    return
                
            if self._dataType[3] == "TimedAcceleration3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.ax = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.ay = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.az = ans
                else:
                    return
                
            if self._dataType[3] == "TimedAngularAcceleration3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.aax = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.aay = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.aaz = ans
                else:
                    return
                
            if self._dataType[3] == "TimedPoseVel3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.z = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.orientation.r = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.orientation.p = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.orientation.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.vx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.vy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.vz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.vr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.vp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.velocities.va = ans
                else:
                    return
                
            if self._dataType[3] == "TimedSize3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.l = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.w = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.h = ans
                else:
                    return
                    
            if self._dataType[3] == "TimedGeometry3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.x = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.position.z = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.orientation.r = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.orientation.p = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pose.orientation.y = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.size.l = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.size.w = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.size.h = ans
                else:
                    return
                
            if self._dataType[3] == "TimedCovariance3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.xa = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.yy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.yz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.yr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.ya = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.zz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.za = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.rr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.rp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.ra = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.pa = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.aa = ans
                else:
                    return
                
            if self._dataType[3] == "TimedSpeedHeading3D":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.speed = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.direction.r = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.direction.p = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.direction.y = ans
                else:
                    return
                
            if self._dataType[3] == "TimedOAP":
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.vx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.vy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.vz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.vr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.vp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.orientation.va = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.approach.vx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.approach.vy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.approach.vz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.approach.vr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.approach.vp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.approach.va = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.vx = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.vy = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.vz = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.vr = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.vp = ans
                else:
                    return
                ans = self.putDataEx(count, val, m_value)
                if ans != None:
                    self._data.data.position.va = ans
                else:
                    return
            OpenRTM_aist.setTimestamp(self._data)
            self._port.write()
            if self.state:
                self._num = self._num + 1

            
        
        
        
        
    ##
    # @brief 
    # @param count カウンター
    # @param val データ配列
    # @param d_type データタイプ
    def putDataEx(self, count, val, d_type):
        
        if count[0] < len(val):
            count[0] += 1
            
            return val[count[0]-1]
        else:
            return None


