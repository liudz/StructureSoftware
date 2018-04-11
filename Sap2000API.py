#! python3  
# -*- coding: gbk -*- 

import os
import sys
import comtypes.client
import pdb  #调整用的，如果不用的话，注释掉它，n(ext)

#新建一个工程
AttachToInstance = False

#指定路径。如果有多版本，令最新版本的sap2000启动
SpecifyPath = False

#指定路径为真时的路径
ProgramPath='D:\Program Files\Computers and Structures\SAP2000 19\SAP2000.exe'

#工程保存目录
#如果指定的文件夹不存在，就创建一个
APIPath = 'D:\CSiAPI案例'
if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass
ModelPath = APIPath + os.sep + 'API_1-001.sdb'

if AttachToInstance:
    try:
        #获取当前sap2000对象
        mySapObject = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
    except (OSError, comtypes.COMError):
        print("没有找到或者附加到程序的实例上。")
        sys.exit(-1)
else:
    #创建 api helper 对象
    helper = comtypes.client.CreateObject('SAP2000v19.Helper')
    helper = helper.QueryInterface(comtypes.gen.SAP2000v19.cHelper)
    if SpecifyPath:
        try:
            #从指定路径创建一个sap对象的实例
            mySapObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("不能从" + ProgramPath + "路径创建一个程序新实例")
            sys.exit(-1)
    else:
        try:
            #从最新安装的sap2000创建一个sap2000对象实例
            mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
        except (OSError, comtypes.COMError):
            print("不能创建一个新的程序实例。")
            sys.exit(-1)
    
    #启动sap2000程序
    mySapObject.ApplicationStart()
    
    #创建SapModel对象
    SapModel = mySapObject.SapModel
    
    #初始化模型，设置单位制
    #kN_mm_C = 5
    #kN_m_C = 6
    #N_mm_C = 9
    #N_m_C = 10
    N_mm_C = 9
    SapModel.InitializeNewModel(N_mm_C)
    
    #创建新的空白模型
    ret = SapModel.File.NewBlank()
    
    #定义材料属性
    MATERIAL_CONCRETE = 2
    ret = SapModel.PropMaterial.SetMaterial('CONC', MATERIAL_CONCRETE)
    
    #设置材料的各向同性力学属性，取《混规》里C30的属性
    #名称，弹性模量，泊松比，线膨胀系数
    ret = SapModel.PropMaterial.SetMPIsotropic('CONC', 30000, 0.2, 0.00001)
    
    #设置矩形截面属性，高×宽
    ret = SapModel.PropFrame.SetRectangle('R1', 'CONC', 300, 300)
    
    #设置截面属性调整系数
    #横截的(轴向)面积，2方向的剪切面积，3方向的剪切面积，扭转常数，围绕2轴的惯性矩，围绕3轴的惯性矩，质量，重量
    ModValue = [1000, 0, 0, 1, 1, 1, 1, 1]
    ret = SapModel.PropFrame.SetModifiers('R1', ModValue)
    
    #改变单位制
    kN_m_C = 6
    ret = SapModel.SetPresentUnits(kN_m_C)
    
    #通过坐标定位，加框架
    #xi, yi, zi, xj, yj, zj, Name, PropName, CSys
    FrameName1 = ' '
    FrameName2 = ' '
    FrameName3 = ' '
    [FrameName1, ret] = SapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 3, FrameName1, 'R1', '1', 'Global')
    [FrameName2, ret] = SapModel.FrameObj.AddByCoord(0, 0, 3, 2.5, 0, 5, FrameName1, 'R1', '2', 'Global')
    [FrameName3, ret] = SapModel.FrameObj.AddByCoord(-1.2, 0, 3, 0, 0, 3, FrameName1, 'R1', '3', 'Global')

    #在支座处设置节点约束支座
    PointName1 = ' '
    PointName2 = ' '
    #U1, U2, U3, R1, R2, R3
    Restraint = [True, True, True, True, False, False]
    [PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName1, PointName1, PointName2)
    #在0,0,0处
    ret = SapModel.PointObj.SetRestraint(PointName1, Restraint)
    
    #在顶部设置节点约束支座
    Restraint = [True, True, False, False, False, False]
    [PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName2, PointName1, PointName2)
    ret = SapModel.PointObj.SetRestraint(PointName2, Restraint)

    #刷新视图，更新到默认缩放值
    ret = SapModel.View.RefreshView(0, False)

    #增加荷载类型，把sap2000的荷载都列上了，它的荷载类型真多呀
    #LTYPE_DEAD = 1
    #LTYPE_SUPERDEAD = 2
    #LTYPE_LIVE = 3
    #LTYPE_REDUCELIVE = 4
    #LTYPE_QUAKE = 5
    #LTYPE_WIND= 6
    #LTYPE_SNOW = 7
    #LTYPE_OTHER = 8
    #LTYPE_MOVE = 9
    #LTYPE_TEMPERATURE = 10
    #LTYPE_ROOFLIVE = 11
    #LTYPE_NOTIONAL = 12
    #LTYPE_PATTERNLIVE = 13
    #LTYPE_WAVE= 14
    #LTYPE_BRAKING = 15
    #LTYPE_CENTRIFUGAL = 16
    #LTYPE_FRICTION = 17
    #LTYPE_ICE = 18
    #LTYPE_WINDONLIVELOAD = 19
    #LTYPE_HORIZONTALEARTHPRESSURE = 20
    #LTYPE_VERTICALEARTHPRESSURE = 21
    #LTYPE_EARTHSURCHARGE = 22
    #LTYPE_DOWNDRAG = 23
    #LTYPE_VEHICLECOLLISION = 24
    #LTYPE_VESSELCOLLISION = 25
    #LTYPE_TEMPERATUREGRADIENT = 26
    #LTYPE_SETTLEMENT = 27
    #LTYPE_SHRINKAGE = 28
    #LTYPE_CREEP = 29
    #LTYPE_WATERLOADPRESSURE = 30
    #LTYPE_LIVELOADSURCHARGE = 31
    #LTYPE_LOCKEDINFORCES = 32
    #LTYPE_PEDESTRIANLL = 33
    #LTYPE_PRESTRESS = 34
    #LTYPE_HYPERSTATIC = 35
    #LTYPE_BOUYANCY = 36
    #LTYPE_STREAMFLOW = 37
    #LTYPE_IMPACT = 38
    #LTYPE_CONSTRUCTION = 39
    #Name, MyType, SelfWeightMultiplier, AddLoadCase(线性静力荷载类型)
    LTYPE_OTHER = 8
    ret = SapModel.LoadPatterns.Add('1', LTYPE_OTHER, 1, True)
    ret = SapModel.LoadPatterns.Add('2', LTYPE_OTHER, 0, True)
    ret = SapModel.LoadPatterns.Add('3', LTYPE_OTHER, 0, True)
    ret = SapModel.LoadPatterns.Add('4', LTYPE_OTHER, 0, True)
    ret = SapModel.LoadPatterns.Add('5', LTYPE_OTHER, 0, True)
    ret = SapModel.LoadPatterns.Add('6', LTYPE_OTHER, 0, True)
    ret = SapModel.LoadPatterns.Add('7', LTYPE_OTHER, 0, True)
    
    #为荷载类型2，指定荷载
    [PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
    #点荷载
    #F1, F2, F3, M1, M2, M3, 默认整体坐标系
    PointLoadValue = [0, 0, -2, 0, 0, 0]
    ret = SapModel.PointObj.SetLoadForce(PointName1, '2', PointLoadValue)
    
    #线荷载
    #Name, LoadPat, MyType(1=单位长度分布力, 2=单位长度分布弯矩)
    #Dir，确定一个方向，选项也真多呀
    #1=局部1轴
    #2=局部2轴
    #3=局部3轴
    #4=整体X向
    #5=整体Y向
    #6=整体Z向
    #7=整体投影X向
    #8=整体投影Y向
    #9=整体投影Z向
    #10=整体重力方向
    #11=整体投影重力方向
    #Dist1(从框架起点到荷载起点的距离), Dist2(从框架起点到荷载终点的距离)
    #Val1(均布荷载的起始值), Val2(均布荷载的终止值), CSys, RelDist(默认是相对距离)
    ret = SapModel.FrameObj.SetLoadDistributed(FrameName3, '2', 1, 10, 0, 1, 2.5, 2.5)
    
    #为荷载类型3，指定荷载
    #点荷载，同时加力和弯矩，默认整体坐标系
    [PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
    PointLoadValue = [0, 0, -3, 0, -3.5, 0]
    ret = SapModel.PointObj.SetLoadForce(PointName2, '3', PointLoadValue)
    
    #为荷载类型4，指定荷载
    #Dir(11=投影重力方向)，默认整体坐标系
    ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '4', 1, 11, 0, 1, 4, 4)

    #为荷载类型5，指定荷载
    #Dir(2=局部2轴方向)
    ret = SapModel.FrameObj.SetLoadDistributed(FrameName1, '5', 1, 2, 0, 1, 5, 5, 'Local')
    ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '5', 1, 2, 0, 1, -5.5, -5.5, 'Local')

    #从这里开始debug
    pdb.set_trace()
    
    #为荷载类型6，指定荷载
    ret = SapModel.FrameObj.SetLoadDistributed(FrameName1, '6', 1, 2, 0, 1, 6.5, 6, 'Local')
    ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '6', 1, 2, 0, 1, -6, 0, 'Local')

    #为荷载类型7，指定荷载
    #线单元上加点荷载
    #Name, LoadPat, Mytype(1=力, 2=弯矩), Dir(2=局部2轴), Dist(相对绝对距离), Val, CSys, RelDist
    ret = SapModel.FrameObj.SetLoadPoint(FrameName2, '7', 1, 2, 0.5, -7, 'Local')

    #转换单位制
    N_mm_C = 9
    ret = SapModel.SetPresentUnits(N_mm_C)

    #保存模型
    ret = SapModel.File.Save(ModelPath)

    #运行模型(这步骤会创建分析模型)
    ret = SapModel.Analyze.RunAnalysis()

    #初始化Sap2000结果
    SapResult= [0, 0, 0, 0, 0, 0, 0]
    [PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName2, PointName1, PointName2)

    #获得1~7荷载类型的Sap2000结果
    for i in range(0, 7):
        NumberResults = 0
        Obj = []
        Elm = []
        ACase = []
        StepType = []
        StepNum = []
        U1 = []
        U2 = []
        U3 = []
        R1 = []
        R2 = []
        R3 = []
        ObjectElm = 0
        #取消所有输出的选择的荷载和组合
        ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        #输出1~7号荷载
        ret = SapModel.Results.Setup.SetCaseSelectedForOutput(str(i + 1))
        #0~3类型，
        if i <= 3:
            #点对象的相对位移
            [NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret] = SapModel.Results.JointDispl(PointName2, ObjectElm, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
            SapResult[i] = U3[0]
        #4~6类型，
        else:
            [NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret] = SapModel.Results.JointDispl(PointName1, ObjectElm, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
            SapResult[i] = U1[0]
    #关闭Sap2000
    ret = mySapObject.ApplicationExit(False)
    SapModel = None
    mySapObject = None

    #填充独立结果
    IndResult= [0, 0, 0, 0, 0, 0, 0]
    IndResult[0] = -0.02639
    IndResult[1] = 0.06296
    IndResult[2] = 0.06296
    IndResult[3] = -0.2963
    IndResult[4] = 0.3125
    IndResult[5] = 0.11556
    IndResult[6] = 0.00651

    #填充百分比差
    PercentDiff = [0, 0, 0, 0, 0, 0, 0]
    for i in range(0, 7):
        PercentDiff[i] = (SapResult[i] / IndResult[i]) - 1

    #显示结果
    for i in range(0,7):
        print()
        print(SapResult[i])
        print(IndResult[i])
        print(PercentDiff[i]) 
