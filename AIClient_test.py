from pythonnet import load
load("coreclr")

import clr
clr.AddReference("System")
import System
from datetime import datetime
from System.Collections.Generic import List
from System import *
import time
dll_root = "DLL/"
references = [
    #"Newtonsoft.Json",  # dependents: ArbinCTI, ArbinMQ, AIClient
    "ArbinDataModel",   # dependents: ArbinMQ, AIClient
    #"NetMQ",            # dependents: ArbinMQ 
    #"NaCl",             # dependents: ArbinMQ
    #"AsyncIO",          # dependents: ArbinMQ
    "ArbinMQ",
    "ArbinCTI",         # dependents: AIClient
    "AIClient"          
]

def copy_dll_to_runtime(source_dll):
    import os, shutil, site
    # Find the pythonnet runtime folder
    site_packages = site.getsitepackages()
    for path in site_packages:
        runtime_path = os.path.join(path, 'pythonnet', 'runtime')
        if os.path.exists(runtime_path):
            destination = os.path.join(runtime_path, 'ArbinDataModel.dll')
            try:
                shutil.copy2(source_dll, destination)
                print(f"[Demo] Successfully copied ArbinDataModel.dll to {destination}")
                return True
            except Exception as e:
                print(f"[Demo] Error copying DLL: {e}")
    
    print("[Demo] Could not find pythonnet runtime folder")
    return False

#copy_dll_to_runtime("DLL/ArbinDataModel.dll")

for ref in references:
    try:
        # System.Reflection.Assembly.LoadFrom(ref + ".dll")
        clr.AddReference(dll_root + ref)
        print(f"[Demo] {ref}.dll loaded successfully.")
    except System.IO.FileNotFoundException as e:
        print(f"[Demo] Error: {ref}.dll not found.")
        raise e
    except System.IO.FileLoadException as e:
        print(f"[Demo] Error: {ref}.dll or its dependencies failed to load.")
        raise ref
    except Exception as e:
        print(f"[Demo] Error: Failed to load {ref}.dll due to an unexpected error.")
        raise e
    
import AIClient
import Arbin        # ArbinDataModel


###################Config###################
SN = "oct21"
username = "admin"
def mycreateContinueChannelArgs():
    args = Arbin.Library.DataModel.ChannelManagement.ContinueChannelArgs()
    args.SN = SN
    return args

def mycreateAIClientArgs():
    args = AIClient.Core.CreateAIClientArgs()
    args.IPAddress = "127.0.0.1"
    args.UserName = username
    args.Password = "000000"
    args.Timeout = 15000
    return args

def createScheduleArgs(scheduleName):
    args = Arbin.Library.DataModel.TestManagement.AssignFileArgs()
    args.FileName = scheduleName + '.sdx'
    args.SN = SN
    args.FileType = Arbin.Library.DataModel.EAIFileType.Schedule
    args.ChannelIDs = List[System.Int32]()
    args.ChannelIDs.Add(1)
    return args

def createBrowseFileListArgs():
    args = Arbin.Library.DataModel.TestManagement.BrowseFileListArgs()
    args.SN = SN
    args.FileType = Arbin.Library.DataModel.EAIFileType.Schedule
    return args

def createBarcodeArgs(barcode):
    barcodeInfo = Arbin.Library.DataModel.Common.BarcodeInfo()
    barcodeInfo.GlobalID = 1
    barcodeInfo.Barcode = barcode
    barcodeInfo.BarcodeType = Arbin.Library.DataModel.EBarcodeType.IV

    args = Arbin.Library.DataModel.TestManagement.AssignBarcodeInfoArgs()
    args.SN = SN
    args.BarcodeInfos = List[Arbin.Library.DataModel.Common.BarcodeInfo]()
    args.BarcodeInfos.Add(barcodeInfo)
    return args

def createStartChannelArgs(scheduleName,barcode):
    channelResumeData = Arbin.Library.DataModel.Common.ChannelResumeData()
    channelResumeData.ChannelID = 1
    channelResumeData.TestNames = List[String]()
    now = datetime.now()
    formatted_datetime = now.strftime('%Y-%m-%d_%H-%M')
    testName = f"{barcode}_{formatted_datetime}"
    channelResumeData.TestNames.Add(f"{barcode}_{formatted_datetime}")
    channelResumeData.ScheduleName = scheduleName

    args = Arbin.Library.DataModel.ChannelManagement.StartChannelArgs()
    args.Creators = username
    args.SN = SN
    args.ChannelResumeDatas = List[Arbin.Library.DataModel.Common.ChannelResumeData]()
    args.ChannelResumeDatas.Add(channelResumeData)
    return args,testName

###############Event Listener#######################
listenerStatus = 0  # 0 =  not recived, 1 = received success, string = error message
fileList = []
def barCodeEventHandler(*args):
    global listenerStatus
    result = next((x for x in args[0].BarcodeInfos if x.Result != "Success"), None)
    if result == None:
        listenerStatus = 1
    else:
        listenerStatus = result.Result

def fileEventHandler(*args):
    global listenerStatus
    result = args[0].IsSuccess
    if result:
        listenerStatus = 1
    else:
        listenerStatus = args[0].FailedResults[0].Result

def startChannelEventHandler(*args):
    global listenerStatus
    result = args[0].IsSuccess
    if result:
        listenerStatus = 1
    else:
        listenerStatus = args[0].FailedResults[0].Result

def startBrowseFileEventHandler(*args):
    global listenerStatus
    global fileList
    result = args[0].Result
    if result == "Success":
        fileList = []
        for file in args[0].DirFileInfoList:
            fileList.append(file.DirFileName)
        listenerStatus = 1
    else:
        listenerStatus = args[0].FailedResults[0].Result
########################Request################################
args = mycreateAIClientArgs()
err = System.Int32(0)

def getSchedule(client):
    global listenerStatus
    client.OnBrowseFileList += startBrowseFileEventHandler
    bGetSchedule = client.BrowseFileList(createBrowseFileListArgs())
    if bGetSchedule == False:
        print("[Demo] Get schedule failed,please restart and try again.")
        return None
    while listenerStatus == 0:
        time.sleep(0.5)
    if listenerStatus == 1:
        listenerStatus = 0
        print("[Demo] Get schedule succeeded!")
        print("[Demo] Schedule list:")
        for file in fileList:
            print(file)
        return input("\n[Demo] Please enter the schedule name for the test: ")
    else:
        print(f"[Demo] Get schedule failed, error message:{listenerStatus},please restart and try again.")
        return None
    
def assignSchedule(client):
    global listenerStatus
    client.OnAssignFile += fileEventHandler
    userInput = input("\n[Demo] To start the test with schedule \"test.sdx\" (Enter 'yes' or 'no' to show all the schedules): ")
    if userInput == "yes":
        scheduleName = 'test'
    else:
        scheduleName = getSchedule(client)
    if scheduleName == None:
        return None
    bAssign = client.AssignFile(createScheduleArgs(scheduleName))
    if bAssign == False:
        print("[Demo] Assign schedule failed,please restart and try again.")
        return None
    while listenerStatus == 0:
        time.sleep(0.5)
    if listenerStatus == 1:
        listenerStatus = 0
        print(f"[Demo] Assign schedule succeeded! Schedule name: {scheduleName}.sdx")
        return scheduleName
    else:
        print(f"[Demo] Assign schedule failed, error message: {listenerStatus},please restart and try again.")  
        return None
    
def startChannel(client,scheduleName,barcode):
    global listenerStatus
    client.OnStartChannel += startChannelEventHandler
    args,testName = createStartChannelArgs(scheduleName+".sdx",barcode)
    bStartChannel = client.StartChannel(args)
    if bStartChannel == False:
        print("[Demo] Start Channel failed,please restart and try again.")
    while listenerStatus == 0:
        time.sleep(0.5)
    if listenerStatus == 1:
        print(f"[Demo] Start Channel succeeded! Test Name: {testName}")
    else:
        print(f"[Demo] Start Channel failed, error message {listenerStatus},please restart and try again.")
    
def getBarcode(client):
    global listenerStatus
    client.OnAssignBarcodeInfo += barCodeEventHandler
    while True:
        userInputBarcode = input("[Demo] Please enter a barcode:").strip()
        if userInputBarcode.startswith( 'a' ):
            barcodeArg = createBarcodeArgs(userInputBarcode)
            bSend = client.AssignBarcodeInfo(barcodeArg)
            if bSend == False:
                print("[Demo] Assign Barcode failed,please try again.")
                return None
            while listenerStatus == 0:
                time.sleep(0.5)
            if listenerStatus == 1:
                listenerStatus = 0
                print("[Demo] Assign Barcode succeeded!")
                return userInputBarcode
            elif listenerStatus == 2:
                print(f"[Demo] Assign Barcode failed, error message: {listenerStatus},please try again.")
                return None
        else:
            print(f"[Demo] Sorry, '{userInputBarcode}' is a invalid barcode, please try again.")

###########################Main#################################
try:
    client = AIClient.Core.AIClient.CreateAIClient(args, err)
    if client[0] == None:
        print(f"\n[Demo] Connection to AIClient failed! Please restart and try again.")
    else:
        bConnected = client[0].IsConnected()
        if bConnected:
            print(f"[Demo] Connection to AIClient established!")
            print(f"\n[Demo] Welcome to HTE demo!")
            scheduleResult = assignSchedule(client[0])
            if scheduleResult != None:
                barcodeResult = getBarcode(client[0])
                if barcodeResult != None:
                    startChannel(client[0],scheduleResult,barcodeResult)
        else:
            print(f"\n[Demo] Connection to AIClient failed! Please restart and try again.")
except Exception as e:
    print(f"\n[Demo] Error! Please restart and try again. Error messsage:{e}")
