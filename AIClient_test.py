from pythonnet import load
load("coreclr")

import clr
clr.AddReference("System")
import System

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


###################CONFIG###################
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

def createStartChannelArgs(scheduleName):
    channelResumeData = Arbin.Library.DataModel.Common.ChannelResumeData()
    channelResumeData.ChannelID = 1
    channelResumeData.TestNames = List[String]()
    channelResumeData.TestNames.Add("Test")
    channelResumeData.ScheduleName = scheduleName

    args = Arbin.Library.DataModel.ChannelManagement.StartChannelArgs()
    args.Creators = username
    args.SN = SN
    args.ChannelResumeDatas = List[Arbin.Library.DataModel.Common.ChannelResumeData]()
    args.ChannelResumeDatas.Add(channelResumeData)
    return args

###############Event Listener#######################
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

listenerStatus = 0  # 0 =  not recived, 1 = received success, string = error message
####################################
args = mycreateAIClientArgs()
err = System.Int32(0)
def getBarcode(client):
    global listenerStatus
    client.OnAssignBarcodeInfo += barCodeEventHandler
    client.OnAssignFile += fileEventHandler
    client.OnStartChannel += startChannelEventHandler
    while True:
        user_input = input("[Demo] Please enter a barcode (or 'q' to quit):").strip()
        if user_input == "q":
            print("\n[Demo] Goodbye!")
            break
        elif user_input.startswith( 'a' ):
            barcodeArg = createBarcodeArgs(user_input)
            bSend = client.AssignBarcodeInfo(barcodeArg)
            if bSend == False:
                print("[Demo] Assign Barcode failed,please try again.")
                continue
            while listenerStatus == 0:
                time.sleep(0.5)
            if listenerStatus == 1:
                listenerStatus = 0
                print("[Demo] Assign Barcode succeeded!")
                more_info = input("\n[Demo] To start the test with \"Test\" (Enter 'yes' or other schedule name): ")
                bAssign,scheduleName = assignSchedule(more_info,client)
                if bAssign == False:
                    print("[Demo] Assign schedule failed,please restart and try again.")
                while listenerStatus == 0:
                    time.sleep(0.5)
                if listenerStatus == 1:
                    listenerStatus = 0
                    print("[Demo] Assign schedule succeeded! Start channel request sent.")

                    bStartChannel = client.StartChannel(createStartChannelArgs(scheduleName+".sdx"))
                    if bStartChannel == False:
                        print("[Demo] Start Channel failed,please restart and try again.")
                    while listenerStatus == 0:
                        time.sleep(0.5)
                    if listenerStatus == 1:
                        print("[Demo] Start Channel succeeded! Test starts")
                    else:
                        print(f"[Demo] Start Channel failed, error message{listenerStatus},please restart and try again.")
                    break
                else:
                    print(f"[Demo] Assign schedule failed, error message: {listenerStatus},please restart and try again.")  
                    break 
            elif listenerStatus == 2:
                print(f"[Demo] Assign Barcode failed, error message: {listenerStatus},please try again.")
        else:
            print(f"[Demo] Sorry, '{user_input}' invalid barcode, please try again.")

def assignSchedule(more_info,client):
    scheduleName = "Test"
    if more_info.lower() == "yes" or more_info == "":
        scheduleName = "Test"
    else:
        scheduleName = more_info
    args = createScheduleArgs(scheduleName)
    bSend = client.AssignFile(args)
    return (bSend,scheduleName)
try:
    #client = AIClient.Core.AIClient()
    client = AIClient.Core.AIClient.CreateAIClient(args, err)
    if client[0] == None:
        print(f"\n[Demo] Connection to AIClient failed! Please restart and try again.")
    else:
        bConnected = client[0].IsConnected()
        if bConnected:
            print(f"[Demo] Connection to AIClient established!")
            print(f"\n[Demo] Welcome to HTE demo!")
            getBarcode(client[0])
        else:
            print(f"\n[Demo] Connection to AIClient failed! Please restart and try again.")
except Exception as e:
    print(f"\n[Demo] Error! Please restart and try again. Error messsage:{e}")
