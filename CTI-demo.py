import os
import sys
import clr
import time
newpath = os.path.dirname(os.path.abspath(__file__)) + '\\DLL'
sys.path.append(newpath)
clr.AddReference("System.Collections")
from System.Collections.Generic import List
from System import Int32
from System import UInt16
from datetime import datetime
classDLL = clr.AddReference('ArbinCTI')
from ArbinCTI.Core.Control import ArbinControlLabView
from ArbinCTI.Core import *

################### Config ###################
class ProgramConst(object):
    __slots__ = ()
    USER = "admin"
    PASSWORD = "000000"
    IP = "192.168.3.8"
    CTI_PORT = 9031

ProgramConst = ProgramConst()
g_programRunning = True
g_IVChannelCount = 0
g_bLogin = False
g_bSchedule = False
g_bBrowse = False
g_bBarcode = False
g_bChannel = False
g_bStop = False
g_client = None
g_bError = False
g_scheduleName = ""
g_channelIndex = 0
g_ctrl = ArbinControlLabView()
g_ctrl.Start()

######################## Response Status ################################
loginResultTokenMap = {
    ArbinCommandLoginFeed.LOGIN_RESULT.CTI_LOGIN_SUCCESS : 'CTI_LOGIN_SUCCESS',
    ArbinCommandLoginFeed.LOGIN_RESULT.CTI_LOGIN_FAILED : 'CTI_LOGIN_FAILED',
    ArbinCommandLoginFeed.LOGIN_RESULT.CTI_LOGIN_BEFORE_SUCCESS : 'CTI_LOGIN_BEFORE_SUCCESS',
}
browseDirectoryResultMap = {
    ArbinCommandBrowseDirectoryFeed.BROWSE_DIRECTORY_RESULT.CTI_BROWSE_DIRECTORY_SUCCESS : 'BROWSE_DIRECTORY_SUCCESS',
    ArbinCommandBrowseDirectoryFeed.BROWSE_DIRECTORY_RESULT.CTI_BROWSE_SCHEDULE_VERSION1_SUCCESS:"CTI_BROWSE_SCHEDULE_VERSION1_SUCCESS",
    ArbinCommandBrowseDirectoryFeed.BROWSE_DIRECTORY_RESULT.CTI_BROWSE_SCHEDULE_SUCCESS : 'BROWSE_SCHEDULE_SUCCESS',
    ArbinCommandBrowseDirectoryFeed.BROWSE_DIRECTORY_RESULT.CTI_BROWSE_DIRECTORY_FAILED : 'BROWSE_DIRECTORY_FAILED',
}
assignChannelTokenMap = {
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_SUCCESS : 'CTI_ASSIGN_SUCCESS',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_INDEX : 'CTI_ASSIGN_INDEX',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_ERROR : 'CTI_ASSIGN_ERROR',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_SCHEDULE_NAME_EMPTY_ERROR : 'CTI_ASSIGN_SCHEDULE_NAME_EMPTY_ERROR',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_SCHEDULE_NOT_FIND_ERROR : 'CTI_ASSIGN_SCHEDULE_NOT_FIND_ERROR',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_CHANNEL_RUNNING_ERROR : 'CTI_ASSIGN_CHANNEL_RUNNING_ERROR',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_CHANNEL_DOWNLOAD_ERROR : 'CTI_ASSIGN_CHANNEL_DOWNLOAD_ERROR',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_BACTH_FILE_OPENED : 'CTI_ASSIGN_BATCH_FILE_OPENED',
    ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_SDU_SAVE_FAILED : 'CTI_ASSIGN_SDU_SAVE_FAILED',
}
assignBarcodeTokenMap = {
    ArbinCommandAssignBarcodeInfoFeed.ASSIGN_BARCODE_RESULT.CTI_ASSIGN_BARCODE_SUCCESS : 'CTI_ASSIGN_BARCODE_SUCCESS',
    ArbinCommandAssignBarcodeInfoFeed.ASSIGN_BARCODE_RESULT.CTI_ASSIGN_BARCODE_CHANNEL_TYPE_NOT_SUPPORT : 'CTI_ASSIGN_BARCODE_CHANNEL_TYPE_NOT_SUPPORT',
    ArbinCommandAssignBarcodeInfoFeed.ASSIGN_BARCODE_RESULT.CTI_ASSIGN_BARCODE_CHANNEL_RUNNING : 'CTI_ASSIGN_BARCODE_CHANNEL_RUNNING',
    ArbinCommandAssignBarcodeInfoFeed.ASSIGN_BARCODE_RESULT.CTI_ASSIGN_BARCODE_CHANNEL_INDEX : 'CTI_ASSIGN_BARCODE_CHANNEL_INDEX'
}
startChannelTokenMap = {
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_SUCCESS : 'CTI_START_SUCCESS',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_INDEX :'CTI_START_INDEX',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_ERROR : 'CTI_START_ERROR',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_CHANNEL_RUNNING : 'CTI_START_CHANNEL_RUNNING',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_CHANNEL_NOT_CONNECT : 'CTI_START_CHANNEL_NOT_CONNECT',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_SCHEDULE_VALID : 'CTI_START_SCHEDULE_VALID',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_NO_SCHEDULE_ASSIGNED : 'CTI_START_NO_SCHEDULE_ASSIGNED',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_SCHEDULE_VERSION : 'CTI_START_SCHEDULE_VERSION',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_POWER_PROTECTED : 'CTI_START_POWER_PROTECTED',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_RESULTS_FILE_SIZE_LIMIT : 'CTI_START_RESULTS_FILE_SIZE_LIMIT',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_STEP_NUMBER : 'CTI_START_STEP_NUMBER',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_NO_CAN_CONFIGURATON_ASSIGNED : 'CTI_START_NO_CAN_CONFIGURATON_ASSIGNED',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_AUX_CHANNEL_MAP : 'CTI_START_AUX_CHANNEL_MAP',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_BUILD_AUX_COUNT : 'CTI_START_BUILD_AUX_COUNT',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_POWER_CLAMP_CHECK : 'CTI_START_POWER_CLAMP_CHECK',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_AI : 'CTI_START_AI',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_SAFOR_GROUPCHAN : 'CTI_START_SAFOR_GROUPCHAN',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_BT6000RUNNINGGROUP : 'CTI_START_BT6000RUNNINGGROUP',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_CHANNEL_DOWNLOADING_SCHEDULE : 'CTI_START_CHANNEL_DOWNLOADING_SCHEDULE',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_DATABASE_QUERY_TEST_NAME_ERROR : 'CTI_START_DATABASE_QUERY_TEST_NAME_ERROR',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_TEXTNAME_EXITS : 'CTI_START_TEXTNAME_EXITS',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_GO_STEP : 'CTI_START_GO_STEP',
    ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_INVALID_PARALLEL : 'CTI_START_INVALID_PARALLEL',
}
stopChannelTokenMap = {
    ArbinCommandStopChannelFeed.STOP_TOKEN.SUCCESS : 'SUCCESS',
    ArbinCommandStopChannelFeed.STOP_TOKEN.STOP_INDEX : 'STOP_INDEX',
    ArbinCommandStopChannelFeed.STOP_TOKEN.STOP_NOT_RUNNING: 'STOP_NOT_RUNNING',
    ArbinCommandStopChannelFeed.STOP_TOKEN.STOP_CHANNEL_NOT_CONNECT : 'STOP_CHANNEL_NOT_CONNECT'
}

############### Event Listener #######################
def LoginFeedbackEvent(cmd):
    if(cmd is  None):
        print("[Demo] Login failed!")
        return
    global g_IVChannelCount
    global g_bLogin
    if(cmd.Result == ArbinCommandLoginFeed.LOGIN_RESULT.CTI_LOGIN_SUCCESS):
        g_bLogin = True 
    if(cmd is not None):
        g_IVChannelCount = cmd.ChannelNum

def AssignSchFeedBackEvent(cmd):
    global g_bSchedule
    global g_bError
    if(cmd.Result == ArbinCommandAssignScheduleFeed.ASSIGN_TOKEN.CTI_ASSIGN_SUCCESS):
        print("[Demo] Assign Schedule Success!")
        g_bSchedule = True
    else:
        print("[Demo] Assign Schedule Failed! Error: {}".format(assignChannelTokenMap[cmd.Result]))
        g_bError = True

def StartFeedEvent(cmd):
    global g_bChannel
    global g_bError
    if(cmd.Result == ArbinCommandStartChannelFeed.START_TOKEN.CTI_START_SUCCESS):
        print(f"[Demo] Start Channel succeeded!")
        g_bChannel = True
    else:
        print("[Demo] Start Channel Failed! Error: {}".format(startChannelTokenMap[cmd.Result]))
        g_bError = True

def StopFeedEvent(cmd):
    global g_bStop
    global g_bError
    if(cmd.Result == ArbinCommandStopChannelFeed.STOP_TOKEN.SUCCESS):
        print(f"[Demo] Channel Stop succeeded!")
        g_bStop = True
    else:
        print("[Demo] Channel Stop Failed! Error: {}".format(stopChannelTokenMap[cmd.Result]))
        g_bError = True

def AssignBarcodeInfoFeedBackEvent(cmd):
    global g_bBarcode
    global g_bError
    if(cmd.BarcodeInfos[0].Error == ArbinCommandAssignBarcodeInfoFeed.ASSIGN_BARCODE_RESULT.CTI_ASSIGN_BARCODE_SUCCESS):
        print(f"[Demo] Assign Barcode succeeded!")
        g_bBarcode = True
    else:
        print("[Demo] Assign Barcode Failed! Error: {}".format(assignBarcodeTokenMap[cmd.Result]))
        g_bError = True

def BrowseFeedEvent(cmd):
    global g_bError
    global g_bBrowse
    global g_scheduleName
    DirInfolist = cmd.DirFileInfoList
    if(cmd.Result != ArbinCommandBrowseDirectoryFeed.BROWSE_DIRECTORY_RESULT.CTI_BROWSE_DIRECTORY_FAILED ):
        print(f"[Demo] Browse schedule succeeded!")
    else:
        print("[Demo] Browse schedule Failed! Error: {}".format(browseDirectoryResultMap[cmd.Result]))
        g_bError = True
    while True:
        print("[Demo] Select schedule file:")
        for idx, file in enumerate(DirInfolist):
            print(f"{idx + 1}: {file.DirFileName}")
        user_input = input("[Demo] Enter the number of the schedule you want to use: ")
        selected_index = int(user_input) - 1
        if 0 <= selected_index < len(DirInfolist):
            selected_file = DirInfolist[selected_index]
            print(f"[Demo] Selected file: {selected_file.DirFileName}")
            g_scheduleName = selected_file.DirFileName
            g_bBrowse = True
            break


######################## Request ################################
def PostLogin():
    global g_client
    global g_ctrl
    if(g_client is None or g_client.IsConnected()):
        g_client = ArbinClient()
        g_client.OnConnectionChanged += Client_OnConnectionChanged
        g_ctrl.LoginFeedEvent += LoginFeedbackEvent
        g_ctrl.AssignSchFeedBackEvent += AssignSchFeedBackEvent
        g_ctrl.ArbinCommandAssignBarcodeInfoFeedBackEvent += AssignBarcodeInfoFeedBackEvent
        g_ctrl.StartFeedEvent += StartFeedEvent
        g_ctrl.StopFeedEvent += StopFeedEvent
        g_ctrl.BrowseFeedEvent += BrowseFeedEvent

        g_ctrl.ListenSocketRecv( g_client )
        result, err = g_client.ConnectAsync(ProgramConst.IP, ProgramConst.CTI_PORT, 0, Int32(0) ); 
        if(result != 0):
            print("[Demo] Please check the network or IP address")
    else:
        print("[Demo] Connected...")

def Client_OnConnectionChanged(socket, e):
    global g_ctrl
    if(e.Connected):
        print("[Demo] Connection to CTI established!")
        # Calling the login command after a network connection
        g_ctrl.PostUserLogin( socket, ProgramConst.USER, ProgramConst.PASSWORD )

def Assign(barcode):
    global g_client
    global g_ctrl
    global g_scheduleName
    g_ctrl.PostAssignSchedule(g_client , g_scheduleName, barcode, 0.0, 0.0, 0.0, 0.0, 0.0, False, g_channelIndex)

def AssignBarcodeInfo(barcode):
    global g_client
    global g_ctrl
    channel_type = ArbinCommandAssignBarcodeInfoFeed.EChannelType.IV
    barcodeinfo = List[ArbinCommandAssignBarcodeInfoFeed.ChannelBarcodeInfo]()
    # Create and populate ChannelBarcodeInfo objects
    info = ArbinCommandAssignBarcodeInfoFeed.ChannelBarcodeInfo()
    info.GlobalIndex = UInt16(g_channelIndex)
    info.Barcode = barcode
    barcodeinfo.Add(info)
    g_ctrl.PostAssignBarcodeInfo(g_client, channel_type, barcodeinfo)

def StartChannel(barcode):
    global g_client
    global g_ctrl
    if(g_client is None or not g_client.IsConnected()):
        print("[Demo] Failed Connect Status.")
        return
    now = datetime.now()
    formatted_datetime = now.strftime('%Y-%m-%d_%H-%M')
    testName = f"{barcode}_{formatted_datetime}"
    channels = List[UInt16]()
    channels.Add(UInt16(g_channelIndex))
    input(f"[Demo] Test will start with test name: {testName}, press Enter to start the test now")
    g_ctrl.PostStartChannel(g_client, testName, channels)
    

def StopChannel():
    global g_client
    global g_ctrl
    global g_channelIndex
    inputStr = (input("[Demo] Press Enter to stop the test now"))
    g_ctrl.PostStopChannel(g_client,g_channelIndex, True)

def BrowseDirectory():
    global g_client
    global g_ctrl    
    g_ctrl.PostBrowseDirectory(g_client, 'SCHEDULE')

########################### Main #################################
try:
    PostLogin()
    while g_bLogin == False:
        time.sleep(0.5)
    print("[Demo] Login succeeded!")

    print(f"[Demo] Welcome to HTE demo!")
    barcode = input("Please the Barcode:")
    AssignBarcodeInfo(barcode)
    while g_bBarcode == False:
        if g_bError:
            print("[Demo] Error! Please restart and try again.")
            exit()
        time.sleep(0.5)

    BrowseDirectory()
    while g_bBrowse == False:
        if g_bError:
            print("[Demo] Error! Please restart and try again.")
            exit()
        time.sleep(0.5)

    Assign(barcode)
    while g_bSchedule == False:
        if g_bError:
            print("[Demo] Error! Please restart and try again.")
            exit()
        time.sleep(0.5)

    StartChannel(barcode)
    while g_bChannel == False:
        if g_bError:
            print("[Demo] Error! Please restart and try again.")
            exit()
        time.sleep(0.5)

    StopChannel()
    while g_bStop == False:
        if g_bError:
            print("[Demo] Error! Please restart and try again.")
            exit()
        time.sleep(0.5)
    print(f"[Demo] Demo ends, thank you.")  
except Exception as e:
    print(f"\n[Demo] Error! Please restart and try again. Error messsage:{e}")
