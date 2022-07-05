import win32com.client
import uiautomation as ua
import subprocess, win32gui, time
import pywintypes


class SAPApplication:

    def OpenSAP(self):

        sap_app = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(sap_app)

    def Connect2SAP_API(self):

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return False

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return 0

        connection = application.Children(0)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return 0

        flag = 0
        while flag == 0:
            try:
                session = connection.Children(0)
                flag = 1
            except:
                time.sleep(0.5)

        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return 0

        return session

    def LogInMenu(self):

        saplogonHwnd = 0
        while saplogonHwnd == 0:
            saplogonHwnd = win32gui.FindWindow("#32770", "SAP Logon 740")  # 借助spy++工具提前得到其类名#32770，和窗口标题SAP Logon 740
            time.sleep(0.1)
        '''如果担心句柄捕获到后，sap界面依然没加载好，可以使用IsWindowVisible进一步确认，直至窗口可见'''
        visibleFlag = False
        while visibleFlag == False:
            time.sleep(0.1)
            visibleFlag = win32gui.IsWindowVisible(saplogonHwnd)

        sapLogonDialog = ua.WindowControl(searchDepth=2, Name='SAP Logon 740')
        testControl = sapLogonDialog.Control(searchDepth=12, Name='T01 [HQ_PRODUCTION]')
        testControl.DoubleClick()
    
    def WindowResize(self, session):

        session.findById("wnd[0]").resizeWorkingPane(164, 36, 0)
    
    def LogInUser(self, session):

        session.findById("wnd[0]/usr/tblSAPMSYSTTC_IUSRACL/btnIUSRACL-BNAME[1,1]").setFocus()
        session.findById("wnd[0]/usr/tblSAPMSYSTTC_IUSRACL/btnIUSRACL-BNAME[1,1]").press()
    
    def DownloadSpecs(self, session, part_no):

        if not session:

            return 'API connect wrongly'
        try:
            self._DownloadSpecs(session, part_no)

        except pywintypes.com_error:
            error_msg = session.findById('/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]').text
            session.findById("wnd[0]").close()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            return error_msg

    def _DownloadSpecs(self, session, part_no):

        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = part_no
        session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 12
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = -1
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/mbar/menu[4]/menu[6]").select()

        groupMembers = session.ActiveWindow.findAllByName("DRAW-DOKAR", "")
        ROH_ID = ''

        for groupMember in groupMembers:
            if groupMember.text.startswith("ROH"):
                ROH_ID = groupMember.id
                break

        while not ROH_ID:
            ROH_ID = self.ScrollAndFind(session, "DRAW-DOKAR", '')

        groupMembers = session.ActiveWindow.findAllByName("DRAW-DOKAR", "")
        for groupMember in groupMembers:
            if groupMember.text.startswith("ROH"):
                ROH_ID = groupMember.id
                break

        session.findById(ROH_ID).setFocus()
        session.findById(ROH_ID).caretPosition = 0

        session.findById("wnd[1]/usr/subSCREEN:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/btnX4").press()
        session.findById("wnd[1]/tbar[0]/btn[12]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    def ScrollAndFind(self, session, name, label):

        session.findById(
            "wnd[1]/usr/subSCREEN:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC").verticalScrollbar.position = 3

        groupMembers = session.ActiveWindow.findAllByName(name, label)
        for groupMember in groupMembers:
            if groupMember.text.startswith("ROH"):
                ROH_ID = groupMember.id
                return ROH_ID

        return False

    def DownloadBOMForm(self, session,part_no):

        if not session:
            return 'API connect wrongly'
        try:
            self._DownloadBOMForm(session, part_no)
        except pywintypes.com_error:
            error_msg = session.findById('/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]').text
            session.findById("wnd[0]").close()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            return error_msg

    def _DownloadBOMForm(self, session, part_no):

        session.findById("wnd[0]/tbar[0]/okcd").text = "YP43"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtP_MATNR").text = part_no
        session.findById("wnd[0]/usr/ctxtP_MATNR").setFocus()
        session.findById("wnd[0]/usr/ctxtP_MATNR").caretPosition = 12
        session.findById("wnd[0]").resizeWorkingPane(164, 36, 0)
        session.findById("wnd[0]/usr/btn%_S_PREFIX_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = '{}.xls'.format(part_no)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()


if __name__=='__main__':
    
    pass