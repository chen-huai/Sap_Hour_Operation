import sys, win32com.client, time
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import QApplication, QMainWindow
# from PyQt5.QtWidgets import *
# from PyQt5.QtCore import *

def Sap_Operate():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return

        connection = application.Children(0)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        global num
        global m1
        # session.findById("wnd[0]").resizeWorkingPane(172, 38, 0)
        # session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
        # session.findById("wnd[0]").sendVKey(0)
        # session.findById("wnd[0]").sendVKey(0)
        # session.findById(
        #     "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
        # session.findById(
        #     "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
        # session.findById("wnd[0]").sendVKey(2)
        # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
        # num = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]").text
        # m1 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text
        # # session.findById("wnd[0]/tbar[0]/btn[3]").press()
        # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09").select()
        # session.findById(
        #     "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = "Text"
        # session.findById(
        #     "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
        #     4, 4)

        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").key = "ZG"
        a = session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").text
        b = session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,5]").text
        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").text = "01500236"

        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,5]").text = ""
        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW").columns.elementAt(
            4).width = 23
        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").caretPosition = 8
        session.findById("wnd[0]").sendVKey(0)
        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,5]").text = "cai"
        # session.findById(
        #     "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,5]").setFocus()
        # session.findById(
        #     "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,5]").caretPosition = 3
        # session.findById("wnd[0]").sendVKey(0)
        # session.findById("wnd[1]/usr/lbl[5,7]").setFocus()
        # session.findById("wnd[1]/usr/lbl[5,7]").caretPosition = 2
        # session.findById("wnd[1]").sendVKey(0)
        # print(num)
        # print(m1)
        print(a)
        print(b)

    except:
        print(sys.exc_info()[0])

    finally:
        # print(111,num,num)
        session = None
        connection = None
        application = None
        SapGuiAuto = None


if __name__ == "__main__":
    # sapNo = 5010850954
    # projectNo = '66.405.22.1045.01'
    # materialCode = 'T75-405-A2'
    # today = time.strftime('%Y.%m.%d')
    # currencyType = "USD"
    # exchangeRate = 6.3746
    # globalPartnerCode = 1500890
    # csSurname = 6375108
    # salesSurname = 6375108
    revenue = 230
    # m1='1'
    # n1=''
    # cost = 0
    # if ('405' in materialCode) and ("A2" in materialCode):
    #     chmCost = format((revenue * exchangeRate - cost) * 0.3 * 0.5, '.2f')
    #     phyCost = format((revenue * exchangeRate - cost) * 0.3 * 0.5, '.2f')
    #     chmRe = format(revenue * 0.5, '.2f')
    #     phyRe = format(revenue * 0.5, '.2f')
    # elif ('441' in materialCode) and ("A2" in materialCode):
    #     chmCost = format((revenue * exchangeRate - cost) * 0.3 * 0.8, '.2f')
    #     phyCost = format((revenue * exchangeRate - cost) * 0.3 * 0.2, '.2f')
    #     chmRe = format(revenue * 0.8, '.2f')
    #     phyRe = format(revenue * 0.2, '.2f')
    # else:
    #     chmCost = format((revenue * exchangeRate - cost) * 0.3, '.2f')
    #     phyCost = format((revenue * exchangeRate - cost) * 0.3, '.2f')
    #     chmRe = revenue
    #     phyRe = revenue
    # print(chmCost, phyCost, revenue, chmRe, phyRe, revenue * exchangeRate)
    Sap_Operate()
