on ERROR RESUME NEXT
do
    if Not IsObject(application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
        Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject application, "on"
    End If

    if err.number <> 0 then
        resSAP = msgbox("SAP was not detected please open a SAP session.  Thank you.", vbOkCancel)
        err.clear
        if resSAP = vbCancel then
            wscript.quit
        end if
    else
        exit do
    end if
loop
on error goto 0

isSPEX = False
newSPEX = False
version10 = False
STAYLABOREDON = False
isConverted = False
isExchange = False
isDPMI = False
OOPS = False
dim tmpServOrder
dim tmpCust
dim tmpCustomerNUM
dim salesvalue
dim tmpDELBLCK
dim tmpZHStatus
dim tmpZGStatus
dim tmpMODSIN
dim tmpMODSOUT
dim tmpSFMODSIN
dim tmpSFMODSOUT
dim tmpSWVERSIONSIN
dim tmpSWVERSIONSOUT
dim tmpPN


UserName = session.Info.User

strText = "Oh no"
set objVoice = CreateObject("SAPI.SpVoice")

Call Main

Sub Main
    dim msgBoxResult

    dim tmpSN

    Call Welcome

    isSPEX = False
    newSPEX = False
    version10 = False
    STAYLABOREDON = False

    strText = "Oh no"

    call GetServiceOrderInput(tmpServOrd)

    Call OpenZiwbn(tmpServOrd)

    Call GetIW32Information(tmpServOrd, isSPEX)

    Call chkIfConverted

    Call LaborOnOlathe

    Call GetPartNumberSerialNumber(tmpPn, tmpSn)
    on error resume next
    Call AskAboutPartNumberSerialNumber(tmpPn, tmpSn)

    on ERROR GOTO 0

    Call chkREPAIRLEVEL

    Call CheckANLYCMPL

    Call AskAboutPaperwork

    Call AskAboutHardware

    if isSPEX = False then
        Call checkQT

        Call AskAboutCustomerReq

        Call Z8Notifications

        Call checkWARRANTY

        Call checkZTASKS

    End if

    Call AskAboutUSR

    Call AskAboutDoesTestSheetMatchUnit

    Call AskAboutDoesTestSheetShowAnyFails

    Call AskAboutDateAndSignatureOnTestSheet

    Call AskAboutOperatorComments

    Call CheckGAT

    Call CheckAuthDocs

    Call doAsMaintain

    'Call doPartsLinking

    Call getActualFindingInfo(tmpMODSIN, tmpMODSOUT)

    Call UpdateWandingStatus

    Call UpdateWSUPDComments

    Call LaborOffCompleteOlathe

    Call openFinalInspection

    Call PrintServiceReport
end sub
'end sub main



sub Welcome
    res = msgbox("Welcome to the Close-Up script." & vbcrlf & vbcrlf & "Please read the SETUP document located on the Share drive -> CSC DATABASES-> Close-UP Script for instructions on how to avoid issues while using this script." & vbcrlf & vbcrlf & "Rev O-3", vbOKCancel, "Welcome")
    if res = VBCancel then
        wscript.quit
    End if
End Sub

Sub GetServiceOrderInput(servOrd)
    servOrd = InputBox("Enter service order number.", "Close-Up")
    If servOrd = "" then
        wscript.quit
    End If
End Sub

Sub OpenZIWBN(tmpServOrd)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZIWBN"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text = tmpServOrd
    session.findById("wnd[0]").sendVKey 0
End Sub

Sub GetIW32Information(servOrd, boolSPEX)

    tmpCustomerNUM = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM").text
    tmpSuperiorOrder = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text

    if tmpCustomerNUM = "PLANT1133" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "SLSR01" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1057" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1052" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1013" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1103" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1116" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1005" then
        boolSPEX = TRUE
    end if

End Sub

sub chkIfConverted

    dim checkPASS
    dim tmpSTATUS2
    checkPASS = "CONV"
    checkExch = "EXCH"

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    tmpSTATUS2 = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
    result2 = InStr(tmpSTATUS2, checkPass)
    result3 = Instr(tmpSTATUS2, checkExch)
    IF result2 > 0 THEN
        isConverted = True
    END IF
    if result3 > 0 then
        isExchange = True
    end if
end sub

sub LaborOnOlathe
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellRow = -1
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&FILTER"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Close Up*"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 6
    session.findById("wnd[1]").sendVKey 0


    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton "LABON"

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").setCurrentCell - 1, "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&FILTER"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = ""
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 0
    session.findById("wnd[1]/tbar[0]/btn[0]").press

end sub

Sub GetPartNumberSerialNumber(pn, sn)
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H").select
    on error resume next
    err.clear
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").currentCellColumn = "MATNR"
    if err.number <> 0 then
        version10 = True
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").currentCellColumn = "MATNR"
        pn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "MATNR")

        sn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "SERNR")

        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    else
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").currentCellColumn = "MATNR"
        pn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "MATNR")

        sn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "SERNR")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select

    end if
    on error goto 0
    err.clear
End Sub

Sub AskAboutPartNumberSerialNumber(pn, sn)
    snTries = 0
    pnTries = 0
    inputPNRes = ""
    inputSNRes = ""
    res = ""

    res1 = msgbox("Does the Part Number match the ID plate on the unit and the outgoing Part Number in SAP?", VBYesNo, "Check")
    if res1 = vbNo then
        objVoice.Speak strText
        wscript.quit
    end if

    res2 = msgbox("Does the Serial Number match the ID plate on the unit and the outgoing Serial Number in SAP?", VBYesNo, "Check")
    if res1 = vbNo then
        objVoice.Speak strText
        wscript.quit
    end if

    Do While pn <> inputPNRes
        inputPNRes = InputBox("Please enter the Part Number from the Unit being inspected.")
        if pn = inputPNRes then
            Exit Do
        else

            if pnTries > 1 then
                objVoice.Speak strText
                MsgBox("The Part Number being tried does not appear to match SAP. This script will now terminate.")
                wscript.quit
            end if
            objVoice.Speak strText
            MsgBox("The Part Number entered does not match the Part Number of this record in SAP. Please double check your entry.")
            pnTries = pnTries + 1
        end if
    Loop

    dim result
    dim tmpSTATUS
    dim SearchChar

    SearchChar = "DPMI"

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    tmpSTATUS = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
    result = InStr(tmpSTATUS, SearchChar)
    If result = 0 Then
        Do While UCASE(sn) <> UCASE(inputSNRes)
            inputSNRes = InputBox("Please enter the Serial Number from the unit being inspected.")
            If UCASE(sn) = UCASE(inputSNRes) then
                Exit Do
            else
                if snTries > 1 Then
                    objVoice.Speak strText
                    Msgbox("The Serial Number being tried does not appear to match SAP. This script will now terminate.")
                    wscript.quit
                end if
                objVoice.Speak strText
                MsgBox("The Serial Number entered does not match the Serial Number of this record in SAP. Please double check your entry.")
                snTries = snTries + 1
            End If
        Loop
    end if
    If result <> 0 Then
        isDPMI = True
    End If
End Sub

sub chkREPAIRLEVEL
    do
        tmpREPLEVEL = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtW_REPAIR_LEVEL").text
        if tmpREPLEVEL = "" then
            res = msgbox("The repair level is empty, please input the repair level.", vbokcancel)
            if res = vbcancel then
                wscript.quit
            end if
        else
            Exit Do
        end if
    loop

End Sub

Sub CheckANLYCMPL
    dim result
    dim tmpSTATUS
    dim SearchChar

    SearchChar = "ANLY"

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    tmpSTATUS = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
    result = InStr(tmpSTATUS, SearchChar)
    If result = 0 Then
        resANLY = msgbox("!!ALERT!!" & vbcrlf & "ANLY CMPL has not been done, is this a test only?", vbYesNo, "ANLY CMPL")
        if resANLY = vbNo then
            do
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
                tmpSTATUS = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
                result = InStr(tmpSTATUS, SearchChar)
                If result = 0 Then
                    OOPSres = msgbox("ANLY CMPL must be done to continue, please hit OK when fixed.", vbOkCancel, "ANLY CMPL")
                    if OOPSres = vbCancel then
                        wscript.quit
                    else
                    end if
                else
                    exit do
                end if
            loop
        End If
    end if
End Sub

Sub AskAboutPaperwork
    res = msgbox("Are all needed paperwork present and correct?" & vbcrlf & "(Pick Tickets, Test Sheets, Customer Documents, Potential DOA form filled out if needed, etc..,)", vbYesNo, "Check")
    if res = VBNo then
        objVoice.Speak strText
        msgbox("!!ALERT!!" & vbcrlf & "Paperwork needs to be present and correct, TERMINATING SCRIPT.")
        wscript.quit
    end if
End Sub

Sub AskAboutHardware
    res = MsgBox("Has all hardware been checked?" & vbcrlf & "(No loose/missing screws, knobs turn freely, handles tight, etc...)", vbYesNo, "Hardware Status")
    If res = vbNo then
        objVoice.Speak strText
        MsgBox("!!ALERT!!" & vbcrlf & "Hardware Failure. TERMINATING SCRIPT.")
        wscript.quit
    End If
End Sub

sub checkQT

    dim tmpQTStatus
    dim SearchChar

    SearchChar = "ACC"

    tmpQTStatus = ""

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpADMIN_H").select
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpADMIN_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0214/txtYAFS_ZIWBN_HEADER-ADMIN_QUOTE").setFocus
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpADMIN_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0214/txtYAFS_ZIWBN_HEADER-ADMIN_QUOTE").caretPosition = 1
    on error resume next
    err.clear
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    if err.number <> 0 then
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        err.clear
        Exit Sub
    end if
    err.clear
    on error goto 0
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10").select
    tmpQTStatus = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4305/txtRV45A-ASTTX").text

    if tmpQTStatus = SearchChar then
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        Exit Sub
    else
        msgbox("!!ALERT!!" & vbcrlf & "Unit is not quote approved!  Please get quote approval before shipping unit! Ending script.")
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        wscript.quit
    end if

end sub

Sub AskAboutCustomerReq
    CRDres = msgbox("Have you reviewed the PO and Customer Requirements Database for special instructions?", vbYesNo, "Customer Requirements")
    if CRDres = vbNo then
        objVoice.Speak strText
        msgbox("!!ALERT" & vbcrlf & "You must review these for customers, TERMINATING SCRIPT.")
        wscript.quit
    end if

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SALES_ORD").setFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]").sendVKey 0

    salesvalue = session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBAK-NETWR").text


    if CNA = True then
        if salesvalue <> 0 then
            RES = msgbox("Does Net value match traveler?", vbYesNo, "Net Value")
            if RES = vbNo then
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                wscript.quit
            end if
        else
            msgbox("Script ending, needs money in sales order before continuing.")
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            wscript.quit
        end if
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    else
        res = MsgBox("Does the dollar amount in the Sales Order match the funded quote or funded PO?" & vbcrlf & "(Units that are warranty or MSA should be a 0$ value)", vbYesNo, "Funded Sales Order")
        if res = vbNo then
            objVoice.Speak strText
            MsgBox("!!ALERT!!" & vbcrlf & "The sales order is either not funded, or does not match the notes on the traveler. TERMINATING SCRIPT.")
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            wscript.quit
        End If

        tmpDELBLCK = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key
        if tmpDELBLCK = "" then

        else
            msgbox("There is a Delivery Block, before going to QA this Delivery Block must be removed, see your workflow/supervisor for help.")
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            wscript.quit
        end if

        session.findById("wnd[0]/tbar[0]/btn[3]").press
    End if
End Sub

Sub Z8Notifications
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/btnG_H_SERORD_BT_ALERT").press


    res = MsgBox("Have all applicable Z8 Notifications been complied with?", vbYesNo, "Z8 Compliance")
    If res = vbNo then
        objVoice.Speak strText
        MsgBox("!!ALERT!!" & vbcrlf & "Z8 Notifications have not been fully complied with. TERMINATING SCRIPT.")
        session.findById("wnd[1]").close
        wscript.quit
    End If

    session.findById("wnd[1]").close
End Sub

sub checkWARRANTY
    if version10 = True then
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H").select
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont[0]/shell").pressToolbarButton "&MB_FILTER"
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont[0]/shell").currentCellRow = 2
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont[0]/shell").selectedRows = "2"
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
        session.findById("wnd[2]").sendVKey 4
        session.findById("wnd[3]").sendVKey 2
        session.findById("wnd[2]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont[0]/shell").currentCellColumn = ""
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont[0]/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont[0]/shell").pressToolbarButton "CHANGE"
        session.findById("wnd[0]").sendVKey 0
    else
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H").select
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").pressToolbarButton "&MB_FILTER"
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").currentCellRow = 2
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").selectedRows = "2"
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
        session.findById("wnd[2]").sendVKey 4
        session.findById("wnd[3]").sendVKey 2
        session.findById("wnd[2]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").currentCellColumn = ""
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").pressToolbarButton "CHANGE"
        session.findById("wnd[0]").sendVKey 0

    end if

end sub

sub checkZTASKS
    IF version10 = True THEN
        session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").select
        session.findById("wnd[0]").sendVKey 0
        zres = msgbox("Are all tasks closed?", vbYesNo, "ZTasks")
        if zres = vbNo then
            objVoice.Speak strText
            msgbox("!!ALERT!!" & vbcrlf & "Seek guidance on what to do with the open task, TERMINATING SCRIPT.")
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont[0]/shell").pressToolbarButton "&MB_FILTER"
            session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_FL_SING").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
            script.quit
        end if
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont[0]/shell").pressToolbarButton "&MB_FILTER"
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_FL_SING").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    else
        session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").select
        session.findById("wnd[0]").sendVKey 0
        zres = msgbox("Are all tasks closed?", vbYesNo, "ZTasks")
        if zres = vbNo then
            objVoice.Speak strText
            msgbox("!!ALERT!!" & vbcrlf & "Seek guidance on what to do with the open task, TERMINATING SCRIPT.")
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").pressToolbarButton "&MB_FILTER"
            session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_FL_SING").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
            wscript.quit
        end if
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").pressToolbarButton "&MB_FILTER"
        session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_FL_SING").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    END IF
end sub

sub AskAboutUSR
    USRres = msgbox("Are all fields complete on the Unit Status Report?" & vbcrlf & "(Sales $, Mods, Inprocess Inspections, squawk addressed and special requirements have been accomplished, etc ..,)", vbYesNo, "USR Details")
    If USRres = vbNo then
        objVoice.Speak strText
        MsgBox("!!ALERT!!" & vbcrlf & "Unit Status Report must be filled out correctly. TERMINATING SCRIPT.")
        wscript.quit
    End If
end sub

Sub AskAboutDoesTestSheetMatchUnit
    res = MsgBox("Is the test sheet header information correct? P/N, S/N, Cal Dates, etc...", vbYesNo, "Test Sheet Header Status")
    If res = vbNo then
        MsgBox("Test Sheet Header Failure. Script terminated.")
        wscript.quit
    End If
End Sub

Sub AskAboutDoesTestSheetShowAnyFails
    res = MsgBox("Did all tests pass or have been corrected on test sheet?", vbYesNo, "Test Sheet Data Status")
    If res = vbNo then
        MsgBox("Test Sheet Data Failure. Script terminated.")
        wscript.quit
    End If
End Sub

Sub AskAboutDateAndSignatureOnTestSheet
    res = MsgBox("Was the test sheet signed and dated?", vbYesNo, "Test Sheet Signature Status")
    If res = vbNo then
        MsgBox("Test Sheet Signature Failure. Script terminated.")
        wscript.quit
    End If
End Sub

sub AskAboutOperatorComments
    USRres = msgbox("Are the Operation comments filled out on each labor line?", vbYesNo, "OP comments")
    If USRres = vbNo then
        objVoice.Speak strText
        MsgBox("!!ALERT!!" & vbcrlf & "Operation comments must be filled out. TERMINATING SCRIPT.")
        wscript.quit
    End If
end sub

Sub CheckGAT
    dim result
    dim tmpSTATUS
    dim SearchChar

    SearchChar = "GAT3"

    do
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        tmpSTATUS = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
        result = InStr(tmpSTATUS, SearchChar)
        If result = 0 Then
            resANLY = msgbox("!!ALERT!!" & vbcrlf & "Unit is not in GAT3.  Please fix before continuing, then hit OK.", vbOkCancel, "GAT3")
            if resANLY = vbCancel then
                wscript.quit
            end if
        else
            exit do
        End If
    loop
End Sub

Sub CheckAuthDocs
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpAUTHDOCS_I").select
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpAUTHDOCS_I").select
    authdocsres = MsgBox("Do the documents used in AuthDocs match the work performed? Are all Auth docs completed?", VBYesNo, "Check Auth Docs")
    if authdocsres = VBNo then
        objVoice.Speak strText
        msgbox("!!ALERT!!" & vbcrlf & "Script will be terminated")
        wscript.quit
    end if
End Sub

Sub doAsMaintain
    dim checkPASS
    dim tmpSTATUS2
    checkPASS = "Z107"
    checkOVERRIDE = "Z109"

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    tmpSTATUS2 = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
    result2 = InStr(tmpSTATUS2, checkPass)
    result3 = InStr(tmpSTATUS2, checkOVERRIDE)
    IF result2 > 0 THEN
        EXIT SUB
    END IF
    IF result3 > 0 THEN
        EXIT SUB
    END IF

    if isConverted = True then
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I").select
        res = msgbox("This is a converted order and As Maintain must be done manually, please hit ok when finished.", vbokcancel)
        if res = vbCancel then
            wscript.quit
        end if
        exit sub
    end if

    if version10 = False THEN
        do
            on ERROR RESUME NEXT
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I").select
            on ERROR RESUME NEXT
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "MPL_ERROR"
            if err.number <> 0 then
                err.clear
                MSGBOX("There is a problem with the Auth Docs please fix and hit Ok")
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I").select
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "MPL_ERROR"
                on ERROR GOTO 0
            ELSE
                exit DO
            end if
        LOOP
        on ERROR GOTO 0

        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "MPL_ERROR"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "&SORT_ASC"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectAll
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "EXPAND"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "MPL_ERROR"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "MPL_ERROR"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "&SORT_ASC"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectAll
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "LOCK_SELECT"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "STRUPDT"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "FINAL_CHK"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[1]/btn[17]").press
    ELSE
        do
            on ERROR RESUME NEXT
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I").select
            on ERROR RESUME NEXT
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "MPL_ERROR"
            if err.number <> 0 then
                err.clear
                res = MSGBOX("There is a problem with the Auth Docs please fix and hit Ok", vbokcancel)
                if res = vbcancel then
                    wscript.quit
                end if
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I").select
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "MPL_ERROR"
                on ERROR GOTO 0
            ELSE
                exit DO
            end if
        LOOP
        on ERROR GOTO 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "MPL_ERROR"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "&SORT_ASC"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectAll
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "EXPAND"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "MPL_ERROR"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "MPL_ERROR"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "&SORT_ASC"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectAll
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "LOCK_SELECT"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "STRUPDT"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpMAINTAINED_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "FINAL_CHK"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[1]/btn[17]").press
    END IF

    dim result
    dim tmpSTATUS1
    dim SearchChar

    SearchChar = "Z108"
    DO
        tmpSTATUS1 = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
        result = InStr(tmpSTATUS1, SearchChar)

        If result > 0 Then
            RAIres = msgbox("Is this a RAI or Scrap?", vbYesNo)
            if RAIres = vbYes then
                exit DO
            end if
            ALERTres = msgbox("!!ALERT!!" & VBCRLF & "Unit failed As Maintain, Do you wish to continue?", vbYesNo, "ALERT")
            if ALERTres = vbNo then
                wscript.quit
            end if
        ELSE
            exit DO
        END IF
    LOOP
End Sub


sub doPartsLinking
    tmpFilter = "MREP"
    if isExchange = True then
        tmpFilter = "EXCH"
    end if
    if isDPMI = True then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7026421-904" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7026421-905" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7026421-906" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7023497-903" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7026236-906" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7025327-905" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7027508-901" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7035751-904" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7023507-903" then
        tmpFilter = "MREP"
    end if
    if tmpPN = "7032402-906" then
        tmpFilter = "MREP"
    end if

    if version10 = true then
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I").select
        if tmpPN = "300-60173-0502" then
            tmpA = msgbox("Does this unit have MREP as the top line in Disp column?", vbyesno)
            if tmpA = vbyes then
                tmpFilter = "MREP"
            else
                tmpFilter = "EXCH"
            end if
        end if
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = tmpFilter
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 4
        session.findById("wnd[1]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = ""
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "PRTLNK"
        session.findById("wnd[1]/usr/cntlG_CNTL_FAIL/shellcont/shell").currentCellColumn = ""
        session.findById("wnd[1]/usr/cntlG_CNTL_FAIL/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").pressToolbarButton "Select all secondary"

        'message box pause to do defect linking manually
        partsres = MsgBox("Are all defects properly selected?", vbYesNo, "Parts Linking")
        if partsres = vbNo then
            MSGBOX("TERMINATING SCRIPT")
            ERR.CLEAR
            On Error RESUME NEXT
            session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").setCurrentCell - 1, ""
            session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").selectAll
            session.findById("wnd[2]/usr/btnW_PUSH").press
            session.findById("wnd[3]").sendVKey 0
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
            session.findById("wnd[1]").sendVKey 16
            session.findById("wnd[1]").sendVKey 0
            if ERR.NUMBER <> 0 then
                session.findById("wnd[0]").sendVKey 2
                session.findById("wnd[0]").sendVKey 12
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
                session.findById("wnd[1]").sendVKey 16
                session.findById("wnd[1]").sendVKey 0
            end if
            wscript.quit
        end if
        ERR.CLEAR
        On Error RESUME NEXT
        session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").setCurrentCell - 1, ""
        session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").selectAll
        session.findById("wnd[2]/usr/btnW_PUSH").press
        session.findById("wnd[3]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
        session.findById("wnd[1]").sendVKey 16
        session.findById("wnd[1]").sendVKey 0
        if ERR.NUMBER <> 0 then
            session.findById("wnd[0]").sendVKey 2
            session.findById("wnd[0]").sendVKey 12
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
            session.findById("wnd[1]").sendVKey 16
            session.findById("wnd[1]").sendVKey 0
        end if
    else
        on error resume next
        err.clear
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I").select
        if tmpPN = "300-60173-0502" then
            tmpA = msgbox("Does this unit have MREP as the top line in Disp column?", vbyesno)
            if tmpA = vbyes then
                tmpFilter = "MREP"
            else
                tmpFilter = "EXCH"
            end if
        end if
        if err.number <> 0 then
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I").select
        end if
        err.clear
        on error goto 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").setCurrentCell - 1, "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = tmpFilter
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 4
        session.findById("wnd[1]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = ""
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").pressToolbarButton "PRTLNK"
        session.findById("wnd[1]/usr/cntlG_CNTL_FAIL/shellcont/shell").currentCellColumn = ""
        session.findById("wnd[1]/usr/cntlG_CNTL_FAIL/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").pressToolbarButton "Select all secondary"

        'message box pause to do defect linking manually
        partsres = MsgBox("Are all defects properly selected?", vbYesNo, "Parts Linking")
        if partsres = vbNo then
            MSGBOX("TERMINATING SCRIPT")
            ERR.CLEAR
            On Error RESUME NEXT
            session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").setCurrentCell - 1, ""
            session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").selectAll
            session.findById("wnd[2]/usr/btnW_PUSH").press
            session.findById("wnd[3]").sendVKey 0
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
            session.findById("wnd[1]").sendVKey 16
            session.findById("wnd[1]").sendVKey 0
            if ERR.NUMBER <> 0 then
                session.findById("wnd[0]").sendVKey 2
                session.findById("wnd[0]").sendVKey 12
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
                session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
                session.findById("wnd[1]").sendVKey 16
                session.findById("wnd[1]").sendVKey 0
            end if
            wscript.quit
        end if
        ERR.CLEAR
        On Error RESUME NEXT
        session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").setCurrentCell - 1, ""
        session.findById("wnd[2]/usr/cntlG_DEFECT_LINK/shellcont/shell").selectAll
        session.findById("wnd[2]/usr/btnW_PUSH").press
        session.findById("wnd[3]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
        session.findById("wnd[1]").sendVKey 16
        session.findById("wnd[1]").sendVKey 0
        if ERR.NUMBER <> 0 then
            session.findById("wnd[0]").sendVKey 2
            session.findById("wnd[0]").sendVKey 12
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").currentCellColumn = "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectColumn "DISPO"
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").contextMenu
            session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpEVALUATION_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0311/cntlG_CNTR_ITM_EVAL/shellcont/shell").selectContextMenuItem "&FILTER"
            session.findById("wnd[1]").sendVKey 16
            session.findById("wnd[1]").sendVKey 0
        end if
    end if
End Sub

sub PrintServiceReport
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS").getAbsoluteRow(7).selected = true
    session.findById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS/txtWWORKPAPER-WORKPAPER[0,7]").setFocus
    session.findById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS/txtWWORKPAPER-WORKPAPER[0,7]").caretPosition = 0
    session.findById("wnd[1]/tbar[0]/btn[16]").press

    SRres = MsgBox("Is everything correct on the Service Report? Click Yes to print.", vbYesNo, "Service Report")
    if SRres = vbNo then
        strText = "Close up Complete. Thank you"
        objVoice.Speak strText

        MsgBox("Close-Up Complete")
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I").select
        'session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        'session.findById("wnd[0]").sendVKey 0
        'session.findById("wnd[0]").sendVKey 0
        exit SUB
    end if
    session.findById("wnd[0]/mbar/menu[0]/menu[0]").select
    session.findById("wnd[0]/tbar[0]/btn[3]").press

    if STAYLABOREDON = True then
        msgbox("Close-Up will be complete after Final Test has been Labored Off Complete.  Thank you!")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I").select
    else
        strText = "Close up Complete. Thank you"
        objVoice.Speak strText

        MsgBox("Close-Up Complete")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I").select
        'session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        'session.findById("wnd[0]").sendVKey 0
        'session.findById("wnd[0]").sendVKey 0
    end if

End Sub

sub openFinalInspection
       session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpFINALINSP_I").select
    if isSPEX = False then
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpFINALINSP_I/ssubG_IWB_ITEMS:SAPLYAFF_GDBE:0110/btnW_SIM").press
    end if
    MsgBox("Check red indicators for errors.")
    if newSPEX = True then
        exit sub
    end if

    if STAYLABOREDON = True then
        msgbox("You will have to labor off complete of Final before sending to QA.")
    end if

    if newSPEX = True then
        exit sub
    end if


End Sub

Sub UpdateWandingStatus
    ON ERROR RESUME NEXT
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    IF ERR.NUMBER <> 0 THEN
        msgbox("Oops can you fix that")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        err.clear
        on error goto 0
    END IF

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/ctxtYAFS_ZIWBN_HEADER-SRV_LOCATION").text = "FBOHOLD"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/ctxtYAFS_ZIWBN_HEADER-SRV_LOCATION").caretPosition = 7
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/btnG_H_SRV_BT_SAVE").press

End Sub

Sub UpdateWSUPDComments
    dim tmpDateTime
    tmpDateTime = Now

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/btnG_H_SERORD_BT_LTXT").press
    session.findById("wnd[1]/usr/cntlW_TEXT_LTXT/shellcont/shell").text = "QA FINAL, Rev 3, " + CStr(tmpDateTime) + "  " + UserName + vbCr+session.findById("wnd[1]/usr/cntlW_TEXT_LTXT/shellcont/shell").text
    session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub

sub recordSO(tmpServOrder, tmpUSER_RESPONSIBLE)
    Const fsoForAppend = 8

    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    'Open the text file
    Dim objTextStream
    Set objTextStream = objFSO.OpenTextFile("S:\CSC DataBases\Close-Up Script\archive\DATA FOR SCRIPT\autorun.TXT", fsoForAppend)

    'Display the contents of the text file
    objTextStream.WriteLine(tmpServOrder & ", " & tmpUSER_RESPONSIBLE & ", " & NOW & ", " & "REV -")
    'objTextStream.Write Now

    'Close the file and clean up
    objTextStream.Close
    Set objTextStream = Nothing
    Set objFSO = Nothing

end sub

sub LaborOffIncompleteOlathe
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellRow = -1
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&FILTER"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Final Test*"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 6
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton "LABOFFI"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").setCurrentCell - 1, "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&FILTER"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = ""
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 0
    session.findById("wnd[1]/tbar[0]/btn[0]").press
end sub

sub LaborOffCompleteOlathe
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellRow = -1
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&FILTER"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Close Up*"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 6
    session.findById("wnd[1]").sendVKey 0

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = "OPER_COMM"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").clickCurrentCell
    session.findById("wnd[1]/usr/cntlG_CNTL_OPER_COMM/shellcont/shell").setUnprotectedTextPart 0, "SSOE COMPLETE" + vbCr+""
    session.findById("wnd[1]/usr/cntlG_CNTL_OPER_COMM/shellcont/shell").setSelectionIndexes 99, 99
    session.findById("wnd[1]").sendVKey 11

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton "LABOFFC"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").setCurrentCell - 1, "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn "LTXA1"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&FILTER"
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = ""
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 0
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    res = msgbox("Are all Operation lines closed except for the Final Audit operation?", vbyesno)
    if res = vbno then
        msgbox("All Operation lines must be closed before sending unit to QA, terminating script")
    end if


end sub




sub getActualFindingInfo(tmpMODSIN, tmpMODSOUT)
    wscript.sleep(5000)
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I").select

    dim tmpSealsRec
    dim tmpTeamSealRec
    dim tmpCAUSEtxt



    'OPEN FINDINGS
    session.findById("wnd[0]/tbar[1]/btn[9]").press

    'SELECT REMOVAL DATA TAB
    session.findById("wnd[1]/usr/tabsTC_CNTR/tabpLASSA").SELECT

    res = msgbox("Input the required fields if you haven't already." & vbcrlf & "(Zam codes, User Responsible, etc..,)", vbokcancel)
    if res = vbcancel then
        msgbox("Exiting Script")
        session.findById("wnd[1]").sendVKey 11
        wscript.quit
    end if

    DO
        tmpPRWRKSCPE = session.findById("wnd[1]/usr/tabsTC_CNTR/tabpLASSA/ssubSSA:SAPLYAFF_AM_PERSONAS_REQ:9004/ctxtW_AVRMVD-PW").text

        IF tmpPRWRKSCPE = "" THEN
            resRAI = msgbox("Is this an RAI or SCRAP?", vbYesNo)
            if resRAI = vbYes then
                exit do
            end if
            res = msgbox("Please input a Primary Workscope or hit cancel to end script.", vbokcancel)
            if res = vbcancel then
                msgbox("Ending script now.")
                wscript.quit
            end if
        else
            exit do
        END IF
    LOOP


    'Force secondary workscope uncheck
    session.findById("wnd[1]/usr/tabsTC_CNTR/tabpLASSA/ssubSSA:SAPLYAFF_AM_PERSONAS_REQ:9004/chkW_AVRMVD-SWM").selected = false
    session.findById("wnd[1]/usr/tabsTC_CNTR/tabpLASSA/ssubSSA:SAPLYAFF_AM_PERSONAS_REQ:9004/chkW_AVRMVD-SWR").selected = false
    session.findById("wnd[1]/usr/tabsTC_CNTR/tabpLASSA/ssubSSA:SAPLYAFF_AM_PERSONAS_REQ:9004/chkW_AVRMVD-SWI").selected = false
    session.findById("wnd[1]/usr/tabsTC_CNTR/tabpLASSA/ssubSSA:SAPLYAFF_AM_PERSONAS_REQ:9004/chkW_AVRMVD-SWT").selected = false

    'SWITCH TO CONDITION DATA TAB
    session.findById("wnd[1]/usr/tabsTC_CNTR/tabpHASSA").select

    msgbox("Verify the required fields are filled out correctly, then hit OK." & vbcrlf & "(Work accomplished, Mods in and out, etc..,)")


    res2 = msgbox("Does the MODS match the ID plate on the unit and the outgoing MODS in SAP?", VBYesNo, "Check")
    if res1 = vbNo then
        objVoice.Speak strText
        wscript.quit
    end if

    session.findById("wnd[1]/usr/tabsTC_CNTR/tabpHASSA").select
    session.findById("wnd[1]").sendVKey 11
    wscript.sleep(5000)

end sub



