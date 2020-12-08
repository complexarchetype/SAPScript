' Multiple Email Address Update

' This script allows upto 20 email address to be maintained in customer master data.
' Contains core module code only
' For data upload format refer to the attached xlsm file.

' Changelog 1.1
' Support for updating email address after initial run

' Release 1.0
' Initial Code 

Option Explicit
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As SAPFEWSELib.GuiConnection
Public objSess As SAPFEWSELib.GuiSession
Public objSBar As GuiStatusbar
Public objSheet As Worksheet
Dim W_System
Dim iCtr As Integer
Const tcode = "xd02"


Function Attach_Session(iRow, Optional mysystem As String) As Boolean
    Dim il, it
    Dim W_conn, W_Sess
    Dim SapGuiAuto As Object
    Dim i As Integer
    Dim iSession As Long
    Dim SAP_APP As Object
    Dim Connection As SAPFEWSELib.GuiConnection
    Dim Session As SAPFEWSELib.GuiSession
    Dim strMSG As String
    Dim strSessions$

    i = 1
    ' There may be bad entries in the ROT from previous crashes
    While i < 10 And SapGuiAuto Is Nothing
       i = i + 1
       On Error Resume Next
       Set SapGuiAuto = GetObject("SAPGUI")
       On Error GoTo 0
    Wend
    
    If SapGuiAuto Is Nothing Then
        MsgBox "Could not connect to SAPlogon process. Did you start it?"
    End If
    
    On Error Resume Next
    Set SAP_APP = SapGuiAuto.GetScriptingEngine
    Set SapGuiAuto = Nothing
    On Error GoTo 0
    
    If SAP_APP Is Nothing Then
    
        MsgBox "Could not access GuiApplication. Maybe Scripting is disabled?"
        
    End If
    
    Set SapGuiAuto = Nothing
    iSession = 0
    
    For Each Connection In SAP_APP.Children
        If Not Connection.DisabledByServer Then
            For Each Session In Connection.Children
                If Session.Busy = False Then
                    iSession = iSession + 1
                    strMSG = strMSG & (Session.Info.SystemName & " (" & CStr(Session.Info.SessionNumber) & ") (" & Session.Info.Client & ") | User: " & Session.Info.User & " | Transaction: " & Session.Info.Transaction & " | Session ID: " & Session.ID) & vbCrLf & vbCrLf
                End If
            Next
            Set objSess = Connection.Children(0)
        End If
    Next
    strMSG = strMSG & vbCrLf & "Sessions counted: " & iSession
    MsgBox strMSG, vbInformation + vbOKOnly, "Information"
    
    
    If IsObject(WScript) Then
        WScript.ConnectObject objSess, "on"
        WScript.ConnectObject objGui, "on"
    End If
    
    Set objSBar = objSess.FindById("wnd[0]/sbar")
     objSess.FindById("wnd[0]").Maximize

End Function

Public Sub StartProcessing()

Dim W_Obj1, W_Obj2, W_Obj3, W_Obj4, iRow
Dim W_Func
Dim W_Src_Ord
Dim W_Ret As Boolean
Dim itemcount As Integer
Dim itemmax As Integer
Const startrow As Integer = 11 'First row with actual data

Set objSheet = ActiveWorkbook.ActiveSheet
' Connect to a system stored in cell A8
W_Ret = Attach_Session(8)
'If Not W_Ret Then
'    MsgBox "Not connected to client"
'    GoTo MyEnd
'End If

itemcount = 0
itemmax = 0

' Determine the number of items to be processed: where the status is zero
For iRow = startrow To objSheet.UsedRange.Rows.Count
    If objSheet.Cells(iRow, 20) = "0" Then
        itemmax = itemmax + 1
    End If
Next
' Update the counter in cell A9
objSheet.Cells(9, 1) = itemcount & "/" & itemmax

' Cycle through the rows with status 0 and call the ProcessRow function to process them
For iRow = startrow To objSheet.UsedRange.Rows.Count
    If objSheet.Cells(iRow, 20) = "0" Then
        Call ProcessRow(iRow)
        itemcount = itemcount + 1
        objSheet.Cells(9, 1) = itemcount & "/" & itemmax
    End If
Next
           

MyEnd:
' destory the objects, free up memory
    Set objSess = Nothing
    Set objGui = Nothing
    Set SapGuiAuto = Nothing
    
    MsgBox "Script completed.", vbInformation + vbOKOnly
       
End Sub

Function ProcessRow(iRow)
Dim Cust_ID, Comp, Email_M, Email_2, Email_3, Email_4, Email_5, Email_6, Email_7, Email_8, Email_9, Email_10, Email_11, Email_12, Email_13, Email_14, Email_15, Email_16, Email_17
Dim lineitems As Long

' Set the line status to "processing..."
objSheet.Cells(iRow, 20) = 1

'Cust_ID
If objSheet.Cells(iRow, 1) <> "" Then
    Cust_ID = objSheet.Cells(iRow, 1)
Else
    Cust_ID = "xxxxxx"
End If

'Comp
If objSheet.Cells(iRow, 2) <> "" Then
    Comp = objSheet.Cells(iRow, 2)
Else
    Comp = "xxxxxx"
End If

'Email_M
If objSheet.Cells(iRow, 3) <> "" Then
    Email_M = objSheet.Cells(iRow, 3)
Else
    Email_M = "xxxxxx"
End If

'Email_2
If objSheet.Cells(iRow, 4) <> "" Then
    Email_2 = objSheet.Cells(iRow, 4)
Else
    Email_2 = "xxxxxx"
End If

'Email_3
If objSheet.Cells(iRow, 5) <> "" Then
    Email_3 = objSheet.Cells(iRow, 5)
Else
    Email_3 = "xxxxxx"
End If

'Email_4
If objSheet.Cells(iRow, 6) <> "" Then
    Email_4 = objSheet.Cells(iRow, 6)
Else
    Email_4 = "xxxxxx"
End If

'Email_5
If objSheet.Cells(iRow, 7) <> "" Then
    Email_5 = objSheet.Cells(iRow, 7)
Else
    Email_5 = "xxxxxx"
End If

'Email_6
If objSheet.Cells(iRow, 8) <> "" Then
    Email_6 = objSheet.Cells(iRow, 8)
Else
    Email_6 = "xxxxxx"
End If

'Email_7
If objSheet.Cells(iRow, 9) <> "" Then
    Email_7 = objSheet.Cells(iRow, 9)
Else
    Email_7 = "xxxxxx"
End If

'Email_8
If objSheet.Cells(iRow, 10) <> "" Then
    Email_8 = objSheet.Cells(iRow, 10)
Else
    Email_8 = "xxxxxx"
End If

'Email_9
If objSheet.Cells(iRow, 11) <> "" Then
    Email_9 = objSheet.Cells(iRow, 11)
Else
    Email_9 = "xxxxxx"
End If

'Email_10
If objSheet.Cells(iRow, 12) <> "" Then
    Email_10 = objSheet.Cells(iRow, 12)
Else
    Email_10 = "xxxxxx"
End If

'Email_11
If objSheet.Cells(iRow, 13) <> "" Then
    Email_11 = objSheet.Cells(iRow, 13)
Else
    Email_11 = "xxxxxx"
End If

'Email_12
If objSheet.Cells(iRow, 14) <> "" Then
    Email_12 = objSheet.Cells(iRow, 14)
Else
    Email_12 = "xxxxxx"
End If

'Email_13
If objSheet.Cells(iRow, 15) <> "" Then
    Email_13 = objSheet.Cells(iRow, 15)
Else
    Email_13 = "xxxxxx"
End If

'Email_14
If objSheet.Cells(iRow, 16) <> "" Then
    Email_14 = objSheet.Cells(iRow, 16)
Else
    Email_14 = "xxxxxx"
End If

'Email_15
If objSheet.Cells(iRow, 17) <> "" Then
    Email_15 = objSheet.Cells(iRow, 17)
Else
    Email_15 = "xxxxxx"
End If

'Email_16
If objSheet.Cells(iRow, 18) <> "" Then
    Email_16 = objSheet.Cells(iRow, 18)
Else
    Email_16 = "xxxxxx"
End If

'Email_17
If objSheet.Cells(iRow, 19) <> "" Then
    Email_17 = objSheet.Cells(iRow, 19)
Else
    Email_17 = "xxxxxx"
End If

' Set error handling to the below code where we can capture if a line had failed in the GUI script
'On Error GoTo myerr

' SAP GUI Script starts here

objSess.FindById("wnd[0]/tbar[0]/okcd").Text = "xd02"
objSess.FindById("wnd[0]").sendVKey 0
objSess.FindById("wnd[1]/usr/ctxtRF02D-KUNNR").Text = Cust_ID
objSess.FindById("wnd[1]/usr/ctxtRF02D-BUKRS").Text = Comp
objSess.FindById("wnd[1]/usr/ctxtRF02D-BUKRS").SetFocus
objSess.FindById("wnd[1]/usr/ctxtRF02D-BUKRS").caretPosition = 4
objSess.FindById("wnd[1]").sendVKey 0

If objSheet.Cells(8, 6) = "Update" Then
    objSess.FindById("wnd[0]").sendVKey 0
    objSess.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_ICON_SMTP").press
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_M
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 22
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
    
    If objSheet.Cells(iRow, 4) <> "" Then
        objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_2
        objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
        objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
    End If
    
Else
    objSess.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").Text = Email_M
    objSess.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").SetFocus
    objSess.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").caretPosition = 17
    objSess.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_ICON_SMTP").press
    On Error Resume Next
    objSess.FindById("wnd[0]").sendVKey 0
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
    
    If objSheet.Cells(iRow, 4) <> "" Then
        objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,1]").Text = Email_2
        objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,1]").caretPosition = 17
        objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
    End If

End If

If objSheet.Cells(iRow, 5) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_3
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 6) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_4
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 7) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_5
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 8) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_6
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 9) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_7
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 10) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_8
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 11) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_9
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 12) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_10
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 13) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_11
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 14) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_12
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 15) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_13
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 16) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_14
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 17) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_15
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 18) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_16
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

If objSheet.Cells(iRow, 19) <> "" Then
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").Text = Email_17
    objSess.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,2]").caretPosition = 17
    objSess.FindById("wnd[1]/tbar[0]/btn[13]").press
End If

objSess.FindById("wnd[1]/tbar[0]/btn[0]").press
objSess.FindById("wnd[0]/tbar[0]/btn[11]").press
objSess.FindById("wnd[1]").Close
objSess.FindById("wnd[0]/tbar[0]/okcd").Text = "/n"
objSess.FindById("wnd[0]").sendVKey 0

' Gets the message from the status bar and save it in column B
objSheet.Cells(iRow, 22) = objSBar.Text

' Update the Status to "Completed" and exit
objSheet.Cells(iRow, 20) = 2
Exit Function

myerr:
' Update the status to "Error"
objSheet.Cells(iRow, 20) = 3

End Function
