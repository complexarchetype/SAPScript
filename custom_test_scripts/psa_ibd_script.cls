Option Explicit
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As SAPFEWSELib.GuiConnection
Public objSess As SAPFEWSELib.GuiSession
Public objSBar As GuiStatusbar
Public objSheet As Worksheet
Dim W_System


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
    
    Set objSBar = objSess.findById("wnd[0]/sbar")
     objSess.findById("wnd[0]").maximize

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

itemcount = 0
itemmax = 0

' Determine the number of items to be processed: where the status is zero
For iRow = startrow To objSheet.UsedRange.Rows.Count
    If objSheet.Cells(iRow, 6) = "0" Then
        itemmax = itemmax + 1
    End If
Next
' Update the counter in cell A9
objSheet.Cells(9, 1) = itemcount & "/" & itemmax

' Cycle through the rows with status 0 and call the ProcessRow function to process them
For iRow = startrow To objSheet.UsedRange.Rows.Count
    If objSheet.Cells(iRow, 6) = "0" Then
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
Dim Ven_ID, Mat, Qty, St_Loc, Comp
Dim lineitems As Long
Dim po, SC, SB As String
Dim Del, Del_1, Del_2 As String

' Set the line status to "processing..."
objSheet.Cells(iRow, 6) = 1

'Ven_ID
If objSheet.Cells(iRow, 1) <> "" Then
    Ven_ID = objSheet.Cells(iRow, 1)
Else
    Ven_ID = "xxxxxx"
End If

'Mat
If objSheet.Cells(iRow, 2) <> "" Then
    Mat = objSheet.Cells(iRow, 2)
Else
    Mat = "xxxxxx"
End If

'Qty
If objSheet.Cells(iRow, 3) <> "" Then
    Qty = objSheet.Cells(iRow, 3)
Else
    Qty = "xxxxxx"
End If

'St_Loc
If objSheet.Cells(iRow, 4) <> "" Then
    St_Loc = objSheet.Cells(iRow, 4)
Else
    St_Loc = "xxxxxx"
End If

'Comp
If objSheet.Cells(iRow, 5) <> "" Then
    Comp = objSheet.Cells(iRow, 5)
Else
    Comp = "xxxxxx"
End If

' SAP GUI Script starts here

objSess.findById("wnd[0]").maximize
objSess.findById("wnd[0]/tbar[0]/okcd").Text = "me21n"
objSess.findById("wnd[0]").sendVKey 0
On Error Resume Next
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/chkMEPO1211-UMSON[23,0]").Selected = True
On Error Resume Next
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").Text = Mat
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").Text = Qty
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-LGOBE[16,0]").Text = St_Loc
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = Ven_ID
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").SetFocus
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").caretPosition = 13
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT18/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB1:SAPLMEGUI:1334/cmbMEPO1334-BSTAE").Key = "0004"
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT18/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB1:SAPLMEGUI:1334/cmbMEPO1334-BSTAE").SetFocus
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").press
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT7").Select
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT7/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1229/txtMEPO1229-SUBMI").Text = Comp
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT7/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1229/txtMEPO1229-SUBMI").SetFocus
objSess.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT7/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1229/txtMEPO1229-SUBMI").caretPosition = 4
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/tbar[0]/btn[11]").press
objSess.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press

' Gets the message from the status bar and save it in column B
po = objSBar.Text

objSheet.Cells(iRow, 8) = Right(po, 10)

objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/nme21n"
objSess.findById("wnd[0]").sendVKey 0

If objSheet.Cells(iRow, 11) = "PO" Then
    
    GoTo po
    
End If

' Call Transaction VL31N - Del, TO
objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl31n"
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/usr/ctxtLIKP-LIFNR").Text = Ven_ID
objSess.findById("wnd[0]/usr/txtLV50C-BSTNR").Text = Right(po, 10)
objSess.findById("wnd[0]/usr/txtLV50C-BSTNR").SetFocus
objSess.findById("wnd[0]/usr/txtLV50C-BSTNR").caretPosition = 10
objSess.findById("wnd[0]").sendVKey 0
'objSess.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1202/tblSAPMV50ATC_LIPS_OVER_INB/ctxtLIPS-LGORT[23,1]").Text = St_Loc
'objSess.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1202/tblSAPMV50ATC_LIPS_OVER_INB/ctxtLIPS-LGORT[23,1]").SetFocus
'objSess.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1202/tblSAPMV50ATC_LIPS_OVER_INB/ctxtLIPS-LGORT[23,1]").caretPosition = 4
objSess.findById("wnd[0]").sendVKey 0
'objSess.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
'objSess.findById("wnd[1]/usr/btnSPOP-OPTION1").press
objSess.findById("wnd[0]/tbar[0]/btn[11]").press

Del = objSBar.Text

Del_1 = Mid(Del, 17, 10)
Del_2 = "0" & Right(Del_1, 9)

objSheet.Cells(iRow, 10) = Del_2

objSheet.Cells(iRow, 9) = Mid(Del, 17, 10)

If objSheet.Cells(iRow, 11) = "DEL" Then
    
    GoTo Del
    
End If

' Call Transaction VL32N - Handling Unit Update

If Comp = "1970" Then

objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl32n"
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = Del_1
objSess.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = 9
objSess.findById("wnd[0]/tbar[1]/btn[18]").press
objSess.findById("wnd[0]").sendVKey 2

    Do While objSBar.Text <> "There are no other handling units"
    objSess.findById("wnd[0]/usr/txtVEKP-EXIDV2").Text = RandString(9)
    objSess.findById("wnd[0]/usr/txtVEKP-EXIDV2").SetFocus
    objSess.findById("wnd[0]/usr/txtVEKP-EXIDV2").caretPosition = 10
    objSess.findById("wnd[0]/tbar[1]/btn[5]").press
    Loop

objSess.findById("wnd[0]/tbar[0]/btn[11]").press

End If

' Call Transaction VL32N - TO

objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl32n"
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = Del_1
objSess.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/tbar[1]/btn[16]").press
objSess.findById("wnd[0]/tbar[1]/btn[6]").press
objSess.findById("wnd[0]/tbar[0]/btn[11]").press
objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
objSess.findById("wnd[0]").sendVKey 0

' Call Transaction VL32N - PGI
objSess.findById("wnd[0]/tbar[0]/okcd").Text = "vl32n"
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = Del_1
objSess.findById("wnd[0]/tbar[1]/btn[20]").press
objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
objSess.findById("wnd[0]").sendVKey 0

If objSheet.Cells(iRow, 11) = "PGR" Then
    
    GoTo pgr
    
End If

' Call Transaction LT10 - Stock Transfer

objSess.findById("wnd[0]/tbar[0]/okcd").Text = "LT10"
objSess.findById("wnd[0]").sendVKey 0
objSess.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = "IA1"
objSess.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").Text = "917"
objSess.findById("wnd[0]/usr/ctxtS1_LGPLA-LOW").Text = Del_2
objSess.findById("wnd[0]/usr/ctxtS1_LGPLA-LOW").SetFocus
objSess.findById("wnd[0]/usr/ctxtS1_LGPLA-LOW").caretPosition = 10
objSess.findById("wnd[0]/tbar[1]/btn[8]").press
objSess.findById("wnd[0]/tbar[1]/btn[45]").press
objSess.findById("wnd[0]/tbar[1]/btn[48]").press

If Comp = "1970" Then
    SC = "PLN"
    SB = "PLN-KDDC"
Else
    SC = "PLN"
    SB = "PLN-TEST1"
End If

objSess.findById("wnd[1]/usr/ctxtLAGP-LGTYP").Text = SC
objSess.findById("wnd[1]/usr/ctxtLAGP-LGPLA").Text = SB
objSess.findById("wnd[1]/usr/ctxtLAGP-LGPLA").SetFocus
objSess.findById("wnd[1]/usr/ctxtLAGP-LGPLA").caretPosition = 8
objSess.findById("wnd[1]/tbar[0]/btn[0]").press
objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
objSess.findById("wnd[0]").sendVKey 0


' Update the Status to "Completed" and exit
objSheet.Cells(iRow, 6) = 2
Exit Function

myerr:
' Update the status to "Error"
objSheet.Cells(iRow, 6) = 3

po:
objSheet.Cells(iRow, 6) = 2

Del:
objSess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
objSess.findById("wnd[0]").sendVKey 0
objSheet.Cells(iRow, 6) = 2

pgr:
objSheet.Cells(iRow, 6) = 2

End Function


Function RandString(n As Long) As String
    Dim i As Long, j As Long, m As Long, s As String, pool As String
    pool = "123456789C"
    m = Len(pool)
    For i = 1 To n
        j = 1 + Int(m * Rnd())
        s = s & Mid(pool, j, 1)
    Next i
    RandString = s
End Function
