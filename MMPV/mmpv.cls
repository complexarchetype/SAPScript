'Refer to the xlsm file for Excel file format

Option Explicit
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As SAPFEWSELib.GuiConnection
Public objSess As SAPFEWSELib.GuiSession
Public objSBar As GuiStatusbar
Public objSheet As Worksheet
Dim W_System
Dim iCtr As Integer
Const tcode = "mmpv"


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
    If objSheet.Cells(iRow, 5) = "0" Then
        itemmax = itemmax + 1
    End If
Next
' Update the counter in cell A9
objSheet.Cells(9, 1) = itemcount & "/" & itemmax

' Cycle through the rows with status 0 and call the ProcessRow function to process them
For iRow = startrow To objSheet.UsedRange.Rows.Count
    If objSheet.Cells(iRow, 5) = "0" Then
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
Dim F_Comp, T_Comp, Period, Year
Dim lineitems As Long

' Set the line status to "processing..."
objSheet.Cells(iRow, 5) = 1

'F_Comp
If objSheet.Cells(iRow, 1) <> "" Then
    F_Comp = objSheet.Cells(iRow, 1)
Else
    F_Comp = "xxxxxx"
End If

'T_Comp
If objSheet.Cells(iRow, 2) <> "" Then
    T_Comp = objSheet.Cells(iRow, 2)
Else
    T_Comp = "xxxxxx"
End If

'Period
If objSheet.Cells(iRow, 3) <> "" Then
    Period = objSheet.Cells(iRow, 3)
Else
    Period = "xxxxxx"
End If

'Year
If objSheet.Cells(iRow, 4) <> "" Then
    Year = objSheet.Cells(iRow, 4)
Else
    Year = "xxxxxx"
End If


' Set error handling to the below code where we can capture if a line had failed in the GUI script
'On Error GoTo myerr

' SAP GUI Script starts here

objSess.FindById("wnd[0]").Maximize
objSess.FindById("wnd[0]/tbar[0]/okcd").Text = "mmpv"
objSess.FindById("wnd[0]").sendVKey 0
objSess.FindById("wnd[0]/usr/ctxtI_VBUKR").Text = F_Comp
objSess.FindById("wnd[0]/usr/ctxtI_BBUKR").Text = T_Comp
objSess.FindById("wnd[0]/usr/txtI_LFMON").Text = Period
objSess.FindById("wnd[0]/usr/txtI_LFGJA").Text = Year
objSess.FindById("wnd[0]/usr/txtI_LFGJA").SetFocus
objSess.FindById("wnd[0]/usr/txtI_LFGJA").caretPosition = 4
objSess.FindById("wnd[0]/tbar[1]/btn[8]").press
objSess.FindById("wnd[0]/tbar[0]/okcd").Text = "/n"
objSess.FindById("wnd[0]").sendVKey 0

' Gets the message from the status bar and save it in column B
objSheet.Cells(iRow, 7) = objSBar.Text

' Update the Status to "Completed" and exit
objSheet.Cells(iRow, 5) = 2
Exit Function

myerr:
' Update the status to "Error"
objSheet.Cells(iRow, 5) = 3

End Function

