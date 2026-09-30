Attribute VB_Name = "RequisitionSAP"
'==============================================================================
' Stage 03 - Purchase requisition creation in SAP (ME51N)
'
' Reads the requisition workbook produced by stage 02, creates one requisition
' per sheet through SAP GUI scripting, and writes the number SAP returns back
' into the workbook.
'
' The script never invents data. A line missing a required field is skipped and
' logged; the run continues with the remaining lines.
'
' SETUP
'   1. SAP GUI scripting must be enabled on the client (Options > Accessibility
'      & Scripting > Scripting) and allowed on the server (profile parameter
'      sapgui/user_scripting).
'   2. Log into SAP before running. The script attaches to the open session; it
'      does not handle credentials and never stores them.
'   3. The screen paths in the constants below are specific to an SAP release
'      and screen variant. Record them once in your own system with the SAP GUI
'      script recorder and replace the values here. Paths copied from another
'      installation will not resolve.
'==============================================================================

Option Explicit

'--- Workbook layout ----------------------------------------------------------
Private Const SHEET_TOOLING As String = "Tooling"
Private Const SHEET_SERVICE As String = "Service"
Private Const FIRST_DATA_ROW As Long = 8

Private Const COL_CODE As Long = 3
Private Const COL_QUANTITY As Long = 4
Private Const COL_VALUE As Long = 5
Private Const COL_DESCRIPTION As Long = 6
Private Const COL_ITEM_TYPE As Long = 7
Private Const COL_MATERIAL_GROUP As Long = 8
Private Const COL_BUDGET_CODE As Long = 9
Private Const COL_BUDGET_LINE As Long = 10
Private Const COL_PROJECT_CODE As Long = 11
Private Const COL_LOG As Long = 13

Private Const CELL_DELIVERY_DATE As String = "E3"
Private Const CELL_PLANT As String = "E4"
Private Const CELL_PURCHASING_GROUP As String = "E5"
Private Const CELL_REQUESTER As String = "E6"
Private Const CELL_REQUISITION_NUMBER As String = "K2"
Private Const CELL_STATUS As String = "K4"

'--- SAP transaction and screen paths ----------------------------------------
' Re-record these in your own system. See SETUP above.
Private Const TRANSACTION_CODE As String = "/nME51N"
Private Const PATH_COMMAND_FIELD As String = "wnd[0]/tbar[0]/okcd"
Private Const PATH_STATUS_BAR As String = "wnd[0]/sbar"
Private Const PATH_SAVE_BUTTON As String = "wnd[0]/tbar[0]/btn[11]"
Private Const PATH_ITEM_GRID As String = "wnd[0]/usr/subSUB0:SAPLMEGUI:0014/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"

'--- Grid column names --------------------------------------------------------
Private Const GRID_MATERIAL As String = "MATNR"
Private Const GRID_QUANTITY As String = "MENGE"
Private Const GRID_PRICE As String = "PREIS"
Private Const GRID_SHORT_TEXT As String = "TXZ01"
Private Const GRID_MATERIAL_GROUP As String = "MATKL"
Private Const GRID_PLANT As String = "NAME1"
Private Const GRID_DELIVERY_DATE As String = "EEIND"

'--- Module state -------------------------------------------------------------
Private SapSession As Object


'==============================================================================
' Entry points
'==============================================================================

Public Sub CreateToolingRequisition()
    CreateRequisition SHEET_TOOLING
End Sub


Public Sub CreateServiceRequisition()
    CreateRequisition SHEET_SERVICE
End Sub


Public Sub CreateAllRequisitions()
    CreateRequisition SHEET_TOOLING
    CreateRequisition SHEET_SERVICE
End Sub


'==============================================================================
' Main routine
'==============================================================================

Private Sub CreateRequisition(ByVal SheetName As String)

    Dim Sheet As Worksheet
    Dim LastRow As Long
    Dim RowIndex As Long
    Dim GridRow As Long
    Dim Written As Long
    Dim Skipped As Long
    Dim Problem As String

    Set Sheet = ThisWorkbook.Worksheets(SheetName)
    LastRow = LastFilledRow(Sheet)

    If LastRow < FIRST_DATA_ROW Then
        MsgBox "No lines to process on sheet " & SheetName & ".", vbInformation
        Exit Sub
    End If

    If Not ConnectToSap() Then Exit Sub
    ClearLog Sheet, LastRow

    OpenTransaction
    SetHeaderFields Sheet

    GridRow = 0
    For RowIndex = FIRST_DATA_ROW To LastRow

        Problem = ValidateLine(Sheet, RowIndex)

        If Len(Problem) > 0 Then
            Sheet.Cells(RowIndex, COL_LOG).Value = "Skipped: " & Problem
            Skipped = Skipped + 1
        Else
            On Error Resume Next
            FillGridLine Sheet, RowIndex, GridRow

            If Err.Number <> 0 Then
                Sheet.Cells(RowIndex, COL_LOG).Value = "Error: " & Err.Description
                Err.Clear
                Skipped = Skipped + 1
            Else
                Sheet.Cells(RowIndex, COL_LOG).Value = "Sent"
                GridRow = GridRow + 1
                Written = Written + 1
            End If
            On Error GoTo 0
        End If

    Next RowIndex

    If Written = 0 Then
        Sheet.Range(CELL_STATUS).Value = "Nothing sent"
        MsgBox "No valid lines on sheet " & SheetName & ". Nothing was sent to SAP.", vbExclamation
        Exit Sub
    End If

    SaveAndCapture Sheet, Written, Skipped

End Sub


'==============================================================================
' SAP session
'==============================================================================

Private Function ConnectToSap() As Boolean

    Dim SapGui As Object
    Dim Application As Object
    Dim Connection As Object

    On Error GoTo NoSession

    Set SapGui = GetObject("SAPGUI")
    Set Application = SapGui.GetScriptingEngine
    Set Connection = Application.Children(0)
    Set SapSession = Connection.Children(0)

    SapSession.findById("wnd[0]").maximize
    ConnectToSap = True
    Exit Function

NoSession:
    MsgBox "Could not attach to an SAP session." & vbCrLf & vbCrLf & _
           "Log into SAP first, and make sure GUI scripting is enabled " & _
           "on both the client and the server.", vbCritical
    ConnectToSap = False

End Function


Private Sub OpenTransaction()
    SapSession.findById(PATH_COMMAND_FIELD).Text = TRANSACTION_CODE
    SapSession.findById("wnd[0]").sendVKey 0
End Sub


'==============================================================================
' Filling the requisition
'==============================================================================

Private Sub SetHeaderFields(ByVal Sheet As Worksheet)
    ' Header fields such as plant, purchasing group and requester sit on the
    ' item detail tabs rather than a single header screen in ME51N. Record the
    ' paths for your own screen variant and set them here.
End Sub


Private Sub FillGridLine(ByVal Sheet As Worksheet, ByVal RowIndex As Long, _
                         ByVal GridRow As Long)

    Dim Grid As Object
    Set Grid = SapSession.findById(PATH_ITEM_GRID)

    Grid.modifyCell GridRow, GRID_MATERIAL, _
                    CStr(Sheet.Cells(RowIndex, COL_CODE).Value)
    Grid.modifyCell GridRow, GRID_QUANTITY, _
                    CStr(Sheet.Cells(RowIndex, COL_QUANTITY).Value)
    Grid.modifyCell GridRow, GRID_SHORT_TEXT, _
                    CStr(Sheet.Cells(RowIndex, COL_DESCRIPTION).Value)
    Grid.modifyCell GridRow, GRID_MATERIAL_GROUP, _
                    CStr(Sheet.Cells(RowIndex, COL_MATERIAL_GROUP).Value)
    Grid.modifyCell GridRow, GRID_PRICE, _
                    FormatForSap(Sheet.Cells(RowIndex, COL_VALUE).Value)
    Grid.modifyCell GridRow, GRID_DELIVERY_DATE, _
                    Format(Sheet.Range(CELL_DELIVERY_DATE).Value, "dd.mm.yyyy")

    SapSession.findById("wnd[0]").sendVKey 0

End Sub


Private Sub SaveAndCapture(ByVal Sheet As Worksheet, ByVal Written As Long, _
                           ByVal Skipped As Long)

    Dim Message As String
    Dim Number As String

    SapSession.findById(PATH_SAVE_BUTTON).press
    Message = SapSession.findById(PATH_STATUS_BAR).Text

    Number = ExtractNumber(Message)

    If Len(Number) > 0 Then
        Sheet.Range(CELL_REQUISITION_NUMBER).Value = Number
        Sheet.Range(CELL_STATUS).Value = "Created"
        MsgBox "Requisition " & Number & " created." & vbCrLf & _
               Written & " lines sent, " & Skipped & " skipped.", vbInformation
    Else
        Sheet.Range(CELL_STATUS).Value = "Not created"
        MsgBox "SAP did not return a requisition number." & vbCrLf & vbCrLf & _
               "Status bar: " & Message, vbExclamation
    End If

End Sub


'==============================================================================
' Helpers
'==============================================================================

Private Function ValidateLine(ByVal Sheet As Worksheet, _
                              ByVal RowIndex As Long) As String

    Dim Quantity As Variant
    Dim Value As Variant

    If Len(Trim(Sheet.Cells(RowIndex, COL_CODE).Value)) = 0 Then
        ValidateLine = "code is empty"
        Exit Function
    End If

    Quantity = Sheet.Cells(RowIndex, COL_QUANTITY).Value
    If Not IsNumeric(Quantity) Then
        ValidateLine = "quantity is not a number"
        Exit Function
    ElseIf Quantity <= 0 Then
        ValidateLine = "quantity is not positive"
        Exit Function
    End If

    Value = Sheet.Cells(RowIndex, COL_VALUE).Value
    If Not IsNumeric(Value) Then
        ValidateLine = "value is not a number"
        Exit Function
    ElseIf Value <= 0 Then
        ValidateLine = "value is not positive"
        Exit Function
    End If

    If Len(Trim(Sheet.Cells(RowIndex, COL_DESCRIPTION).Value)) = 0 Then
        ValidateLine = "description is empty"
        Exit Function
    End If

    ValidateLine = ""

End Function


Private Function LastFilledRow(ByVal Sheet As Worksheet) As Long

    Dim RowIndex As Long
    LastFilledRow = FIRST_DATA_ROW - 1

    For RowIndex = FIRST_DATA_ROW To Sheet.Cells(Sheet.Rows.Count, COL_CODE).End(xlUp).Row
        If Len(Trim(Sheet.Cells(RowIndex, COL_CODE).Value)) > 0 Then
            LastFilledRow = RowIndex
        End If
    Next RowIndex

End Function


Private Sub ClearLog(ByVal Sheet As Worksheet, ByVal LastRow As Long)
    Sheet.Range(Sheet.Cells(FIRST_DATA_ROW, COL_LOG), _
                Sheet.Cells(LastRow, COL_LOG)).ClearContents
    Sheet.Range(CELL_STATUS).ClearContents
End Sub


Private Function FormatForSap(ByVal Value As Variant) As String
    ' SAP expects the decimal separator configured in the user profile. Sending
    ' a locale-formatted string is the commonest cause of a rejected line.
    FormatForSap = Replace(Format(Value, "0.00"), ",", ".")
End Function


Private Function ExtractNumber(ByVal Message As String) As String

    Dim Position As Long
    Dim Digits As String
    Dim Character As String

    For Position = 1 To Len(Message)
        Character = Mid(Message, Position, 1)
        If Character >= "0" And Character <= "9" Then
            Digits = Digits & Character
        ElseIf Len(Digits) > 0 Then
            Exit For
        End If
    Next Position

    If Len(Digits) >= 8 Then ExtractNumber = Digits

End Function
