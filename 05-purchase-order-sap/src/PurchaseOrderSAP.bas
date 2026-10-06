Attribute VB_Name = "PurchaseOrderSAP"
'==============================================================================
' Stage 05 - Purchase order creation in SAP (ME21N)
'
' Reads a purchase order input file produced by stage 04 and creates the order
' in SAP by adopting the lines from the requisitions it references, rather than
' re-typing items SAP already holds.
'
' The number SAP returns is written back into the workbook, and every line is
' logged as adopted, skipped or failed.
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
Private Const SHEET_NAME As String = "Purchase Order"
Private Const FIRST_DATA_ROW As Long = 9

Private Const COL_ID As Long = 1
Private Const COL_REQUISITION As Long = 2
Private Const COL_DESCRIPTION As Long = 3
Private Const COL_QUANTITY As Long = 4
Private Const COL_GROSS As Long = 5
Private Const COL_NET As Long = 6
Private Const COL_LOG As Long = 8

Private Const CELL_DELIVERY_DATE As String = "E2"
Private Const CELL_PURCHASING_GROUP As String = "E3"
Private Const CELL_COMPANY As String = "E4"
Private Const CELL_ORDER_TYPE As String = "E5"
Private Const CELL_SUPPLIER As String = "E6"
Private Const CELL_ORDER_NUMBER As String = "H2"
Private Const CELL_STATUS As String = "H3"

'--- SAP transaction and screen paths ----------------------------------------
' Re-record these in your own system. See SETUP above.
Private Const TRANSACTION_CODE As String = "/nME21N"
Private Const PATH_COMMAND_FIELD As String = "wnd[0]/tbar[0]/okcd"
Private Const PATH_STATUS_BAR As String = "wnd[0]/sbar"
Private Const PATH_SAVE_BUTTON As String = "wnd[0]/tbar[0]/btn[11]"
Private Const PATH_DOCUMENT_OVERVIEW As String = "wnd[0]/tbar[1]/btn[16]"
Private Const PATH_SELECTION_VARIANT As String = "wnd[0]/tbar[1]/btn[17]"
Private Const PATH_REQUISITION_FIELD As String = "wnd[1]/usr/ctxtENT_BANFN-LOW"
Private Const PATH_SUPPLIER_FIELD As String = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD"
Private Const PATH_ITEM_GRID As String = "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"

'--- Grid column names --------------------------------------------------------
Private Const GRID_QUANTITY As String = "MENGE"
Private Const GRID_PRICE As String = "NETPR"

'--- Module state -------------------------------------------------------------
Private SapSession As Object


'==============================================================================
' Entry point
'==============================================================================

Public Sub CreatePurchaseOrder()

    Dim Sheet As Worksheet
    Dim LastRow As Long
    Dim RowIndex As Long
    Dim GridRow As Long
    Dim Adopted As Long
    Dim Skipped As Long
    Dim Problem As String
    Dim Requisitions As Collection

    Set Sheet = ThisWorkbook.Worksheets(SHEET_NAME)
    LastRow = LastFilledRow(Sheet)

    If LastRow < FIRST_DATA_ROW Then
        MsgBox "No lines to process.", vbInformation
        Exit Sub
    End If

    Problem = ValidateHeader(Sheet)
    If Len(Problem) > 0 Then
        MsgBox "The header is incomplete: " & Problem, vbExclamation
        Exit Sub
    End If

    If Not ConnectToSap() Then Exit Sub
    ClearLog Sheet, LastRow

    OpenTransaction
    SetSupplier Sheet

    ' Adopt the lines of every requisition this order references. A purchase
    ' order may draw on more than one requisition, so the list is collected
    ' first and each number is pulled in once.
    Set Requisitions = DistinctRequisitions(Sheet, LastRow)
    AdoptRequisitions Requisitions

    GridRow = 0
    For RowIndex = FIRST_DATA_ROW To LastRow

        Problem = ValidateLine(Sheet, RowIndex)

        If Len(Problem) > 0 Then
            Sheet.Cells(RowIndex, COL_LOG).Value = "Skipped: " & Problem
            Skipped = Skipped + 1
        Else
            On Error Resume Next
            ApplyLineValues Sheet, RowIndex, GridRow

            If Err.Number <> 0 Then
                Sheet.Cells(RowIndex, COL_LOG).Value = "Error: " & Err.Description
                Err.Clear
                Skipped = Skipped + 1
            Else
                Sheet.Cells(RowIndex, COL_LOG).Value = "Adopted"
                GridRow = GridRow + 1
                Adopted = Adopted + 1
            End If
            On Error GoTo 0
        End If

    Next RowIndex

    If Adopted = 0 Then
        Sheet.Range(CELL_STATUS).Value = "Nothing sent"
        MsgBox "No valid lines. Nothing was sent to SAP.", vbExclamation
        Exit Sub
    End If

    SaveAndCapture Sheet, Adopted, Skipped

End Sub


'==============================================================================
' SAP session
'==============================================================================

Private Function ConnectToSap() As Boolean

    Dim SapGui As Object
    Dim GuiApplication As Object
    Dim Connection As Object

    On Error GoTo NoSession

    Set SapGui = GetObject("SAPGUI")
    Set GuiApplication = SapGui.GetScriptingEngine
    Set Connection = GuiApplication.Children(0)
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
' Building the order
'==============================================================================

Private Sub SetSupplier(ByVal Sheet As Worksheet)
    SapSession.findById(PATH_SUPPLIER_FIELD).Text = _
        CStr(Sheet.Range(CELL_SUPPLIER).Value)
    SapSession.findById("wnd[0]").sendVKey 0
End Sub


Private Function DistinctRequisitions(ByVal Sheet As Worksheet, _
                                      ByVal LastRow As Long) As Collection

    Dim Numbers As New Collection
    Dim RowIndex As Long
    Dim Number As String

    On Error Resume Next
    For RowIndex = FIRST_DATA_ROW To LastRow
        Number = Trim(CStr(Sheet.Cells(RowIndex, COL_REQUISITION).Value))
        If Len(Number) > 0 Then
            ' Adding with the number as key raises on a repeat, which is how
            ' duplicates are discarded without a second pass.
            Numbers.Add Number, Number
            Err.Clear
        End If
    Next RowIndex
    On Error GoTo 0

    Set DistinctRequisitions = Numbers

End Function


Private Sub AdoptRequisitions(ByVal Numbers As Collection)

    Dim Number As Variant

    SapSession.findById(PATH_DOCUMENT_OVERVIEW).press

    For Each Number In Numbers
        SapSession.findById(PATH_SELECTION_VARIANT).press
        SapSession.findById(PATH_REQUISITION_FIELD).Text = CStr(Number)
        SapSession.findById("wnd[1]").sendVKey 0
        SapSession.findById("wnd[0]").sendVKey 0
    Next Number

End Sub


Private Sub ApplyLineValues(ByVal Sheet As Worksheet, ByVal RowIndex As Long, _
                            ByVal GridRow As Long)

    ' The adopted line already carries the item and its text. Only quantity and
    ' price are reapplied, because the order may differ from what was requested.

    Dim Grid As Object
    Set Grid = SapSession.findById(PATH_ITEM_GRID)

    Grid.modifyCell GridRow, GRID_QUANTITY, _
                    CStr(Sheet.Cells(RowIndex, COL_QUANTITY).Value)
    Grid.modifyCell GridRow, GRID_PRICE, _
                    FormatForSap(Sheet.Cells(RowIndex, COL_NET).Value)

    SapSession.findById("wnd[0]").sendVKey 0

End Sub


Private Sub SaveAndCapture(ByVal Sheet As Worksheet, ByVal Adopted As Long, _
                           ByVal Skipped As Long)

    Dim Message As String
    Dim Number As String

    SapSession.findById(PATH_SAVE_BUTTON).press
    Message = SapSession.findById(PATH_STATUS_BAR).Text

    Number = ExtractNumber(Message)

    If Len(Number) > 0 Then
        Sheet.Range(CELL_ORDER_NUMBER).Value = Number
        Sheet.Range(CELL_STATUS).Value = "Created"
        MsgBox "Purchase order " & Number & " created." & vbCrLf & _
               Adopted & " lines adopted, " & Skipped & " skipped.", vbInformation
    Else
        Sheet.Range(CELL_STATUS).Value = "Not created"
        MsgBox "SAP did not return a purchase order number." & vbCrLf & vbCrLf & _
               "Status bar: " & Message, vbExclamation
    End If

End Sub


'==============================================================================
' Helpers
'==============================================================================

Private Function ValidateHeader(ByVal Sheet As Worksheet) As String

    If Len(Trim(Sheet.Range(CELL_SUPPLIER).Value)) = 0 Then
        ValidateHeader = "supplier is empty"
    ElseIf Len(Trim(Sheet.Range(CELL_ORDER_TYPE).Value)) = 0 Then
        ValidateHeader = "order type is empty"
    ElseIf Len(Trim(Sheet.Range(CELL_PURCHASING_GROUP).Value)) = 0 Then
        ValidateHeader = "purchasing group is empty"
    ElseIf Not IsDate(Sheet.Range(CELL_DELIVERY_DATE).Value) Then
        ValidateHeader = "delivery date is not a date"
    Else
        ValidateHeader = ""
    End If

End Function


Private Function ValidateLine(ByVal Sheet As Worksheet, _
                              ByVal RowIndex As Long) As String

    Dim Quantity As Variant
    Dim Price As Variant

    If Len(Trim(Sheet.Cells(RowIndex, COL_REQUISITION).Value)) = 0 Then
        ValidateLine = "requisition number is empty"
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

    Price = Sheet.Cells(RowIndex, COL_NET).Value
    If Not IsNumeric(Price) Then
        ValidateLine = "net value is not a number"
        Exit Function
    ElseIf Price <= 0 Then
        ValidateLine = "net value is not positive"
        Exit Function
    End If

    ValidateLine = ""

End Function


Private Function LastFilledRow(ByVal Sheet As Worksheet) As Long

    Dim RowIndex As Long
    LastFilledRow = FIRST_DATA_ROW - 1

    For RowIndex = FIRST_DATA_ROW To _
                   Sheet.Cells(Sheet.Rows.Count, COL_REQUISITION).End(xlUp).Row
        If Len(Trim(Sheet.Cells(RowIndex, COL_REQUISITION).Value)) > 0 Then
            LastFilledRow = RowIndex
        End If
    Next RowIndex

End Function


Private Sub ClearLog(ByVal Sheet As Worksheet, ByVal LastRow As Long)
    Sheet.Range(Sheet.Cells(FIRST_DATA_ROW, COL_LOG), _
                Sheet.Cells(LastRow, COL_LOG)).ClearContents
    Sheet.Range(CELL_ORDER_NUMBER).ClearContents
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
