Attribute VB_Name = "modStationRelay"
' =============================================================================
' modStationRelay  --  Data logic for StationRelay Excel macro app
'
' Public API used by frmStationRelay:
'   AppendRow(text, errMsg)  -> Boolean   write a new row to the sheet
'   GetDataSheet()           -> Worksheet reference to the data sheet
'   ShowAndQuit()                         restore Excel + close workbook
'
' Constants controlling which sheet/column to write to:
'   SHEET_NAME   default "Sheet1"
'   DATA_COL     default 1  (column A, 1-based)
'   START_ROW    default 1  (first data row; set to 2 if you have a header row)
' =============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Configuration constants — adjust to match your workbook
' ---------------------------------------------------------------------------
Public Const SHEET_NAME  As String  = "Sheet1"
Public Const DATA_COL    As Long    = 1      ' 1 = column A
Public Const START_ROW   As Long    = 1      ' set to 2 if row 1 is a header

' ---------------------------------------------------------------------------
' AppendRow
' Writes text + display name + timestamp to the next empty row.
' Returns True on success; on failure sets errMsg and returns False.
' ---------------------------------------------------------------------------
Public Function AppendRow(ByVal txt As String, ByRef errMsg As String) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = GetDataSheet()
    If ws Is Nothing Then
        errMsg = "Sheet '" & SHEET_NAME & "' not found in this workbook."
        AppendRow = False
        Exit Function
    End If

    ' Find first empty row in the data column
    Dim nextRow As Long
    nextRow = NextEmptyRow(ws)

    ' Write the three fields
    ws.Cells(nextRow, DATA_COL).Value     = txt
    ws.Cells(nextRow, DATA_COL + 1).Value = GetDisplayName()
    ws.Cells(nextRow, DATA_COL + 2).Value = Format(Now, "YYYY-MM-DD HH:MM:SS")

    ' Save the workbook silently
    ThisWorkbook.Save

    AppendRow = True
    Exit Function

Fail:
    errMsg = Err.Description
    AppendRow = False
End Function

' ---------------------------------------------------------------------------
' GetDataSheet
' Returns the target worksheet, or Nothing if it doesn't exist.
' ---------------------------------------------------------------------------
Public Function GetDataSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    Set GetDataSheet = ws
End Function

' ---------------------------------------------------------------------------
' ShowAndQuit
' Restores Excel visibility and closes the workbook (without re-saving).
' Called when the user clicks Close on the form.
' ---------------------------------------------------------------------------
Public Sub ShowAndQuit()
    Application.Visible = True
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=False
End Sub

' ---------------------------------------------------------------------------
' NextEmptyRow  (private helper)
' Returns the row number of the first empty cell in DATA_COL, starting
' from START_ROW.
' ---------------------------------------------------------------------------
Private Function NextEmptyRow(ws As Worksheet) As Long
    Dim lastUsed As Long
    lastUsed = ws.Cells(ws.Rows.Count, DATA_COL).End(xlUp).Row

    ' If the sheet is blank or only has content before START_ROW
    If lastUsed < START_ROW Then
        NextEmptyRow = START_ROW
        Exit Function
    End If

    ' Check whether the last-used cell is actually populated
    If ws.Cells(lastUsed, DATA_COL).Value = "" Then
        NextEmptyRow = lastUsed
    Else
        NextEmptyRow = lastUsed + 1
    End If
End Function

' ---------------------------------------------------------------------------
' GetDisplayName  (private helper)
' Returns "Firstname Lastname" from Windows / Active Directory.
' Falls back to the Windows login username.
' ---------------------------------------------------------------------------
Private Function GetDisplayName() As String
    On Error GoTo Fallback

    ' WScript.Network gives us the full display name on domain machines
    Dim wn As Object
    Set wn = CreateObject("WScript.Network")

    ' Try to get display name via ADSI (works on AD-joined machines)
    Dim adsiUser As Object
    On Error Resume Next
    Set adsiUser = GetObject("WinNT://" & wn.UserDomain & "/" & wn.UserName & ",user")
    On Error GoTo Fallback

    If Not adsiUser Is Nothing Then
        Dim fullName As String
        fullName = adsiUser.FullName
        If Len(Trim(fullName)) > 0 Then
            GetDisplayName = Trim(fullName)
            Exit Function
        End If
    End If

Fallback:
    ' Last resort: just return the login username
    On Error Resume Next
    Dim wn2 As Object
    Set wn2 = CreateObject("WScript.Network")
    GetDisplayName = wn2.UserName
    If Len(GetDisplayName) = 0 Then GetDisplayName = Environ("USERNAME")
End Function
