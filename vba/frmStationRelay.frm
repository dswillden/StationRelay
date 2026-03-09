VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStationRelay
   Caption         =   "StationRelay"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStationRelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' frmStationRelay  --  StationRelay custom application UI
'
' Controls (all created in code so no .frx binary is needed):
'   lblTitle        -- app title
'   txtInput        -- multi-line text entry
'   btnSend         -- submit button  (Ctrl+Enter also works)
'   lblStatus       -- status / error feedback
'   lblHistHeader   -- "Submitted Entries" section label
'   lsvHistory      -- ListBox showing all rows, newest-first
'   btnClose        -- close / exit
' =============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Form-level control references (created dynamically in Initialize)
' ---------------------------------------------------------------------------
Private WithEvents txtInput     As MSForms.TextBox
Private WithEvents btnSend      As MSForms.CommandButton
Private WithEvents btnClose     As MSForms.CommandButton
Private WithEvents lsvHistory   As MSForms.ListBox
Private lblTitle        As MSForms.Label
Private lblStatus       As MSForms.Label
Private lblHistHeader   As MSForms.Label

' ---------------------------------------------------------------------------
' UserForm_Initialize  --  build the UI programmatically
' ---------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' ---- Form dimensions ---------------------------------------------------
    Me.Caption = "StationRelay"
    Me.Width   = 520
    Me.Height  = 600
    Me.BorderStyle = fmBorderStyleSingle

    ' ---- Title label -------------------------------------------------------
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Caption  = "StationRelay"
        .Left     = 10
        .Top      = 8
        .Width    = 480
        .Height   = 22
        .Font.Size = 13
        .Font.Bold = True
        .ForeColor = RGB(0, 84, 166)
    End With

    ' ---- Text input label --------------------------------------------------
    Dim lblInput As MSForms.Label
    Set lblInput = Me.Controls.Add("Forms.Label.1", "lblInput")
    With lblInput
        .Caption = "Paste your text below  (Ctrl+Enter to send)"
        .Left    = 10
        .Top     = 36
        .Width   = 480
        .Height  = 14
        .Font.Size = 9
        .ForeColor = RGB(80, 80, 80)
    End With

    ' ---- Text input area ---------------------------------------------------
    Set txtInput = Me.Controls.Add("Forms.TextBox.1", "txtInput")
    With txtInput
        .Left        = 10
        .Top         = 54
        .Width       = 480
        .Height      = 100
        .MultiLine   = True
        .ScrollBars  = fmScrollBarsVertical
        .WordWrap    = True
        .Font.Size   = 10
        .EnterKeyBehavior = True   ' Enter inserts newline; Ctrl+Enter fires Send
    End With

    ' ---- Send button -------------------------------------------------------
    Set btnSend = Me.Controls.Add("Forms.CommandButton.1", "btnSend")
    With btnSend
        .Caption   = "Send"
        .Left      = 380
        .Top       = 160
        .Width     = 110
        .Height    = 26
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = RGB(0, 120, 212)
        .ForeColor = RGB(255, 255, 255)
    End With

    ' ---- Status label ------------------------------------------------------
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Caption   = "Ready."
        .Left      = 10
        .Top       = 164
        .Width     = 360
        .Height    = 18
        .Font.Size = 9
        .ForeColor = RGB(80, 80, 80)
    End With

    ' ---- History section header -------------------------------------------
    Set lblHistHeader = Me.Controls.Add("Forms.Label.1", "lblHistHeader")
    With lblHistHeader
        .Caption   = "Submitted Entries  (newest first)"
        .Left      = 10
        .Top       = 192
        .Width     = 480
        .Height    = 14
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = RGB(60, 60, 60)
    End With

    ' ---- History list ------------------------------------------------------
    Set lsvHistory = Me.Controls.Add("Forms.ListBox.1", "lsvHistory")
    With lsvHistory
        .Left       = 10
        .Top        = 210
        .Width      = 480
        .Height     = 310
        .Font.Size  = 9
        .BorderStyle = fmBorderStyleSingle
        .ColumnCount = 3
        .ColumnWidths = "100 pt;90 pt;260 pt"   ' Timestamp | Name | Text
        .MultiSelect = fmMultiSelectSingle
    End With

    ' ---- Close button ------------------------------------------------------
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption   = "Close"
        .Left      = 380
        .Top       = 528
        .Width     = 110
        .Height    = 24
        .Font.Size = 9
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(30, 30, 30)
    End With

    ' ---- Populate the history list on first load ---------------------------
    RefreshHistory
End Sub

' ---------------------------------------------------------------------------
' Send button click
' ---------------------------------------------------------------------------
Private Sub btnSend_Click()
    Dim txt As String
    txt = Trim(txtInput.Text)

    If Len(txt) = 0 Then
        SetStatus "Nothing to send — type something first.", RGB(180, 60, 0)
        Exit Sub
    End If

    SetStatus "Saving...", RGB(0, 100, 180)
    Me.Repaint

    Dim errMsg As String
    If modStationRelay.AppendRow(txt, errMsg) Then
        txtInput.Text = ""
        SetStatus "Sent successfully.", RGB(0, 130, 0)
        RefreshHistory
    Else
        SetStatus "Error: " & errMsg, RGB(180, 0, 0)
    End If
End Sub

' ---------------------------------------------------------------------------
' Ctrl+Enter in the text box triggers Send
' ---------------------------------------------------------------------------
Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Shift=2 means Ctrl is held; KeyCode 13 = Enter
    If KeyCode = 13 And Shift = 2 Then
        KeyCode = 0
        btnSend_Click
    End If
End Sub

' ---------------------------------------------------------------------------
' Close button
' ---------------------------------------------------------------------------
Private Sub btnClose_Click()
    modStationRelay.ShowAndQuit
End Sub

' ---------------------------------------------------------------------------
' RefreshHistory  --  reload lsvHistory from the sheet (newest row first)
' ---------------------------------------------------------------------------
Public Sub RefreshHistory()
    lsvHistory.Clear

    Dim ws As Worksheet
    Set ws = modStationRelay.GetDataSheet()
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, modStationRelay.DATA_COL).End(xlUp).Row
    If lastRow < modStationRelay.START_ROW Then Exit Sub

    Dim i As Long
    For i = lastRow To modStationRelay.START_ROW Step -1
        Dim stamp   As String
        Dim name    As String
        Dim display As String
        stamp   = CStr(ws.Cells(i, modStationRelay.DATA_COL + 2).Value)
        name    = CStr(ws.Cells(i, modStationRelay.DATA_COL + 1).Value)
        display = CStr(ws.Cells(i, modStationRelay.DATA_COL).Value)
        If Len(display) > 80 Then display = Left(display, 77) & "..."

        lsvHistory.AddItem stamp
        lsvHistory.List(lsvHistory.ListCount - 1, 1) = name
        lsvHistory.List(lsvHistory.ListCount - 1, 2) = display
    Next i
End Sub

' ---------------------------------------------------------------------------
' SetStatus  (private helper)
' ---------------------------------------------------------------------------
Private Sub SetStatus(ByVal msg As String, ByVal colour As Long)
    lblStatus.Caption  = msg
    lblStatus.ForeColor = colour
End Sub
