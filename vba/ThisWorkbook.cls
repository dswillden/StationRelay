Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' =============================================================================
' ThisWorkbook  --  Auto-hides Excel and launches the StationRelay form
' =============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Workbook_Open
' Fires automatically when the file is opened.
' Hides the entire Excel application window and shows the custom form.
' ---------------------------------------------------------------------------
Private Sub Workbook_Open()
    ' Give Excel a moment to finish loading before hiding
    Application.Wait Now + TimeSerial(0, 0, 1)

    Application.Visible = False     ' hide the Excel shell
    Application.DisplayAlerts = False

    ' Show the form as modeless so the user can interact with it freely.
    ' If you want to block VBA until the form is closed use vbModal instead.
    frmStationRelay.Show vbModeless
End Sub

' ---------------------------------------------------------------------------
' Workbook_BeforeClose
' If the user closes via the taskbar / title bar (not the form's Close
' button), make Excel visible again so it shuts down cleanly without
' leaving an invisible orphan process.
' ---------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.Visible = True
    Application.DisplayAlerts = True
End Sub
