VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

'    Application.ScreenUpdating = False
'    ActiveWindow.DisplayHorizontalScrollBar = False
'    ActiveWindow.DisplayVerticalScrollBar = False
'    ActiveWindow.DisplayHeadings = False
'    ActiveWindow.DisplayWorkbookTabs = False
'    Application.DisplayFormulaBar = False
'    Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"", False)"
'    Application.WindowState = xlMaximized
'    Application.DisplayStatusBar = False
'
'    meineCloseVar = False
'    Tabelle2.ScrollArea = "A1:A1"
'    Application.ScreenUpdating = True

End Sub

'
'Sub ScreenFull()
'
'Dim wbk As Workbook
'
'
'  'Fullscreen anzeigen
'  If Not wbk Is Nothing Then
'    If MsgBox("Die Fullscreeneinstellungen für """ & wbk.Name _
'      & """ wurden noch nicht zurückgesetzt." _
'      & vbLf & "Makro ""ScreenNormal"" jetzt starten?", _
'            vbQuestion + vbOKCancel, "Fullscreen zurücksetzen") = vbOK Then
'      Call ScreenNormal
'    Else
'      Exit Sub
'    End If
'  End If
'  Set wbk = ActiveWorkbook
'  With Application.Windows(wbk.Name)
'    bolHeadings = .DisplayHeadings
'    .DisplayHeadings = False
'    bolHScrollBar = .DisplayHorizontalScrollBar
'    .DisplayHorizontalScrollBar = False
'    bolVScrollbar = .DisplayVerticalScrollBar
'    .DisplayVerticalScrollBar = False
'    bolWorkTabs = .DisplayWorkbookTabs
'    .DisplayWorkbookTabs = False
'  End With
'  With Application
'    .DisplayFormulaBar = False
'    .DisplayFullScreen = True
'  End With
'End Sub
'
'Sub ScreenNormal()
'Dim wbk As Workbook
'Set wbk = ActiveWorkbook
'  'Fullscreen zurücksetzen
'  If wbk Is Nothing Then
'    MsgBox "Makro ""ScreenFull"" wurde noch nicht gestartet.", _
'          vbInformation, "Fullscreen zurücksetzen"
'  Else
'    With Application.Windows(wbk.Name)
'      .DisplayHeadings = bolHeadings
'      .DisplayHorizontalScrollBar = bolHScrollBar
'      .DisplayVerticalScrollBar = bolVScrollbar
'      .DisplayWorkbookTabs = bolWorkTabs
'    End With
'    With Application
'      .DisplayFormulaBar = True
'      .DisplayFullScreen = False
'    End With
'  End If
'  Set wbk = Nothing
'End Sub



Sub Auto_Close()

    Application.ScreenUpdating = False
    
    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
    End If
    
    'Call ScreenNormal
    If Application.Workbooks.Count = 1 Then
        ActiveWorkbook.Saved = True
        Application.Quit
    Else
        ActiveWorkbook.Close SaveChanges:=True
    End If
        
    Application.DisplayFormulaBar = True
    Application.ScreenUpdating = True
  
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

  Cancel = Not meineCloseVar
 
End Sub




