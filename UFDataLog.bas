VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFDataLog 
   ClientHeight    =   15150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   29910
   OleObjectBlob   =   "UFDataLog.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "UFDataLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Option Compare Text

' ************************************************************************************************************************************************************************************
' Konstanten, Parametrisierung

Private Const iCONST_ANZAHL_EINGABEFELDER As Integer = 141          ' Wie viele Textboxen sind auf der UserForm platziert?
Private Const lCONST_STARTZEILENNUMMER_DER_TABELLE As Long = 2      ' In welcher Zeile starten die Eingaben?
Public weitMiet1, weitMiet2 As Boolean                              ' Merker, ob ein oder zwei weitere Mieter eingetragen sind
Public strTmp2, strTmp3 As String                                   ' Variablen zur Prüfung, ob der 2./3. Mieter angezeigt werden muss
Public varQuelle As String                                          ' Variable zur Feststellung, ob der Kopiermodus aktiv ist
Public CopyModeOn, AddModeOn, ProdModeOn As Boolean                 ' Merker, ob Kopiermodus/Erfassen-Modus/Modulbuchung aktiv sind
Public i1, i2, i3, i4, i5, i6, iX, iXneu As Integer                 ' zentrale Variablen zum Zuordnen der gebuchten Module
                                                                    ' Sonderbereich zum Deaktivieren der Funktionen in der Titelleiste
                                                                    ' (Schließen, Minimieren, Maximieren)
Public ErrCount As Integer                                          ' Zähler für Pflichfelder
                                                                    ' (Speichern erst möglich, wenn ErrCount = 0)

Private Declare Function FindWindow Lib "user32" Alias _
      "FindWindowA" (ByVal lpClassName As String, ByVal _
      lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias _
      "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex _
      As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
      "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex _
      As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal _
      hwnd As Long) As Long
Private Const GWL_STYLE As Long = -16
Private Const WS_SYSMENU As Long = &H80000
Private hwndForm As Long
Private bCloseBtn As Boolean                                        ' Ende Sonderbereich Titelleiste




' ************************************************************************************************************************************************************************************
' TESTBEREICH








' ENDE TESTBEREICH
' ************************************************************************************************************************************************************************************


' ************************************************************************************************************************************************************************************
' Userform

Private Sub UserForm_Initialize()                                   ' Startroutine bevor die UserForm angezeigt wird

    If Val(Application.Version) >= 9 Then                           ' Sonderbereich zum Deaktivieren der Funktionen in der Titelleiste
        hwndForm = FindWindow("ThunderDFrame", Me.Caption)
    Else
        hwndForm = FindWindow("ThunderXFrame", Me.Caption)
    End If
    
    bCloseBtn = False
    SET_USERFORM_STYLE
    
   Call LISTE_LADEN_UND_INITIALISIEREN                             ' Aufruf der entsprechenden Verarbeitungsroutine
'
'    Call BUTTON_STANDARD                                             ' Buttons auf Grundeinstellung setzen
    
End Sub

Private Sub UserForm_Activate()                                     ' Ereignisroutine beim Anzeigen der UserForm
    
    With UFDataUpload                                               ' Anpassen der Größe der Userform auf die Größe des aktuellen Anwendung
        .Top = Application.Top                                      ' (da Excel im Vollbildmodus gestartet wird, wird dieser dann auch hier übernommen)
        .Left = Application.Left
        .Height = Application.Height
        .Width = Application.Width
    End With
    
    If ListBox1a.ListCount > 0 Then ListBox1a.ListIndex = 0           ' 1. Eintrag selektieren
    
'    Call BUTTON_STANDARD                                             ' Buttons auf Grundeinstellung setzen

End Sub


    Private Sub UserForm_layout()
    With Me
        .StartUpPosition = 0
        .Top = -15
        .Left = ActiveWindow.Left + 3
        .Width = 1250
        .Height = 780
    End With
End Sub
' ************************************************************************************************************************************************************************************
' Listbox(en)

Private Sub Listbox1aa_Click()                                        ' ListBox Ereignisroutine
    
'    Call EINTRAG_LADEN_UND_ANZEIGEN                                 ' Aufruf der entsprechenden Verarbeitungsroutine
    
End Sub


' ************************************************************************************************************************************************************************************
' Button(s)



' ************************************************************************************************************************************************************************************
' Verarbeitungsroutinen

Private Sub LISTE_LADEN_UND_INITIALISIEREN()                        ' Routine um die ListBox zu leeren, einzustellen und neu zu füllen
    
    Dim lZeile As Long                                              ' erforderliche Variablen definieren
    Dim lZeileMaximum As Long
    Dim i As Integer
    Dim Log, varQuelle As String
    Dim GesamtFord
    
    
    ErrCount = 0                                                    ' Zähler für nicht ausgefüllte Pflichtfelder auf 0 setzen
    varQuelle = "LOG"
    
    tbx0a = CStr(Worksheets("PARAM").Cells(17, 6).Text)             ' Eintrag Mandantennummer aus Parametern
    tbx0a.ForeColor = RGB(72, 209, 204)                             ' Schriftfarbe setzen
  

    ListBox1a.Clear                                                 ' Listbox leeren
    
    ListBox1a.ColumnCount = 2                                       ' = Anzahl der Spalten (mehr als 10 bei ungebundenen Listboxen nicht möglich)
                                                                    ' Spaltenbreiten der Liste anpassen (0=ausblenden, nichts=automatisch)
    ListBox1a.ColumnWidths = "0;100"                                ' (<Breite Spalte 1>;<Breite Spalte 2>;etc.)
                                                                        
    lZeileMaximum = Worksheets(varQuelle).UsedRange.Rows.Count      ' letzte verwendete Zeile ermitteln und benutzten Bereich auslesen
    
    For lZeile = lCONST_STARTZEILENNUMMER_DER_TABELLE To lZeileMaximum
    
    Log = CStr(Worksheets(varQuelle).Cells(lZeile, 1).Text)
    ListBox1a.AddItem lZeile
    ListBox1a.List(ListBox1a.ListCount - 1, 1) = Log

     
    Next lZeile

End Sub




' ************************************************************************************************************************************************************************************
' Hilfsfunktionen /-prozeduren


Private Sub SET_USERFORM_STYLE()                                                      ' zum Deaktivieren der Funktionen in der Titelleiste
                                                                                    ' (siehe Sonderbereich im Kopf des Codes)
    Dim frmStyle As Long
    
    If hwndForm = 0 Then Exit Sub
    
    frmStyle = GetWindowLong(hwndForm, GWL_STYLE)
    
    If bCloseBtn Then
      frmStyle = frmStyle Or WS_SYSMENU
    Else
      frmStyle = frmStyle And Not WS_SYSMENU
    End If
    
    SetWindowLong hwndForm, GWL_STYLE, frmStyle
    DrawMenuBar hwndForm
    
End Sub






Private Sub cBnBckLog_Click()
        Unload Me
End Sub
