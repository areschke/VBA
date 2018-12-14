VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFDataUpload 
   ClientHeight    =   15150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   29910
   OleObjectBlob   =   "UFDataUpload.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UFDataUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Compare Text

' ###################################################################################################################################################################################
' +++++++++++++ Konstanten, Parametrisierung ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################

Private arrData As Variant                                          ' Variablen für Filterfunktionen
Private wksData As Worksheet
Private arrList As Variant

Private Const iCONST_ANZAHL_EINGABEFELDER As Integer = 126          ' Wie viele Textboxen sind auf der UserForm platziert?
Private Const lCONST_STARTZEILENNUMMER_DER_TABELLE As Long = 2      ' In welcher Zeile starten die Eingaben?
Public weitMiet1, weitMiet2 As Boolean                              ' Merker, ob ein oder zwei weitere Mieter eingetragen sind
Public strTmp2, strTmp3 As String                                   ' Variablen zur Prüfung, ob der 2./3. Mieter angezeigt werden muss
Public varQuelle As String                                          ' Variable zur Feststellung, ob der Kopiermodus aktiv ist
Public CopyModeOn, AddModeOn, ProdModeOn As Boolean                 ' Merker, ob Kopiermodus/Erfassen-Modus/Modulbuchung aktiv sind
Public i1, i2, i3, i4, i5, i6, iX, iC, iXneu As Integer             ' zentrale Variablen zum Zuordnen der gebuchten Module
                                                                    ' Sonderbereich zum Deaktivieren der Funktionen in der Titelleiste
                                                                    ' (Schließen, Minimieren, Maximieren)
Public ErrCount As Integer                                          ' Zähler für Pflichfelder
                                                                    ' (Speichern erst möglich, wenn ErrCount = 0)
Public varError, varTipText As String                               ' Variablen zum Füllen mit dem jeweiligen Objekt, dass als Pflichteingabe definiert ist
Public SavOK As Boolean                                             ' Variable, ob der OK Button geklickt wurde, um beim Aktualisierung nicht wieder auf den ersten Datensatz
                                                                    ' zu springen
Public i108, i109, i33, i110, i119, i120, i122, i123 As Integer     ' Variablen zum Auswerten der Änderungen beim Datumsfeldern
Public i33_M, i33_A, i33_R, i33_S As Integer                        ' dto.
Public lZeileMaximum As Long
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
Private bCloseBtn As Boolean

Public bypass As Boolean

' ###############################################################################################################################################################################
' TESTBEREICH

Private Sub cBnUeber_Click()

    Call ARCHIV_UEBERNAHME
    CopyModeOn = False
    Call LISTE_LADEN_UND_INITIALISIEREN

End Sub

' ENDE TESTBEREICH
' ###################################################################################################################################################################################






' ###################################################################################################################################################################################
' +++++++++++++ Verarbeitungsroutinen +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################


'   LISTE_LADEN_UND_INITIALISIEREN(Routine um die ListBox zu leeren, einzustellen und neu zu füllen)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LISTE_LADEN_UND_INITIALISIEREN()
    
    Dim lZeile As Long                                              ' erforderliche Variablen definieren
    Dim i As Integer
    Dim MieterKey, Name, Anschrift, varQuelle As String
    Dim GesamtFord
    
    tbx0 = CStr(Worksheets("PARAM").Cells(17, 6).Text)              ' Eintrag Mandantennummer aus Parametern
    tbx0.ForeColor = RGB(72, 209, 204)                              ' Schriftfarbe setzen
  
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' alle Textboxen leeren
        Me.Controls("tbx" & i) = ""
        On Error Resume Next
    Next i

    ListBox1.Clear                                                  ' Listbox leeren
    
    ListBox1.ColumnCount = 10                                       ' = Anzahl der Spalten (mehr als 10 bei ungebundenen Listboxen nicht möglich)
                                                                    ' Spaltenbreiten der Liste anpassen (0=ausblenden, nichts=automatisch)
    ListBox1.ColumnWidths = "0;150;150;300;95;95;95;95;115;95"      ' (<Breite Spalte 1>;<Breite Spalte 2>;etc.)
                                                                        
    If CopyModeOn = True Then
        varQuelle = "DATA_UPLOAD_ARCHIV"
    Else
        varQuelle = "DATA_UPLOAD"
    End If
                                                                      
    lZeileMaximum = Worksheets(varQuelle).UsedRange.Rows.Count      ' letzte verwendete Zeile ermitteln und benutzten Bereich auslesen
    
    For lZeile = lCONST_STARTZEILENNUMMER_DER_TABELLE To lZeileMaximum
                                                                    ' Zusammensetzen der Spalteninhalte der Listbox für die Anzeige
                                                                    ' a) MieterKey =  WE.HausNr.WohnNr.WohnNrZus.Folgenummer (wenn WohnZusNr nicht leer ist) bzw.
                                                                    '                 WE.HausNr.WohnNr.Folgenummer (wenn WohnZusNr leer ist)
    If CStr(Worksheets(varQuelle).Cells(lZeile, 51).Text) = "" Then
        MieterKey = CStr(Worksheets(varQuelle).Cells(lZeile, 22).Text) _
                & "." & CStr(Worksheets(varQuelle).Cells(lZeile, 50).Text) _
                & "." & CStr(Worksheets(varQuelle).Cells(lZeile, 25).Text) _
                & "." & CStr(Worksheets(varQuelle).Cells(lZeile, 12).Text)
                Else
        MieterKey = CStr(Worksheets(varQuelle).Cells(lZeile, 22).Text) _
                & "." & CStr(Worksheets(varQuelle).Cells(lZeile, 50).Text) _
                & "." & CStr(Worksheets(varQuelle).Cells(lZeile, 25).Text) _
                & "." & CStr(Worksheets(varQuelle).Cells(lZeile, 51).Text) _
                & "." & CStr(Worksheets(varQuelle).Cells(lZeile, 12).Text)
    End If

                                                                    ' b) Name = Name1, Name2
    Name = CStr(Worksheets(varQuelle).Cells(lZeile, 60).Text) & ", " & CStr(Worksheets(varQuelle).Cells(lZeile, 61).Text)
        
                                                                    ' c) Anschrift = Strasse, PLZ Ort

    Anschrift = CStr(Worksheets(varQuelle).Cells(lZeile, 64).Text) _
        & " " & CStr(Worksheets(varQuelle).Cells(lZeile, 65).Text) _
        & ", " & CStr(Worksheets(varQuelle).Cells(lZeile, 66).Text) _
        & " " & CStr(Worksheets(varQuelle).Cells(lZeile, 67).Text) _
        & ", " & CStr(Worksheets(varQuelle).Cells(lZeile, 68).Text)
                                                                    ' d) Gesamtforderung
    GesamtFord = Format(((Worksheets(varQuelle).Cells(lZeile, 26) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 29) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 32) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 35) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 38) * 1)) _
        , "##,##0.00")
        
    If IST_ZEILE_LEER(lZeile) = False Then
    
        ListBox1.AddItem lZeile                                     ' Spalte 1 der Liste mit der Zeilennummer füllen
        ListBox1.List(ListBox1.ListCount - 1, 1) = MieterKey        ' Spalten 2 bis 7 der Liste füllen
        ListBox1.List(ListBox1.ListCount - 1, 2) = Name
        ListBox1.List(ListBox1.ListCount - 1, 3) = Anschrift
        ListBox1.List(ListBox1.ListCount - 1, 4) = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 26).Text), "##,##0.00")
        ListBox1.List(ListBox1.ListCount - 1, 5) = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 29).Text), "##,##0.00")
        ListBox1.List(ListBox1.ListCount - 1, 6) = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 32).Text), "##,##0.00")
        ListBox1.List(ListBox1.ListCount - 1, 7) = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 35).Text), "##,##0.00")
        ListBox1.List(ListBox1.ListCount - 1, 8) = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 38).Text), "##,##0.00")
        ListBox1.List(ListBox1.ListCount - 1, 9) = GesamtFord
    End If

    Next lZeile
    
    mPg1.Value = 0                                                  ' Fokus auf 1. Seite der Multipage setzen
    ListBox1.MultiSelect = fmMultiSelectSingle                      ' Mehrfachauswahl deaktivieren
    ListBox1.ListStyle = fmListStylePlain
    
    Call SUMMENFELDER                                               ' Aufruf Verarbeitungsroutine zum Berechnen der Summenfelder

End Sub

'   EINTRAG_LADEN_UND_ANZEIGEN(Füllen der Textboxen anhand des ausgewählten Datensatzes in der Listbox)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EINTRAG_LADEN_UND_ANZEIGEN()

    Dim lZeile As Long                                              ' erforderliche Variablen definieren
    Dim i As Integer
    Dim GesamtFord
    Dim varQuelle As String

    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Eingabefelder resetten
        Me.Controls("tbx" & i) = ""
        On Error Resume Next
    Next i
    
    If ListBox1.ListIndex >= 0 Then                                 ' Prüfung, ob ein Eintrag selektiert/markiert ist
        lZeile = ListBox1.List(ListBox1.ListIndex, 0)               ' Zugriff über die Zeilennummer des Datensatzes in der ersten ausgeblendeten Spalte
    End If

    If lZeile = 0 Then
        CSVempty = True
    Else
        CSVempty = False
    End If
    
    If CopyModeOn = True Then
        varQuelle = "DATA_UPLOAD_ARCHIV"
    Else
        varQuelle = "DATA_UPLOAD"
    End If
    
    GesamtFord = Format(((Worksheets(varQuelle).Cells(lZeile, 26) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 29) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 32) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 35) * 1) _
        + (Worksheets(varQuelle).Cells(lZeile, 38) * 1)) _
        , "##,##0.00")
                                                                                    ' Addition von Hauptforderung (F26), Mahnkosten (F29), Auskunftskosten (F32),
                                                                                    ' RLS-Gebühren (F35) und sonstigen Nebenforderungen (F38)

    tbx0 = CStr(Worksheets(varQuelle).Cells(lZeile, 1).Text)                        ' Mandantennummer des Endkunden bei SEPA Collect
    tBx1 = CStr(Worksheets(varQuelle).Cells(lZeile, 48).Text)                       ' Mandant
    tBx2 = CStr(Worksheets(varQuelle).Cells(lZeile, 49).Text)                       ' Unternehmen
    tBx3 = CStr(Worksheets(varQuelle).Cells(lZeile, 22).Text)                       ' WE
    tBx4 = CStr(Worksheets(varQuelle).Cells(lZeile, 50).Text)                       ' HausNr
    tBx5 = CStr(Worksheets(varQuelle).Cells(lZeile, 25).Text)                       ' WohnNr
    tBx6 = CStr(Worksheets(varQuelle).Cells(lZeile, 51).Text)                       ' WohnNrZus
    tBx7 = CStr(Worksheets(varQuelle).Cells(lZeile, 12).Text)                       ' Folgenummer
    tBx8 = CStr(Worksheets(varQuelle).Cells(lZeile, 13).Text)                       ' OP-Nummer (Belegnummer)
    tBx9 = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 26).Text), "##,##0.00")  ' Hauptforderung
    tBx10 = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 29).Text), "##,##0.00") ' Mahnkosten
    tBx11 = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 32).Text), "##,##0.00") ' Auskunftskosten
    tBx12 = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 35).Text), "##,##0.00") ' RLS-Gebühren
    tBx13 = Format(CStr(Worksheets(varQuelle).Cells(lZeile, 38).Text), "##,##0.00") ' Sonstige Nebenforderungen
    tBx14 = GesamtFord                                                              ' Gesamtforderung (Berechnung s. o.)
    tBx15 = CStr(Worksheets(varQuelle).Cells(lZeile, 59).Text)                      ' Anrede
    cBxAnrede = CStr(Worksheets(varQuelle).Cells(lZeile, 59).Text)
    tBx16 = CStr(Worksheets(varQuelle).Cells(lZeile, 60).Text)                      ' Name
    tBx17 = CStr(Worksheets(varQuelle).Cells(lZeile, 61).Text)                      ' Vorname
    tBx18 = CStr(Worksheets(varQuelle).Cells(lZeile, 70).Text)                      ' Adresszusatz (c/o Zeile)
    tBx19 = CStr(Worksheets(varQuelle).Cells(lZeile, 58).Text)                      ' AdressNr
    tBx20 = CStr(Worksheets(varQuelle).Cells(lZeile, 64).Text)                      ' Zustelladresse: Straße
    tBx66 = CStr(Worksheets(varQuelle).Cells(lZeile, 65).Text)                      ' Zustelladresse: Hausnummer
    tBx21 = CStr(Worksheets(varQuelle).Cells(lZeile, 69).Text)                      ' Zustelladresse: Nation
    tBx22 = CStr(Worksheets(varQuelle).Cells(lZeile, 66).Text)                      ' Zustelladresse: PLZ
    tBx23 = CStr(Worksheets(varQuelle).Cells(lZeile, 67).Text)                      ' Zustelladresse: Ort
    tBx24 = CStr(Worksheets(varQuelle).Cells(lZeile, 68).Text)                      ' Zustelladresse: Ortsteil
    tBx25 = CStr(Worksheets(varQuelle).Cells(lZeile, 72).Text)                      ' Zustelladresse: E-Mail
    tBx26 = CStr(Worksheets(varQuelle).Cells(lZeile, 73).Text)                      ' Zustelladresse: Mobil
    tBx27 = CStr(Worksheets(varQuelle).Cells(lZeile, 74).Text)                      ' Zustelladresse: Telefon (Festnetz)
    tBx28 = CStr(Worksheets(varQuelle).Cells(lZeile, 75).Text)                      ' Zustelladresse: Telefax
    tBx29 = CStr(Worksheets(varQuelle).Cells(lZeile, 76).Text)                      ' Zustelladresse: IBAN
    tBx30 = CStr(Worksheets(varQuelle).Cells(lZeile, 77).Text)                      ' Zustelladresse: BIC
    tBx31 = CStr(Worksheets(varQuelle).Cells(lZeile, 78).Text)                      ' Zustelladresse: Vermerk weitere IBANs
    If CStr(Worksheets(varQuelle).Cells(lZeile, 71).Text) = "1" Then                ' Zustelladresse: unbekannt verzogen
        oPb3.Value = False
        oPb4.Value = True
    End If
    If CStr(Worksheets(varQuelle).Cells(lZeile, 71).Text) = "0" Then                ' Zustelladresse: unbekannt verzogen
        oPb3.Value = True
        oPb4.Value = False
    End If
    tBx33 = CStr(Worksheets(varQuelle).Cells(lZeile, 28).Text)                      ' Valuta Hauptforderung
    tBxValDatum_M = CStr(Worksheets(varQuelle).Cells(lZeile, 31).Text)              ' Valuta Mahnkosten
    tBxValDatum_A = CStr(Worksheets(varQuelle).Cells(lZeile, 34).Text)              ' Valuta Auskunftskosten
    tBxValDatum_R = CStr(Worksheets(varQuelle).Cells(lZeile, 37).Text)              ' Valuta RLS-Gebühren
    tBxValDatum_S = CStr(Worksheets(varQuelle).Cells(lZeile, 40).Text)              ' Valuta Sonst. Nebenforderungen
    tBx34 = CStr(Worksheets(varQuelle).Cells(lZeile, 87).Text)                      ' Anrede
    cBxAnrede2 = CStr(Worksheets(varQuelle).Cells(lZeile, 87).Text)
    tBx35 = CStr(Worksheets(varQuelle).Cells(lZeile, 88).Text)                      ' Name
    tBx36 = CStr(Worksheets(varQuelle).Cells(lZeile, 89).Text)                      ' Vorname
    tBx37 = CStr(Worksheets(varQuelle).Cells(lZeile, 98).Text)                      ' Adresszusatz (c/o Zeile)
    tBx121 = CStr(Worksheets(varQuelle).Cells(lZeile, 86).Text)                     ' AdressNr
    tBx38 = CStr(Worksheets(varQuelle).Cells(lZeile, 92).Text)                      ' Zustelladresse: Straße
    tBx67 = CStr(Worksheets(varQuelle).Cells(lZeile, 93).Text)                      ' Zustelladresse: Hausnummer
    tBx39 = CStr(Worksheets(varQuelle).Cells(lZeile, 97).Text)                      ' Zustelladresse: Nation
    tBx40 = CStr(Worksheets(varQuelle).Cells(lZeile, 94).Text)                      ' Zustelladresse: PLZ
    tBx41 = CStr(Worksheets(varQuelle).Cells(lZeile, 95).Text)                      ' Zustelladresse: Ort
    tBx42 = CStr(Worksheets(varQuelle).Cells(lZeile, 96).Text)                      ' Zustelladresse: Ortsteil
    tBx43 = CStr(Worksheets(varQuelle).Cells(lZeile, 100).Text)                     ' Zustelladresse: E-Mail
    tBx44 = CStr(Worksheets(varQuelle).Cells(lZeile, 101).Text)                     ' Zustelladresse: Mobil
    tBx45 = CStr(Worksheets(varQuelle).Cells(lZeile, 102).Text)                     ' Zustelladresse: Telefon (Festnetz)
    tBx46 = CStr(Worksheets(varQuelle).Cells(lZeile, 103).Text)                     ' Zustelladresse: Telefax
    tBx47 = CStr(Worksheets(varQuelle).Cells(lZeile, 104).Text)                     ' Zustelladresse: IBAN
    tBx48 = CStr(Worksheets(varQuelle).Cells(lZeile, 105).Text)                     ' Zustelladresse: BIC
    tBx49 = CStr(Worksheets(varQuelle).Cells(lZeile, 106).Text)                     ' Zustelladresse: Vermerk weitere IBANs
    tBx50 = CStr(Worksheets(varQuelle).Cells(lZeile, 115).Text)                     ' Anrede
    cBxAnrede3 = CStr(Worksheets(varQuelle).Cells(lZeile, 115).Text)
    tBx51 = CStr(Worksheets(varQuelle).Cells(lZeile, 126).Text)                     ' Name
    tBx52 = CStr(Worksheets(varQuelle).Cells(lZeile, 116).Text)                     ' Vorname
    tBx124 = CStr(Worksheets(varQuelle).Cells(lZeile, 114).Text)                    ' AdressNr
    tBx54 = CStr(Worksheets(varQuelle).Cells(lZeile, 120).Text)                     ' Zustelladresse: Straße
    tBx68 = CStr(Worksheets(varQuelle).Cells(lZeile, 121).Text)                     ' Zustelladresse: Hausnummer
    tBx55 = CStr(Worksheets(varQuelle).Cells(lZeile, 125).Text)                     ' Zustelladresse: Nation
    tBx56 = CStr(Worksheets(varQuelle).Cells(lZeile, 122).Text)                     ' Zustelladresse: PLZ
    tBx57 = CStr(Worksheets(varQuelle).Cells(lZeile, 123).Text)                     ' Zustelladresse: Ort
    tBx58 = CStr(Worksheets(varQuelle).Cells(lZeile, 124).Text)                     ' Zustelladresse: Ortsteil
    tBx59 = CStr(Worksheets(varQuelle).Cells(lZeile, 128).Text)                     ' Zustelladresse: E-Mail
    tBx60 = CStr(Worksheets(varQuelle).Cells(lZeile, 129).Text)                     ' Zustelladresse: Mobil
    tBx61 = CStr(Worksheets(varQuelle).Cells(lZeile, 130).Text)                     ' Zustelladresse: Telefon (Festnetz)
    tBx62 = CStr(Worksheets(varQuelle).Cells(lZeile, 131).Text)                     ' Zustelladresse: Telefax
    tBx63 = CStr(Worksheets(varQuelle).Cells(lZeile, 132).Text)                     ' Zustelladresse: IBAN
    tBx64 = CStr(Worksheets(varQuelle).Cells(lZeile, 133).Text)                     ' Zustelladresse: BIC
    tBx65 = CStr(Worksheets(varQuelle).Cells(lZeile, 134).Text)                     ' Zustelladresse: Vermerk weitere IBANs
    tBx96 = CStr(Worksheets(varQuelle).Cells(lZeile, 2).Text)                       ' Produkt 1
    tBx97 = CStr(Worksheets(varQuelle).Cells(lZeile, 3).Text)                       ' Produkt 2
    tBx98 = CStr(Worksheets(varQuelle).Cells(lZeile, 4).Text)                       ' Produkt 3
    tBx99 = CStr(Worksheets(varQuelle).Cells(lZeile, 5).Text)                       ' Produkt 4
    tBx100 = CStr(Worksheets(varQuelle).Cells(lZeile, 6).Text)                      ' Produkt 5
    tBx101.Text = Worksheets(varQuelle).Cells(lZeile, 7)                            ' A07_Dummy1
    tBx102.Text = Worksheets(varQuelle).Cells(lZeile, 8)                            ' A08_Dummy2
    tBx103.Text = Worksheets(varQuelle).Cells(lZeile, 9)                            ' A09_Dummy3
    tBx104.Text = Worksheets(varQuelle).Cells(lZeile, 10)                           ' A10_Dummy4
    tBx105.Text = Worksheets(varQuelle).Cells(lZeile, 11)                           ' A11_Dummy5
    tBx69.Text = Worksheets(varQuelle).Cells(lZeile, 52)                            ' F41_Dummy5
    tBx70.Text = Worksheets(varQuelle).Cells(lZeile, 53)                            ' F42_Dummy6
    tBx71.Text = Worksheets(varQuelle).Cells(lZeile, 54)                            ' F43_Dummy7
    tBx72.Text = Worksheets(varQuelle).Cells(lZeile, 55)                            ' F44_Dummy8
    tBx73.Text = Worksheets(varQuelle).Cells(lZeile, 56)                            ' F45_Dummy9
    tBx74.Text = Worksheets(varQuelle).Cells(lZeile, 57)                            ' F46_Dummy10
    tBx75.Text = Worksheets(varQuelle).Cells(lZeile, 79)                            ' M1_22_Dummy1
    tBx76.Text = Worksheets(varQuelle).Cells(lZeile, 80)                            ' M1_23_Dummy2
    tBx77.Text = Worksheets(varQuelle).Cells(lZeile, 81)                            ' M1_24_Dummy3
    tBx78.Text = Worksheets(varQuelle).Cells(lZeile, 82)                            ' M1_25_Dummy4
    tBx79.Text = Worksheets(varQuelle).Cells(lZeile, 83)                            ' M1_26_Dummy5
    tBx80.Text = Worksheets(varQuelle).Cells(lZeile, 84)                            ' M1_27_Dummy6
    tBx81.Text = Worksheets(varQuelle).Cells(lZeile, 85)                            ' M1_28_Dummy7
    tBx82.Text = Worksheets(varQuelle).Cells(lZeile, 107)                           ' M2_22_Dummy1
    tBx83.Text = Worksheets(varQuelle).Cells(lZeile, 108)                           ' M2_23_Dummy2
    tBx84.Text = Worksheets(varQuelle).Cells(lZeile, 109)                           ' M2_24_Dummy3
    tBx85.Text = Worksheets(varQuelle).Cells(lZeile, 110)                           ' M2_25_Dummy4
    tBx86.Text = Worksheets(varQuelle).Cells(lZeile, 111)                           ' M2_26_Dummy5
    tBx87.Text = Worksheets(varQuelle).Cells(lZeile, 112)                           ' M2_27_Dummy6
    tBx88.Text = Worksheets(varQuelle).Cells(lZeile, 113)                           ' M2_28_Dummy7
    tBx89.Text = Worksheets(varQuelle).Cells(lZeile, 135)                           ' M3_22_Dummy1
    tBx90.Text = Worksheets(varQuelle).Cells(lZeile, 136)                           ' M3_23_Dummy2
    tBx91.Text = Worksheets(varQuelle).Cells(lZeile, 137)                           ' M3_24_Dummy3
    tBx92.Text = Worksheets(varQuelle).Cells(lZeile, 138)                           ' M3_25_Dummy4
    tBx93.Text = Worksheets(varQuelle).Cells(lZeile, 139)                           ' M3_26_Dummy5
    tBx94.Text = Worksheets(varQuelle).Cells(lZeile, 140)                           ' M3_27_Dummy6
    tBx95.Text = Worksheets(varQuelle).Cells(lZeile, 141)                           ' M3_28_Dummy7
    tBx106.Text = Worksheets(varQuelle).Cells(lZeile, 14)                           ' Geschäftsjahr
    tBx107.Text = Worksheets(varQuelle).Cells(lZeile, 15)                           ' LfdNr
    tBx108.Text = Worksheets(varQuelle).Cells(lZeile, 17)                           ' Anspruch von
    tBx109.Text = Worksheets(varQuelle).Cells(lZeile, 18)                           ' Anspruch bis
    tBx110.Text = Worksheets(varQuelle).Cells(lZeile, 20)                           ' Vertragsdatum
    If CStr(Worksheets(varQuelle).Cells(lZeile, 21).Text) = "gewerblich" Then       ' Vertragsart = gewerblich
        oPb1.Value = True
        oPb2.Value = False
    End If
    If CStr(Worksheets(varQuelle).Cells(lZeile, 21).Text) = "privat" Then           ' Vertragsart = privat
        oPb1.Value = False
        oPb2.Value = True
    End If
    tBx111.Text = Worksheets(varQuelle).Cells(lZeile, 23)                           ' Etage
    tBx112.Text = Worksheets(varQuelle).Cells(lZeile, 24)                           ' Lage
    tBx114.Text = Worksheets(varQuelle).Cells(lZeile, 41)                           ' Leistungsadresse: Straße und Hausnummer
    tBx117.Text = Worksheets(varQuelle).Cells(lZeile, 42)                           ' Leistungsadresse: PLZ
    tBx118.Text = Worksheets(varQuelle).Cells(lZeile, 43)                           ' Leistungsadresse: Ort
    tBx116.Text = Worksheets(varQuelle).Cells(lZeile, 44)                           ' Leistungsadresse: Nation
    tBx113.Text = Worksheets(varQuelle).Cells(lZeile, 45)                           ' Leistungsadresse: Adresszusatz
    tBx119.Text = Worksheets(varQuelle).Cells(lZeile, 62)                           ' Geburtsdatum Hauptmieter
    tBx120.Text = Worksheets(varQuelle).Cells(lZeile, 63)                           ' Todesdatum
    tBx122.Text = Worksheets(varQuelle).Cells(lZeile, 90)                           ' Geburtsdatum Mieter 2
    tBx123.Text = Worksheets(varQuelle).Cells(lZeile, 91)                           ' Todesdatum
    tBx125.Text = Worksheets(varQuelle).Cells(lZeile, 118)                          ' Geburtsdatum Mieter 3
    tBx126.Text = Worksheets(varQuelle).Cells(lZeile, 119)                          ' Todesdatum
    If CStr(Worksheets(varQuelle).Cells(lZeile, 71).Text) = "1" Then                ' Zustelladresse: unbekannt verzogen
        oPb3.Value = False
        oPb4.Value = True
    End If
    If CStr(Worksheets(varQuelle).Cells(lZeile, 71).Text) = "0" Then                ' Zustelladresse: unbekannt verzogen
        oPb3.Value = True
        oPb4.Value = False
    End If
    If CStr(Worksheets(varQuelle).Cells(lZeile, 99).Text) = "1" Then                ' Zustelladresse: unbekannt verzogen
        oPb5.Value = False
        oPb6.Value = True
    End If
    If CStr(Worksheets(varQuelle).Cells(lZeile, 99).Text) = "0" Then                ' Zustelladresse: unbekannt verzogen
        oPb5.Value = True
        oPb6.Value = False
    End If
    If CStr(Worksheets(varQuelle).Cells(lZeile, 127).Text) = "1" Then                ' Zustelladresse: unbekannt verzogen
        oPb7.Value = False
        oPb8.Value = True
    End If
    If CStr(Worksheets(varQuelle).Cells(lZeile, 127).Text) = "0" Then                ' Zustelladresse: unbekannt verzogen
        oPb7.Value = True
        oPb8.Value = False
    End If
    cBxKatNr.Value = Worksheets(varQuelle).Cells(lZeile, 16)                         ' Katalognummer
    cBxAnsprMM.Value = Worksheets(varQuelle).Cells(lZeile, 19)                       ' Anspruchsgrund
       
    tBxNotes = Worksheets("NOTIZEN").Cells(lZeile, 2)                                ' Notizen
                                                                          
    If IST_MIETER2_LEER(lZeile) = False Then                                         ' Vorabprüfung, ob Felder zum 2. Mieter gefüllt sind, dann:
            mPg1.Pages(1).Visible = True                                             ' Anzeige des Tabs 2. Mieter und
            bTnMieter2.Visible = False                                               ' Button 'weiterer Mieter?' (nebst Label)
            tBxMieter2.Visible = False                                               ' auf Seite 'Hauptmieter' deaktivieren und Felder füllen
    Else:   mPg1.Pages(1).Visible = False                                            ' ansonsten ist dieser Tab unsichtbar
            bTnMieter2.Visible = True                                                ' und Button 'weiterer Mieter' nebst Label
            tBxMieter2.Visible = True                                                ' wird aktiviert
    End If
                    
    If IST_MIETER3_LEER(lZeile) = False Then                                         ' Vorabprüfung, ob Felder zum 3. Mieter gefüllt sind, dann:
            mPg1.Pages(2).Visible = True                                             ' Anzeige des Tabs 3. Mieter und
            bTnMieter3.Visible = False                                               ' Button 'weiterer Mieter?' (nebst Label)
            tBxMieter3.Visible = False                                               ' auf Seite '2. Mieter' deaktivieren und Felder füllen
    Else:   mPg1.Pages(2).Visible = False                                            ' ansonsten ist dieser Tab unsichtbar
            bTnMieter3.Visible = True                                                ' und Button 'weiterer Mieter' nebst Label
            tBxMieter3.Visible = True                                                ' wird aktiviert
    End If

    If IST_VALUTA_GLEICH(lZeile) = False Then                                        ' Vorabprüfung, ob unterschiedliche Valutadaten je Teilforderung
        Call oPb10_Click                                                             ' vorliegen.
        oPb10.Value = True                                                           ' Wenn ja, wird der entsprechende Optionsbutton auf 'J' gestellt und
    Else                                                                             ' die zus. Valutafelder werden sichtbar. Anderenfalls steht der Button
        Call oPb9_Click                                                              ' auf 'N', die anderen Felder sind disabled und das Valutadatum der
        oPb9.Value = True                                                            ' Hauptforderung wird auch für alle anderen Teilforderungen gespeichert.
    End If

End Sub

'   EINTRAG_ANLEGEN(Routine zum Speichern eines neuen Datensatzes)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EINTRAG_ANLEGEN()

    Dim lZeile As Long                                              ' erforderliche Variablen definieren
    Dim i As Integer
    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
        Me.Controls("tbx" & i).Enabled = True
        On Error Resume Next
    Next i
    
    cBxAnrede.Enabled = True
    cBxAnrede2.Enabled = True
    cBxAnrede3.Enabled = True

    For i = 1 To 8                                                  ' Optionbuttons aktivieren
        Me.Controls("oPb" & i).Enabled = True
        On Error Resume Next
    Next i
    
    AddModeOn = True                                                ' Merker Erfassungsmodus "AN" setzen

    oPb2.Value = True                                               ' Initialisierung der Optionsbox "Vertragsart" mit 'gewerblich'
    oPb3.Value = True                                               ' Initialisierung der Optionsbox "Unbekannt verzogen" mit 'nein'


    lZeile = lCONST_STARTZEILENNUMMER_DER_TABELLE                   ' Schleife bis eine leere ungebrauchte Zeile gefunden wird
    Do While IST_ZEILE_LEER(lZeile) = False
        lZeile = lZeile + 1                                         ' Nächste Zeile bearbeiten
    Loop
    
    Worksheets("DATA_UPLOAD").Cells(lZeile, 1) = _
    CStr(Worksheets("PARAM").Cells(17, 6).Text)                     ' Nach Durchlauf dieser Schleife steht lZeile in der ersten
                                                                    ' leeren Zeile von Worksheets("DATA_UPLOAD")
    ListBox1.AddItem lZeile                                         ' neuen Eintrag in die UserForm eintragen
    ListBox1.List(ListBox1.ListCount - 1, 1) = CStr("Neuer Eintrag Zeile " & lZeile)
    ListBox1.List(ListBox1.ListCount - 1, 2) = ""
    ListBox1.List(ListBox1.ListCount - 1, 3) = ""
    ListBox1.List(ListBox1.ListCount - 1, 4) = ""
    ListBox1.List(ListBox1.ListCount - 1, 5) = ""
    ListBox1.List(ListBox1.ListCount - 1, 6) = ""
    ListBox1.List(ListBox1.ListCount - 1, 7) = ""
    
    ListBox1.ListIndex = ListBox1.ListCount - 1                     ' Den neuen Eintrag markieren mit Hilfe des ListIndex,
                                                                    ' durch das Click Ereignis der ListBox werden die Daten automatisch geladen
    tBx1.SetFocus                                                   ' Cursor in das erste Eingabefeld stellen und alles vorselektieren,
    tBx1.SelStart = 0                                               ' so kann der Benutzer direkt loslegen mit der Dateneingabe.
    tBx1.SelLength = Len(tBx1)
    
End Sub

'   EINTRAG_ANLEGEN_AUS_COPY(Routine zum Speichern eines kopierten Datensatzes)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EINTRAG_ANLEGEN_AUS_COPY()
                                                                    
    Dim lZeile As Long                                              ' erforderliche Variablen definieren
    Dim i As Integer
    
    CopyModeOn = False
    AddModeOn = True                                                ' Merker Erfassungsmodus "AN" setzen
    lZeile = lCONST_STARTZEILENNUMMER_DER_TABELLE                   ' Schleife bis eine leere ungebrauchte Zeile gefunden wird
    Do While IST_ZEILE_LEER(lZeile) = False
        lZeile = lZeile + 1                                         ' Nächste Zeile bearbeiten
    Loop

    oPb2.Value = True                                               ' Initialisierung der Optionsbox "Vertragsart" mit 'gewerblich'
    oPb3.Value = True                                               ' Initialisierung der Optionsbox "Unbekannt verzogen" mit 'nein'
    Worksheets("DATA_UPLOAD").Cells(lZeile, 58) = tBx19.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 59) = tBx15.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 70) = tBx18.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 60) = tBx16.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 61) = tBx17.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 64) = tBx20.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 65) = tBx66.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 69) = tBx21.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 66) = tBx22.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 67) = tBx23.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 68) = tBx24.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 72) = tBx25.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 73) = tBx26.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 74) = tBx27.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 75) = tBx28.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 76) = tBx29.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 77) = tBx30.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 78) = tBx31.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 86) = tBx121.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 87) = tBx34.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 88) = tBx35.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 89) = tBx36.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 92) = tBx38.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 93) = tBx67.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 94) = tBx40.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 95) = tBx41.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 96) = tBx42.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 97) = tBx39.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 98) = tBx37.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 100) = tBx43.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 101) = tBx44.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 102) = tBx45.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 103) = tBx46.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 104) = tBx47.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 105) = tBx48.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 106) = tBx49.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 114) = tBx124.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 115) = tBx50.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 116) = tBx51.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 117) = tBx53.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 120) = tBx54.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 121) = tBx68.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 122) = tBx56.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 123) = tBx57.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 124) = tBx58.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 125) = tBx55.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 126) = tBx51.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 128) = tBx59.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 129) = tBx60.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 130) = tBx61.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 131) = tBx62.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 132) = tBx63.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 133) = tBx64.Text
    Worksheets("DATA_UPLOAD").Cells(lZeile, 134) = tBx65.Text
    
    ListBox1.AddItem lZeile                                         ' neuen Eintrag in die UserForm eintragen
    ListBox1.List(ListBox1.ListCount - 1, 1) = ""
    ListBox1.List(ListBox1.ListCount - 1, 2) = CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 60).Text) & ", " & _
    CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 61).Text)
    ListBox1.List(ListBox1.ListCount - 1, 3) = CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 64).Text) & " " & _
    CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 65).Text) & ", " & CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 69).Text) & " " & _
    CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 66).Text) & " " & CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 67).Text) & _
    CStr(Worksheets("DATA_UPLOAD").Cells(lZeile, 28).Text)
    ListBox1.List(ListBox1.ListCount - 1, 4) = ""
    ListBox1.List(ListBox1.ListCount - 1, 5) = ""
    ListBox1.List(ListBox1.ListCount - 1, 6) = ""
    ListBox1.List(ListBox1.ListCount - 1, 7) = ""

    bTnCopy.Enabled = True
    bTnAdd.Enabled = True
    bTnCan.Visible = True
    bTnBB.Visible = True
    lBlCopy.Visible = False

    Call UserForm_Initialize
    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
        Me.Controls("tbx" & i).Enabled = True
        On Error Resume Next
    Next i
    If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
        Me.Controls("tbx" & i).Enabled = True
        On Error Resume Next
    Next i
    
    ListBox1.ListIndex = ListBox1.ListCount - 1                     ' Den neuen Eintrag markieren mit Hilfe des ListIndex,
                                                                    ' durch das Click Ereignis der ListBox werden die Daten automatisch geladen
    tBx1.SetFocus                                                   ' Cursor in das erste Eingabefeld stellen und alles vorselektieren,
    tBx1.SelStart = 0                                               ' so kann der Benutzer direkt loslegen mit der Dateneingabe.
    tBx1.SelLength = Len(tBx1)

    
    bTnBB.BackColor = &HC0&
    bTnBck.Caption = "<"
    bTnBck.ControlTipText = "zurück ins Hauptmenü"
    bTnBck.ForeColor = &H80000012
    cBxAnrede.Enabled = True
    cBxAnrede2.Enabled = True
    cBxAnrede2.Enabled = True
    bTnValDatum.Enabled = True
    bTnMietbeginn.Enabled = True
    bTnAnVonDatum.Enabled = True
    bTnAnBisDatum.Enabled = True
    bTnGebDatum.Enabled = True
    bTnTodDatum.Enabled = True
        
    AddModeOn = True
    
End Sub

'   EINTRAG_SPEICHERN(Routine zum Speichern des aktuellen Datensatzes)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EINTRAG_SPEICHERN()
    
    If CopyModeOn = True Then
        varQuelle = "DATA_UPLOAD_ARCHIV"
    Else
        varQuelle = "DATA_UPLOAD"
    End If

    If ErrCount > 0 Then
        If MsgBox("PFLICHTEINGABEN | Verarbeitungshinweis" & vbCr & vbCr _
        & "Pflichtfelder nicht vollständig gefüllt. Abbruch?" & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect") = vbOK Then
        Call EINTRAG_LOESCHEN2
        End If
    End If

    Dim lZeile As Long                                              ' erforderliche Variablen definieren
    Dim i As Integer
    
    If ListBox1.ListIndex = -1 Then Exit Sub                        ' Wenn kein Datensatz in der ListBox markiert wurde, wird die Routine beendet
   
    lZeile = ListBox1.List(ListBox1.ListIndex, 0)                   ' Ermittlung Zeilennummer zum Speichern
    
    Worksheets(varQuelle).Cells(lZeile, 1) = CStr(Worksheets("PARAM").Cells(17, 6).Text)  ' Mandatennummer des Endkunden bei SEPA Collect
    Worksheets(varQuelle).Cells(lZeile, 48) = tBx1.Text             ' Mandant
    Worksheets(varQuelle).Cells(lZeile, 49) = tBx2.Text             ' Unternehmen
    Worksheets(varQuelle).Cells(lZeile, 22) = tBx3.Text             ' WE
    Worksheets(varQuelle).Cells(lZeile, 50) = tBx4.Text             ' HausNr
    Worksheets(varQuelle).Cells(lZeile, 25) = tBx5.Text             ' WohnNr
    Worksheets(varQuelle).Cells(lZeile, 51) = tBx6.Text             ' WohnNrZus
    Worksheets(varQuelle).Cells(lZeile, 12) = tBx7.Text             ' Folgenummer
    Worksheets(varQuelle).Cells(lZeile, 13) = tBx8.Text             ' OP-Nummer (Belegnummer)
    Worksheets(varQuelle).Cells(lZeile, 26) = tBx9.Text             ' Hauptforderung
    Worksheets(varQuelle).Cells(lZeile, 29) = tBx10.Text            ' Mahnkosten
    Worksheets(varQuelle).Cells(lZeile, 32) = tBx11.Text            ' Auskunftskosten
    Worksheets(varQuelle).Cells(lZeile, 35) = tBx12.Text            ' RLS-Gebühren
    Worksheets(varQuelle).Cells(lZeile, 38) = tBx13.Text            ' Sonstige Nebenforderungen
    Worksheets(varQuelle).Cells(lZeile, 17) = tBx108.Text           ' Anspruch von
    Worksheets(varQuelle).Cells(lZeile, 18) = tBx109.Text           ' Anspruch bis
    Worksheets(varQuelle).Cells(lZeile, 14) = tBx106.Text           ' Geschäftsjahr
    Worksheets(varQuelle).Cells(lZeile, 20) = tBx110.Text           ' Vertragsdatum
    If oPb1.Value = True Then                                       ' Vertragsart = gewerblich
    Worksheets(varQuelle).Cells(lZeile, 21) = "gewerblich"
    End If
    If oPb2.Value = True Then                                       ' Vertragsart = privat
    Worksheets(varQuelle).Cells(lZeile, 21) = "privat"
    End If
    Worksheets(varQuelle).Cells(lZeile, 23) = tBx111.Text           ' Etage
    Worksheets(varQuelle).Cells(lZeile, 24) = tBx112.Text           ' Lage
    Worksheets(varQuelle).Cells(lZeile, 28) = tBx33.Text            ' Valuta Hauptforderung
    Worksheets(varQuelle).Cells(lZeile, 31) = tBxValDatum_M.Text    ' Valuta Mahnkosten
    Worksheets(varQuelle).Cells(lZeile, 34) = tBxValDatum_A.Text    ' Valuta Auskunftskosten
    Worksheets(varQuelle).Cells(lZeile, 37) = tBxValDatum_R.Text    ' Valuta RLS-Gebühren
    Worksheets(varQuelle).Cells(lZeile, 40) = tBxValDatum_S.Text    ' Valuta Sonstige Nebenforderungen
    Worksheets(varQuelle).Cells(lZeile, 15) = tBx107.Text           ' LfdNr
    Worksheets(varQuelle).Cells(lZeile, 59) = cBxAnrede.Value       ' Anrede
    Worksheets(varQuelle).Cells(lZeile, 60) = tBx16.Text            ' Name
    Worksheets(varQuelle).Cells(lZeile, 61) = tBx17.Text            ' Vorname
    Worksheets(varQuelle).Cells(lZeile, 70) = tBx18.Text            ' Adresszusatz (c/o Zeile)
    Worksheets(varQuelle).Cells(lZeile, 58) = tBx19.Text            ' AdressNr
    Worksheets(varQuelle).Cells(lZeile, 64) = tBx20.Text            ' Zustelladresse: Straße
    Worksheets(varQuelle).Cells(lZeile, 65) = tBx66.Text            ' Zustelladresse: Hausnummer
    Worksheets(varQuelle).Cells(lZeile, 69) = tBx21.Text            ' Zustelladresse: Nation
    Worksheets(varQuelle).Cells(lZeile, 66) = tBx22.Text            ' Zustelladresse: PLZ
    Worksheets(varQuelle).Cells(lZeile, 67) = tBx23.Text            ' Zustelladresse: Ort
    Worksheets(varQuelle).Cells(lZeile, 68) = tBx24.Text            ' Zustelladresse: Ortsteil
    Worksheets(varQuelle).Cells(lZeile, 72) = tBx25.Text            ' Zustelladresse: E-Mail
    Worksheets(varQuelle).Cells(lZeile, 73) = tBx26.Text            ' Zustelladresse: Mobil
    Worksheets(varQuelle).Cells(lZeile, 74) = tBx27.Text            ' Zustelladresse: Telefon (Festnetz)
    Worksheets(varQuelle).Cells(lZeile, 75) = tBx28.Text            ' Zustelladresse: Telefax
    Worksheets(varQuelle).Cells(lZeile, 76) = tBx29.Text            ' Zustelladresse: IBAN
    Worksheets(varQuelle).Cells(lZeile, 77) = tBx30.Text            ' Zustelladresse: BIC
    Worksheets(varQuelle).Cells(lZeile, 78) = tBx31.Text            ' Zustelladresse: Vermerk weitere IBANs
    If oPb3.Value = True Then                                       ' Zustelladresse: unbekannt verzogen = Nein
    Worksheets(varQuelle).Cells(lZeile, 71) = "0"
    End If
    If oPb4.Value = True Then                                       ' Zustelladresse: unbekannt verzogen = Ja
    Worksheets(varQuelle).Cells(lZeile, 71) = "1"
    End If
    Worksheets(varQuelle).Cells(lZeile, 62) = tBx119.Text           ' Geburtsdatum
    Worksheets(varQuelle).Cells(lZeile, 63) = tBx120.Text           ' Todesdatum
    Worksheets(varQuelle).Cells(lZeile, 41) = tBx114.Text           ' Leistungsadresse: Straße und Hausnummer
    Worksheets(varQuelle).Cells(lZeile, 42) = tBx117.Text           ' Leistungsadresse: PLZ
    Worksheets(varQuelle).Cells(lZeile, 43) = tBx118.Text           ' Leistungsadresse: Ort
    Worksheets(varQuelle).Cells(lZeile, 44) = tBx116.Text           ' Leistungsadresse: Nation
    Worksheets(varQuelle).Cells(lZeile, 45) = tBx113.Text           ' Leistungsadresse: Adresszusatz
    Worksheets(varQuelle).Cells(lZeile, 87) = cBxAnrede2.Value      ' Anrede
    Worksheets(varQuelle).Cells(lZeile, 88) = tBx35.Text            ' Name
    Worksheets(varQuelle).Cells(lZeile, 89) = tBx36.Text            ' Vorname
    Worksheets(varQuelle).Cells(lZeile, 98) = tBx37.Text            ' Adresszusatz (c/o Zeile)
    Worksheets(varQuelle).Cells(lZeile, 86) = tBx121.Text           ' AdressNr
    Worksheets(varQuelle).Cells(lZeile, 92) = tBx38.Text            ' Zustelladresse: Straße
    Worksheets(varQuelle).Cells(lZeile, 93) = tBx67.Text            ' Zustelladresse: Hausnummer
    Worksheets(varQuelle).Cells(lZeile, 97) = tBx39.Text            ' Zustelladresse: Nation
    Worksheets(varQuelle).Cells(lZeile, 94) = tBx40.Text            ' Zustelladresse: PLZ
    Worksheets(varQuelle).Cells(lZeile, 95) = tBx41.Text            ' Zustelladresse: Ort
    Worksheets(varQuelle).Cells(lZeile, 96) = tBx42.Text            ' Zustelladresse: Ortsteil
    Worksheets(varQuelle).Cells(lZeile, 100) = tBx43.Text           ' Zustelladresse: E-Mail
    Worksheets(varQuelle).Cells(lZeile, 101) = tBx44.Text           ' Zustelladresse: Mobil
    Worksheets(varQuelle).Cells(lZeile, 102) = tBx45.Text           ' Zustelladresse: Telefon (Festnetz)
    Worksheets(varQuelle).Cells(lZeile, 103) = tBx46.Text           ' Zustelladresse: Telefax
    Worksheets(varQuelle).Cells(lZeile, 104) = tBx47.Text           ' Zustelladresse: IBAN
    Worksheets(varQuelle).Cells(lZeile, 105) = tBx48.Text           ' Zustelladresse: BIC
    Worksheets(varQuelle).Cells(lZeile, 106) = tBx49.Text           ' Zustelladresse: Vermerk weitere IBANs
    If oPb5.Value = True Then                                       ' Zustelladresse: unbekannt verzogen = Nein
    Worksheets(varQuelle).Cells(lZeile, 99) = "0"
    End If
    If oPb6.Value = True Then                                       ' Zustelladresse: unbekannt verzogen = Ja
    Worksheets(varQuelle).Cells(lZeile, 99) = "1"
    End If
    Worksheets(varQuelle).Cells(lZeile, 90) = tBx122.Text           ' Geburtsdatum
    Worksheets(varQuelle).Cells(lZeile, 91) = tBx123.Text           ' Todesdatum
    Worksheets(varQuelle).Cells(lZeile, 115) = cBxAnrede3.Value     ' Anrede
    Worksheets(varQuelle).Cells(lZeile, 126) = tBx51.Text           ' Name
    Worksheets(varQuelle).Cells(lZeile, 116) = tBx52.Text           ' Vorname
    Worksheets(varQuelle).Cells(lZeile, 117) = tBx53.Text           ' Adresszusatz (c/o Zeile)
    Worksheets(varQuelle).Cells(lZeile, 114) = tBx124.Text          ' AdressNr
    Worksheets(varQuelle).Cells(lZeile, 120) = tBx54.Text           ' Zustelladresse: Straße
    Worksheets(varQuelle).Cells(lZeile, 121) = tBx68.Text           ' Zustelladresse: Hausnummer
    Worksheets(varQuelle).Cells(lZeile, 125) = tBx55.Text           ' Zustelladresse: Nation
    Worksheets(varQuelle).Cells(lZeile, 122) = tBx56.Text           ' Zustelladresse: PLZ
    Worksheets(varQuelle).Cells(lZeile, 123) = tBx57.Text           ' Zustelladresse: Ort
    Worksheets(varQuelle).Cells(lZeile, 124) = tBx58.Text           ' Zustelladresse: Ortsteil
    Worksheets(varQuelle).Cells(lZeile, 128) = tBx59.Text           ' Zustelladresse: E-Mail
    Worksheets(varQuelle).Cells(lZeile, 129) = tBx60.Text           ' Zustelladresse: Mobil
    Worksheets(varQuelle).Cells(lZeile, 130) = tBx61.Text           ' Zustelladresse: Telefon (Festnetz)
    Worksheets(varQuelle).Cells(lZeile, 131) = tBx62.Text           ' Zustelladresse: Telefax
    Worksheets(varQuelle).Cells(lZeile, 132) = tBx63.Text           ' Zustelladresse: IBAN
    Worksheets(varQuelle).Cells(lZeile, 133) = tBx64.Text           ' Zustelladresse: BIC
    Worksheets(varQuelle).Cells(lZeile, 134) = tBx65.Text           ' Zustelladresse: Vermerk weitere IBANs
    If oPb7.Value = True Then                                       ' Zustelladresse: unbekannt verzogen = Nein
    Worksheets(varQuelle).Cells(lZeile, 127) = "0"
    End If
    If oPb8.Value = True Then                                       ' Zustelladresse: unbekannt verzogen = Ja
    Worksheets(varQuelle).Cells(lZeile, 127) = "1"
    End If
    Worksheets(varQuelle).Cells(lZeile, 118) = tBx125.Text          ' Geburtsdatum
    Worksheets(varQuelle).Cells(lZeile, 119) = tBx126.Text          ' Todesdatum
    Worksheets(varQuelle).Cells(lZeile, 2) = tBx96.Text             ' Produkt 1
    Worksheets(varQuelle).Cells(lZeile, 3) = tBx97.Text             ' Produkt 2
    Worksheets(varQuelle).Cells(lZeile, 4) = tBx98.Text             ' Produkt 3
    Worksheets(varQuelle).Cells(lZeile, 5) = tBx99.Text             ' Produkt 4
    Worksheets(varQuelle).Cells(lZeile, 6) = tBx100.Text            ' Produkt 5
    Worksheets(varQuelle).Cells(lZeile, 7) = tBx101.Text            ' A07_Dummy1
    Worksheets(varQuelle).Cells(lZeile, 8) = tBx102.Text            ' A08_Dummy2
    Worksheets(varQuelle).Cells(lZeile, 9) = tBx103.Text            ' A09_Dummy3
    Worksheets(varQuelle).Cells(lZeile, 10) = tBx104.Text           ' A10_Dummy4
    Worksheets(varQuelle).Cells(lZeile, 11) = tBx105.Text           ' A11_Dummy5
    Worksheets(varQuelle).Cells(lZeile, 52) = tBx69.Text            ' F41_Dummy5
    Worksheets(varQuelle).Cells(lZeile, 53) = tBx70.Text            ' F42_Dummy6
    Worksheets(varQuelle).Cells(lZeile, 54) = tBx71.Text            ' F43_Dummy7
    Worksheets(varQuelle).Cells(lZeile, 55) = tBx72.Text            ' F44_Dummy8
    Worksheets(varQuelle).Cells(lZeile, 56) = tBx73.Text            ' F45_Dummy9
    Worksheets(varQuelle).Cells(lZeile, 57) = tBx74.Text            ' F46_Dummy10
    Worksheets(varQuelle).Cells(lZeile, 79) = tBx75.Text            ' M1_22_Dummy1
    Worksheets(varQuelle).Cells(lZeile, 80) = tBx76.Text            ' M1_23_Dummy2
    Worksheets(varQuelle).Cells(lZeile, 81) = tBx77.Text            ' M1_24_Dummy3
    Worksheets(varQuelle).Cells(lZeile, 82) = tBx78.Text            ' M1_25_Dummy4
    Worksheets(varQuelle).Cells(lZeile, 83) = tBx79.Text            ' M1_26_Dummy5
    Worksheets(varQuelle).Cells(lZeile, 84) = tBx80.Text            ' M1_27_Dummy6
    Worksheets(varQuelle).Cells(lZeile, 85) = tBx81.Text            ' M1_28_Dummy7
    Worksheets(varQuelle).Cells(lZeile, 107) = tBx82.Text           ' M2_22_Dummy1
    Worksheets(varQuelle).Cells(lZeile, 108) = tBx83.Text           ' M2_23_Dummy2
    Worksheets(varQuelle).Cells(lZeile, 109) = tBx84.Text           ' M2_24_Dummy3
    Worksheets(varQuelle).Cells(lZeile, 110) = tBx85.Text           ' M2_25_Dummy4
    Worksheets(varQuelle).Cells(lZeile, 111) = tBx86.Text           ' M2_26_Dummy5
    Worksheets(varQuelle).Cells(lZeile, 112) = tBx87.Text           ' M2_27_Dummy6
    Worksheets(varQuelle).Cells(lZeile, 113) = tBx88.Text           ' M2_28_Dummy7
    Worksheets(varQuelle).Cells(lZeile, 135) = tBx89.Text           ' M3_22_Dummy1
    Worksheets(varQuelle).Cells(lZeile, 136) = tBx90.Text           ' M3_23_Dummy2
    Worksheets(varQuelle).Cells(lZeile, 137) = tBx91.Text           ' M3_24_Dummy3
    Worksheets(varQuelle).Cells(lZeile, 138) = tBx92.Text           ' M3_25_Dummy4
    Worksheets(varQuelle).Cells(lZeile, 139) = tBx93.Text           ' M3_26_Dummy5
    Worksheets(varQuelle).Cells(lZeile, 140) = tBx94.Text           ' M3_27_Dummy6
    Worksheets(varQuelle).Cells(lZeile, 141) = tBx95.Text           ' M3_28_Dummy7
    If Worksheets(varQuelle).Cells(lZeile, 51) <> "" Then
        Worksheets("NOTIZEN").Cells(lZeile, 1) = CStr(Worksheets("PARAM").Cells(17, 6).Text) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 48) & "." & Worksheets(varQuelle).Cells(lZeile, 49) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 22) & "." & Worksheets(varQuelle).Cells(lZeile, 50) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 25) & "." & Worksheets(varQuelle).Cells(lZeile, 51) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 12) & "." & Worksheets(varQuelle).Cells(lZeile, 13)
    Else
        Worksheets("NOTIZEN").Cells(lZeile, 1) = CStr(Worksheets("PARAM").Cells(17, 6).Text) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 48) & "." & Worksheets(varQuelle).Cells(lZeile, 49) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 22) & "." & Worksheets(varQuelle).Cells(lZeile, 50) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 25) & "." & Worksheets(varQuelle).Cells(lZeile, 12) & "." & _
        Worksheets(varQuelle).Cells(lZeile, 13)
    End If
    Worksheets("NOTIZEN").Cells(lZeile, 2) = tBxNotes.Text           ' manuelle Notizen zum Fall ohne Übertrag an Mahnfabrik
    If oPb10.Value = True Then                                       ' Merker setzen, wenn die Valutadaten unterschiedlich sind,
        Worksheets(varQuelle).Cells(lZeile, 31) = tBxValDatum_M.Text
        Worksheets(varQuelle).Cells(lZeile, 34) = tBxValDatum_A.Text
        Worksheets(varQuelle).Cells(lZeile, 37) = tBxValDatum_R.Text
        Worksheets(varQuelle).Cells(lZeile, 40) = tBxValDatum_S.Text
    End If
    
    If oPb9.Value = True Then                                       ' Merker setzen, wenn die Valutadaten unterschiedlich sind,
        Worksheets(varQuelle).Cells(lZeile, 31) = tBx33.Text
        Worksheets(varQuelle).Cells(lZeile, 34) = tBx33.Text
        Worksheets(varQuelle).Cells(lZeile, 37) = tBx33.Text
        Worksheets(varQuelle).Cells(lZeile, 40) = tBx33.Text
    End If
    
    Worksheets(varQuelle).Cells(lZeile, 16) = cBxKatNr.Value        ' Katalognummer
    Worksheets(varQuelle).Cells(lZeile, 19) = cBxAnsprMM.Value      ' Anspruchsgrund
                                                                                                                               
    Call LISTE_LADEN_UND_INITIALISIEREN
  
    AddModeOn = False                                               ' Merker Erfassungsmodus "AUS" setzen
    CSVempty = False                                                ' Merker, dass jetzt Module gebucht werden können
    bTnPplusAct.Enabled = False                                     ' Modulbuchung ermöglichen
    
    'If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
    
    Call BUTTON_STANDARD

End Sub

'   EINTRAG_LOESCHEN(Routine zum Löschen des aktuellen Datensatzes)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EINTRAG_LOESCHEN()

    Dim lZeile As Long                                                              ' erforderliche Variablen definieren
    
    If ListBox1.ListIndex = -1 Then Exit Sub                                        ' Wenn kein Datensatz in der ListBox markiert wurde, wird die Routine beendet
  
    If CopyModeOn = False Then                                                      ' Sicherheitsabfrage, abhängig vom Kopiermodus
                                                                                    ' wenn "AUS":
                                                                                    
        If MsgBox("LÖSCHEN | Verarbeitungshinweis" & vbCr & vbCr _
            & "Soll der markierte Datensatz wirklich gelöscht werden?" & vbCr _
            , vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect") = vbOK Then
            lZeile = ListBox1.List(ListBox1.ListIndex, 0)                           ' Ermittlung Zeilenummner zum Löschen
            Worksheets("DATA_UPLOAD").Rows(CStr(lZeile & ":" & lZeile)).Delete      ' Zeile löschen
            ListBox1.RemoveItem ListBox1.ListIndex                                  ' Eintrag aus Liste entfernen
        End If
    End If
  
    If CopyModeOn = True Then                                                       ' Sicherheitsabfrage, abhängig vom Kopiermodus
                                                                                    ' wenn "AN":
        If MsgBox("KOPIEREN | Verarbeitungshinweis" & vbCr & vbCr _
            & "Soll der Kopiermodus abgebrochen werden?" & vbCr _
            , vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect") = vbOK Then
            lZeile = ListBox1.List(ListBox1.ListIndex, 0)                           ' Ermittlung Zeilenummner zum Löschen
            Worksheets("DATA_UPLOAD").Rows(CStr(lZeile & ":" & lZeile)).Delete      ' Zeile löschen
            ListBox1.RemoveItem ListBox1.ListIndex                                  ' Eintrag aus Liste entfernen
            CopyModeOn = False                                                      ' Kopiermodus ausschalten
            lBlCopy.Visible = False                                                 ' Hinweise auf Kopiermodus unsichtbar machen
        Else                                                                    ' ansonsten nichts machen
            Exit Sub
        End If
    End If
  
    Call BUTTON_STANDARD                                                             ' Buttons auf Grundeinstellung setzen
  
End Sub

'   EINTRAG_LOESCHEN2(Routine zum Löschen nach Abbruch der Erfassung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EINTRAG_LOESCHEN2()

    Dim lZeile As Long                                                              ' erforderliche Variablen definieren
   
    If ListBox1.ListIndex = -1 Then Exit Sub                                        ' Wenn kein Datensatz in der ListBox markiert wurde, wird die Routine beendet
  
            lZeile = ListBox1.List(ListBox1.ListIndex, 0)                           ' Ermittlung Zeilenummner zum Löschen
            Worksheets("DATA_UPLOAD").Rows(CStr(lZeile & ":" & lZeile)).Delete      ' Zeile löschen
            ListBox1.RemoveItem ListBox1.ListIndex                                  ' Eintrag aus Liste entfernen
 
    Call BUTTON_STANDARD                                                             ' Buttons auf Grundeinstellung setzen
  
End Sub


















' ###################################################################################################################################################################################
' +++++++++++++ Userform ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################
                
                
'   UserForm_Initialize(Startroutine bevor die UserForm angezeigt wird)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

    Dim i As Integer

    If Val(Application.Version) >= 9 Then                           ' Sonderbereich zum Deaktivieren der Funktionen in der Titelleiste
        hwndForm = FindWindow("ThunderDFrame", Me.Caption)
    Else
        hwndForm = FindWindow("ThunderXFrame", Me.Caption)
    End If
    
    bCloseBtn = False
    SET_USERFORM_STYLE
    
    Call LISTE_LADEN_UND_INITIALISIEREN                             ' Aufruf der entsprechenden Verarbeitungsroutine
    
    Call BUTTON_STANDARD                                             ' Buttons auf Grundeinstellung setzen
    
    If SavOK = False Then
        If AddModeOn = False And ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
    End If
    
    If ListBox1.ListCount = 0 Then                                  ' Funktionen deaktivieren, wenn Tabelle DATA_UPLOAD leer ist
        For i = 1 To iCONST_ANZAHL_EINGABEFELDER                    ' und keine Datensätze angezeigt werden können
            Me.Controls("tbx" & i).Enabled = False
            On Error Resume Next
        Next i
        
        cBxAnrede.Enabled = False
        cBxAnrede2.Enabled = False
        cBxAnrede3.Enabled = False
        bTnPplus.Enabled = False
        bTnDel.Enabled = False
        bTnSav.Enabled = False
        bTnUeberAdr.Enabled = False
        bTnGebDatum.Enabled = False
        bTnTodDatum.Enabled = False
        bTnValDatum.Enabled = False
        bTnAnVonDatum.Enabled = False
        bTnAnBisDatum.Enabled = False
        bTnMietbeginn.Enabled = False
        bTnMieter2.Enabled = False
        oPb9.Value = True
        oPb10.Value = False
        cBxAnrede.BackColor = RGB(255, 255, 255)
                
        For i = 1 To 8                                              ' Optionbuttons deaktivieren
            Me.Controls("oPb" & i).Enabled = False
            On Error Resume Next
        Next i
        
        AddModeOn = False
        
    End If
    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER
        Me.Controls("tbx" & i).ForeColor = RGB(0, 0, 0)
        On Error Resume Next
    Next i
    
    SavOK = False




End Sub

'   UserForm_Activate(Ereignisroutine beim Anzeigen der UserForm)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
    
    With UFDataLog                                                  ' Anpassen der Größe der Userform auf die Größe des aktuellen Anwendung
        .Top = Application.Top                                      ' (da Excel im Vollbildmodus gestartet wird, wird dieser dann auch hier übernommen)
        .Left = Application.Left
        .Height = Application.Height
        .Width = Application.Width
    End With
    
    Call BUTTON_STANDARD                                             ' Buttons auf Grundeinstellung setzen
    
    If CSVempty = True Then
        bTnPplus.Enabled = True
    End If

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


' ###################################################################################################################################################################################
' +++++++++++++ Listbox (en) ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################

'   ListBox1_Click (ListBox Ereignisroutine)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ListBox1_Click()
    SavOK = False
    Call EINTRAG_LADEN_UND_ANZEIGEN                                 ' Aufruf der entsprechenden Verarbeitungsroutine
    
End Sub

'   ListBox1_MouseUp (ListBox Ereignisroutine)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ListBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Zeile As Long
    Dim i As Long

'    If CopyModeOn = False Or ProdModeOn = True Or AddModeOn = False Then               ???? noch zu klären ????
'        With Me.ListBox1
'            For i = 0 To .ListCount - 1
'                If .Selected(i) Then
'                    Zeile = .List(i, .ColumnCount - 1)
'                    bTnPplusAct.Enabled = True
'                    Else
'                    bTnPplusAct.Enabled = False
'                End If
'            Next
'        End With
'    End If

    If CopyModeOn = True Then
        iC = 1
        With Me.ListBox1
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    Zeile = .List(i, .ColumnCount - 1)
                    bTnBck.Enabled = False
                    Else
                    bTnBck.Enabled = True
                End If
            Next
        End With
    End If

End Sub


' ###################################################################################################################################################################################
' +++++++++++++ Textbox (en)+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################


'   tBx1_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx1.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx2_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx2.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx3_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx3.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx4_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx4.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx5_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx5.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx6_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx6.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx7_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx7.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx8_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx8.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx107_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx107_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx107.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx106_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx106_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx106.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx9_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx9.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx10_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx10_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx10.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx11_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx11_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx11.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx12_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx12_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx12.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx13_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx13_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx13.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx16_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx16_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx16.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx17_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx17_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx17.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx18_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx18_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx18.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx20_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx20_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx20.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx66_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx66_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx66.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx21_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx21_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With tBx21
        .Value = UCase(.Value)
        .ForeColor = RGB(255, 0, 0)
    End With
    SavOK = True
End Sub

'   tBx22_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx22_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx22.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx23_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx23_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx23.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx24_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx24_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx24.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx25_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx25_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx25.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx26_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx26_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx26.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx27_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx27_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx27.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx28_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx28_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx28.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx29_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx29_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx29.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx30_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx30_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx30.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx31_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx31_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx31.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx113_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx113_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx113.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx114_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx114_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx114.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx116_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx116_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With tBx116
        .Value = UCase(.Value)
        .ForeColor = RGB(255, 0, 0)
    End With
    SavOK = True
End Sub

'   tBx117_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx117_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx117.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx118_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx118_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx118.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx119_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx19_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx19.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx111_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx111_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx111.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx112_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx112_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx112.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx35_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx35_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx35.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx36_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx36_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx36.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx37_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx37_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx37.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx38_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx38_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx38.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx67_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx67_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx67.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx39_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx39_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With tBx39
        .Value = UCase(.Value)
        .ForeColor = RGB(255, 0, 0)
    End With
    SavOK = True
End Sub

'   tBx40_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx40_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx40.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx41_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx41_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx41.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx42_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx42_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx42.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx43_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx43_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx43.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx44_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx44_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx44.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx45_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx45_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx45.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx46_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx46_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx46.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx47_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx47_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx47.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx48_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx48_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx48.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx49_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx49_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx49.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx121_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx121_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx121.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx51_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx51_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx51.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx52_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx52_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx52.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx53_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx53_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx53.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx54_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx54_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx54.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx68_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx68_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx68.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx55_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx55_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With tBx55
        .Value = UCase(.Value)
        .ForeColor = RGB(255, 0, 0)
    End With
    SavOK = True
End Sub

'   tBx56_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx56_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx56.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx57_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx57_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx57.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx58_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx58_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx58.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx59_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx59_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx59.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx60_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx60_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx60.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx61_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx61_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx61.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx62_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx62_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx62.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx63_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx63_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx63.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx64_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx64_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx64.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tBx65_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx65_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx65.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub

'   tB124_MouseUp(Forecolor = Rot bei Änderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tB124_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tBx24.ForeColor = RGB(255, 0, 0)
    SavOK = True
End Sub



' ###################################################################################################################################################################################
' +++++++++++++ Button (s) ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################


'   bTnAdd_Click(Neuen Datensatz erfassen ( + ))
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnAdd_Click()
                                                                    
    CopyModeOn = False                                              ' Kopiermodus = AUS
    AddModeOn = True                                                ' Erfassen-Modus = EIN
    
    bTnCopy.Visible = True                                          ' Buttonsteuerung
    bTnCopy.Enabled = False
    bTnAdd.Visible = True
    bTnAdd.Enabled = False
    bTnDel.Visible = True
    bTnDel.Enabled = False
    bTnSav.Visible = True
    bTnSav.Enabled = False
    bTnPplus.Enabled = False
    lBlModule.Visible = True
    bTnCan.Visible = True
    bTnCan.Enabled = True
    bTnBck.Visible = False
    bTnBB.Visible = False
    bTnMieter2.Enabled = True
    
    bTnUeberAdr.Enabled = True
    bTnGebDatum.Enabled = True
    bTnTodDatum.Enabled = True
    bTnValDatum.Enabled = True
    bTnAnVonDatum.Enabled = True
    bTnAnBisDatum.Enabled = True
    bTnMietbeginn.Enabled = True
    bTnMieter2.Enabled = True
         
    Call EINTRAG_ANLEGEN                                            ' Aufruf der entsprechenden Verarbeitungsroutine
    
End Sub

'   bTnCopy_Click(Vorhandenen Datensatz kopieren ( ++ ))
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnCopy_Click()
                                                                    
    Dim i As Integer                                                ' erforderliche Variablen definieren
    
    CopyModeOn = True                                               ' Merker für "Kopiermodus aktiv" setzen
    lBlCopy.Visible = True                                          ' Hinweis auf Kopiermodus aktivieren
    Worksheets("ERROR").Range("A1").Value = "1"
    
    iC = 0                                                          ' Merker, dass im Kopiermodus noch keine Auswahl getroffen wurde
            
    Call UserForm_Initialize
    
    ListBox1.MultiSelect = fmMultiSelectSingle                      ' Einzelauswahl aktivieren
    ListBox1.ListStyle = fmListStyleOption

    bTnCopy.Enabled = False
    bTnAdd.Enabled = False
    bTnDel.Enabled = False
    bTnSav.Enabled = False
    bTnCan.Visible = True
    bTnCan.Enabled = True
    bTnPplusAct.Visible = True
    bTnPplusAct.Enabled = False
    cBxAnrede.Enabled = False
    cBxAnrede2.Enabled = False
    cBxAnrede3.Enabled = False
    
    
    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                            ' Hintergrundfarbe anpassen
        Me.Controls("tbx" & i).Enabled = False
        On Error Resume Next
    Next i
    
    bTnBB.BackColor = &H8000&
    bTnBck.Caption = ">"
    bTnBck.ControlTipText = "Übernahme des ausgewählten Datensatzes"
    bTnBck.ForeColor = &H808000
    bTnBck.Enabled = False
        
End Sub

'   bTnDel_Click(vorhandenen Datensatz löschen ( - ))
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnDel_Click()

    If Worksheets("ERROR").Range("B1").Value = "1" Then
        Worksheets("ERROR").Range("A1").Value = "0"
        Worksheets("ERROR").Range("B1").Value = "0"
    End If
                                                                    
    bTnCopy.Visible = True                                              ' Buttonsteuerung
    bTnAdd.Visible = True
    bTnDel.Visible = True
    bTnSav.Visible = True
    bTnBck.Visible = True
    bTnBB.Visible = True
    
    Call EINTRAG_LOESCHEN                                               ' Aufruf der entsprechenden Verarbeitungsroutine
    
    Call UserForm_Initialize
    
End Sub

'   bTnSav_Click(aktuellen Datensatz speichern ( ok ))
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnSav_Click()

    ' Speichern-Button
    ' ----------------
    ' Wenn geänderte Feldinhalte vorliegen (die Schriftfarbe dieser Textboxen wird beim Klick in dieselben auf ROT gesetzt)
    ' wird vorab nachgefragt, ob die Änderungen auch tatsächlich gespeichert werden sollen.
    ' Liegen keine Veränderungen bewirkt der Button lediglich ein (nochmaliges) Speichern der angezeigten Werte.
    ' Ist der Fehlerwert (errCount) > 0, wird aufgrund der fehlerhaften Inhalte der Speichern-Button deaktiviert um zu
    ' verhindern, dass falsche Daten gespeichert werden können.

    Dim i As Integer

    If SavOK = True Then
        If ERR_COUNT_ERMITTELN = True Then
            MsgBox "SPEICHERN | Verarbeitungshinweis" & vbCr & vbCr _
            & "Vervollständigen Sie bitte Ihre Eingaben." & vbCr _
            & "Nicht alle erforderlichen Felder sind gefüllt." & vbCr & vbCr, vbInformation, "MAHNFABRIK.DE  powered by SEPA Collect"
            Worksheets("ERROR").Range("A1").Value = "1"
            Exit Sub
        Else
            If MsgBox("SPEICHERN | Verarbeitungshinweis" & vbCr & vbCr _
                & "Sie haben Veränderungen an den Daten vorgenommen." & vbCr _
                & "Sollen diese gespeichert werden?", vbQuestion + vbOKCancel, "MAHNFABRIK.DE  powered by SEPA Collect") = vbOK Then
                SavOK = False
                Call EINTRAG_SPEICHERN                                  ' Aufruf der entsprechenden Verarbeitungsroutine
                Call UserForm_Initialize
                For i = 1 To iCONST_ANZAHL_EINGABEFELDER                ' Hintergrundfarbe anpassen
                    Me.Controls("tbx" & i).ForeColor = RGB(0, 0, 0)
                    On Error Resume Next
                Next i
                tBxValDatum_M.ForeColor = RGB(0, 0, 0)
                tBxValDatum_A.ForeColor = RGB(0, 0, 0)
                tBxValDatum_R.ForeColor = RGB(0, 0, 0)
                tBxValDatum_S.ForeColor = RGB(0, 0, 0)
                cBxAnrede.ForeColor = RGB(0, 0, 0)
                Worksheets("ERROR").Range("A1").Value = "0"
            Else
                Exit Sub
            End If
        End If
    Else
        If ERR_COUNT_ERMITTELN = True Then
            MsgBox "SPEICHERN | Verarbeitungshinweis" & vbCr & vbCr _
            & "Vervollständigen Sie bitte Ihre Eingaben." & vbCr _
            & "Nicht alle erforderlichen Felder sind gefüllt." & vbCr & vbCr, vbInformation, "MAHNFABRIK.DE  powered by SEPA Collect"
            Worksheets("ERROR").Range("A1").Value = "1"
            Exit Sub
        Else
            Call EINTRAG_SPEICHERN                                      ' Aufruf der entsprechenden Verarbeitungsroutine
            Call UserForm_Initialize
            For i = 1 To iCONST_ANZAHL_EINGABEFELDER                    ' Hintergrundfarbe anpassen
                Me.Controls("tbx" & i).ForeColor = RGB(0, 0, 0)
                On Error Resume Next
            Next i
            tBxValDatum_M.ForeColor = RGB(0, 0, 0)
            tBxValDatum_A.ForeColor = RGB(0, 0, 0)
            tBxValDatum_R.ForeColor = RGB(0, 0, 0)
            tBxValDatum_S.ForeColor = RGB(0, 0, 0)
            cBxAnrede.ForeColor = RGB(0, 0, 0)
            Worksheets("ERROR").Range("A1").Value = "0"
        End If
    End If
    
End Sub

'   bTnBck_Click(zurück ( < )  auf die Hauptseite)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnBck_Click()
    
    If CopyModeOn = False Then
        Unload Me                                                       ' Datenerfassung schließen
    Else
        Call EINTRAG_ANLEGEN_AUS_COPY
    End If

End Sub

'   bTnMieter2_Click('Weiterer Mieter?' auf dem Tab 'Hauptmieter')
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnMieter2_Click()
                                                                    ' Tab '2. Mieter' aktivieren und neuen Tab anzeigen
    weitMiet1 = True                                                ' Merker erster weiterer Mieter
    mPg1.Pages(1).Visible = True                                    ' Anzeige des Tabs 2. Mieter
    mPg1.Value = 1                                                  ' Tab 2. Mieter in den Vordergrund holen
     
End Sub

'   bTnMieter3_Click('Weiterer Mieter?' auf dem Tab '2. Mieter')
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnMieter3_Click()
                                                                    ' Tab '3. Mieter' aktivieren und neuen Tab anzeigen
    weitMiet2 = True                                                ' Merker zweiter weiterer Mieter
    mPg1.Pages(2).Visible = True                                    ' Anzeige des Tabs 3. Mieter
    mPg1.Value = 2                                                  ' Tab 3. Mieter in den Vordergrund holen
    
End Sub

'   bTnPplus_Click(Produktauswahl aktivieren)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnPplus_Click()
    
    Dim i As Integer                                                ' Variablen definieren
    
    ProdModeOn = True                                               ' Buchungsmodus Module = AN
    ListBox1.MultiSelect = fmMultiSelectMulti                       ' Mehrfachauswahl aktivieren
    ListBox1.ListStyle = fmListStyleOption
    mPg1.Pages(3).Visible = False                                   ' Anzeige des Tabs Notizen ausblenden
    lBlProdAusw.Visible = True                                      ' Hinweis einblenden

    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
        Me.Controls("tbx" & i).BackColor = &H80000004
        On Error Resume Next
        tBxValDatum_M.BackColor = &H80000004
        tBxValDatum_A.BackColor = &H80000004
        tBxValDatum_R.BackColor = &H80000004
        tBxValDatum_S.BackColor = &H80000004
    Next i

    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Schriftfarbe anpassen
        Me.Controls("tbx" & i).ForeColor = &H80000004
        On Error Resume Next
        tBxValDatum_M.ForeColor = &H80000004
        tBxValDatum_A.ForeColor = &H80000004
        tBxValDatum_R.ForeColor = &H80000004
        tBxValDatum_S.ForeColor = &H80000004
    Next i

    bTnCopy.Visible = True                                          ' Buttonsteuerung
    bTnCopy.Enabled = False
    bTnAdd.Visible = True
    bTnAdd.Enabled = False
    bTnDel.Visible = True
    bTnDel.Enabled = False
    bTnSav.Visible = True
    bTnSav.Enabled = False
    bTnBck.Visible = True
    bTnBck.Enabled = True
    bTnBB.Visible = True
    bTnBB.Enabled = True
    lBlModule.Visible = True
    bTnCan.Visible = True
    bTnCan.Enabled = True
    bTnPplusAct.Enabled = False
    cBxAnrede.Enabled = False

End Sub

'   bTnPplusAct_Click(Produkte auswählen)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnPplusAct_Click()
    
    Dim lZeile As Long                                              ' erforderliche Variablen definieren
    Dim i As Integer
    
    lBlProdAusw.Visible = False
    
    UFProd.Show                                                     ' Prouktauswahl (neue Userform) öffnen
                                                                    ' Variablen für die gebuchten Modul aus UFProd übernehmen
    
    With Me.ListBox1                                                ' Werte entsprechend der Auswahl den öffentlichen Variablen zuweisen
    
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                lZeile = i + lCONST_STARTZEILENNUMMER_DER_TABELLE   ' Startzeilennummer festlegen
                Worksheets("DATA_UPLOAD").Cells(lZeile, 2) = i1
                Worksheets("DATA_UPLOAD").Cells(lZeile, 3) = i2
                Worksheets("DATA_UPLOAD").Cells(lZeile, 4) = i3
                Worksheets("DATA_UPLOAD").Cells(lZeile, 5) = i4
                Worksheets("DATA_UPLOAD").Cells(lZeile, 6) = i5
            End If
        Next i
    
    End With
    
    Call bTnPminus                                                  ' Aufruf der entsprechenden Verarbeitungsroutine
    
End Sub

'   bTnCan_Click(Abbruch | hier: Standardfall)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnCan_Click()
    
    Dim i As Integer
    
    If ProdModeOn = True Then
    bTnCan.Visible = False
    bTnPplusAct.Enabled = False
    cBxAnrede.Enabled = True
    ListBox1.MultiSelect = fmMultiSelectSingle                              ' Mehrfachauswahl deaktivieren
    ListBox1.ListStyle = fmListStylePlain
    
    Call bTnCanProd
    Exit Sub
    End If
    
    If CopyModeOn = True Then                                               ' Sicherheitsabfrage beim Abbruch während des Kopieren eines vorhandenen Datensatzes
                                                                            ' ohne dass zu diesem Zeitpunkt bereits eine Auswahl vorgenommen wurde
        If MsgBox("KOPIEREN | Verarbeitungshinweis" & vbCr & vbCr _
            & "Soll der Kopiervorgang abgebrochen werden?", _
            vbQuestion + vbYesNo, "MAHNFABRIK.DE  powered by SEPA Collect") = vbYes Then
            bTnCopy.Enabled = True
            bTnAdd.Enabled = True
            bTnBck.Visible = True
            bTnBB.Visible = True
            bTnCan.Visible = False
            lBlCopy.Visible = False
            CopyModeOn = False
            Call UserForm_Initialize
            If iC = 1 Then
                Call EINTRAG_LOESCHEN2
            End If
            For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
            Me.Controls("tbx" & i).Enabled = True
            On Error Resume Next
            Next i
            If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
            
            For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
                Me.Controls("tbx" & i).Enabled = True
            On Error Resume Next
            Next i
            
            bTnBB.BackColor = &HC0&
            bTnBck.Caption = "<"
            bTnBck.ControlTipText = "zurück ins Hauptmenü"
            bTnBck.ForeColor = &H80000012
            cBxAnrede.Enabled = True
            cBxAnrede2.Enabled = True
            cBxAnrede3.Enabled = True
        Else
            Exit Sub
        End If
        Else
            If MsgBox("ERFASSEN | Verarbeitungshinweis" & vbCr & vbCr _
                & "Soll die Erfassung abgebrochen werden?", _
                vbQuestion + vbYesNo, "MAHNFABRIK.DE  powered by SEPA Collect") = vbYes Then
                    bTnCopy.Enabled = True
                    bTnAdd.Enabled = True
                    bTnBck.Visible = True
                    bTnBB.Visible = True
                    lBlCopy.Visible = False
                    bTnCan.Visible = False
                    Call EINTRAG_LOESCHEN2
                    Call UserForm_Initialize
                    ListBox1.SetFocus
            Else
                bTnCan.Visible = True
                Exit Sub
            End If
            
            If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
            
        
    End If
    
    Call bTnPminus                                                  ' Aufruf der entsprechenden Verarbeitungsroutine
    
    CopyModeOn = False
    
    lBlUeber.Enabled = False                                        ' Buttonsteuerung
    lBlUeber.Visible = False
    cBnUeber.Enabled = False
    cBnUeber.Visible = False
    
    Call LISTE_LADEN_UND_INITIALISIEREN                             ' Aufruf der entsprechenden Verarbeitungsroutine
    
    bTnPplusAct.Enabled = False
    
    ErrCount = 0                                                    ' Bei Abbruch Zähler für Pflichtfelder zurücksetzen
    
End Sub

'   bTnCanProd(Abbruch | hier: Sonderfall Modulbuchung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnCanProd()
           
    Dim i As Integer

    If ProdModeOn = True Then                                               ' Sicherheitsabfrage beim Abbruch während der Modulbuchung
        If MsgBox("MODULE | Verarbeitungshinweis" & vbCr & vbCr _
            & "Soll die Modulauswahl abgebrochen werden?", _
            vbQuestion + vbYesNo, "MAHNFABRIK.DE  powered by SEPA Collect") = vbYes Then
            Call BUTTON_STANDARD
            lBlProdAusw.Visible = False
            ProdModeOn = False
            If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
                
            For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
                Me.Controls("tbx" & i).BackColor = &HFFFFFF
                On Error Resume Next
            Next i
            tBxValDatum_M.BackColor = &HFFFFFF
            tBxValDatum_A.BackColor = &HFFFFFF
            tBxValDatum_R.BackColor = &HFFFFFF
            tBxValDatum_S.BackColor = &HFFFFFF
            
            For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Schriftfarbe anpassen
                Me.Controls("tbx" & i).ForeColor = &H0&
                On Error Resume Next
            Next i
            tBxValDatum_M.ForeColor = &H0&
            tBxValDatum_A.ForeColor = &H0&
            tBxValDatum_R.ForeColor = &H0&
            tBxValDatum_S.ForeColor = &H0&
            Else
            Call bTnPplus_Click
        End If
   
    End If

    
End Sub

'   cBnUeberAdr_Click(Abbruch | hier: Sonderfall Modulbuchung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnUeberAdr_Click()

    If MsgBox("ADRESS-ÜBERNAHME | Verarbeitungshinweis" & vbCr & vbCr _
        & "Soll die Zustelladresse auch als Leistungsadresse verwendet werden?", _
        vbQuestion + vbYesNo, "MAHNFABRIK.DE  powered by SEPA Collect") = vbYes Then
        tBx113.Text = tBx18.Text
        tBx114.Text = tBx20.Text & " " & tBx66.Text
        tBx116.Text = tBx21.Text
        tBx117.Text = tBx22.Text
        tBx118.Text = tBx23.Text
    End If

End Sub

'   bTnLOG_Click(LOG-Einträge anzeigen)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnLOG_Click()

    UFDataLog.Show
    
End Sub

'   bTnGebDatum(Kalendersteuerelement Geburtsdatum Hauptmieter)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnGebDatum_Click()

    Dim i As Integer

    frmCalendar.Show
    i119 = 1
    With tBx119
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    tBxGeb.Text = g_datCalendarDate
    tBxMiet.Text = tBx110.Text
    tBxTod.Text = tBx120.Text
    
    If tBxMiet.Text = "" Then
        tBxDiff = 0
    Else
        tBxMiet.Text = tBx110.Text
        tBxGeb.Text = tBx119.Text
        tBxDiff = (tBxMiet - tBxGeb)
    End If
    
    If tBxTod.Text = "" Then
        tBxDiff = 0
    Else
        tBxTod.Text = tBx120.Text
        tBxGeb.Text = tBx119.Text
        tBxDiff = (tBxMiet - tBxGeb)
    End If
    
    
    If tBxDiff < 0 Then
        MsgBox "DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
        & "Das Datum des Mietbeginns kann nicht VOR dem Geburtsdatum liegen." & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect"
        SavOK = False
        tBx108.ForeColor = RGB(255, 0, 0)
        tBx109.ForeColor = RGB(255, 0, 0)
        bTnSav.Enabled = False
            If tBxDiffGT < 0 Then
                MsgBox "DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
                & "Das Todesdatum kann nicht VOR dem Geburtsdatum liegen." & vbCr _
                & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect"
                SavOK = False
                tBx108.ForeColor = RGB(255, 0, 0)
                tBx109.ForeColor = RGB(255, 0, 0)
                bTnSav.Enabled = False
            End If
    Else
        If tBxDiffGT < 0 Then
                MsgBox "DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
                & "Das Todesdatum kann nicht VOR dem Geburtsdatum liegen." & vbCr _
                & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect"
            SavOK = False
            tBx108.ForeColor = RGB(255, 0, 0)
            tBx109.ForeColor = RGB(255, 0, 0)
            bTnSav.Enabled = False
        End If
        SavOK = True
    End If
    
    If i108 = 1 Then
        tBx108.ForeColor = RGB(255, 0, 0)
    Else
        If i109 = 1 Then
            tBx109.ForeColor = RGB(255, 0, 0)
        Else
            If i33 = 1 Then
                tBx33.ForeColor = RGB(255, 0, 0)
            Else
                If i120 = 1 Then
                    tBx120.ForeColor = RGB(255, 0, 0)
                Else
                    If i110 = 1 Then
                        tBx110.ForeColor = RGB(255, 0, 0)
                    Else
                        If i122 = 1 Then
                            tBx122.ForeColor = RGB(255, 0, 0)
                        Else
                            If i123 = 1 Then
                                tBx123.ForeColor = RGB(255, 0, 0)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    tBx119.ForeColor = RGB(255, 0, 0)

End Sub

'   bTnTodDatum(Kalendersteuerelement Todesdatum Hauptmieter)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnTodDatum_Click()

    Dim i As Integer
    
    frmCalendar.Show
    i120 = 1
    With tBx120
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    tBxTod.Text = g_datCalendarDate
    tBxGeb.Text = tBx119.Text
    
    If tBxGeb.Text = "" Then
        tBxDiff = 0
    Else
        tBxTod.Text = tBx120.Text
        tBxGeb.Text = tBx119.Text
        tBxDiffGT = (tBxTod - tBxGeb)
    End If
    
    If tBxDiffGT < 0 Then
        MsgBox "DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
        & "Das Todesdatum kann nicht VOR dem Geburtsdatum liegen." & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect"
        SavOK = False
        tBx108.ForeColor = RGB(255, 0, 0)
        tBx109.ForeColor = RGB(255, 0, 0)
        bTnSav.Enabled = False
    Else
        SavOK = True
    End If
    
    If i108 = 1 Then
        tBx108.ForeColor = RGB(255, 0, 0)
    Else
        If i109 = 1 Then
            tBx109.ForeColor = RGB(255, 0, 0)
        Else
            If i119 = 1 Then
                tBx119.ForeColor = RGB(255, 0, 0)
            Else
                If i33 = 1 Then
                    tBx33.ForeColor = RGB(255, 0, 0)
                Else
                    If i110 = 1 Then
                        tBx110.ForeColor = RGB(255, 0, 0)
                    Else
                        If i122 = 1 Then
                            tBx122.ForeColor = RGB(255, 0, 0)
                        Else
                            If i123 = 1 Then
                                tBx123.ForeColor = RGB(255, 0, 0)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    tBx120.ForeColor = RGB(255, 0, 0)

End Sub

'   bTnMietbeginn(Kalendersteuerelement Mietbeginn Hauptmieter)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnMietbeginn_Click()

    Dim i As Integer

    frmCalendar.Show
    i110 = 1
    With tBx110
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    tBxMiet.Text = g_datCalendarDate
    tBxGeb.Text = tBx119.Text
    
    If tBxGeb.Text = "" Then
        tBxDiff = 0
    Else
        tBxMiet.Text = tBx110.Text
        tBxGeb.Text = tBx119.Text
        tBxDiff = (tBxMiet - tBxGeb)
    End If
    
    If tBxDiff < 0 Then
        MsgBox "DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
        & "Das Datum des Mietbeginns kann nicht VOR dem Geburtsdatum liegen." & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect"
        SavOK = False
        tBx110.Text = ""
        tBx108.ForeColor = RGB(255, 0, 0)
        tBx109.ForeColor = RGB(255, 0, 0)
        bTnSav.Enabled = False
    Else
        SavOK = True
    End If
    
    If i108 = 1 Then
        tBx108.ForeColor = RGB(255, 0, 0)
    Else
        If i109 = 1 Then
            tBx109.ForeColor = RGB(255, 0, 0)
        Else
            If i119 = 1 Then
                tBx119.ForeColor = RGB(255, 0, 0)
            Else
                If i120 = 1 Then
                    tBx120.ForeColor = RGB(255, 0, 0)
                Else
                    If i33 = 1 Then
                        tBx33.ForeColor = RGB(255, 0, 0)
                    Else
                        If i122 = 1 Then
                            tBx122.ForeColor = RGB(255, 0, 0)
                        Else
                            If i123 = 1 Then
                                tBx123.ForeColor = RGB(255, 0, 0)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    tBx110.ForeColor = RGB(255, 0, 0)

End Sub

'   bTnValDatum(Kalendersteuerelement Valuta Hauptforderung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnValDatum_Click()

    frmCalendar.Show
    i33 = 1
    
    With tBx33
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    
    SavOK = True
    
    If i108 = 1 Then
        tBx108.ForeColor = RGB(255, 0, 0)
    Else
        If i109 = 1 Then
            tBx109.ForeColor = RGB(255, 0, 0)
        Else
            If i119 = 1 Then
                tBx119.ForeColor = RGB(255, 0, 0)
            Else
                If i120 = 1 Then
                    tBx120.ForeColor = RGB(255, 0, 0)
                Else
                    If i110 = 1 Then
                        tBx110.ForeColor = RGB(255, 0, 0)
                    Else
                        If i122 = 1 Then
                            tBx122.ForeColor = RGB(255, 0, 0)
                        Else
                            If i123 = 1 Then
                                tBx123.ForeColor = RGB(255, 0, 0)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
   
End Sub


'   bTnValDatum_M(Kalendersteuerelement Valuta Mahnkosten)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnValDatum_M_Click()

    frmCalendar.Show
    i33_M = 1
    
    With tBxValDatum_M
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    
    SavOK = True
   
End Sub

'   bTnValDatum_A(Kalendersteuerelement Valuta Auskunftskosten)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnValDatum_A_Click()

    frmCalendar.Show
    i33_A = 1
    
    With tBxValDatum_A
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    
    SavOK = True
   
End Sub


'   bTnValDatum_R(Kalendersteuerelement Valuta RLS-Gebühren)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnValDatum_R_Click()

    frmCalendar.Show
    i33_R = 1
    
    With tBxValDatum_R
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    
    SavOK = True
   
End Sub

'   bTnValDatum_S(Kalendersteuerelement Valuta Mahnkosten)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnValDatum_S_Click()

    frmCalendar.Show
    i33_S = 1
    
    With tBxValDatum_S
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    
    SavOK = True
   
End Sub

'   bTnAnVonDatum(Kalendersteuerelement Anspruch von - Datum Hauptmieter)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnAnVonDatum_Click()

    Dim i As Integer
    
    frmCalendar.Show
    i108 = 1
    With tBx108
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    tBxVon.Text = g_datCalendarDate
    
    If tBx109.Text = "" Then
        tBxDiff = 0
    Else
        tBxBis.Text = tBx109.Text
        tBxVon.Text = tBx108.Text
        tBxDiff = (tBxBis - tBxVon)
    End If
    
    If tBxDiff < 0 Then
        MsgBox "DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
        & "Das Anspruchsdatum 'bis' liegt VOR dem Anspruchsdatum 'von'." & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect"
        SavOK = False
        tBx108.ForeColor = RGB(255, 0, 0)
        tBx109.ForeColor = RGB(255, 0, 0)
        bTnSav.Enabled = False
        i108 = 1
    Else
        SavOK = True
    End If

    If i33 = 1 Then
        tBx33.ForeColor = RGB(255, 0, 0)
    Else
        If i109 = 1 Then
            tBx109.ForeColor = RGB(255, 0, 0)
        Else
            If i119 = 1 Then
                tBx119.ForeColor = RGB(255, 0, 0)
            Else
                If i120 = 1 Then
                    tBx120.ForeColor = RGB(255, 0, 0)
                Else
                    If i110 = 1 Then
                        tBx110.ForeColor = RGB(255, 0, 0)
                    Else
                        If i122 = 1 Then
                            tBx122.ForeColor = RGB(255, 0, 0)
                        Else
                            If i123 = 1 Then
                                tBx123.ForeColor = RGB(255, 0, 0)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    tBx108.ForeColor = RGB(255, 0, 0)
    tBx109.ForeColor = RGB(255, 0, 0)

    
End Sub

'   bTnAnBisDatum(Kalendersteuerelement Anspruch bis - Datum Hauptmieter)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnAnBisDatum_Click()

    Dim i As Integer
    
    frmCalendar.Show
    i109 = 1
    With tBx109
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    tBxVon.Text = g_datCalendarDate
    
    If tBx108.Text = "" Then
        tBxDiff = 0
    Else
        tBxBis.Text = tBx109.Text
        tBxVon.Text = tBx108.Text
        tBxDiff = (tBxBis - tBxVon)
    End If

    If tBxDiff < 0 Then
        MsgBox "DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
        & "Das Anspruchsdatum 'bis' liegt VOR dem Anspruchsdatum 'von'." & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect"
        SavOK = False
        tBx108.ForeColor = RGB(255, 0, 0)
        tBx109.ForeColor = RGB(255, 0, 0)
        bTnSav.Enabled = False
    Else
        SavOK = True
    End If
    
        If i33 = 1 Then
        tBx33.ForeColor = RGB(255, 0, 0)
    Else
        If i109 = 1 Then
            tBx109.ForeColor = RGB(255, 0, 0)
        Else
            If i119 = 1 Then
                tBx119.ForeColor = RGB(255, 0, 0)
            Else
                If i120 = 1 Then
                    tBx120.ForeColor = RGB(255, 0, 0)
                Else
                    If i110 = 1 Then
                        tBx110.ForeColor = RGB(255, 0, 0)
                    Else
                        If i122 = 1 Then
                            tBx122.ForeColor = RGB(255, 0, 0)
                        Else
                            If i123 = 1 Then
                                tBx123.ForeColor = RGB(255, 0, 0)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    tBx109.ForeColor = RGB(255, 0, 0)
    tBx108.ForeColor = RGB(255, 0, 0)

End Sub

'   bTnGebDatum2(Kalendersteuerelement Geburtsdatum 2. Mieter)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnGebDatum2_Click()

    Dim i As Integer
    
    frmCalendar.Show
    i122 = 1
    With tBx122
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    tBxGeb2.Text = g_datCalendarDate
    
    If tBxTod2.Text = "" Then
        tBxTod2.Text = tBxGeb2.Value + 1
    End If
    
    tBxDiff2 = (tBxTod2 - tBxGeb2)
    
    If tBxDiff2 < 0 Then
        If MsgBox("DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
        & "Das Anspruchsdatum 'bis' liegt VOR dem Anspruchsdatum 'von'." & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect") = vbOK Then
            SavOK = False
            tBx122.ForeColor = RGB(255, 0, 0)
            tBx123.ForeColor = RGB(255, 0, 0)
        Else
            Exit Sub
        End If

    End If
    
    mPg1.Pages(1).Visible = True                                    ' Anzeige des Tabs 2. Mieter
    mPg1.Value = 1                                                  ' Tab 2. Mieter in den Vordergrund holen
    SavOK = True
    
End Sub

'   bTnTodDatum2(Kalendersteuerelement Todesdatum 2. Mieter)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnTodDatum2_Click()

    Dim i As Integer
    i123 = 1
    frmCalendar.Show
    
    With tBx123
        .Value = g_datCalendarDate
        .ForeColor = RGB(255, 0, 0)
    End With
    tBxGeb2.Text = g_datCalendarDate
    tBxTod2.Text = tBx122.Text
    tBxDiff2 = (tBxGeb2 - tBxTod2)
    
    If tBxDiff2 < 0 Then
        If MsgBox("DATUMSFEHLER | Verarbeitungshinweis" & vbCr & vbCr _
        & "Das Anspruchsdatum 'bis' liegt VOR dem Anspruchsdatum 'von'." & vbCr _
        & "Bitte ändern", vbExclamation + vbOK, "MAHNFABRIK.DE  powered by SEPA Collect") = vbOK Then
            SavOK = False
            tBx122.ForeColor = RGB(255, 0, 0)
            tBx123.ForeColor = RGB(255, 0, 0)
        Else
            Exit Sub
        End If

    End If
    
    mPg1.Pages(1).Visible = True                                    ' Anzeige des Tabs 2. Mieter
    mPg1.Value = 1                                                  ' Tab 2. Mieter in den Vordergrund holen
    SavOK = True
    
End Sub

Private Sub bTnLbReset_Click()

    Cont1.Clear
    Call UserForm_Initialize

End Sub

' ###################################################################################################################################################################################
' +++++++++++++ Optionbutton (s) ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################

'   oPb10(ein Valutadatum für alle Teilforderungen = Nein)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub oPb9_Click()

    lBlValDatum_M.Visible = False
    lBlValDatum_A.Visible = False
    lBlValDatum_R.Visible = False
    lBlValDatum_S.Visible = False
    tBxValDatum_M.Visible = False
    tBxValDatum_A.Visible = False
    tBxValDatum_R.Visible = False
    tBxValDatum_S.Visible = False
    tBx2ValDatum_M.Visible = False
    tBx2ValDatum_A.Visible = False
    tBx2ValDatum_R.Visible = False
    tBx2ValDatum_S.Visible = False
    bTnValDatum_M.Visible = False
    bTnValDatum_A.Visible = False
    bTnValDatum_R.Visible = False
    bTnValDatum_S.Visible = False

End Sub


'   oPb10(ein Valutadatum für alle Teilforderungen = Nein)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub oPb10_Click()

    lBlValDatum_M.Visible = True
    lBlValDatum_A.Visible = True
    lBlValDatum_R.Visible = True
    lBlValDatum_S.Visible = True
    tBxValDatum_M.Visible = True
    tBxValDatum_A.Visible = True
    tBxValDatum_R.Visible = True
    tBxValDatum_S.Visible = True
    tBx2ValDatum_M.Visible = True
    tBx2ValDatum_A.Visible = True
    tBx2ValDatum_R.Visible = True
    tBx2ValDatum_S.Visible = True
    bTnValDatum_M.Visible = True
    bTnValDatum_A.Visible = True
    bTnValDatum_R.Visible = True
    bTnValDatum_S.Visible = True

End Sub


' ###################################################################################################################################################################################
' +++++++++++++ Comboboxen für Katalogauswahl +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################


'   cBxAnrede(Auswahl Adressanrede)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cBxAnrede_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    cBxAnrede.ForeColor = RGB(255, 0, 0)
End Sub

'   cBxKatNr(Auswahl Katalognummern; in Abhängigkeit zum AnsprGrundMM)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cBxKatNr_Change()

    With cBxKatNr
    
        If .Value < 4 Or .Value = 7 Then
            cBxAnsprMM.RowSource = "Katalog!d3:e4"
            Else
            If .Value = 4 Then
                cBxAnsprMM.RowSource = "Katalog!d6:e7"
                Else
                If .Value > 4 And .Value < 7 Then
                    cBxAnsprMM.RowSource = "Katalog!d9:e11"
                    Else
                    If .Value = "" Then
                        cBxAnsprMM.RowSource = "Katalog!d13:e16"
                    End If
                End If
            End If
        End If
    
    End With
        
End Sub

'   cBxAnsprMM(Auswahl AnsprGrundMM; in Abhängigkeit zur Katalognummern)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cBxAnsprMM_Change()

    With cBxAnsprMM
    
        If .Value = 1 Then
            cBxKatNr.RowSource = "Katalog!b3:c7"
            Else
            If .Value = 2 Then
                cBxKatNr.RowSource = "Katalog!b9:c12"
                Else
                If .Value = 3 Then
                    cBxKatNr.RowSource = "Katalog!b14:c16"
                    Else
                    If .Value = "" Then
                        cBxKatNr.RowSource = "Katalog!b18:c25"
                    End If
                End If
            End If
        End If
    
    End With
        
End Sub

' ###################################################################################################################################################################################
' +++++++++++++ Hilfsfunktionen / -prozeduren  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################

'   bTnPminus(Produktauswahl deaktivieren)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub bTnPminus()                                             ' Produktauswahl deaktivieren                                                                   ' -------------------------
          
    Dim i As Integer                                                ' Variablen definieren
                                                                    
    ListBox1.MultiSelect = fmMultiSelectSingle                      ' Mehrfachauswahl deaktivieren
    ListBox1.ListStyle = fmListStylePlain
    mPg1.Pages(3).Visible = True                                    ' Anzeige des Tabs Notizen einblenden
    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Hintergrundfarbe anpassen
        Me.Controls("tbx" & i).BackColor = &HFFFFFF
        On Error Resume Next
    Next i
    tBxValDatum_M.BackColor = &HFFFFFF
    tBxValDatum_A.BackColor = &HFFFFFF
    tBxValDatum_R.BackColor = &HFFFFFF
    tBxValDatum_S.BackColor = &HFFFFFF
    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                        ' Schriftfarbe anpassen
        Me.Controls("tbx" & i).ForeColor = &H0&
        On Error Resume Next
    Next i
    tBxValDatum_M.ForeColor = &H0&
    tBxValDatum_A.ForeColor = &H0&
    tBxValDatum_R.ForeColor = &H0&
    tBxValDatum_S.ForeColor = &H0&
    
    bTnPplusAct.Enabled = False
    
End Sub

'   ERR_COUNT_ERMITTELN(Ermittlung, ob der Fehlerwert err_Count > 0 ist, und damit das Speichern unterbunden werden muss)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ERR_COUNT_ERMITTELN(Optional ByVal lZeile As Long) As Boolean

    Dim errCt_1, errCt_2, errCt_3, errCt_4, errCt_5, errCt_6, errCt_7, errCt_8, errCt_107, errCt_106, _
        errCt_9, errCt_10, errCt_11, errCt_12, errCt_13, errCt_cBxAnr, errCt_16, errCt_17, errCt_18, _
        errCt_20, errCt_66, errCt_21, errCt_22, errCt_23, errCt_114, errCt_116, errCt_117, errCt_118, _
        errCt_119, errCt_19, errCt_110, errCt_cBxAnr2, errCt_35, errCt_36, errCt_37, errCt_38, errCt_67, _
        errCt_39, errCt_40, errCt_41, errCt_42, errCt_122, errCt_cBxAnr3, errCt_51, errCt_52, errCt_53, _
        errCt_54, errCt_68, errCt_55, errCt_56, errCt_57, errCt_125 As Integer
    
    If tBx1.Text = "" Then errCt_1 = 1 Else errCt_1 = 0
    If tBx2.Text = "" Then errCt_2 = 1 Else errCt_2 = 0
    If tBx3.Text = "" Then errCt_3 = 1 Else errCt_3 = 0
    If tBx4.Text = "" Then errCt_4 = 1 Else errCt_4 = 0
    If tBx5.Text = "" Then errCt_5 = 1 Else errCt_5 = 0
    If tBx7.Text = "" Then errCt_7 = 1 Else errCt_7 = 0
    If tBx8.Text = "" Then errCt_8 = 1 Else errCt_8 = 0
    If tBx107.Text = "" Then errCt_107 = 1 Else errCt_107 = 0
    If tBx106.Text = "" Then errCt_106 = 1 Else errCt_106 = 0
    If tBx9.Text = "" Then errCt_9 = 1 Else errCt_9 = 0
    If cBxAnrede.Text = "" Then errCt_cBxAnr = 1 Else errCt_cBxAnr = 0
    If tBx16.Text = "" Then errCt_16 = 1 Else errCt_16 = 0
    If tBx17.Text = "" Then errCt_17 = 1 Else errCt_17 = 0
    If tBx20.Text = "" Then errCt_20 = 1 Else errCt_20 = 0
    If tBx66.Text = "" Then errCt_66 = 1 Else errCt_66 = 0
    If tBx21.Text = "" Then errCt_21 = 1 Else errCt_21 = 0
    If tBx22.Text = "" Then errCt_22 = 1 Else errCt_22 = 0
    If tBx23.Text = "" Then errCt_23 = 1 Else errCt_23 = 0
    If tBx114.Text = "" Then errCt_114 = 1 Else errCt_114 = 0
    If tBx116.Text = "" Then errCt_116 = 1 Else errCt_116 = 0
    If tBx117.Text = "" Then errCt_117 = 1 Else errCt_117 = 0
    If tBx118.Text = "" Then errCt_118 = 1 Else errCt_118 = 0
    If tBx119.Text = "" Then errCt_119 = 1 Else errCt_119 = 0
    If tBx19.Text = "" Then errCt_19 = 1 Else errCt_19 = 0
    If tBx110.Text = "" Then errCt_110 = 1 Else errCt_110 = 0
    If weitMiet1 = True And cBxAnrede2.Text = "" Then errCt_cBxAnr2 = 1 Else errCt_cBxAnr2 = 0
    If weitMiet1 = True And tBx35.Text = "" Then errCt_35 = 1 Else errCt_35 = 0
    If weitMiet1 = True And tBx36.Text = "" Then errCt_36 = 1 Else errCt_36 = 0
    If weitMiet1 = True And tBx37.Text = "" Then errCt_37 = 1 Else errCt_37 = 0
    If weitMiet1 = True And tBx38.Text = "" Then errCt_38 = 1 Else errCt_38 = 0
    If weitMiet1 = True And tBx67.Text = "" Then errCt_67 = 1 Else errCt_67 = 0
    If weitMiet1 = True And tBx39.Text = "" Then errCt_39 = 1 Else errCt_39 = 0
    If weitMiet1 = True And tBx40.Text = "" Then errCt_40 = 1 Else errCt_40 = 0
    If weitMiet1 = True And tBx41.Text = "" Then errCt_41 = 1 Else errCt_41 = 0
    If weitMiet1 = True And tBx122.Text = "" Then errCt_122 = 1 Else errCt_122 = 0
    If weitMiet2 = True And cBxAnrede3.Text = "" Then errCt_cBxAnr3 = 1 Else errCt_cBxAnr3 = 0
    If weitMiet2 = True And tBx51.Text = "" Then errCt_51 = 1 Else errCt_51 = 0
    If weitMiet2 = True And tBx52.Text = "" Then errCt_52 = 1 Else errCt_52 = 0
    If weitMiet2 = True And tBx53.Text = "" Then errCt_53 = 1 Else errCt_53 = 0
    If weitMiet2 = True And tBx54.Text = "" Then errCt_54 = 1 Else errCt_54 = 0
    If weitMiet2 = True And tBx68.Text = "" Then errCt_68 = 1 Else errCt_68 = 0
    If weitMiet2 = True And tBx55.Text = "" Then errCt_55 = 1 Else errCt_55 = 0
    If weitMiet2 = True And tBx56.Text = "" Then errCt_56 = 1 Else errCt_56 = 0
    If weitMiet2 = True And tBx57.Text = "" Then errCt_57 = 1 Else errCt_57 = 0
    If weitMiet2 = True And tBx125.Text = "" Then errCt_125 = 1 Else errCt_125 = 0
    
    ErrCount = errCt_1 + errCt_2 + errCt_3 + errCt_4 + errCt_5 + errCt_6 + errCt_7 + errCt_8 + errCt_107 + errCt_106 + _
        errCt_9 + errCt_10 + errCt_11 + errCt_12 + errCt_13 + errCt_cBxAnr + errCt_16 + errCt_17 + errCt_18 + _
        errCt_20 + errCt_66 + errCt_21 + errCt_22 + errCt_23 + errCt_114 + errCt_116 + errCt_117 + errCt_118 + _
        errCt_119 + errCt_19 + errCt_110 + errCt_cBxAnr2 + errCt_35 + errCt_36 + errCt_37 + errCt_38 + errCt_67 + _
        errCt_39 + errCt_40 + errCt_41 + errCt_42 + errCt_122 + errCt_cBxAnr3 + errCt_51 + errCt_52 + errCt_53 + _
        errCt_54 + errCt_68 + errCt_55 + errCt_56 + errCt_57 + errCt_125
    
    If ErrCount > 0 Then ERR_COUNT_ERMITTELN = True
    
End Function




Private Function IST_ZEILE_LEER(ByVal lZeile As Long) As Boolean

    If CopyModeOn = True Then
        varQuelle = "DATA_UPLOAD_ARCHIV"
    Else
        varQuelle = "DATA_UPLOAD"
    End If
    
    Dim i As Long                                                                   ' erforderliche Variablen definieren
    Dim stemp As String
      
    stemp = ""                                                                      ' Hilfsvariable initialisieren
                                                                                    
    For i = 1 To iCONST_ANZAHL_EINGABEFELDER                                        ' Um zu erkennen, ob eine Zeile komplett leer/ungebraucht ist
        stemp = stemp & Trim(CStr(Worksheets(varQuelle).Cells(lZeile, i).Text)) ' werden alle Spalteninhalte der Zeile miteinander verkettet.
    Next i                                                                          ' Ist die zusammengesetzte Zeichenkette aller Spalten leer,
                                                                                    ' ist die Zeile nicht genutzt.
    If Trim(stemp) = "" Then                                                        ' Rückgabewert festlegen:
        IST_ZEILE_LEER = True                                                       ' hier: Zeile ist leer
    Else
        IST_ZEILE_LEER = False                                                      ' hier: Zeile ist mindestens in einer Spalte gefüllt
    End If
    
End Function

'   IST_VALUTA_GLEICH(Ermittlung, ob die Valutadaten der Teilforderungen identisch sind)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function IST_VALUTA_GLEICH(ByVal lZeile As Long) As Boolean
                                                                                   
    If tBxValDatum_M.Text = tBx33.Text _
        And tBxValDatum_A.Text = tBx33.Text _
        And tBxValDatum_R.Text = tBx33.Text _
        And tBxValDatum_S.Text = tBx33.Text Then
        IST_VALUTA_GLEICH = True
    Else
        IST_VALUTA_GLEICH = False
    End If
    
End Function




'   IST_MIETER2_LEER(Ermittlung, ob auf der Seite "2. Mieter" Einträge vorhanden sind)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function IST_MIETER2_LEER(ByVal lZeile As Long) As Boolean
                                                                                   
    If CopyModeOn = True Then
        varQuelle = "DATA_UPLOAD_ARCHIV"
    Else
        varQuelle = "DATA_UPLOAD"
    End If
                                                                                   
    strTmp2 = CStr(Worksheets(varQuelle).Cells(lZeile, 86).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 87).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 98).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 88).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 89).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 92).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 93).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 97).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 94).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 95).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 96).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 100).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 101).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 102).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 103).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 104).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 105).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 106).Text)
    
    If Trim(strTmp2) <> "" Then                                                     ' ja = Seite Anzeigen, nein = Seite bleibt ausgeblendet
        IST_MIETER2_LEER = False
    Else
        IST_MIETER2_LEER = True
    End If
    
End Function

'   IST_MIETER3_LEER(Ermittlung, ob auf der Seite "3. Mieter" Einträge vorhanden sind)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function IST_MIETER3_LEER(ByVal lZeile As Long) As Boolean

    If CopyModeOn = True Then
        varQuelle = "DATA_UPLOAD_ARCHIV"
    Else
        varQuelle = "DATA_UPLOAD"
    End If
    
    strTmp3 = CStr(Worksheets(varQuelle).Cells(lZeile, 114).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 115).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 126).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 116).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 117).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 120).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 121).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 125).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 122).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 123).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 124).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 128).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 129).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 130).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 131).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 132).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 133).Text) & _
              CStr(Worksheets(varQuelle).Cells(lZeile, 134).Text)
                    
    If Trim(strTmp3) <> "" Then                                                     ' ja = Seite Anzeigen, nein = Seite bleibt ausgeblendet
        IST_MIETER3_LEER = False
    Else
        IST_MIETER3_LEER = True
    End If
    
End Function

'   SUMMENFELDER(Ermittlung der Spaltensummen getrennt für jede Spalte in der Listbox)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SUMMENFELDER()

    Dim i As Integer                                ' Variablen definieren
    Dim SpSum As Single
   
    With ListBox1                                   ' Summe über Spalte 4 (= Hauptforderung)
         SpSum = 0
            For i = 0 To .ListCount - 1
                If (ListBox1.List(i, 4)) = "" Then
                    SpSum = SpSum
                Else
                    SpSum = SpSum + .List(i, 4)
                End If
            Next i
    End With
    tBxSumHF = Format(SpSum, "##,##0.00")

    i = 0                                           ' Zähler zurücksetzen
 
    With ListBox1                                   ' Summe über Spalte 5 (= Mahnkosten)
         SpSum = 0
            For i = 0 To .ListCount - 1
                If (ListBox1.List(i, 5)) = "" Then
                    SpSum = SpSum
                 Else
                    SpSum = SpSum + .List(i, 5)
                 End If
            Next i
    End With
    tBxSumMahn = Format(SpSum, "##,##0.00")
 
    i = 0                                           ' Zähler zurücksetzen
 
    With ListBox1                                   ' Summe über Spalte 6 (= Auskunftskosten)
         SpSum = 0
            For i = 0 To .ListCount - 1
                If (ListBox1.List(i, 6)) = "" Then
                    SpSum = SpSum
                Else
                    SpSum = SpSum + .List(i, 6)
                End If
            Next i
    End With
    tBxSumAusk = Format(SpSum, "##,##0.00")

    i = 0                                           ' Zähler zurücksetzen
 
    With ListBox1                                   ' Summe über Spalte 7 (= Bankrücklastschriftgebühren)
         SpSum = 0
            For i = 0 To .ListCount - 1
                If (ListBox1.List(i, 7)) = "" Then
                    SpSum = SpSum
                Else
                    SpSum = SpSum + .List(i, 7)
                End If
            Next i
    End With
    tBxSumRLS = Format(SpSum, "##,##0.00")
 
    i = 0                                            ' Zähler zurücksetzen
    
    With ListBox1                                    ' Summe über Spalte 8 (= Sonstige Nebenforderungen)
         SpSum = 0
             For i = 0 To .ListCount - 1
                 If (ListBox1.List(i, 8)) = "" Then
                     SpSum = SpSum
                 Else
                     SpSum = SpSum + .List(i, 8)
                 End If
             Next i
     End With
     tBxSumSoNF = Format(SpSum, "##,##0.00")
 
    i = 0                                            ' Zähler zurücksetzen
    
    With ListBox1                                    ' Summe über Spalte 9 (= Gesamtforderung)
         SpSum = 0
            For i = 0 To .ListCount - 1
                If (ListBox1.List(i, 9)) = "" Then
                    SpSum = SpSum
                Else
                    SpSum = SpSum + .List(i, 9)
                End If
            Next i
    End With
    tBxSumGes = Format(SpSum, "##,##0.00")
  
End Sub

'   SET_USERFORM_STYLE(zum Deaktivieren der Funktionen in der Titelleiste; siehe Sonderbereich im Kopf des Codes)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SET_USERFORM_STYLE()                                                      '

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

'   BUTTON_STANDARD(Standard Buttonsteuerung)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub BUTTON_STANDARD()                                                        '

    bTnCopy.Visible = True
    bTnCopy.Enabled = True
    bTnAdd.Visible = True
    bTnAdd.Enabled = True
    bTnDel.Visible = True
    bTnDel.Enabled = True
    bTnSav.Visible = True
    bTnSav.Enabled = True
    lBlModule.Visible = True
    bTnBck.Visible = True
    bTnBck.Enabled = True
    bTnBB.Visible = True
    
    tBx14.Enabled = False
    
End Sub

'   LISTBOX1_CHANGE(Zählen, ob Datensätze zum Buchen markiert sind; erst danach wird der Buchen-Button aktiv)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LISTBOX1_CHANGE()
    
    SavOK = False
    
    If CopyModeOn = True Then
        lBlUeber.Enabled = True
        cBnUeber.Enabled = True
    End If

    Dim iSelCnt As Integer                                                      ' Erforderliche Variablen definieren
    Dim iX As Integer

    iSelCnt = 0
    
    If ListBox1.MultiSelect = fmMultiSelectMulti Then
    
        For iX = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(iX) = True Then iSelCnt = iSelCnt + 1
        Next
        iX = iSelCnt
        MsgBox iX
    Else
        'iX = 0
    End If
'    MsgBox "Zähler iX: " & iX
'    MsgBox "Zustand Lbox1: " & ListBox1.MultiSelect
'    MsgBox "ProdmodeOn: " & ProdModeOn
'    If iX > 0 And ListBox1.MultiSelect = fmMultiSelectMulti And ProdModeOn = True Then
'        bTnPplusAct.Enabled = True
'    End If

    If ListBox1.MultiSelect = fmMultiSelectMulti Then
       
        If iX > 0 Then
            
            If ProdModeOn = True Then
                
                bTnPplusAct.Enabled = True
            End If
        End If
    End If


End Sub


'   ARCHIV_UEBERNAHME(Füllen der "Copy-Felder" (als Zwischenablage zum Kopieren der aktuell angezeigten Werte)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ARCHIV_UEBERNAHME()

    tBx15Copy.Text = tBx15.Text
    tBx16Copy.Text = tBx16.Text
    tBx17Copy.Text = tBx17.Text
    tBx18Copy.Text = tBx18.Text
    tBx19Copy.Text = tBx19.Text
    tBx20Copy.Text = tBx20.Text
    tBx21Copy.Text = tBx21.Text
    tBx22Copy.Text = tBx22.Text
    tBx23Copy.Text = tBx23.Text
    tBx24Copy.Text = tBx24.Text
    tBx25Copy.Text = tBx25.Text
    tBx26Copy.Text = tBx26.Text
    tBx27Copy.Text = tBx27.Text
    tBx28Copy.Text = tBx28.Text
    tBx29Copy.Text = tBx29.Text
    tBx30Copy.Text = tBx30.Text
    tBx31Copy.Text = tBx31.Text
    tBx66Copy.Text = tBx66.Text
    tBx119Copy.Text = tBx119.Text
    tBx120Copy.Text = tBx120.Text
    tBx34Copy.Text = tBx34.Text
    tBx35Copy.Text = tBx35.Text
    tBx36Copy.Text = tBx36.Text
    tBx37Copy.Text = tBx37.Text
    tBx121Copy.Text = tBx121.Text
    tBx38Copy.Text = tBx38.Text
    tBx39Copy.Text = tBx39.Text
    tBx40Copy.Text = tBx40.Text
    tBx41Copy.Text = tBx41.Text
    tBx42Copy.Text = tBx42.Text
    tBx43Copy.Text = tBx43.Text
    tBx44Copy.Text = tBx44.Text
    tBx45Copy.Text = tBx45.Text
    tBx46Copy.Text = tBx46.Text
    tBx47Copy.Text = tBx47.Text
    tBx48Copy.Text = tBx48.Text
    tBx49Copy.Text = tBx49.Text
    tBx67Copy.Text = tBx67.Text
    tBx122Copy.Text = tBx122.Text
    tBx123Copy.Text = tBx123.Text
    tBx50Copy.Text = tBx50.Text
    tBx51Copy.Text = tBx51.Text
    tBx52Copy.Text = tBx52.Text
    tBx53Copy.Text = tBx53.Text
    tBx124Copy.Text = tBx124.Text
    tBx54Copy.Text = tBx54.Text
    tBx55Copy.Text = tBx55.Text
    tBx56Copy.Text = tBx56.Text
    tBx57Copy.Text = tBx57.Text
    tBx58Copy.Text = tBx58.Text
    tBx59Copy.Text = tBx59.Text
    tBx60Copy.Text = tBx60.Text
    tBx61Copy.Text = tBx61.Text
    tBx62Copy.Text = tBx62.Text
    tBx63Copy.Text = tBx63.Text
    tBx64Copy.Text = tBx64.Text
    tBx65Copy.Text = tBx65.Text
    tBx68Copy.Text = tBx68.Text

    Call EINTRAG_ANLEGEN_AUS_COPY                  ' Aufruf der entsprechenden Verarbeitungsroutine (Anlage neuer Datensatzes auf Bssis Copy-Felder)

End Sub


' ###################################################################################################################################################################################
' +++++++++++++ Pflichteingaben / Pflichtformatierungen +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ###################################################################################################################################################################################
'                                                                                   - Pflichtfelder werden, wenn leer, rot hinterlegt
'                                                                                   - Felder mit fehlerhaften Formaten bei der Eingabe werden immer geleert,
'                                                                                   - bei Pflichtfeldern rot, ansonsten gelb hinterlegt
'                                                                                   - Formatprüfung erfolgt nur beim Erfassen
'   Zahlen OHNE Nachkommastellen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   tBx1 - Mandant
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx1_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If ListBox1.ListCount = 0 Then                                  ' Funktionen deaktivieren, wenn Tabelle DATA_UPLOAD leer ist
        AddModeOn = False
        Exit Sub
    End If

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx1
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With

End Sub

'   tBx2 - Unternehmen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx2_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx2
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx3 - WE
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx3_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx3
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub
        
        End If

    End With
    
End Sub

'   tBx4 - HausNr
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx4_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx4
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx5 - WohnNr
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx5_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx5
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx7 - Folgenummer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx7_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx7
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx8 - Belegnummer/OP-Nummer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx8_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx8
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx107 - lfdNr
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx107_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx107
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub


'   tBx106 - Geschäftsjahr
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tBx106_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx106
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub


'   tBx19 - AdressNr des Hauptmieters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx19_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx19
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub


'   tBx121 - AdressNr des 2. Mieters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx121_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx121
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx124 - AdressNr des 3. Mieters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx124_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx124
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx22 - PLZ Hauptmieter Zustelladresse
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx22_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx22
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub


'   tBx117 - PLZ Hauptmieter Leistungsadresse
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx117_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx117
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                 .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx40 - PLZ 2. Mieter Zustelladresse
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx40_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx40
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx56 - PLZ 3. Mieter Zustelladresse
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx56_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx56
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "0")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub


'   Zahlen MIT 2 Nachkommastellen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   tBx9 - Hauptforderung
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx9_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx9
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "#,##0.00")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx10 - Mahnkosten
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx10_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx10
    
        If .Value = "" Then
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "#,##0.00")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With

End Sub

'   tBx11 - Auskunftskosten
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx11_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx11
    
        If .Value = "" Then
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "#,##0.00")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
 
End Sub

'   tBx12 - Bankrücklastschriftgebühren
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx12_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx12
    
        If .Value = "" Then
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "#,##0.00")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx13 - Sonstige Nebenforderungen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx13_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx13
    
        If .Value = "" Then
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "#,##0.00")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With

End Sub

'   Datum
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   tBx33 - Valuta Datum                     .Text = Format(.Text, "dd.mm.yyyy")
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx33_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx33
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub


'   tBx108 - Anspruch-von Datum
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx108_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx108
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx109 - Anspruch-bis Datum
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx109_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx109
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx119 - Geburtsdatum Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx119_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx119
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx122 - Geburtsdatum 2. Mieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx122_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx122
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx125 - Geburtsdatum 3. Mieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx125_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx125
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx120 - Todesdatum Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx120_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx120
    
        If .Value = "" Then
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx123 - Todesdatum 2. Mieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx123_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx123
    
        If .Value = "" Then
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub

'   tBx126 - Todesdatum 3. Mieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx126_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx126
    
        If .Value = "" Then
            Exit Sub
        Else
            ErrCount = ErrCount
            If IsDate(.Text) Then
                .Text = Format(.Text, "dd.mm.yyyy")
                ErrCount = ErrCount - 1
            Else
                ErrCount = ErrCount + 1
                .Text = ""
                Cancel = False
            End If

            Exit Sub

        End If

    End With
    
End Sub


'   Texteingaben
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   tBx16 - Zustelladresse Nachname Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx16_exit(ByVal Cancel As MSForms.ReturnBoolean)

    With tBx16
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx17 - Zustelladresse Vorname Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx17_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx17
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx20 - Zustelladresse Straße Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx20_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx20
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx66 - Zustelladresse Hausnummer Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx66_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx66
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx21 - Zustelladresse Land Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx21_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx21
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

Private Sub tBx21_change()

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx21
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx23 - Zustelladresse Ort Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx23_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx23
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx114 - Leistungsadresse Straße + Hausnummer Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx114_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx114
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx116 - Leistungsadresse Land Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx116_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx116
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

Private Sub tBx116_change()

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx116
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx118 - Leistungsadresse Ort Hauptmieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx118_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx118
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx39 - Zustelladresse Land 2. Mieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx39_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx39
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

Private Sub tBx39_change()

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx39
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   tBx55 - Zustelladresse Land 2. Mieter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbx55_exit(ByVal Cancel As MSForms.ReturnBoolean)

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx55
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
            Cancel = False
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

Private Sub tBx55_change()

    If AddModeOn = False Then
        Exit Sub
    End If

    With tBx55
    
        If .Value = "" Then
            ErrCount = ErrCount + 1
        Else
            UCase (.Text)
            ErrCount = ErrCount - 1
        End If
    
    End With
    
End Sub

'   zzzz '''''''''''''''''
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Cont1_Change()

'        Cont2.Clear

        Dim i As Integer
        Dim inhalt As String

'        For i = 0 To ListBox1.ListCount - 1
'            If ListBox1.List(i, 1) = Cont1.Value Then
'                inhalt = ListBox1.List(i, 2)
'
'                Dim gefunden As Boolean
'
'                If (Not inhalt_existiert(inhalt, Cont2)) Or (Cont2.ListCount = 0) Then 'Wert aus function
'                    Cont2.AddItem (inhalt)
'                End If
'
'
'            End If
'        Next i
'
'
        Dim anz_geloescht As Integer


        Do
            i = 0
            anz_geloescht = 0
            For i = 0 To ListBox1.ListCount - 1
                If ListBox1.List(i, 1) <> Cont1.Value Then
                    ListBox1.RemoveItem i
                    anz_geloescht = 1
                    Exit For
                End If
            Next i
        Loop While anz_geloescht <> 0

  Call SUMMENFELDER



End Sub

'Private Sub Cont1_Click()
'Dim lIndxA   As Long      ' For/Next Index - außen
'Dim lIndxI   As Long      ' For/Next Index - innen
'Dim stemp As String
''+++ ComboBox sortieren :
'   For lIndxA = 0 To Me.Cont1.ListCount - 1
'      For lIndxI = 0 To lIndxA - 1
'         If Me.Cont1.List(lIndxI) > Me.Cont1.List(lIndxA) Then
'            stemp = Me.Cont1.List(lIndxI)
'            Me.Cont1.List(lIndxI) = Me.Cont1.List(lIndxA)
'            Me.Cont1.List(lIndxA) = stemp
'         End If
'      Next lIndxI
'   Next lIndxA
'End Sub


Private Sub Cont1_Enter()

'   If Cont2.Value = "" Then

    Cont1.Clear

    Dim inhalt As String

    Dim i As Integer

    For i = 0 To ListBox1.ListCount - 1
     inhalt = ListBox1.List(i, 1)            'Mieterkey besorgen


     Dim n As Integer
     Dim gefunden As Boolean

     gefunden = False
     For n = 0 To Cont1.ListCount - 1
         If Cont1.List(n, 1) = inhalt Then
             gefunden = True
             Exit For
         End If
     Next

     If (Not gefunden) Or (Cont1.ListCount = 0) Then
         Cont1.AddItem inhalt
     End If

    Next i
   
End Sub


'Private Sub Cont1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Cont2.Enabled = False
'    Cont3.Enabled = False
'
'End Sub

'Function inhalt_existiert(inhalt As String, listbox As ComboBox) As Boolean
'    Dim ret As Boolean
'    ret = False
'
'    Dim gefunden As Boolean
'
'
'    Dim n As Integer
'    For n = 0 To listbox.ListCount - 1
'        If listbox.List(n, 0) = inhalt Then
'            ret = True
'            Exit For
'        End If
'    Next
'
'
'    inhalt_existiert = ret
'End Function



'Private Sub Cont2_Change()
'
''    Cont1.Clear
'
'    Dim i As Integer
'    Dim inhalt As String
'
'    For i = 0 To ListBox1.ListCount - 1
'        If ListBox1.List(i, 2) = Cont2.Value Then
'
'            inhalt = ListBox1.List(i, 1)
'
'            Dim gefunden As Boolean
'
'            gefunden = False
'            Dim n As Integer
'            For n = 0 To Cont1.ListCount - 1
'                If Cont1.List(n, 0) = inhalt Then
'                    gefunden = True
'                    Exit For
'                End If
'            Next
'
'            If (Not gefunden) Or (Cont1.ListCount = 0) Then
'                Cont1.AddItem (inhalt)
'            End If
'
'
'        End If
'    Next i
'        Dim anz_geloescht As Integer
'
'
'        Do
'            i = 0
'            anz_geloescht = 0
'            For i = 0 To ListBox1.ListCount - 1
'                If ListBox1.List(i, 2) <> Cont1.Value Then
'                    ListBox1.RemoveItem i
'                    anz_geloescht = 1
'                    Exit For
'                End If
'            Next i
'        Loop While anz_geloescht <> 0
'
'
'Call SUMMENFELDER
'
'
'End Sub
'
'Private Sub Cont2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Cont1.Enabled = False
'    Cont3.Enabled = False
'
'End Sub
'
'Private Sub Cont3_Change()
'
'
'End Sub
'
'
'
'
'Private Sub Cont2_Enter()
'
'    If Cont1.Value = "" Then
'
'        Cont2.Clear
'
'        Dim inhalt As String
'
'       Dim i As Integer
'
'       For i = 0 To ListBox1.ListCount - 1
'        inhalt = ListBox1.List(i, 2)            'Name und Vorname besorgen
'
'        Dim n As Integer
'        Dim gefunden As Boolean
'
'        gefunden = False
'        For n = 0 To Cont2.ListCount - 1
'            If Cont2.List(n, 0) = inhalt Then
'                gefunden = True
'                Exit For
'            End If
'        Next
'
'        If (Not gefunden) Or (Cont2.ListCount = 0) Then
'            Cont2.AddItem inhalt
'        End If
'
'       Next i
'
'    End If
'
'
'
'End Sub
''
'
'
'
'
'
'
