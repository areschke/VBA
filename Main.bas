Attribute VB_Name = "Main"
Option Explicit

Public closeVar As Boolean

Public meineCloseVar  As Boolean

Public CSVempty, CSVDirekt, CSVError As Boolean

Public DatAngabe, strMappenpfad, strKundennummer, strDateiname As String


Sub Datenerfassung_Klicken()            ' Datenerfassung: Userform 'UFDataUpload' öffnen
                                        ' Die Startautomatismen sind in der Userform hinterlegt.
    UFDataUpload.Show
   
End Sub

Sub CSV_Klicken()                                                       ' Bildung der CSV-Datei aus den Einträgen im Sheet 'DATA_UPLOAD'

If Worksheets("ERROR").Range("A1").Value = "1" Then
    MsgBox "Es liegen noch unvollständige Datensätze vor. Bitte erst ergänzen oder löschen." & vbCr _
            , vbInformation, "MAHNFABRIK.DE  powered by SEPA Collect"
    Worksheets("ERROR").Range("B1").Value = "1"
    Exit Sub
End If

Call EXPORT_LEER

DatAngabe = Format(Now, "YYYY" & "_" & "MM" & "_" & "DD" & "_" & "hh" & "_" & "mm" & "_" & "ss")

If CSVempty = True Then                                                 ' Vorabprüfung, ob überhaupt Daten für den Export vorhanden sind.
    MsgBox "Es liegen keine Daten zum Upload vor." & vbCr _
            , vbInformation, "MAHNFABRIK.DE  powered by SEPA Collect"
    Exit Sub
    End If

' Variablen definieren
    Dim Bereich As Object, Zeile As Object, Zelle As Object
    Dim strTemp, strTrennzeichen, strKundennummer, intMessage1, intMessage2  As String
    Dim blnAnfuehrungszeichen As Boolean
    Dim ObjShell
    Dim Wscript
    
' Speicherort festlegen
    strMappenpfad = Worksheets("PARAM").Cells(11, 6).Value
    strKundennummer = Worksheets("PARAM").Cells(17, 6).Value
    strDateiname = strMappenpfad & strKundennummer & "_" & DatAngabe & ".csv"
    strTrennzeichen = ";"
     
    blnAnfuehrungszeichen = False
    Set Bereich = Worksheets("DATA_UPLOAD").UsedRange
     
    Open strDateiname For Output As #1
     
    For Each Zeile In Bereich.Rows
        For Each Zelle In Zeile.Cells
            If blnAnfuehrungszeichen = True Then
                strTemp = strTemp & """" & CStr(Zelle.Text) & """" & strTrennzeichen
            Else
                strTemp = strTemp & CStr(Zelle.Text) & strTrennzeichen
            End If
        Next
        If Right(strTemp, 1) = strTrennzeichen Then strTemp = Left(strTemp, Len(strTemp) - 1)
        Print #1, strTemp
        strTemp = ""
    Next
    
    Close #1
    Set Bereich = Nothing
      
    Set ObjShell = CreateObject("Wscript.Shell")
    
    intMessage1 = MsgBox("D A T E I - E X P O R T" & vbCr _
            & "==============================================" & vbCr _
            & "Der Export der Daten in die Datei" & vbCr _
            & " *** " & strDateiname & " *** " & vbCr _
            & "ist erfolgt." & vbCr _
            & vbCr _
            & "Wir empfehlen, direkt den Upload zur Mahnfabrik vorzunehmen." & vbCr & vbCr _
            & "Hinweise:" & vbCr _
            & "--------------" & vbCr _
            & "Die jetzt übermittelten Datensätze werden in den Archivbereich verschoben und stehen Ihnen über die Kopierfunktion (' ++ ') zur Erfassung weiterer Forderungen " _
            & "bei bereits bekannten Schuldnern zur Verfügung." & vbCr & vbCr _
            & "Die CSV-Datei wird jetzt in den Ordner  " & strMappenpfad & "Versendet" & "  verschoben." & vbCr & vbCr _
            & "Wollen Sie den Upload jetzt durchführen?", vbOKCancel, "MAHNFABRIK.DE powered by SEPA Collect")
    
    If intMessage1 = 1 Then
        Call Archivierung
        Call File_verschieben
        ObjShell.Run ("https://portal.sepacollect.de/")
        CSVDirekt = True
        Call LogDetails
    Else
        Call Archivierung
        MsgBox "Vergessen Sie bitte nicht, den Upload nachzuholen." & vbCr _
             & "Die CSV-Datei wird jetzt in den Ordner  " & strMappenpfad & "  verschoben." & vbCr & vbCr, vbInformation, "MAHNFABRIK.DE  powered by SEPA Collect"
        CSVDirekt = False
        Call LogDetails
    End If
     
    Worksheets("START").Select
    
End Sub

Sub DOKU()
' Dokumentation öffnen
'
  Sheets("DOKU").Select
  ActiveWindow.DisplayHeadings = False

End Sub

Sub START()
' Startseite öffnen
' zum Rücksprung auf die Hauptmaske

  Sheets("START").Select

End Sub
Sub SYSTEM()
'
' Systemeinstellungen öffnen
'
  Sheets("PARAM").Select
  ActiveWindow.DisplayHeadings = False

End Sub

Public Sub UFDataUpload_activate()                                  ' Ereignisroutine beim Anzeigen der UserForm
Dim ListBox1

    
    With UFDataUpload                                               ' Anpassen der Größe der Userform auf die Größe des aktuellen Anwendung
        .Top = Application.Top                                      ' (da Excel im Vollbildmodus gestartet wird, wird dieser dann auch hier übernommen)
        .Left = Application.Left
        .Height = Application.Height
        .Width = Application.Width
    End With
     
    If UFDataUpload.ListBox1.ListCount > 0 Then UFDataUpload.ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
    
    UFDataUpload.bTnCopy.Visible = True
    
End Sub


Public Sub Grafik19()                                               ' Beim Beenden wieder alles zurücksetzen

    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
    End If
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFormulaBar = True
    Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"", True)"
    Application.DisplayStatusBar = True
    meineCloseVar = True

    ThisWorkbook.Save
    ThisWorkbook.Saved = True
    Application.Quit

End Sub

Public Sub Grafik20()                                               ' siehe Grafik19

    Call Grafik19
 
End Sub


Public Sub Archivierung()                                                      ' Übertragen der in der CSV-Datei enthaltenen Datensätze
                                                                                ' nach der Speicherung der Datei ins Archiv
                                                                                ' a) als interne Sicherung und
                                                                                ' b) als Kopiervorlage bei neuen Datensätzen

    Dim iRow, iRowArc As Integer                                                ' erforderliche Variablen definieren
    Dim lZeileMaximum, lZeileMaximumARC, lZeileMaximumLOG As Long
    Dim objData As New DataObject
    Dim varVar As Variant
    
    Application.ScreenUpdating = False
    
    lZeileMaximum = Worksheets("DATA_UPLOAD").UsedRange.Rows.Count              ' Ermittlung letzte Zeile in aktuellen Daten
    lZeileMaximumARC = Worksheets("DATA_UPLOAD_ARCHIV").UsedRange.Rows.Count    ' Ermittlung letzte Zeile in Archivdaten
    lZeileMaximumLOG = Worksheets("LOG").UsedRange.Rows.Count                   ' Ermittlung letzte Zeile in LOG-Bereich
    iRow = 2                                                                    ' Zähler aktuelle Daten auf erste Datenzeile (= Zeile 2) setzen
    
    Worksheets("DATA_UPLOAD").Select                                            ' auf aktuelle Daten zugreifen
    
    If Range("A2") = "" Then                                                    ' Prüfen, ob überhaupt aktuelle Daten zum Archivieren verfügbar sind
        Exit Sub
    End If
    
    For iRow = 2 To lZeileMaximum                                               ' Kopieren aller Datenzeilen aus dem Sheet "DATA_UPLOAD" in das
        Worksheets("DATA_UPLOAD").Select                                        ' Sheet "DATA_UPLOAD_ARCHIV" (Daten werden angehängt)
        Range("A" & iRow).EntireRow.Copy
        objData.GetFromClipboard
        varVar = objData.GetText
        Worksheets("DATA_UPLOAD_ARCHIV").Select
        lZeileMaximumARC = Worksheets("DATA_UPLOAD_ARCHIV").UsedRange.Rows.Count
        Range("A" & lZeileMaximumARC + 1).EntireRow.PasteSpecial
        Set objData = Nothing
    Next iRow
    
    Worksheets("DATA_UPLOAD").Select                                            ' Zurück zu den aktuellen Daten
                                                                                ' um diese zu löschen
    For iRow = 2 To lZeileMaximum
        ActiveSheet.Rows(2).Delete
    Next iRow
    
    Worksheets("LOG").Select
    
    If CSVDirekt = True Then
        Range("A" & lZeileMaximumLOG + 1).Value = "Speichern in '" & strDateiname & "\Vensendet' ERFOLGREICH."
    Else
        Range("A" & lZeileMaximumLOG + 1).Value = "Speichern in '" & strDateiname & "' ERFOLGREICH."
    End If
    
    CSVempty = True
    
    Application.ScreenUpdating = True
End Sub

Public Sub EXPORT_LEER()                                                        ' Vorabprüfung um Variable zu füllen

    Application.ScreenUpdating = False

    If Worksheets("DATA_UPLOAD").Range("A2").Value = "" Then
        CSVempty = True
    Else
        CSVempty = False
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Public Sub File_verschieben()
    Dim Quelle$, Ziel$, fso As Object
    Quelle = strMappenpfad & "*.csv"
    If Dir(Quelle) = "" Then
        Exit Sub
    Else
        Ziel = strMappenpfad & "Versendet"
        Set fso = CreateObject("Scripting.FileSystemObject")
        Application.DisplayAlerts = False
        fso.MoveFile Quelle, Ziel
        Application.DisplayAlerts = True
        Set fso = Nothing
    End If
End Sub


Public Function LogDetails()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim logFile As Object
    Dim logFileName, myFilePath As String

    Application.ScreenUpdating = False
    
    logFileName = "Mahnfabrik_CSVlog" & ".txt"
    myFilePath = strMappenpfad & logFileName
    


    If fso.FileExists(myFilePath) Then
        Set logFile = fso.OpenTextFile(myFilePath, 8)
    Else
        Set logFile = fso.CreateTextFile(myFilePath, True)
    End If

    If CSVDirekt = True Then
        logFile.WriteLine strKundennummer & "_" & DatAngabe & " " & "Gespeichert in " & " " & strMappenpfad & "Versendet\"
    Else
        logFile.WriteLine strKundennummer & "_" & DatAngabe & " " & "Gespeichert in " & " " & strMappenpfad
    End If
    
    logFile.Close
   
    Application.ScreenUpdating = True
    
End Function


'Public Function CheckDatum(ByVal Datum As String) As String
'  Dim Result As Boolean
'  Dim TT As Integer
'  Dim MM As Integer
'  Dim JJ As Integer
'  Dim sPos As Integer
'  Dim Entry As String
'
'  Result = False
'  ' Ist überhaupt ein Datum eingetragen?
'  If Trim$(Datum) <> ".  ." And Datum <> "" Then
'
'     ' keine Punkte im Datum vorhanden
'    If InStr(Datum, ".") = 0 Then
'
'      ' Eingabelänge mindestens 4-stellig
'      If Len(Datum) > 3 Then
'        TT = Val(Left$(Datum, 2))
'        MM = Val(Mid$(Datum, 3, 2))
'
'        If Trim$(Mid$(Datum, 5)) = "" Then
'          ' Jahresangabe fehlt -> aktuelles Jahr annehmen
'          JJ = Val(Right$(Date$, 4))
'        Else
'          JJ = Val(Mid$(Datum, 5))
'        End If
'        Result = True
'      End If
'    Else
'
'      ' Eingabe enthält Punktangaben
'      Entry = Datum
'
'      ' Tag ermitteln
'      sPos = InStr(Entry, ".")
'      TT = Val(Left$(Entry, sPos - 1))
'      Entry = Mid$(Entry, sPos + 1)
'
'      ' Monat ermitteln
'      sPos = InStr(Entry, ".")
'      If sPos Then
'        MM = Val(Left$(Entry, sPos - 1))
'        Entry = Mid$(Entry, sPos + 1)
'
'        ' keine Jahresangabe -> aktuelles Jahr annehmen
'        If Trim$(Entry) = "" Then
'          Entry = Right$(Date$, 4)
'        End If
'
'        JJ = Val(Entry)
'        Result = True
'      End If
'    End If
'
'    If Result Then
'      Result = False
'
'      ' Tag prüfen (Bereich 1-31)
'      If TT > 0 And TT < 32 Then
'
'        ' Monat prüfen (Bereich 1-12)
'        If MM > 0 And MM < 13 Then
'
'          ' wenn Jahresangabe zweistellig
'          If JJ < 100 Then
'            ' wenn kleiner 30
'            If JJ < 30 Then
'              ' Jahr 2000 annehmen
'              JJ = 2000 + JJ
'            Else
'              ' Jahr 1900 annehmen
'              JJ = 1900 + JJ
'            End If
'          End If
'
'          ' wenn Tag größer als maximale Anzahl Tage
'          ' im angegeben Monat
'          If TT > Day(DateSerial(JJ, MM + 1, 1) - 1) Then
'            ' Tagesangabe korrigieren
'            TT = Day(DateSerial(JJ, MM + 1, 1) - 1)
'          End If
'
'          Result = True
'        End If
'      End If
'
'      If Result Then
'        ' wenn alles OK - Datum formatieren
'        ' tt.mm.jjjj
'        Datum = Format$(DateSerial(JJ, MM, TT), _
'          "dd.mm.yyyy")
'      End If
'    End If
'  Else
'    Result = True
'    Datum = ""
'  End If
'
'  If Not Result Then Datum = ""
'
'  CheckDatum = Datum
'End Function
'



