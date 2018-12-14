Attribute VB_Name = "modFunctions"
Option Explicit

'=============================================================================================================================
' Name:     Kalender-Formular
' Zweck:    Alternative zum Kalender-Steuerelement "Calendar Control 2007" aus MSCOMCTL.OCX, da unter Office 2010 x64 nicht mehr per default unterstützt
' Author:   Edgar Frei, clearByte GmbH
' Datum:    17.03.2016
' Version:  1.2
' Lizenz:   Zur freien Verwendung. Dieser Header muss belassen werden. Über Feedback freue ich mich: edgar.frei@clearbyte.ch
'=============================================================================================================================
' Änderungen in der Version 1.2:
'   Bug behoben (Beim Klick auf den 29.2.2016 wechselt der Kalender zum Monat März, anstatt den 29.02. auszuwählen).
'   Änderung in der Prozedur "Sub lblk5d1_Click()" in der fälschlicherweise der Wert von lblk5d2.Caption anstatt lblk5d1.Caption abgefragt wurde.
'=============================================================================================================================
' Informationen zur Verwendung:
'   Das Formular besteht aus dem Modul "modFunctions" und der UserForm "frmCalendar" und kommt gänzlich ohne DLLs und OCXs aus.
'   Um ein Datum auszuwählen, muss lediglich das Formular aufgerufen werden. Nach Auswahl eines Datums wird das gewählte
'   Datum in eine globale Variable (g_datCalendarDate) geschrieben, welche abgefragt werden kann.
'=============================================================================================================================
' Implementation:
'   1. Import Modul 'modFunctions'
'   2. Import UserForm 'modCalendar'
'   3. Aufruf via:      frmCalendar.Show
'   4. Rückgabewert:    g_datCalendarDate
'=============================================================================================================================


' Public Variables
Public g_datCalendarDate As Date
Public g_bolInitialize As Boolean
Public g_bolMonthChange As Boolean

'=============================================================================================================================
' Index
'-----------------------------------------------------------------------------------------------------------------------------
' 1. fPlaus - Plausibilitätsprüfungen
' 2. fSetMonthText - Umwandlung Ganzzahl zu Monatsname
' 3. fChangeStrToInt - Umwandlung Monatsname zu Ganzzahl
' 4. fGetKW - Berechnung Kalenderwoche pro Datum
' 5. fLastDayInMonth - Suche des letzten Tages eines Monats


'=============================================================================================================================
'Functions
'=============================================================================================================================
' 1. fPlaus - Plausibilitätsprüfungen
'-----------------------------------------------------------------------------------------------------------------------------
' Eingabeparameter:
'        strPlaus               -   String - Monatsname
'        intPlausType       -  Integer - Plausibilitätstyp (Mehrere Plausibilitätsprüfungen via unterschiedlichen Codes (intPlausType). Gegenwärtig nur Typ 1 (Manuelle Eingabe Monatsname) implementiert).
'-----------------------------------------------------------------------------------------------------------------------------
Function fPlaus(strPlaus, intPlausType) As Boolean
On Error GoTo err_fPlaus
   With frmCalendar
        Select Case intPlausType
            Case 1 'Monatsplaus
                    If .cmbMonth.Text = "Januar" Or _
                        .cmbMonth.Text = "Februar" Or _
                        .cmbMonth.Text = "März" Or _
                        .cmbMonth.Text = "April" Or _
                        .cmbMonth.Text = "Mai" Or _
                        .cmbMonth.Text = "Juni" Or _
                        .cmbMonth.Text = "Juli" Or _
                        .cmbMonth.Text = "August" Or _
                        .cmbMonth.Text = "September" Or _
                        .cmbMonth.Text = "Oktober" Or _
                        .cmbMonth.Text = "November" Or _
                        .cmbMonth.Text = "Dezember" Then
                        fPlaus = True
                    End If
        End Select
    End With
Exit Function

' Errorhandling
err_fPlaus:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'fPlaus' in 'modFunctions'. Plausibilitätsprüfung konnte nicht vollzogen werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Function
'=============================================================================================================================
' 2. fSetMonthText - Umwandlung Ganzzahl zu Monatsname
' Eingabeparameter:
'        varMonth -   Integer - Monatsnummer
'-----------------------------------------------------------------------------------------------------------------------------
Function fSetMonthText(varMonth)
On Error GoTo err_fSetMonthText
        With frmCalendar
        
            Select Case varMonth
                Case 1
                    .cmbMonth.Text = "Januar"
                Case 2
                    .cmbMonth.Text = "Februar"
                Case 3
                    .cmbMonth.Text = "März"
                Case 4
                    .cmbMonth.Text = "April"
                Case 5
                    .cmbMonth.Text = "Mai"
                Case 6
                    .cmbMonth.Text = "Juni"
                Case 7
                    .cmbMonth.Text = "Juli"
                Case 8
                    .cmbMonth.Text = "August"
                Case 9
                    .cmbMonth.Text = "September"
                Case 10
                    .cmbMonth.Text = "Oktober"
                Case 11
                    .cmbMonth.Text = "November"
                Case 12
                    .cmbMonth.Text = "Dezember"
            End Select
            
        End With
Exit Function

' Errorhandling
err_fSetMonthText:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'fSetMonthText' in 'modFunctions'. Umwandlung Monatsname zu Monatsnummer konnte nicht verarbeitet werden.. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Function

'=============================================================================================================================
' 3. fChangeStrToInt - Umwandlung Monatsname zu Ganzzahl
' Eingabeparameter:
'        varMonth -   String -  Monatsname
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function fChangeStrToInt(varMonth)
    Select Case varMonth
                Case "Januar"
                    fChangeStrToInt = 1
                Case "Februar"
                    fChangeStrToInt = 2
                Case "März"
                    fChangeStrToInt = 3
                Case "April"
                    fChangeStrToInt = 4
                Case "Mai"
                    fChangeStrToInt = 5
                Case "Juni"
                    fChangeStrToInt = 6
                Case "Juli"
                    fChangeStrToInt = 7
                Case "August"
                    fChangeStrToInt = 8
                Case "September"
                    fChangeStrToInt = 9
                Case "Oktober"
                    fChangeStrToInt = 10
                Case "November"
                    fChangeStrToInt = 11
                Case "Dezember"
                    fChangeStrToInt = 12
            End Select
Exit Function

' Errorhandling
err_fChangeStrToInt:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'fChangeStrToInt' in 'modFunctions'. Umwandlung Monatsnummer zu Monatsname konnte nicht verarbeitet werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Function


'=============================================================================================================================
' 4. fGetKW - Berechnung Kalenderwoche pro Datum
' Eingabeparameter:
'        d -   Datum
'-----------------------------------------------------------------------------------------------------------------------------
Function fGetKW(datKW As Date) As Integer
Dim datTemp As Date
On Error GoTo err_fGetKW
    datTemp = DateSerial(Year(datKW + (8 - Weekday(datKW)) Mod 7 - 3), 1, 1)
    fGetKW = (datKW - datTemp - 3 + (Weekday(datTemp) + 1) Mod 7) \ 7 + 1
Exit Function

' Errorhandling
err_fGetKW:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'fGetKW' in 'modFunctions'. Ermittlung Kalenderwoche fehlgeschlagen. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Function


'=============================================================================================================================
' 5. fLastDayInMonth - Suche des letzten Tages eines Monats
' Eingabeparameter:
'        dtmDate -   Datum - Optional (wenn leer dtmDate = Heute)
'-----------------------------------------------------------------------------------------------------------------------------
Function fLastDayInMonth(Optional dtmDate As Date = 0) As Date
On Error GoTo err_fLastDayInMonth
    If dtmDate = 0 Then
        dtmDate = Date
    End If
     fLastDayInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 0)
Exit Function

' Errorhandling
err_fLastDayInMonth:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'fLastDayInMonth' in 'modFunctions'. Ermittlung letzter Tag des Monats fehlgeschlagen. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Function

