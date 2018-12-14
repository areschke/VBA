VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFPwInput 
   Caption         =   "MANHFABRIK - ADMIN"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3330
   OleObjectBlob   =   "UFPwInput.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UFPwInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub UserForm_Initialize()
  If Val(Application.Version) >= 9 Then
    hwndForm = FindWindow("ThunderDFrame", Me.Caption)
  Else
    hwndForm = FindWindow("ThunderXFrame", Me.Caption)
  End If

  bCloseBtn = False
  SET_USERFORM_STYLE
End Sub

Private Sub SET_USERFORM_STYLE()
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
Private Sub btnOK_Click()
If tbxPW.Value = "123" Then
    MsgBox "Hier kommt noch der Code rein, damit Mahnfabrik bestimmte Felder editieren kann."
    Else:
    tbxPW.Value = ""
    MsgBox "Noch mal"
    tbxPW.SetFocus
End If




'If condition [ Then ]
'    [ statements ]
'[ ElseIf elseifcondition [ Then ]
'    [ elseifstatements ] ]
'[ Else
'    [ elsestatements ] ]
'End If


End Sub

Private Sub btnAbbr_Click()
Unload Me
End Sub

