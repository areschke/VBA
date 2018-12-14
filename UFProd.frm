VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFProd 
   Caption         =   "Modulbuchung"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13380
   OleObjectBlob   =   "UFProd.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "UFProd"
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

Private Sub cBnProdBck_Click()

Call ProdSelect
Call ProdClose
Unload Me


End Sub

Private Sub UserForm_Activate()

    cBxProd2.Value = Checked

End Sub

Private Sub cBnProdCan_Click()

Call ProdClose
Unload Me




End Sub

Private Sub ProdSelect()                                                                ' Für jedes gebuchte Modul wird der Wert '1' hinterlegt
    
    If cBxProd1 = True Then UFDataUpload.i1 = 1 Else: UFDataUpload.i1 = ""
    If cBxProd2 = True Then UFDataUpload.i2 = 1 Else: UFDataUpload.i2 = ""
    If cBxProd3 = True Then UFDataUpload.i3 = 1 Else: UFDataUpload.i3 = ""
    If cBxProd4 = True Then UFDataUpload.i4 = 1 Else: UFDataUpload.i4 = ""
    If cBxProd5 = True Then UFDataUpload.i5 = 1 Else: UFDataUpload.i5 = ""
'    If cBxProd6 = True Then UFDataUpload.i6 = 1 Else: UFDataUpload.i6 = ""


End Sub

Private Sub ProdClose()

    UFDataUpload.bTnCopy.Visible = True
    UFDataUpload.bTnCopy.Enabled = True
    UFDataUpload.bTnAdd.Visible = True
    UFDataUpload.bTnAdd.Enabled = True
    UFDataUpload.bTnDel.Visible = True
    UFDataUpload.bTnDel.Enabled = True
    UFDataUpload.bTnSav.Visible = True
    UFDataUpload.bTnSav.Enabled = True
    UFDataUpload.lBlModule.Visible = True
    UFDataUpload.bTnBck.Visible = True
    UFDataUpload.bTnBck.Enabled = True
    UFDataUpload.bTnBB.Visible = True
    
    UFDataUpload.tBx14.Enabled = False
    UFDataUpload.ListBox1.SetFocus
    If UFDataUpload.ListBox1.ListCount > 0 Then UFDataUpload.ListBox1.ListIndex = 0           ' 1. Eintrag selektieren
End Sub
