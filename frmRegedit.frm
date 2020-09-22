VERSION 5.00
Begin VB.Form frmRegedit 
   Caption         =   "VBRegedit"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   4545
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmRegedit.frx":0000
      Left            =   120
      List            =   "frmRegedit.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblPath 
      Caption         =   "Path:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   7695
   End
   Begin VB.Menu mnuString 
      Caption         =   "String"
      Visible         =   0   'False
      Begin VB.Menu mnuNewString 
         Caption         =   "New String"
      End
      Begin VB.Menu mnuDelString 
         Caption         =   "Delete String"
      End
   End
   Begin VB.Menu mnuKey 
      Caption         =   "Key"
      Visible         =   0   'False
      Begin VB.Menu mnuNewKey 
         Caption         =   "New Key"
      End
      Begin VB.Menu mnuDelKey 
         Caption         =   "Delete Key"
      End
   End
End
Attribute VB_Name = "frmRegedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nPath As String 'Normal Path
Dim hPath As Long 'HKEY_Path
Sub UpdateMe()

List1.Clear
List2.Clear
List3.Clear
If InStr(1, nPath, "\") <> 0 And nPath <> "\" Then List1.AddItem ".."

If Len(nPath) > 0 Then
  tempPath = Right(nPath, Len(nPath) - 1)
Else
  tempPath = ""
End If

Call GetKeys(hPath, tempPath, Me.List1)
Call GetValues(hPath, tempPath, Me.List2)

For op = 0 To List2.ListCount - 1
  asy = GetSettingString(hPath, tempPath, List2.List(op))
  List3.AddItem asy
Next op
lblPath.Caption = Combo1.List(Combo1.ListIndex) & nPath


End Sub

Private Sub Combo1_Click()
hPath = Combo1.ItemData(Combo1.ListIndex)
nPath = ""
lblPath.Caption = Combo1.List(Combo1.ListIndex)

List1.Clear
List2.Clear
List3.Clear
Call GetKeys(hPath, "", Me.List1)
Call GetValues(hPath, "", Me.List2)

For op = 0 To List2.ListCount - 1
  asy = GetSettingString(hPath, "", List2.List(op))
  List3.AddItem asy
Next op
End Sub


Private Sub Form_Load()
Combo1.ItemData(0) = HKEY_CLASSES_ROOT
Combo1.ItemData(1) = HKEY_CURRENT_USER
Combo1.ItemData(2) = HKEY_LOCAL_MACHINE
Combo1.ItemData(3) = HKEY_USERS
Combo1.ItemData(4) = HKEY_CURRENT_CONFIG
Combo1.ItemData(5) = HKEY_DYN_DATA
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.Width < 5505 Then Me.Width = 5505
If Me.Height < 1725 Then Me.Height = 1725
List1.Width = Me.Width / 3 + 220
Combo1.Width = List1.Width

List2.Left = List1.Width + 220 + 120
List2.Width = (Me.Width - 1100) / 3
List3.Width = List2.Width
List3.Left = List2.Left + List2.Width
List2.Height = Me.Height - 800
List3.Height = List2.Height
List1.Height = List2.Height - 240
lblPath.Top = Me.Height - 680
lblPath.Width = Me.Width - 240
End Sub

Private Sub List1_DblClick()
If List1.List(List1.ListIndex) = ".." Then
  goUp
Else
 If Right(nPath, 1) = "\" Then
  nPath = nPath & List1.List(List1.ListIndex)
 Else
  nPath = nPath & "\" & List1.List(List1.ListIndex)
 End If
End If

UpdateMe
End Sub

Private Sub goUp()
Dim NewSt, OldSt As Integer
Do
  OldSt = NewSt
  NewSt = InStr(OldSt + 1, nPath, "\")
Loop Until NewSt = 0
If OldSt = 0 Then
  nPath = "\"
Else
  nPath = Left(nPath, OldSt - 1)
End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  If List1.ListIndex = -1 Then
     mnuDelKey.Enabled = False
  Else
     mnuDelKey.Enabled = True
  End If
   PopupMenu mnuKey
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List3.ListIndex = List2.ListIndex
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
List3.ListIndex = List2.ListIndex
If Button = vbRightButton Then
   If List1.ListIndex = -1 Then
     mnuDelString.Enabled = False
   Else
     mnuDelString.Enabled = True
   End If
   PopupMenu mnuString
End If
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List2.ListIndex = List3.ListIndex
End Sub

Private Sub List3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
List2.ListIndex = List3.ListIndex
End Sub

Private Sub mnuDelKey_Click()
Dim vPath As String
If MsgBox("Are you sure you want to delete this Key?", vbExclamation + vbYesNo, "Delete?") = vbYes Then
If List1.ListIndex < 0 Then Exit Sub
If Left(nPath, 1) = "\" Then
vPath = Right(nPath, Len(nPath) - 1) & "\" & List1.List(List1.ListIndex)
ElseIf nPath = "" Then
vPath = List1.List(List1.ListIndex)
Else
vPath = nPath & "\" & List1.List(List1.ListIndex)
End If
DeleteKey hPath, vPath
UpdateMe
End If
End Sub

Private Sub mnuDelString_Click()
If MsgBox("Are you sure you want to delete this String?", vbExclamation + vbYesNo, "Delete?") = vbYes Then
DeleteSettingString hPath, nPath, List2.List(List2.ListIndex)
UpdateMe
End If
End Sub

Private Sub mnuNewKey_Click()
Dim NewKey, vPath As String
NewKey = InputBox("Enter new Key's name:", "New Key", "NewKey")
If NewKey = vbNullString Then Exit Sub
If Right(nPath, 1) = "\" Then
   vPath = nPath & NewKey
Else
   vPath = nPath & "\" & NewKey
End If
CreateKey hPath, vPath
UpdateMe
End Sub

Private Sub mnuNewString_Click()
Dim NewString, NewValue As String
NewString = InputBox("Enter new Strings's name:", "New String", "NewString")
NewValue = InputBox("Enter new Strings's Value:", "New Value", "NewValue")
If NewString = vbNullString Then Exit Sub
SaveSettingString hPath, nPath, NewString, NewValue
UpdateMe
End Sub
