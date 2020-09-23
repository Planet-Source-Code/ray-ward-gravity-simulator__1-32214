VERSION 5.00
Begin VB.Form frmbox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Box Properities"
   ClientHeight    =   1350
   ClientLeft      =   8925
   ClientTop       =   7155
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtgrav 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "9.8"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtbounce 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "50"
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Gravity (m/s):"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Energy loss on bounce:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public boxnum As Integer
Private Sub Form_Load()
Dim boxnum As Integer, boxl As Long, boxt As Long, frmw As Long, frmh As Long

'position the form centered on the box
boxnum = frmmain.mainpic(remindex).Tag
boxw = (frmmain.mainpic(remindex).Width / 2) + frmmain.mainpic(remindex).Left
boxh = (frmmain.mainpic(remindex).Height / 2) + frmmain.mainpic(remindex).Top
frmw = Me.Width / 2
frmh = Me.Height / 2
frmbox.Move (boxw - frmw), (boxh - frmh + 300)


'read box properties
txtbounce.Text = Box(boxnum).EnergyLoss
txtgrav.Text = Box(boxnum).Gravity


End Sub

Private Sub Form_Unload(Cancel As Integer)

'set box properties
Box(boxnum).EnergyLoss = txtbounce.Text
Box(boxnum).Gravity = txtgrav.Text

End Sub

Private Sub txtbounce_Change()
If txtbounce.Text < 0 Or txtbounce.Text > 100 Then
    MsgBox "Value must be between 0 and 100"
    txtbounce.Text = Box(boxnum).EnergyLoss
End If

End Sub

Private Sub txtgrav_Change()
If txtgrav.Text < 0 Then
    MsgBox "Value must be greater than or equal to 0"
    txtgrav.Text = Box(boxnum).Gravity
End If
End Sub
