Attribute VB_Name = "Module1"
Public Type BoxInfo
    Top As Single
    NewTop As Single
    NewBottom As Single
    Left As Single
    Right As Single
    Bottom As Single
    Height As Long
    Width As Long
    Stopped As Boolean
    ControlNum As Integer
    Velocity As Single
    nomove As Boolean
    Time As Long
    EnergyLoss As Byte
    Energy As Single
    OnGround As Boolean
    Gravity As Single
    ResetTop As Single
    ResetLeft As Single
End Type

Public Type MoveOrder
    boxnum As Integer
    Top As Single
End Type

Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long

 


Global Const SND_ASYNC = &H1     ' Play asynchronously
Global Const SND_NODEFAULT = &H2 ' Don't use default sound
Global Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file

Public Box() As BoxInfo
Public Boxmove() As MoveOrder

Public cyc As Long
Public remindex As Integer, nxtbox As Integer
Public numctrl As Integer, bxcount As Integer, picnum As Integer
Public drawbox As Boolean, addbox As Boolean, movepic As Boolean
Public gstrVisibleEverywhere
Global soundfile As String

Sub PlaySound()
On Error Resume Next
Dim Ret As Variant

If frmmain.Check2.Value = 1 Then
    Ret = sndPlaySound(soundfile, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    DoEvents
End If

End Sub
Sub EndSound()
On Error Resume Next
Dim Ret As Variant
If frmmain.Check2.Value = 1 Then
    Ret = sndPlaySound(0&, 0&)
End If
End Sub
Sub Main()

soundfile = StrConv(LoadResData("msound", "sound"), vbUnicode)


Load frmmain
frmmain.WindowState = vbMaximized
'****************************************************
'gradient code by Brian Harper
'create gradient background
Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$

Redval = 0
Blueval = 255
blstep = Blueval / 126

Greenval = 0
Step = (frmmain.Height / 126)
    
FillLeft = 0
FillRight = frmmain.Width
FillBottom = FillTop + Step


frmmain.Show

Redval = 0
Blueval = 255
blstep = Blueval / 126

Greenval = 0
Step = (frmmain.Height / 126)
    
FillLeft = 0
FillRight = frmmain.Width
FillBottom = FillTop + Step

For Reps = 1 To 126
    frmmain.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
    Redval = Redval - 4
    Greenval = Greenval - 4
    Blueval = Blueval - 2
    If Redval <= 0 Then Redval = 0
    If Greenval <= 0 Then Greenval = 0
    If Blueval <= 0 Then Blueval = 0
    FillTop = FillBottom
    FillBottom = FillTop + Step
Next

'****************************************************

'posistion ground
frmmain.Shape1.Top = frmmain.Height - (2 * frmmain.Shape1.Height)
frmmain.Shape1.Width = frmmain.Width
frmmain.Shape1.Left = 0



End Sub


