VERSION 5.00
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   Caption         =   "Gravity Simulator - Ray Ward"
   ClientHeight    =   8700
   ClientLeft      =   3345
   ClientTop       =   2265
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   12225
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Add Box"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   288
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2172
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   720
   End
   Begin VB.PictureBox mainpic 
      Height          =   495
      Index           =   0
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of cycles of timer:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10920
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   -120
      Top             =   8400
      Width           =   12372
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public stx As Integer, sty As Integer, picl As Long, pict As Long


Private Sub Check1_Click()
addbox = True
Me.MousePointer = vbCrosshair
End Sub

Private Sub cmdreset_Click()
Timer1.Enabled = False

numctrl = frmmain.Controls.Count - 1

bcount = -1
For b = 0 To numctrl
    If TypeOf frmmain.Controls(b) Is PictureBox Then
        'recall picture box positions from array
        bcount = bcount + 1
        MoveBox Box(bcount).ResetTop, Box(bcount).ResetLeft, Box(bcount).ControlNum
        
        z = Box(bcount).ControlNum
        Box(bcount).Top = frmmain.Controls(z).Top
        Box(bcount).Height = frmmain.Controls(z).Height
        Box(bcount).Bottom = Box(bcount).Top + Box(bcount).Height
        Box(bcount).Left = frmmain.Controls(z).Left
        Box(bcount).Width = frmmain.Controls(z).Width
        Box(bcount).Right = Box(bcount).Left + Box(bcount).Width
        Box(bcount).NewBottom = 0
        Box(bcount).NewTop = 0
        Box(bcount).Stopped = False
        Box(bcount).OnGround = False
        Box(bcount).Velocity = 0
        Box(bcount).Energy = 0
        Box(bcount).ResetTop = frmmain.Controls(z).Top
        Box(bcount).ResetLeft = frmmain.Controls(z).Left
        
    End If
Next b


Text1.Text = ""

cmdstart.Caption = "&Start"
cmdreset.Enabled = False
Check1.Enabled = True
frmmenu.mnuprop.Enabled = True



End Sub

'Gravity Simulator by Ray Ward.

Private Sub cmdstart_Click()


If cmdstart.Caption = "&Start" Then
        
    cmdreset.Enabled = True
    frmmenu.mnuprop.Enabled = False
    
    Timer1.Enabled = False
    numctrl = frmmain.Controls.Count - 1
    'loop through each picture box on the form
    ReDim posx(numctrl)
    ReDim posy(numctrl)
    
    picnum = 0
    
    For z = 0 To numctrl
        If TypeOf frmmain.Controls(z) Is PictureBox Then
            picnum = picnum + 1
            
            Randomize
            'set the colours of the boxes to random colours
            col1 = Int((255 - 0 + 1) * Rnd + 0)
            col2 = Int((255 - 0 + 1) * Rnd + 0)
            col3 = Int((255 - 0 + 1) * Rnd + 0)
            numctrl1 = numctrl1 + 1 'number of boxes on screen
            frmmain.Controls(z).BackColor = RGB(col1, col2, col3)
        End If
    Next z
        
    ReDim Boxmove(picnum - 1)
    
    lasttop = Me.Height
 
    'Determine which box has the lowest 'top' so u dont get errors, and boxes overlapping each other when they move
    'Can then move the boxes in order from lowest to highest.
    For v = 0 To (picnum - 1)
        For x = 0 To (picnum - 1)
            If Box(x).Top > Boxmove(v).Top And Box(x).Top < lasttop Then
                Boxmove(v).Top = Box(x).Top
                Boxmove(v).boxnum = x
            ElseIf Box(x).Top = lasttop And x <> lastctrl Then
                For n = 0 To v
                    If Box(n).Top = Box(x).Top And Box(n).ControlNum = x Then
                        cnt = cnt + 1
                    End If
                Next n
                If cnt = 0 Then
                    Boxmove(v).Top = Box(x).Top
                    Boxmove(v).boxnum = x
                End If
            End If
        Next x
    
        lastctrl = Boxmove(v).boxnum
        lasttop = Boxmove(v).Top
    Next v
        
       
    nxtbox = mainpic.Count
       
    'set variables
    bxcount = 0
    cyc = 0
    
    Timer1.Enabled = True
    cmdstart.Caption = "&Pause"
    Check1.Enabled = False
ElseIf cmdstart.Caption = "&Pause" Then
    Timer1.Enabled = False
    cmdstart.Caption = "R&esume"
ElseIf cmdstart.Caption = "R&esume" Then
    Timer1.Enabled = True
    cmdstart.Caption = "&Pause"
End If
End Sub





Private Sub Form_Load()
Dim tmpbx As Integer, z As Integer

'load up a few picture boxes
Load mainpic(1)
mainpic(1).Visible = True
mainpic(1).Move 6000, 6000, 1000, 1000

Load mainpic(2)
mainpic(2).Visible = True
mainpic(2).Move 15000, 1000, 500, 6000

Load mainpic(3)
mainpic(3).Visible = True
mainpic(3).Move 9000, 100, 2000, 2000

ReDim Box(3)

tmpbx = -1
numctrl = frmmain.Controls.Count - 1
For z = 0 To numctrl
    If TypeOf frmmain.Controls(z) Is PictureBox Then
       'set colour of boxes to random
        Randomize
        col1 = Int((255 - 0 + 1) * Rnd + 0)
        Randomize
        col2 = Int((255 - 0 + 1) * Rnd + 0)
        Randomize
        col3 = Int((255 - 0 + 1) * Rnd + 0)
        frmmain.Controls(z).BackColor = RGB(col1, col2, col3)
        tmpbx = tmpbx + 1
        
        'set initial properties of default boxes
        Box(tmpbx).ControlNum = z
        Box(tmpbx).Top = frmmain.Controls(z).Top
        Box(tmpbx).Height = frmmain.Controls(z).Height
        Box(tmpbx).Bottom = Box(tmpbx).Top + Box(tmpbx).Height
        Box(tmpbx).Left = frmmain.Controls(z).Left
        Box(tmpbx).Width = frmmain.Controls(z).Width
        Box(tmpbx).Right = Box(tmpbx).Left + Box(tmpbx).Width
        Box(tmpbx).NewBottom = 0
        Box(tmpbx).NewTop = 0
        Box(tmpbx).Stopped = False
        Box(tmpbx).OnGround = False
        Box(tmpbx).EnergyLoss = 100
        Box(tmpbx).Gravity = 9.8
        Box(tmpbx).ResetTop = frmmain.Controls(z).Top
        Box(tmpbx).ResetLeft = frmmain.Controls(z).Left
        
        frmmain.Controls(z).Tag = tmpbx
    
    End If
Next z


Box(0).EnergyLoss = 20
Box(1).EnergyLoss = 0
Box(2).EnergyLoss = 100
Box(3).EnergyLoss = 50

nxtbox = mainpic.Count

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo handler

If addbox = True Then
    stx = x
    sty = y
    
    CreateBox x, y
    
    addbox = False
    drawbox = True
    
End If

Exit Sub
handler:
'nxtbox = nxtbox + 1
Resume

End Sub
Sub CreateBox(x As Single, y As Single)
Dim z As Integer
On Error Resume Next

 
Load mainpic(nxtbox)
mainpic(nxtbox).Visible = True
mainpic(nxtbox).Move x, y

ReDim Preserve Box(nxtbox)

numctrl = frmmain.Controls.Count

For z = 0 To (numctrl - 1)
    If TypeOf frmmain.Controls(z) Is PictureBox Then
        If frmmain.Controls(z).Index = nxtbox Then
            
            'set initial properties
            Box(nxtbox).ControlNum = z
            Box(nxtbox).Top = frmmain.Controls(z).Top
            Box(nxtbox).Height = frmmain.Controls(z).Height
            Box(nxtbox).Bottom = Box(nxtbox).Top + Box(nxtbox).Height
            Box(nxtbox).Left = frmmain.Controls(z).Left
            Box(nxtbox).Width = frmmain.Controls(z).Width
            Box(nxtbox).Right = Box(nxtbox).Left + Box(nxtbox).Width
            Box(nxtbox).NewBottom = 0
            Box(nxtbox).NewTop = 0
            Box(nxtbox).Stopped = False
            Box(nxtbox).OnGround = False
            Box(nxtbox).EnergyLoss = 100
            Box(nxtbox).Gravity = 9.8
            Box(nxtbox).ResetTop = frmmain.Controls(z).Top
            Box(nxtbox).ResetLeft = frmmain.Controls(z).Left
        
            mainpic(nxtbox).Tag = nxtbox
        
        End If
    End If
Next z

'increment box counter
nxtbox = nxtbox + 1


End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If drawbox = True Then
    mainpic(nxtbox - 1).Move stx, sty, (x - stx), (y - sty)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If drawbox = True Then
    drawbox = False
    Check1.Value = 0
    addbox = False
    Me.MousePointer = vbDefault
End If


End Sub


Private Sub Form_Resize()
Dim z As Integer

numctrl = frmmain.Controls.Count - 1
For z = 0 To numctrl
    If TypeOf frmmain.Controls(z) Is TextBox Then
        If frmmain.Controls(z).Name <> "txtgrav" Then
            frmmain.Controls(z).Left = (Me.Width - frmmain.Controls(z).Width - 300)
        End If
    End If
Next z
Label3.Left = (Me.Width - Label3.Width - 300)
End Sub

Private Sub Form_Unload(Cancel As Integer)
EndSound
End Sub


Private Sub mainpic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button = vbLeftButton Then
    mainpic(Index).MousePointer = vbSizeAll
    movepic = True
    picl = x
    pict = y
        
    Box(mainpic(Index).Tag).nomove = True
End If
End Sub

Private Sub mainpic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If movepic = True Then
    'position the box
    mainpic(Index).Top = mainpic(Index).Top + y - pict
    mainpic(Index).Left = mainpic(Index).Left + x - picl
End If


If Timer1.Enabled = False Then
    'when timer is disabled and box is being moved change the 'Reset' values
    boxnum = mainpic(Index).Tag
    Box(boxnum).ResetLeft = mainpic(Index).Left
    Box(boxnum).ResetTop = mainpic(Index).Top
End If

End Sub

Private Sub mainpic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button = vbLeftButton Then
    mainpic(Index).MousePointer = vbDefault
    movepic = False
    Box(mainpic(Index).Tag).nomove = False
    Box(mainpic(Index).Tag).Stopped = False
    Box(mainpic(Index).Tag).Velocity = 0
Else
    remindex = Index
    PopupMenu frmmenu.mnu
End If

End Sub
Sub MoveBox(ByVal BoxTop As Long, ByVal BoxLeft As Long, ByVal ControlNo As Integer)

frmmain.Controls(ControlNo).Top = BoxTop
frmmain.Controls(ControlNo).Left = BoxLeft

End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Dim oldmvtop As Single, newmvtop As Single, q As Integer

'All the actual work done here to move boxes, detect collisions etc :)"

cyc = cyc + 1
q = 0

picnum = mainpic.Count
'loop through ever picture box on the screen and move it accordingly
For loopbox = 0 To (picnum - 1)
    b = Boxmove(loopbox).boxnum
    
    CN = Box(b).ControlNum
    If Box(b).Stopped = False And Box(b).nomove = False Then
        Box(b).Time = Box(b).Time + 1
        tme = Box(b).Time
        
        Box(b).Top = frmmain.Controls(CN).Top
        Box(b).Height = frmmain.Controls(CN).Height
        Box(b).Bottom = Box(b).Top + Box(b).Height
        
        If Box(b).Velocity = 0 And Box(b).Bottom = Shape1.Top Then
            Box(b).Stopped = True
            GoTo nextbox
        End If
       
        Box(b).Energy = ((Box(b).Velocity ^ 2) / 2)
        
        dist = Box(b).Velocity + (Box(b).Gravity / 2)
                   
        Box(b).NewTop = Box(b).Top + dist
        Box(b).NewBottom = Box(b).Top + Box(b).Height + dist
        Box(b).Left = frmmain.Controls(CN).Left
        Box(b).Width = frmmain.Controls(CN).Width
        Box(b).Right = Box(b).Left + Box(b).Width
        
        newmvtop = 0
        oldmvtop = frmmain.Height
        mvbxnum = -1
        
        'loop through other boxes on screen to see which box it will hit and if that box is actually on the ground.
        For q = 0 To (picnum - 1)
            'test to see if box has already stopped moving
            If q <> b And Box(q).Stopped = True Then
                If Box(b).Left < Box(q).Right And Box(b).Right > Box(q).Left And Box(b).NewBottom >= Box(q).Top And Box(b).Bottom <= Box(q).Top Then
                    newmvtop = Box(q).Top
                    'check if there are any other stopped boxes between 'b' and 'q'
                    If (newmvtop < oldmvtop) Then
                        mvbxnum = q
                        oldmvtop = Box(q).Top
                    End If
                End If
            End If
        Next q
                               
        'if any boxes need to be placed on top of other boxes...
        If mvbxnum >= 0 Then
            If Box(mvbxnum).Stopped = True Then
               'distance box will travel from current position to box underneath
                ndist = Box(mvbxnum).Top - Box(b).Bottom
                'move box to top of box beneath
                MoveBox (Box(mvbxnum).Top - Box(b).Height), Box(b).Left, Box(b).ControlNum
                
                'if has really low energy then dont worry about continueing bounce and waste cpu time
                If Box(b).Energy < 1 Then
                    Box(b).Energy = 0
                    Box(b).Velocity = 0
                    Box(b).Stopped = True
                Else
                    Box(b).Velocity = Sqr((Box(b).Velocity ^ 2) + 2 * Box(b).Gravity * ndist)
                    'also update energy
                    Box(b).Energy = (((100 - Box(b).EnergyLoss) / 100) * ((Box(b).Velocity ^ 2) / 2))
                    'reverse velocity for upwards motion because of bounce
                    Box(b).Velocity = -1 * Sqr(2 * Box(b).Energy)
                End If
                
                Box(b).OnGround = True
                
                'play ground hit sound
                soundfile = StrConv(LoadResData("gsound", "sound"), vbUnicode)
                PlaySound
              
                GoTo nextbox 'save time by skipping other statements to process the next box
            End If
        End If

        'if box aint gonna hit another box but hit the ground instead
        If Box(b).Bottom <= Shape1.Top And Box(b).NewBottom >= Shape1.Top Then
            'distance box will travel from current position to ground
            ndist = Shape1.Top - Box(b).Bottom
            'move box to ground
            MoveBox (Shape1.Top - Box(b).Height), Box(b).Left, CN
            
            If Box(b).Energy < 1 Then
                Box(b).Energy = 0
                Box(b).Velocity = 0
                Box(b).Stopped = True
            Else
                Box(b).Velocity = Sqr((Box(b).Velocity ^ 2) + 2 * Box(b).Gravity * ndist)
                Box(b).Energy = (((100 - Box(b).EnergyLoss) / 100) * ((Box(b).Velocity ^ 2) / 2))
                Box(b).Velocity = -1 * Sqr(2 * Box(b).Energy)
            End If
               
            Box(b).OnGround = True
            
            'play ground hit sound
            soundfile = StrConv(LoadResData("gsound", "sound"), vbUnicode)
            PlaySound
        Else 'if its still moving and aint gonna hit anything at all
            MoveBox Box(b).NewTop, Box(b).Left, CN
                      
            If Box(b).Velocity < 0 And (Box(b).Gravity + Box(b).Velocity) > 0 Then
                Box(b).Velocity = 0
            Else
                Box(b).Velocity = Box(b).Gravity + Box(b).Velocity
            End If
        End If
    End If

nextbox:
    
Next loopbox

Text1.Text = cyc

End Sub

