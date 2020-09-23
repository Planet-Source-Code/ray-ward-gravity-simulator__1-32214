VERSION 5.00
Begin VB.Form frmmenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   11475
   ClientTop       =   1170
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Begin VB.Menu mnuprop 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnurem 
         Caption         =   "&Remove"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuprop_Click()
Load frmbox
frmbox.Show vbModal

End Sub

Private Sub mnurem_Click()
Dim x As Integer, bxar As Integer, bxc As Integer
Dim tempboxes() As BoxInfo

oldtag = frmmain.mainpic(remindex).Tag
'remove the box
Unload frmmain.mainpic(remindex)

bcount = frmmain.mainpic.Count
ReDim tempboxes(bcount)
bxc = 0

numctrl = frmmain.Controls.Count
     
For x = 0 To (numctrl - 1)
    If TypeOf frmmain.Controls(x) Is PictureBox Then
        'If frmmain.Controls(x).Index <> remindex Then
            Debug.Print frmmain.Controls(x).Index
            bxar = frmmain.Controls(x).Tag
            
            'store data still needed in temporary array
            With tempboxes(bxc)
                .Bottom = Box(bxar).Bottom
                Debug.Print Box(bxar).ControlNum
                .ControlNum = x 'this value will change because of the box that has been removed
                .Energy = 0
                .EnergyLoss = Box(bxar).EnergyLoss
                .Gravity = Box(bxar).Gravity
                .Height = Box(bxar).Height
                .Left = Box(bxar).Left
                .NewBottom = 0
                .NewTop = 0
                .nomove = False
                .OnGround = Box(bxar).OnGround
                .ResetLeft = Box(bxar).ResetLeft
                .ResetTop = Box(bxar).ResetTop
                .Right = Box(bxar).Right
                .Stopped = False
                .Time = 0
                .Top = Box(bxar).Top
                .Velocity = Box(bxar).Velocity
                .Width = Box(bxar).Width
            End With
        
            bxc = bxc + 1
        'End If
    End If
Next x


ReDim Box(bxc - 1)

'restore data back in orginal array now with teh exception of the box to be removed
For x = 0 To (bxc - 1)
    With Box(x)
        .Bottom = tempboxes(x).Bottom
        .ControlNum = tempboxes(x).ControlNum
        .Energy = 0
        .EnergyLoss = tempboxes(x).EnergyLoss
        .Gravity = tempboxes(x).Gravity
        .Height = tempboxes(x).Height
        .Left = tempboxes(x).Left
        .NewBottom = 0
        .NewTop = 0
        .nomove = False
        .OnGround = tempboxes(x).OnGround
        .ResetLeft = tempboxes(x).ResetLeft
        .ResetTop = tempboxes(x).ResetTop
        .Right = tempboxes(x).Right
        .Stopped = False
        .Time = 0
        .Top = tempboxes(x).Top
        .Velocity = tempboxes(x).Velocity
        .Width = tempboxes(x).Width
    End With

    frmmain.Controls(Box(x).ControlNum).Tag = x
Next x


'free memory
Erase tempboxes

'current box counter
nxtbox = nxtbox - 1
End Sub

