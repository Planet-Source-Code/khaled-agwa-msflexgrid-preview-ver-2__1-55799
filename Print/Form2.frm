VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Print preview"
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   4230
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   240
      Left            =   5595
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   3990
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   15
      Left            =   15
      TabIndex        =   2
      Top             =   3990
      Visible         =   0   'False
      Width           =   5580
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3990
      LargeChange     =   15
      Left            =   5595
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3900
      Left            =   30
      ScaleHeight     =   3840
      ScaleWidth      =   5430
      TabIndex        =   0
      Top             =   60
      Width           =   5490
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
    'Picture1.Width = Me.ScaleWidth
    'Picture1.Height = Me.ScaleHeight
    
    VScroll1.Visible = False
    HScroll1.Visible = False
    Picture2.Visible = False
    
    If Picture1.Width > Me.ScaleWidth Then
        HScroll1.Visible = True
        HScroll1.Move 0, Me.ScaleHeight - 240, Me.ScaleWidth, 240
    End If
    If Picture1.Height > Me.ScaleHeight Then
        VScroll1.Visible = True
        If HScroll1.Visible Then
            HScroll1.Move 0, Me.ScaleHeight - 240, Me.ScaleWidth - 240, 240
            VScroll1.Move Me.ScaleWidth - 240, 0, 240, Me.ScaleHeight - 240
            Picture2.Move Me.ScaleWidth - 240, Me.ScaleHeight - 240, 240, 240
            Picture2.Visible = True
        Else
            VScroll1.Move Me.ScaleWidth - 240, 0, 240, Me.ScaleHeight
        End If
    End If
    If VScroll1.Visible Then VScroll1.Max = (Me.ScaleHeight + (240 * HScroll1.Visible)) - Picture1.Height
    If HScroll1.Visible Then HScroll1.Max = (Me.ScaleWidth + (240 * VScroll1.Visible)) - Picture1.Width
End Sub

Private Sub HScroll1_Change()
    HorizScroll
End Sub

Private Sub HScroll1_Scroll()
    HorizScroll
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub VScroll1_Change()
    VertScroll
End Sub

Private Sub VScroll1_Scroll()
    VertScroll
End Sub

Private Sub HorizScroll()
    Picture1.Left = HScroll1.Value
End Sub

Private Sub VertScroll()
    Picture1.Top = VScroll1.Value
End Sub
