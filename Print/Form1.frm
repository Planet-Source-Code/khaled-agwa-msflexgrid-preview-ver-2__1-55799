VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "MSFlexGrid Print & Print Preview Demo"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3930
      Left            =   30
      TabIndex        =   0
      Top             =   1905
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   6932
      _Version        =   393216
      Rows            =   5
      Cols            =   3
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3105
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   300
      Width           =   4200
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3105
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   4200
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3105
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1380
      Width           =   4200
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      Height          =   1395
      Left            =   5085
      ScaleHeight     =   1335
      ScaleWidth      =   1365
      TabIndex        =   4
      Top             =   3945
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   240
      Picture         =   "Form1.frx":0000
      Top             =   120
      Width           =   2640
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Print preview"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer
    
    Grid1.Rows = 41
    Grid1.Cols = 6
    
    For i = 1 To 40
    Grid1.TextMatrix(i, 0) = i
    Grid1.TextMatrix(i, 1) = i + 2
    Grid1.TextMatrix(i, 2) = i + 3
    Grid1.TextMatrix(i, 3) = i + 4
    Grid1.TextMatrix(i, 4) = i + 5
    Grid1.TextMatrix(i, 5) = i + 6
  
    Next i
     
        
    Grid1.TextMatrix(0, 0) = "No."
    Grid1.TextMatrix(0, 1) = "      MD"
    Grid1.TextMatrix(0, 2) = " Inclination"
    Grid1.TextMatrix(0, 3) = "   Azimuth"
    Grid1.TextMatrix(0, 4) = "     TVD"
    Grid1.TextMatrix(0, 5) = "     DLS"
    
    Grid1.ColWidth(0) = 400
    Grid1.ColWidth(1) = 860
    Grid1.ColWidth(2) = 860
    Grid1.ColWidth(3) = 860
    Grid1.ColWidth(4) = 860
    Grid1.ColWidth(5) = 860
    
    '******************************************
Grid1.FillStyle = flexFillRepeat
Grid1.CellAlignment = 4
Grid1.FillStyle = flexFillSingle


    Grid1.FillStyle = flexFillRepeat
    Grid1.CellAlignment = 4
    Grid1.FillStyle = flexFillSingle
 Dim lThisRow  As Long
 
   
    Grid1.Row = 1
    Grid1.Col = 1
    SetTextbox
'**************************************
    
End Sub

Private Sub mnuLetter_Click()
    Form3.Text1 = Text1
    Form3.Text2 = Text2
    Form3.Text3 = Text3
    Form3.Show 1, Me
End Sub

Private Sub Form_Resize()
    If Me.Height < 2800 Then Exit Sub
    Grid1.Move 30, 1905, Me.Width - 180, Me.Height - 2750
End Sub

Private Sub Grid11_Click()

End Sub

Private Sub mnuPreview_Click()
    DrawPreview
    Form2.Picture1.Move 0, 0, 11520, 15120
    Form2.Picture1.Picture = picPreview.Image
    Form2.Picture1.Refresh
    Form2.Show 1, Me
End Sub

Private Sub mnuPrint_Click()
    DrawPreview
    On Error GoTo errHandler
    With CommonDialog1
        .CancelError = True
        .Flags = cdlPDReturnIC Or cdlPDReturnDC
        .ShowPrinter
    End With
    
    Printer.PaintPicture picPreview.Image, 0, 0
    Printer.EndDoc
    Exit Sub
    
errHandler:
    MsgBox "Operation cancelled by user.", vbOKOnly, "Cancel"
End Sub

Private Sub DrawPreview()
    Dim i As Integer
    Dim j As Integer
    Dim Grid1Width As Long
    Dim Grid1Height As Long
    Dim Grid1Left As Long
    Dim Grid1Top As Long
    Dim tmp As Long
    Dim tmp2 As Long
    Dim txt As String
    
    'Printer document is 11520 x 12120
    picPreview.Move 0, 0, 11520, 15120
    picPreview.BackColor = vbWhite
    picPreview.ForeColor = vbBlack
    picPreview.Cls
    
    'Paint logo on preview
    picPreview.PaintPicture Image1.Picture, 150, 150
    
    'Print first line
    picPreview.CurrentX = Image1.Width + 300
    picPreview.CurrentY = 430
    picPreview.Print Text1.Text
    
    'Print second line
    picPreview.CurrentX = Image1.Width + 300
    picPreview.CurrentY = 830
    picPreview.Print Text2.Text
    
    'Print third line
    picPreview.CurrentX = Image1.Width + 300
    picPreview.CurrentY = 1230
    picPreview.Print Text3.Text
    
    'Calculate Grid1 width
    Grid1Width = 0
    For i = 0 To Grid1.Cols - 1
        Grid1Width = Grid1Width + Grid1.ColWidth(i)
    Next i
    
    'Calculate Grid1 height
    Grid1Height = 0
    For i = 0 To Grid1.Rows - 1
        Grid1Height = Grid1Height + Grid1.RowHeight(i)
    Next i
    
    'Grid1 position on preview, (change these to move the Grid1)
    Grid1Left = 300
    Grid1Top = Image1.Height + 300
    
    'Draw outer Grid1 frame
    picPreview.Line (Grid1Left, Grid1Top)-(Grid1Left + Grid1Width, Grid1Top + Grid1Height), , B
    
    'Draw horizontal Grid1 lines
    tmp = 0
    For i = 0 To Grid1.Rows - 1
        tmp = tmp + Grid1.RowHeight(i)
        picPreview.Line (Grid1Left, Grid1Top + tmp)-(Grid1Left + Grid1Width, Grid1Top + tmp)
    Next i
    
    'Draw vertical Grid1 lines
    tmp = 0
    For i = 0 To Grid1.Cols - 1
        tmp = tmp + Grid1.ColWidth(i)
        picPreview.Line (Grid1Left + tmp, Grid1Top)-(Grid1Left + tmp, Grid1Top + Grid1Height)
    Next i
    
    'Print MSFlexGrid1 data to the preview Grid1
    tmp = 0
    tmp2 = 0
    For i = 0 To Grid1.Rows - 1
        For j = 0 To Grid1.Cols - 1
            txt = Grid1.TextMatrix(i, j)
            picPreview.CurrentX = Grid1Left + tmp + 90
            picPreview.CurrentY = Grid1Top + tmp2 + 30
            picPreview.Print txt
            tmp = tmp + Grid1.ColWidth(j)
        Next j
        tmp = 0
        tmp2 = tmp2 + Grid1.RowHeight(i)
    Next i
    picPreview.Refresh
End Sub




'*********************************************
'******code required for making msflexgrid Editablte****
Sub NumberCells()
Dim i As Integer

    For i = 1 To Grid.Rows - 1
        Grid1.TextMatrix(0, i) = Format$(i, "000")
    Next
    For i = 1 To Grid1.Cols - 1
        Grid1.TextMatrix(i, 0) = " " & Format$(i, "000")
    Next
   
End Sub

Private Sub grid1_EnterCell()
' Make sure the user doesn't attempt to edit the fixed cells
    If Grid1.MouseRow = 0 Or Grid1.MouseCol = 0 Then
        Text.Visible = False
        Exit Sub
    End If
' clear contents of current cell
    Text.Text = ""
' place Textbox over current cell
    Text.Visible = False
    Text.Top = Grid1.Top + Grid1.CellTop
    Text.Left = Grid1.Left + Grid1.CellLeft
    Text.Width = Grid1.CellWidth
    Text.Height = Grid1.CellHeight
' assing cell's contents to Textbox
    Text.Text = Grid1.Text
' move focus to Textbox
    Text.Visible = True
    Text.SetFocus
End Sub

Private Sub grid1_LeaveCell()
    Grid1.Text = Text.Text
End Sub
Private Sub text_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Grid1.Row = Grid1.Rows - 1 Then
            If Grid1.Col = Grid1.Cols - 1 Then
                Exit Sub
            Else
                Grid1.Col = Grid1.Col + 1
            End If
            Grid1.Row = 1
        Else
            Grid1.Row = Grid1.Row + 1
        End If
    End If
End Sub

Sub SetTextbox()
    Text.Visible = False
    Text.Top = Grid1.Top + Grid1.CellTop
    Text.Left = Grid1.Left + Grid1.CellLeft
    Text.Height = Grid1.CellHeight
    Text.Width = Grid1.CellWidth
    Text.Text = Grid1.Text
    Text.Visible = True
End Sub

Private Sub EditSelect_Click()
    Grid1.Row = 1
    Grid1.Col = 1
    Grid1.RowSel = Grid1.Rows - 1
    Grid1.ColSel = Grid1.Cols - 1
End Sub








