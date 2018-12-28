VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command25 
      Caption         =   "Disable"
      Height          =   1095
      Left            =   8280
      TabIndex        =   39
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Enable"
      Height          =   1095
      Left            =   6720
      TabIndex        =   38
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox richtextbox 
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      TabIndex        =   37
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3600
      Top             =   8760
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H8000000D&
      Caption         =   "Stop"
      Height          =   615
      Left            =   4800
      TabIndex        =   36
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H8000000D&
      Caption         =   "Start"
      Height          =   615
      Left            =   3120
      MaskColor       =   &H00FF0000&
      TabIndex        =   35
      Top             =   7920
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   10200
      TabIndex        =   32
      Top             =   8040
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   10200
      TabIndex        =   31
      Top             =   7200
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   10200
      TabIndex        =   30
      Top             =   6720
      Width           =   3615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Display Date"
      Height          =   375
      Left            =   10320
      TabIndex        =   29
      Top             =   5880
      Width           =   3375
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   10320
      TabIndex        =   28
      Top             =   5160
      Width           =   3375
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Square Root"
      Height          =   495
      Left            =   10320
      TabIndex        =   27
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   10320
      TabIndex        =   26
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   10320
      TabIndex        =   25
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   10320
      TabIndex        =   24
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Length Of The String"
      Height          =   615
      Left            =   10320
      TabIndex        =   23
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   10320
      TabIndex        =   22
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   15240
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14280
      Top             =   0
   End
   Begin VB.CommandButton Command17 
      Caption         =   "1`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   21
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      Caption         =   "3`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   20
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "10`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   19
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "9`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "8`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   17
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "7`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "6`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "5`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "4`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2`s Table"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "OFF"
      Height          =   855
      Left            =   6600
      TabIndex        =   9
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MOD"
      Height          =   855
      Left            =   7920
      TabIndex        =   8
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Division"
      Height          =   855
      Left            =   5160
      TabIndex        =   7
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Multiplication"
      Height          =   855
      Left            =   3360
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Substraction"
      Height          =   855
      Left            =   7920
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Addition"
      Height          =   855
      Left            =   3360
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ON"
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   7080
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "                                   SHADAB"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label Label3 
      Caption         =   "Tables From 1 to 10                 And           Integer Value Finder"
      Height          =   735
      Left            =   960
      TabIndex        =   40
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape12 
      Height          =   3015
      Left            =   6480
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Label lblstop 
      BackColor       =   &H8000000D&
      Height          =   735
      Left            =   4800
      TabIndex        =   34
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblstart 
      BackColor       =   &H8000000D&
      Height          =   735
      Left            =   3120
      TabIndex        =   33
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Shape Shape11 
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   3000
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   13920
      Top             =   4920
      Width           =   6255
   End
   Begin VB.Shape Shape10 
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   10200
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Shape Shape9 
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   10200
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Shape Shape8 
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   10200
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape7 
      Height          =   375
      Left            =   14160
      Top             =   4200
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   14160
      X2              =   19920
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   14160
      X2              =   19920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   14280
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Integer Value Finder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   12
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   9135
      Left            =   840
      Top             =   0
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   6135
      Left            =   3000
      Top             =   360
      Width           =   6975
   End
   Begin VB.Menu MenuEditor 
      Caption         =   "MenuEditor"
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
'this is comment

Private Sub Command1_Click()
Text3.Enabled = True
Text3.Text = Val(Text3.Text)
End Sub

Private Sub Command10_Click()
num = 5

While (num <= 50)

Print num
num = num + 5
Wend
End Sub

Private Sub Command11_Click()
num = 6

While (num <= 60)

Print num
num = num + 6
Wend
End Sub

Private Sub Command12_Click()
num = 7

While (num <= 70)

Print num
num = num + 7
Wend
End Sub

Private Sub Command13_Click()
num = 8

While (num <= 80)

Print num
num = num + 8
Wend
End Sub

Private Sub Command14_Click()
num = 9

While (num <= 90)

Print num
num = num + 9
Wend
End Sub

Private Sub Command15_Click()
num = 10

While (num <= 100)

Print num
num = num + 10
Wend
End Sub

Private Sub Command16_Click()
num = 3

While (num <= 30)

Print num
num = num + 3
Wend
End Sub

Private Sub Command17_Click()
num = 1
Do
Print num
num = num + 1
Loop While (num <= 10)
End Sub

Private Sub Command18_Click()
End
End Sub

Private Sub Command19_Click()
Dim str As String
Dim Length As Integer
str = Text5.Text
Length = Len(str)
Text4.Text = Length
End Sub

Private Sub Command2_Click()
Text3.Text = Val(Text1.Text) + Val(Text2.Text)
End Sub

Private Sub Command20_Click()
Dim num As Integer
Dim r As Integer
num = Val(Text6.Text)
r = Sqr(num)
Text7.Text = r
End Sub

Private Sub Command21_Click()
Text8.Text = Date
End Sub

Private Sub Command22_Click()
starttime = Now
Timer3.Enabled = True
lblstart.Caption = Time
End Sub

Private Sub Command23_Click()
Timer3.Enabled = False
End Sub

Private Sub Command24_Click()
richtextbox.Enabled = True
End Sub

Private Sub Command25_Click()
richtextbox.Enabled = False
End Sub

Private Sub Command3_Click()
Text3.Text = Val(Text1.Text) - Val(Text2.Text)
End Sub

Private Sub Command4_Click()
Text3.Text = Val(Text1.Text) * Val(Text2.Text)
End Sub

Private Sub Command5_Click()
Text3.Text = Val(Text1.Text) / Val(Text2.Text)
End Sub

Private Sub Command6_Click()
Text3.Text = Val(Text1.Text) Mod Val(Text2.Text)
End Sub

Private Sub Command7_Click()
Text3.Text = disabled
Text1.Text = disabled
Text2.Text = disabled
End Sub

Private Sub Command8_Click()
num = 2

While (num <= 20)

Print num
num = num + 2
Wend
End Sub

Private Sub Command9_Click()
num = 4

While (num <= 40)

Print num
num = num + 4
Wend
End Sub

Private Sub Copy_Click()
a = richtextbox.SelText
End Sub

Private Sub Cut_Click()
a = richtextbox.SelText
richtextbox.SelText = Replace(richtextbox.SelText, a, "")
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub File1_Click()
Image1.Picture = LoadPicture(File1.Path & "/" & File1.FileName)
End Sub

Private Sub Form_Load()
Timer3.Enabled = False
End Sub

Private Sub Label2_Click()
num = InputBox("enter any number")
If num < 0 Then
Print "Negative"
Else
Print "Positive"
End If







End Sub

Private Sub Paste_Click()
richtextbox.Text = richtextbox.Text + a
End Sub

Private Sub Timer1_Timer()
If Shape4.Visible = True Then
Shape5.Visible = True
Shape6.Visible = False
Shape4.Visible = False
ElseIf Shape5.Visible = True Then
Shape6.Visible = True
Shape4.Visible = False
Shape5.Visible = False
ElseIf Shape6.Visible = True Then
Shape4.Visible = True
Shape5.Visible = False
Shape6.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
If (Shape5.Visible = True Or Shape6.Visible = True) Then
Shape7.Left = Shape7.Left + 20
ElseIf (Shape4.Visible = True) Then
Shape7.Left = Shape7.Left + 0
End If


End Sub

Private Sub Timer3_Timer()
lblstop.Caption = Format$(Now - starttime, "hh:mm:ss")
End Sub
