VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form 检查单打印 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00004040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检查单浏览"
   ClientHeight    =   6930
   ClientLeft      =   10365
   ClientTop       =   750
   ClientWidth     =   8475
   FillColor       =   &H00D1815F&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "检查单打印.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   8475
   Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
      Height          =   6015
      Left            =   7920
      TabIndex        =   3
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   10610
      _Version        =   393216
      LargeChange     =   100
      Min             =   100
      Max             =   2000
      Orientation     =   1572864
      SmallChange     =   50
      Value           =   280
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   7935
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1680
         Top             =   -120
      End
      Begin VB.PictureBox Pbox1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         ClipControls    =   0   'False
         DragIcon        =   "检查单打印.frx":7CEB8
         ForeColor       =   &H00404040&
         Height          =   5775
         Left            =   600
         ScaleHeight     =   18.585
         ScaleMode       =   0  'User
         ScaleWidth      =   16.792
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "检查单打印.frx":7D442
      Left            =   3720
      List            =   "检查单打印.frx":7D45B
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   375
      Left            =   -1080
      TabIndex        =   2
      Top             =   6480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   100
      Min             =   100
      Max             =   2000
      Orientation     =   1572865
      SmallChange     =   50
      Value           =   280
   End
   Begin VB.Label Label4 
      Caption         =   "4(11)"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   4
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   8880
      TabIndex        =   15
      ToolTipText     =   "就诊时间："
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   8880
      TabIndex        =   14
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "4(10)"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   8880
      TabIndex        =   13
      ToolTipText     =   "就诊日期："
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "4(9)"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   8880
      TabIndex        =   12
      ToolTipText     =   "就诊医师："
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "4(8)"
      Height          =   495
      Index           =   8
      Left            =   8880
      TabIndex        =   11
      ToolTipText     =   "医嘱建议："
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "4(2)"
      Height          =   495
      Index           =   2
      Left            =   8880
      TabIndex        =   10
      ToolTipText     =   "主诉："
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "4(1)"
      Height          =   495
      Index           =   1
      Left            =   8880
      TabIndex        =   9
      ToolTipText     =   "诊断："
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "患者编号"
      Height          =   495
      Index           =   0
      Left            =   8880
      TabIndex        =   8
      ToolTipText     =   "患者编号："
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   3240
      Picture         =   "检查单打印.frx":7D486
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "检查单打印"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End Sub

Private Sub FlatScrollBar1_Change()
Pbox1.Move FlatScrollBar1.Value, FlatScrollBar2.Value
Label2.Caption = "X=" & FlatScrollBar1.Value
End Sub
Private Sub FlatScrollBar2_Change()
Pbox1.Move FlatScrollBar1.Value, -FlatScrollBar2.Value
Label3.Caption = "Y=" & FlatScrollBar2.Value
End Sub
Private Sub Form_Load()
Label5.Caption = MDIForm1.StatusBar1.Panels(6).Text
Me.Left = 10000
End Sub

Private Sub Form_LostFocus()
MsgBox "请先关闭当前窗口！"
End Sub

Private Sub Form_Resize()
Me.FlatScrollBar1.Top = Me.Height - 750
Me.FlatScrollBar2.Left = Me.Width - 500
End Sub

Private Sub Label1_Change()
If Not Label1.Caption = 0.5 Then
Label13.Enabled = True
Else
Label13.Enabled = False
End If

If Not Label1.Caption = 1.5 Then
Label14.Enabled = True
Else
Label14.Enabled = False
End If
End Sub

Private Sub Label12_Click()
Pbox1.Refresh
End Sub

Private Sub Label13_Click()
Label1.Caption = Val(Label1.Caption) - 0.1
Combo1.Text = (Label1.Caption * 100) & "%"
Pbox1.Width = Pbox1.Width - (0.1 * Pbox1.Width)
Pbox1.Height = Pbox1.Height - (0.1 * Pbox1.Height)
a = Label1.Caption
Me.Width = Me.Width - 200
Me.Height = Me.Height - 200

Pbox1.Cls

'Pbox1.PaperSize = 22
Pbox1.ScaleMode = vbCentimeters
'Pbox1.Orientation = 1
Pbox1.FontSize = 16 * a
Pbox1.CurrentX = 1 * a
Pbox1.CurrentY = 1 * a
Pbox1.Print Space(10) & "荒地镇卫生院"
Pbox1.PaintPicture Image1.Picture, 1 * a, 1 * a, 1.5 * a, 1.5 * a
Pbox1.PaintPicture Image1.Picture, 10 * a, 1 * a, 1.5 * a, 1.5 * a
Pbox1.FontSize = 12 * a
Pbox1.CurrentX = 5 * a
Pbox1.CurrentY = 2 * a
Pbox1.Print "挂号单"
Pbox1.DrawStyle = 0   '以实线打印，VbDash 1 虚线 VbDot 2点线
                        'VbDashDot    3         点划线
                        'VbDashDotDot 4       双点划线
                        'VbInvisible  5           无线
                        'VbInsideSolid 6        内收实线
Pbox1.Line (2 * a, 2.5 * a)-(10 * a, 2.5 * a)
Pbox1.Line (2 * a, 2.55 * a)-(10 * a, 2.55 * a)
Pbox1.Line (2 * a, 8.5 * a)-(10 * a, 8.5 * a)
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 3 * a
Pbox1.Print "合作医疗号：" & Space(4) & "2109211211271"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 4 * a
Pbox1.Print "姓名：" & Space(10) & "努尔艾合买提"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 5 * a
Pbox1.Print "身份证号：" & Space(6) & "653125199210053272"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 6 * a
Pbox1.Print "挂号日期：" & Space(6) & Date
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 7 * a
Pbox1.Print "挂号时间：" & Space(6) & Time
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 8 * a
Pbox1.Print "挂号医生：" & Space(6) & "阿里木江"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 9 * a
Pbox1.Print "注：挂号单仅当日有效，下午5点自动作废！"
Pbox1.FontSize = 16 * a
Pbox1.FontBold = True
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 10 * a
Pbox1.Print "门诊号：" & Space(6) & "001"
End Sub

Private Sub Label14_Click()
Label1.Caption = Val(Label1.Caption) + 0.1
Combo1.Text = (Label1.Caption * 100) & "%"
Pbox1.Width = Pbox1.Width + (0.1 * Pbox1.Width)
Pbox1.Height = Pbox1.Height + (0.1 * Pbox1.Height)
Pbox1.Cls
a = Label1.Caption
Me.Width = Me.Width + 200
Me.Height = Me.Height + 200
'Pbox1.PaperSize = 22
Pbox1.ScaleMode = vbCentimeters
'Pbox1.Orientation = 1
Pbox1.FontSize = 16 * a
Pbox1.CurrentX = 1 * a
Pbox1.CurrentY = 1 * a
Pbox1.Print Space(12) & Label5.Caption
Pbox1.PaintPicture Image1.Picture, 1 * a, 1 * a, 1.5 * a, 1.5 * a
Pbox1.PaintPicture Image1.Picture, 10 * a, 1 * a, 1.5 * a, 1.5 * a
Pbox1.FontSize = 12 * a
Pbox1.CurrentX = 6 * a
Pbox1.CurrentY = 2 * a
Pbox1.Print "就诊单"
Pbox1.DrawStyle = 0   '以实线打印，VbDash 1 虚线 VbDot 2点线
                        'VbDashDot    3         点划线
                        'VbDashDotDot 4       双点划线
                        'VbInvisible  5           无线
                        'VbInsideSolid 6        内收实线
Pbox1.Line (2 * a, 2.5 * a)-(10 * a, 2.5 * a)
Pbox1.Line (2 * a, 2.55 * a)-(10 * a, 2.55 * a)
Pbox1.Line (2 * a, 8.5 * a)-(10 * a, 8.5 * a)
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 3 * a
Pbox1.Print "合作医疗号：" & Space(4) & "2109211211271"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 4 * a
Pbox1.Print "姓名：" & Space(10) & "努尔艾合买提"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 5 * a
Pbox1.Print "身份证号：" & Space(6) & "653125199210053272"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 6 * a
Pbox1.Print "挂号日期：" & Space(6) & Date
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 7 * a
Pbox1.Print "挂号时间：" & Space(6) & Time
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 8 * a
Pbox1.Print "挂号医生：" & Space(6) & "阿里木江"
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 9 * a
Pbox1.Print "注：挂号单仅当日有效，下午5点自动作废！"
Pbox1.FontSize = 16 * a
Pbox1.FontBold = True
Pbox1.CurrentX = 2 * a
Pbox1.CurrentY = 10 * a
Pbox1.Print "门诊号：" & Space(6) & "001"
End Sub

Private Sub Label8_Click()
End
End Sub

Private Sub Timer1_Timer()
'For i = 1 To 16
'For Y = 1 To 25
'Pbox1.Line (i, 1)-(i, 25)
'Pbox1.Line (1, Y)-(16, Y)
'Pbox1.CurrentX = i - 0.25
'Pbox1.CurrentY = 0.8
'Pbox1.Print i
'Pbox1.CurrentX = 0.8
'Pbox1.CurrentY = Y - 0.2
'Pbox1.Print Y
'Next Y
'Next i
Label1.Caption = Val(Label1.Caption)
Combo1.Text = (Label1.Caption * 100) & "%"
a = Label1.Caption
Pbox1.FontSize = 16
Pbox1.PSet (6, 3)
Pbox1.ForeColor = vbRed
Pbox1.FontBold = True
Pbox1.Print Label5.Caption
'Pbox1.PaintPicture Image1.Picture, 1, 1, 1.5, 1.5
'Pbox1.PaintPicture Image1.Picture, 10, 1, 1.5, 1.5
Pbox1.FontSize = 12
Pbox1.PSet (6, 4)
Pbox1.ForeColor = vbBlack
Pbox1.Print Space(2) & "就诊记录单"
Pbox1.DrawStyle = 0   '以实线打印，VbDash 1 虚线 VbDot 2点线
                        'VbDashDot    3         点划线
                        'VbDashDotDot 4       双点划线
                        'VbInvisible  5           无线
                        'VbInsideSolid 6        内收实线
 Pbox1.FontBold = False
Pbox1.Line (1, 5)-(16, 5)
Pbox1.Line (1, 5.1)-(16, 5.1)
Pbox1.PSet (1, 6)
Pbox1.Print Label4.Item(0).ToolTipText & Label4.Item(0).Caption
Pbox1.PSet (1, 7)
Pbox1.Print Label4.Item(1).ToolTipText & Label4.Item(1).Caption
Pbox1.PSet (1, 8)
Pbox1.Print Label4.Item(2).ToolTipText & Label4.Item(2).Caption
Pbox1.PSet (1, 9)
Pbox1.Print Label4.Item(8).ToolTipText & Label4.Item(8).Caption
Pbox1.PSet (1, 10)
Pbox1.Print Label4.Item(9).ToolTipText & Label4.Item(9).Caption
Pbox1.PSet (1, 11)
Pbox1.Print Label4.Item(10).ToolTipText & Label4.Item(10).Caption
Pbox1.PSet (1, 12)
Pbox1.Print Label4.Item(11).ToolTipText & Label4.Item(11).Caption
Timer1.Interval = 0
End Sub
