VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ����վ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "����ҽ������վ"
   ClientHeight    =   9090
   ClientLeft      =   3945
   ClientTop       =   450
   ClientWidth     =   14010
   DrawStyle       =   6  'Inside Solid
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MZYSGZZ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   16.034
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   24.712
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   14843
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   176
      TabMaxWidth     =   176
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "MZYSGZZ.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSTab2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "����ѡ��"
         Height          =   465
         Left            =   11640
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   7575
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   13361
         _Version        =   393216
         Tab             =   1
         TabHeight       =   706
         TabMaxWidth     =   2646
         BackColor       =   16777215
         OLEDropMode     =   1
         TabCaption(0)   =   "����"
         TabPicture(0)   =   "MZYSGZZ.frx":03A6
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label9(3)"
         Tab(0).Control(1)=   "Label9(2)"
         Tab(0).Control(2)=   "Label9(0)"
         Tab(0).Control(3)=   "Label10"
         Tab(0).Control(4)=   "Label9(12)"
         Tab(0).Control(5)=   "TextBox1(12)"
         Tab(0).Control(6)=   "Label9(14)"
         Tab(0).Control(7)=   "TextBox1(14)"
         Tab(0).Control(8)=   "Image1"
         Tab(0).Control(9)=   "Label9(1)"
         Tab(0).Control(10)=   "DataGrid2"
         Tab(0).Control(11)=   "Adodc1"
         Tab(0).Control(12)=   "DataGrid1"
         Tab(0).Control(13)=   "Text2"
         Tab(0).Control(14)=   "Text3"
         Tab(0).Control(15)=   "Text5"
         Tab(0).Control(16)=   "Command5"
         Tab(0).Control(17)=   "Adodc9"
         Tab(0).Control(18)=   "Command4"
         Tab(0).Control(19)=   "Command7"
         Tab(0).Control(20)=   "Command8"
         Tab(0).Control(21)=   "Text4"
         Tab(0).Control(22)=   "Combo2"
         Tab(0).Control(23)=   "Combo3"
         Tab(0).Control(24)=   "Text8"
         Tab(0).Control(25)=   "Text12"
         Tab(0).Control(26)=   "Combo4"
         Tab(0).Control(27)=   "Combo6"
         Tab(0).ControlCount=   28
         TabCaption(1)   =   "���ﲡ��"
         TabPicture(1)   =   "MZYSGZZ.frx":03C2
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Line1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label11"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label16"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label14"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label15"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label12"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Line2"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "RichTextBox3"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Adodc6"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Text1"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Frame3"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Frame2"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Combo1"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Adodc7"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "DataCombo4"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Command6"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "DataGrid5"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Adodc8"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Adodc2"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Command10"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Combo5"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "Command11"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "Command15"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "Text14"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "Text13"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "Command16"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).ControlCount=   26
         TabCaption(2)   =   "������"
         TabPicture(2)   =   "MZYSGZZ.frx":03DE
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame6"
         Tab(2).Control(1)=   "Adodc4"
         Tab(2).Control(2)=   "Label27"
         Tab(2).ControlCount=   3
         Begin VB.CommandButton Command16 
            Caption         =   "��Ժ"
            Height          =   495
            Left            =   4680
            TabIndex        =   82
            Top             =   4560
            Width           =   855
         End
         Begin VB.ComboBox Combo6 
            DataField       =   "��λ"
            DataSource      =   "Adodc9"
            Height          =   360
            ItemData        =   "MZYSGZZ.frx":03FA
            Left            =   -73440
            List            =   "MZYSGZZ.frx":0407
            TabIndex        =   2
            Text            =   "��λ"
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   840
            TabIndex        =   80
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox Text14 
            Height          =   360
            Left            =   1680
            TabIndex        =   79
            Top             =   3120
            Width           =   495
         End
         Begin VB.CommandButton Command15 
            Caption         =   "��     ��"
            Height          =   495
            Left            =   3600
            TabIndex        =   78
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton Command11 
            Caption         =   "����Ϊģ��"
            Height          =   495
            Left            =   3600
            TabIndex        =   77
            Top             =   3120
            Width           =   1335
         End
         Begin VB.ComboBox Combo5 
            Height          =   360
            ItemData        =   "MZYSGZZ.frx":0417
            Left            =   2040
            List            =   "MZYSGZZ.frx":0424
            TabIndex        =   74
            Text            =   "��Ժʱ���"
            Top             =   4200
            Width           =   1695
         End
         Begin VB.ComboBox Combo4 
            Height          =   360
            ItemData        =   "MZYSGZZ.frx":0438
            Left            =   -71400
            List            =   "MZYSGZZ.frx":044E
            TabIndex        =   4
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   -72000
            TabIndex        =   3
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox Text8 
            DataField       =   "���"
            Height          =   375
            Left            =   -69480
            TabIndex        =   72
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command10 
            Caption         =   "סԺ����ӡ"
            Height          =   495
            Left            =   3960
            TabIndex        =   68
            Top             =   3960
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
            Height          =   360
            ItemData        =   "MZYSGZZ.frx":046C
            Left            =   -70560
            List            =   "MZYSGZZ.frx":0479
            TabIndex        =   6
            Text            =   "����"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            Height          =   360
            ItemData        =   "MZYSGZZ.frx":0498
            Left            =   -72480
            List            =   "MZYSGZZ.frx":04A5
            TabIndex        =   5
            Text            =   "ʹ�÷���"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Left            =   -68400
            TabIndex        =   62
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   "���洦��"
            Height          =   615
            Left            =   -71280
            TabIndex        =   61
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "��ӡ����"
            Height          =   615
            Left            =   -73200
            TabIndex        =   60
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "ɾ������"
            Height          =   615
            Left            =   -74760
            TabIndex        =   59
            Top             =   2520
            Width           =   1335
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   375
            Left            =   1320
            Top             =   6720
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "סԺ��"
            Caption         =   "Adodc2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc Adodc9 
            Height          =   375
            Left            =   -64320
            Top             =   0
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "ҩƷ���"
            Caption         =   "Adodc9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc Adodc8 
            Height          =   330
            Left            =   3840
            Top             =   2640
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "��������"
            Caption         =   "Adodc8"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid5 
            Bindings        =   "MZYSGZZ.frx":04BF
            Height          =   2175
            Left            =   120
            TabIndex        =   54
            Top             =   5160
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   3836
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   17
            BeginProperty Column00 
               DataField       =   "���߱��"
               Caption         =   "���߱��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "�Ա�"
               Caption         =   "�Ա�"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "סԺ��"
               Caption         =   "סԺ��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "סԺ��"
               Caption         =   "סԺ��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "���"
               Caption         =   "���"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "����ҽ��"
               Caption         =   "����ҽ��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "���֤��"
               Caption         =   "���֤��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "ҽ��֤��"
               Caption         =   "ҽ��֤��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "��ַ"
               Caption         =   "��ַ"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column12 
               DataField       =   "��Ժ����"
               Caption         =   "��Ժ����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column13 
               DataField       =   "��������"
               Caption         =   "��������"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column14 
               DataField       =   "������"
               Caption         =   "������"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column15 
               DataField       =   "�տ�������"
               Caption         =   "�տ�������"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column16 
               DataField       =   "״̬"
               Caption         =   "״̬"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   675.213
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   824.882
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   689.953
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   675.213
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   884.976
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column16 
                  ColumnWidth     =   2099.906
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton Command6 
            Caption         =   "��Ժ"
            Height          =   495
            Left            =   3840
            TabIndex        =   53
            Top             =   4560
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "MZYSGZZ.frx":04D4
            Height          =   390
            Left            =   120
            TabIndex        =   52
            Top             =   4200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   688
            _Version        =   393216
            ListField       =   "������"
            Text            =   "������"
         End
         Begin VB.CommandButton Command5 
            Caption         =   "���"
            Height          =   615
            Left            =   -69480
            TabIndex        =   8
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   -74160
            TabIndex        =   1
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox Text3 
            DataField       =   "ҩƷ��"
            Height          =   375
            Left            =   -71640
            TabIndex        =   51
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   -74160
            TabIndex        =   0
            Top             =   720
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "MZYSGZZ.frx":04E9
            Height          =   2775
            Left            =   -67440
            TabIndex        =   50
            Top             =   480
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4895
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BackColor       =   12632256
            ForeColor       =   128
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "������"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "ҩƷ��"
               Caption         =   "ҩƷ��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "���"
               Caption         =   "���"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "��λ"
               Caption         =   "��λ"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "���"
               Caption         =   "���"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "��ע"
               Caption         =   "��ע"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "״̬"
               Caption         =   "״̬"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1484.787
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   780.095
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   840.189
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   659.906
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1814.74
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adodc7 
            Height          =   375
            Left            =   10680
            Top             =   2760
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "���ﲡ��ģ��"
            Caption         =   "Adodc7"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.ComboBox Combo1 
            Height          =   360
            ItemData        =   "MZYSGZZ.frx":04FE
            Left            =   1320
            List            =   "MZYSGZZ.frx":050E
            TabIndex        =   12
            Text            =   "Combo1"
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Frame Frame6 
            Caption         =   "������"
            Height          =   6375
            Left            =   -74880
            TabIndex        =   34
            Top             =   600
            Width           =   12255
            Begin VB.TextBox Text11 
               Height          =   375
               Left            =   240
               MaxLength       =   16
               TabIndex        =   71
               Top             =   840
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CommandButton Command12 
               Caption         =   "�ԷѴ�ӡ"
               Height          =   495
               Left            =   10080
               TabIndex        =   70
               Top             =   840
               Width           =   1335
            End
            Begin MSDataGridLib.DataGrid DataGrid6 
               Height          =   1095
               Left            =   120
               TabIndex        =   69
               Top             =   1440
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   1931
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ҽ����ӡ"
               Height          =   495
               Left            =   8040
               TabIndex        =   40
               Top             =   840
               Width           =   1455
            End
            Begin VB.TextBox Text7 
               Height          =   375
               Left            =   5160
               TabIndex        =   58
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   5160
               TabIndex        =   57
               Top             =   360
               Width           =   2055
            End
            Begin VB.CommandButton Command3 
               Caption         =   "ɾ��"
               Height          =   495
               Left            =   10080
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton Command2 
               Caption         =   "���"
               Height          =   495
               Left            =   8040
               TabIndex        =   39
               Top             =   240
               Width           =   1455
            End
            Begin MSAdodcLib.Adodc Adodc5 
               Height          =   330
               Left            =   2400
               Top             =   960
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   2
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
               OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "�����Ŀ"
               Caption         =   "5"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSDataListLib.DataCombo DataCombo3 
               DataField       =   "��Ŀ����"
               Height          =   390
               Left            =   2400
               TabIndex        =   37
               Top             =   360
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   688
               _Version        =   393216
               ListField       =   "��Ŀ����"
               Text            =   "�����Ŀ"
            End
            Begin MSDataGridLib.DataGrid DataGrid8 
               Bindings        =   "MZYSGZZ.frx":052A
               Height          =   3375
               Left            =   120
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   2640
               Width           =   12255
               _ExtentX        =   21616
               _ExtentY        =   5953
               _Version        =   393216
               BackColor       =   8421376
               ForeColor       =   -2147483643
               HeadLines       =   1
               RowHeight       =   15
               FormatLocked    =   -1  'True
               AllowAddNew     =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   18
               BeginProperty Column00 
                  DataField       =   "��ˮ��"
                  Caption         =   "��ˮ��"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "�����"
                  Caption         =   "�����"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "�����Ŀ"
                  Caption         =   "�����Ŀ"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "��������"
                  Caption         =   "��������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "����ʱ��"
                  Caption         =   "����ʱ��"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "�������"
                  Caption         =   "�������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "����ҽʦ"
                  Caption         =   "����ҽʦ"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "���"
                  Caption         =   "���"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column08 
                  DataField       =   "�����"
                  Caption         =   "�����"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column09 
                  DataField       =   "������"
                  Caption         =   "������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column10 
                  DataField       =   "�������"
                  Caption         =   "�������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column11 
                  DataField       =   "���ʱ��"
                  Caption         =   "���ʱ��"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column12 
                  DataField       =   "������"
                  Caption         =   "������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column13 
                  DataField       =   "���ҽʦ"
                  Caption         =   "���ҽʦ"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column14 
                  DataField       =   "���ʱ��"
                  Caption         =   "���ʱ��"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column15 
                  DataField       =   "״̬"
                  Caption         =   "״̬"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column16 
                  DataField       =   "�۸�"
                  Caption         =   "�۸�"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column17 
                  DataField       =   "��������"
                  Caption         =   "��������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     ColumnWidth     =   1665.071
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   1275.024
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   1154.835
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   1065.26
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   1184.882
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   1154.835
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   1214.929
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   629.858
                  EndProperty
                  BeginProperty Column08 
                     ColumnWidth     =   1140.095
                  EndProperty
                  BeginProperty Column09 
                     ColumnWidth     =   1110.047
                  EndProperty
                  BeginProperty Column10 
                     ColumnWidth     =   1049.953
                  EndProperty
                  BeginProperty Column11 
                     ColumnWidth     =   1124.787
                  EndProperty
                  BeginProperty Column12 
                     ColumnWidth     =   1200.189
                  EndProperty
                  BeginProperty Column13 
                     ColumnWidth     =   1170.142
                  EndProperty
                  BeginProperty Column14 
                     ColumnWidth     =   1230.236
                  EndProperty
                  BeginProperty Column15 
                     ColumnWidth     =   734.74
                  EndProperty
                  BeginProperty Column16 
                     ColumnWidth     =   780.095
                  EndProperty
                  BeginProperty Column17 
                     ColumnWidth     =   1244.976
                  EndProperty
               EndProperty
            End
            Begin MSAdodcLib.Adodc Adodc3 
               Height          =   375
               Left            =   240
               Top             =   840
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   2
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
               OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "������"
               Caption         =   "3"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Bindings        =   "MZYSGZZ.frx":053F
               Height          =   390
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   688
               _Version        =   393216
               ListField       =   "��������"
               Text            =   "ѡ�����"
            End
            Begin VB.Label Label18 
               Caption         =   "�۸�"
               Height          =   375
               Left            =   4440
               TabIndex        =   56
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "��λ��"
               Height          =   375
               Left            =   4440
               TabIndex        =   55
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "���ղ���ģ��"
            Height          =   3375
            Left            =   6000
            TabIndex        =   15
            Top             =   555
            Width           =   6615
            Begin VB.CommandButton Command14 
               Caption         =   "��   ��"
               Height          =   495
               Left            =   4800
               TabIndex        =   76
               Top             =   1560
               Width           =   1575
            End
            Begin VB.CommandButton Command13 
               Caption         =   "ɾ   ��"
               Height          =   495
               Left            =   4800
               TabIndex        =   75
               Top             =   600
               Width           =   1575
            End
            Begin MSDataGridLib.DataGrid DataGrid3 
               Bindings        =   "MZYSGZZ.frx":0554
               Height          =   2775
               Left            =   120
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   240
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   4895
               _Version        =   393216
               BackColor       =   16776960
               HeadLines       =   1
               RowHeight       =   15
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "���"
                  Caption         =   "���"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "����"
                  Caption         =   "����"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "ҽ������"
                  Caption         =   "ҽ������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     ColumnWidth     =   1470.047
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   1725.165
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   2775.118
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "������¼"
            Height          =   3015
            Left            =   6000
            TabIndex        =   13
            Top             =   4080
            Width           =   7215
            Begin MSDataGridLib.DataGrid DataGrid4 
               Bindings        =   "MZYSGZZ.frx":0569
               Height          =   2655
               Left            =   120
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   240
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4683
               _Version        =   393216
               BackColor       =   8388736
               HeadLines       =   1
               RowHeight       =   15
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   8
               BeginProperty Column00 
                  DataField       =   "���߱��"
                  Caption         =   "���߱��"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "���"
                  Caption         =   "���"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "����"
                  Caption         =   "����"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "ҽ������"
                  Caption         =   "ҽ������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "����ҽʦ"
                  Caption         =   "����ҽʦ"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "��������"
                  Caption         =   "��������"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "��ַ��"
                  Caption         =   "��"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "����ʱ��"
                  Caption         =   "����ʱ��"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     ColumnWidth     =   14.74
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   780.095
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   1124.787
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   1319.811
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   1184.882
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   315.213
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   1574.929
                  EndProperty
               EndProperty
            End
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   840
            MaxLength       =   50
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   840
            Width           =   4935
         End
         Begin MSAdodcLib.Adodc Adodc6 
            Height          =   375
            Left            =   8400
            Top             =   7200
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "���ﲡ��"
            Caption         =   "Adodc6"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   330
            Left            =   -69720
            Top             =   5760
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "��鵥"
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   -70080
            Top             =   8640
            Visible         =   0   'False
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "���ﴦ��"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "MZYSGZZ.frx":057E
            Height          =   3855
            Left            =   -74880
            TabIndex        =   17
            Top             =   3240
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   6800
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "���"
               Caption         =   "���"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "ҩƷ����"
               Caption         =   "ҩƷ����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "���"
               Caption         =   "���"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "��λ"
               Caption         =   "��λ"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "���"
               Caption         =   "���"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "�÷�"
               Caption         =   "�÷�"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "ҽ��"
               Caption         =   "ҽ��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "����"
               Caption         =   "����"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column12 
               DataField       =   "ʱ��"
               Caption         =   "ʱ��"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   2
               Locked          =   -1  'True
               Size            =   664
               BeginProperty Column00 
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2039.811
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2280.189
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   929.764
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   615.118
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   659.906
               EndProperty
            EndProperty
         End
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   855
            Left            =   840
            TabIndex        =   11
            Top             =   1560
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1508
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"MZYSGZZ.frx":0593
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   5880
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Label Label12 
            Caption         =   "��ַ��     ��     ��"
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   3160
            Width           =   2655
         End
         Begin MSForms.Label Label9 
            Height          =   375
            Index           =   1
            Left            =   -72600
            TabIndex        =   73
            Top             =   1920
            Width           =   735
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "����"
            Size            =   "1296;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
            ParagraphAlign  =   3
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   -74640
            Picture         =   "MZYSGZZ.frx":0622
            Stretch         =   -1  'True
            Top             =   8520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSForms.TextBox TextBox1 
            DataField       =   "����"
            Height          =   375
            Index           =   14
            Left            =   -68400
            TabIndex        =   67
            Top             =   720
            Width           =   855
            VariousPropertyBits=   746604571
            Size            =   "1508;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label Label9 
            Height          =   375
            Index           =   14
            Left            =   -69120
            TabIndex        =   66
            Top             =   720
            Width           =   855
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "����"
            Size            =   "1508;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
            ParagraphAlign  =   3
         End
         Begin MSForms.TextBox TextBox1 
            DataField       =   "���"
            Height          =   375
            Index           =   12
            Left            =   -74160
            TabIndex        =   65
            Top             =   1320
            Width           =   1575
            VariousPropertyBits=   746604571
            ScrollBars      =   1
            Size            =   "2778;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label Label9 
            Height          =   375
            Index           =   12
            Left            =   -75000
            TabIndex        =   64
            Top             =   1320
            Width           =   975
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "��  ��"
            Size            =   "1720;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
            ParagraphAlign  =   3
         End
         Begin VB.Label Label10 
            Caption         =   "�ܼ�"
            Height          =   375
            Left            =   -69000
            TabIndex        =   63
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label27 
            Height          =   375
            Left            =   -70800
            TabIndex        =   49
            Top             =   8475
            Width           =   3135
         End
         Begin VB.Label Label15 
            Height          =   375
            Left            =   7320
            TabIndex        =   47
            Top             =   8355
            Width           =   2535
         End
         Begin VB.Label Label14 
            Caption         =   "���:"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "���ߣ�"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            Width           =   1215
         End
         Begin MSForms.Label Label9 
            Height          =   375
            Index           =   0
            Left            =   -74880
            TabIndex        =   21
            Top             =   720
            Width           =   735
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "������"
            Size            =   "1296;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
            ParagraphAlign  =   3
         End
         Begin MSForms.Label Label9 
            Height          =   375
            Index           =   2
            Left            =   -75000
            TabIndex        =   20
            Top             =   1920
            Width           =   975
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "��  ��"
            Size            =   "1720;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
            ParagraphAlign  =   3
         End
         Begin MSForms.Label Label9 
            Height          =   375
            Index           =   3
            Left            =   -72600
            TabIndex        =   19
            Top             =   720
            Width           =   975
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "ҩƷ����"
            Size            =   "1720;661"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
            ParagraphAlign  =   3
         End
         Begin MSForms.Label Label11 
            Height          =   495
            Left            =   120
            TabIndex        =   18
            Top             =   2640
            Width           =   1215
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "��������"
            Size            =   "2143;873"
            FontName        =   "����"
            FontHeight      =   240
            FontCharSet     =   134
            FontPitchAndFamily=   34
         End
         Begin VB.Line Line1 
            X1              =   5880
            X2              =   5880
            Y1              =   480
            Y2              =   9315
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   0
         TabIndex        =   25
         Top             =   -30
         Width           =   13095
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "[���㷽ʽ]"
            Height          =   375
            Left            =   8760
            TabIndex        =   46
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   10080
            TabIndex        =   45
            Top             =   195
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "[��������]"
            Height          =   375
            Index           =   0
            Left            =   2400
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3600
            TabIndex        =   31
            Top             =   195
            Width           =   2775
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   6960
            TabIndex        =   30
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "���䣺"
            Height          =   375
            Left            =   7560
            TabIndex        =   29
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   8160
            TabIndex        =   28
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            Height          =   375
            Left            =   6360
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   195
            Width           =   2295
         End
      End
      Begin VB.Label Label25 
         Caption         =   "��ѡ����ﻼ�ߣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   36
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   2040
         TabIndex        =   48
         Top             =   4080
         Width           =   8775
      End
      Begin VB.Label Label23 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1860
         TabIndex        =   33
         Top             =   8850
         Width           =   375
      End
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label24"
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   960
      TabIndex        =   44
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label22"
      ForeColor       =   &H00FFC0FF&
      Height          =   300
      Left            =   2760
      TabIndex        =   43
      Top             =   8640
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����"
      Height          =   345
      Left            =   2040
      TabIndex        =   42
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "���ң�"
      Height          =   345
      Left            =   240
      TabIndex        =   41
      Top             =   8640
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "����վ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
If Combo5.Text = "��Ժʱ���" Then
MsgBox " ��ѡ����Ժʱ�������ԣ�", vbInformation
Exit Sub
End If

Printer.ScaleMode = vbCentimeters
Printer.PaperSize = 13   '182   x   257   mm
Printer.FontSize = 12
Dim DIZHI As String
If Left(Label19.Caption, 2) = "07" Then
DIZHI = "ɯ���� �ĵ���" & Mid(Label19.Caption, 3, 2) & "��" & Text14.Text & "��"
End If
If Left(Label19.Caption, 2) = "08" Then
DIZHI = "ɯ���� �հ͸���" & Mid(Label19.Caption, 3, 2) & "��" & Text14.Text & "��"
End If
If Not (Left(Label19.Caption, 2) = "07" Or Left(Label19.Caption, 2) = "08") Then
DIZHI = "ɯ����               " & Mid(Label19.Caption, 3, 2) & "��" & Text14.Text & "��"
End If

           
            Printer.CurrentX = 6
            Printer.CurrentY = 2.5
            Printer.FontBold = True
            Printer.FontSize = 16
            Printer.Print "�� �� �� �� �� Ժ"
             
            Printer.CurrentX = 7
            Printer.CurrentY = 3
            Printer.Print "סԺ֪ͨ��"
            
            Printer.FontBold = False
            Printer.FontSize = 12
            Printer.CurrentX = 1
            Printer.CurrentY = 4
            Printer.Print "���߱�ţ�" & Label19.Caption & Space(5) & "����������" & Label2.Caption & Space(3) & "�Ա�" & Label3.Caption

            Printer.CurrentX = 1
            Printer.CurrentY = 5
            Printer.Print "���䣺" & Label5.Caption & Space(10) & "���㷽ʽ:" & Label7.Caption & Space(5) & "��Ժʱ�����" & Combo5.Text
            
            Printer.CurrentX = 1
            Printer.CurrentY = 6
            Printer.Print "��������" & DataCombo4.Text & Space(5) & "��ϣ�" & Text1.Text
            Printer.CurrentX = 1
           Printer.CurrentY = 7
           Printer.Print "��ͥסַ��" & DIZHI
           
           Printer.CurrentX = 1
           Printer.CurrentY = 8
           Printer.Print "����ҽʦ��" & Label22.Caption & Space(10) & "������ң�" & Label24.Caption & "������"

           Printer.FontSize = 14
           Printer.FontBold = True
          
           Printer.CurrentX = 5
           Printer.CurrentY = 10
           Printer.FontSize = 12
           Printer.FontBold = False
           Printer.Print "����ʱ�䣺" & Now
           Call Command6_Click
           Printer.EndDoc
End Sub

Private Sub Command11_Click()
Dim RS As ADODB.Recordset
Set RS = Adodc6.Recordset
    Dim RSs As ADODB.Recordset
    Set RSs = Adodc7.Recordset
    On Error Resume Next
     RSs.AddNew
          RSs!���߱�� = Label19.Caption
          RSs!��� = Text1.Text
          RSs!���� = RichTextBox3.Text
          RSs!ҽ������ = Combo1.Text
          RSs.Update
End Sub

Private Sub Command12_Click()
On Error Resume Next
If (Text1.Text = "" And Text14.Text = "") Then
MsgBox "��Ϻ͵�ַ��Ϣ����Ϊ�գ�"
Exit Sub
End If

 Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
 Printer.ScaleMode = vbCentimeters
 Printer.FontSize = 16
 Printer.Orientation = 1
 Printer.FontBold = True
Printer.CurrentX = 6
Printer.CurrentY = 1
Printer.Print "ɯ���ػĵ�������Ժ"
Printer.CurrentX = 7
Printer.CurrentY = 2
 Printer.FontSize = 12
 If DataCombo1.Text = "������" Then
Printer.Print "�����Ҽ�����뵥"

End If
If DataCombo1.Text = "B����" Then
Printer.Print "��ɫB��������뵥"

End If
If DataCombo1.Text = "ҽѧӰ����" Then
Printer.Print "X�߼�����뵥" '

End If
Printer.FontBold = False

Printer.Line (1, 2.8)-(17, 2.8)
Printer.Line (1, 2.85)-(17, 2.85)

Printer.CurrentX = 1
Printer.CurrentY = 3
Printer.Print "����������" & Label2.Caption & Space(6) & "�Ա�" & Label3.Caption & Space(10) & "���䣺" & Label5.Caption
 Printer.CurrentX = 1
Printer.CurrentY = 4
Printer.Print "���㷽ʽ��" & Label7.Caption & Space(10) & "��ַ��" & Text13.Text & "-" & Text14.Text & Space(10) & "��ϣ�" & Text1.Text
Printer.CurrentX = 1
Printer.CurrentY = 5
If DataCombo3.Text = "������Ŀ" Then
 Printer.Print "�����Ŀ��   " & Text11.Text & Space(3) & "���" & Text6.Text & Space(3) & "�۸񣺴�����     ���۽�________ "
Else
Printer.Print "�����Ŀ��" & DataCombo3.Text & Space(6) & "���" & Text6.Text & Space(6) & "�۸�" & Text7.Text & "Ԫ"
End If

Printer.Line (1, 5.5)-(17, 5.5)
Printer.CurrentX = 1
Printer.CurrentY = 6
Printer.Print "�Էѽ�" & Space(30) & "�Է��ܽ�"
Printer.CurrentX = 1
Printer.CurrentY = 7
Printer.Print "������ң�" & Label24.Caption & Space(2) & "����ҽʦ��" & Label22.Caption & Space(5) & "����Ա��" & Space(10) & "�շ�Ա��"
Printer.CurrentX = 4
Printer.CurrentY = 8
Printer.Print "����ʱ�䣺" & Now
Printer.CurrentX = 1
Printer.CurrentY = 9
Printer.FontSize = 14
Printer.FontBold = True
Printer.Print " ע�����º���Ч��"

Printer.FontSize = 16
 Printer.Orientation = 1
Printer.CurrentX = 6
Printer.CurrentY = 15
Printer.Print "ɯ���ػĵ�������Ժ"
Printer.CurrentX = 7
Printer.CurrentY = 16
 Printer.FontSize = 12
 If DataCombo1.Text = "������" Then
Printer.Print "�����Ҽ�����뵥"
End If

If DataCombo1.Text = "B����" Then
Printer.Print "��ɫB��������뵥"
End If

If DataCombo1.Text = "ҽѧӰ����" Then
Printer.Print "X�߼�����뵥" '
End If
Printer.FontBold = False

Printer.Line (1, 16.8)-(17, 16.8)
Printer.Line (1, 16.85)-(17, 16.85)


Printer.CurrentX = 1
Printer.CurrentY = 17
Printer.Print "����������" & Label2.Caption & Space(6) & "�Ա�" & Label3.Caption & Space(10) & "���䣺" & Label5.Caption
 Printer.CurrentX = 1
Printer.CurrentY = 18
Printer.Print "���㷽ʽ��" & Label7.Caption & Space(10) & "��ַ��" & Text13.Text & "-" & Text14.Text & Space(10) & "��ϣ�" & Text1.Text
Printer.CurrentX = 1
Printer.CurrentY = 19
If DataCombo3.Text = "������Ŀ" Then
 Printer.Print "�����Ŀ��   " & Text11.Text & Space(3) & "���" & Text6.Text & Space(3) & "�۸񣺴�����     ���۽�________ "
Else
 Printer.Print "�����Ŀ��" & DataCombo3.Text & Space(6) & "���" & Text6.Text & Space(6) & "�۸�" & Text7.Text & "Ԫ"
End If

Printer.Line (1, 19.5)-(17, 19.5)
Printer.CurrentX = 1
Printer.CurrentY = 20
Printer.Print "�Էѽ�   " & Space(30) & "�Է��ܽ�"
Printer.CurrentX = 1
Printer.CurrentY = 21
Printer.Print "������ң�" & Label24.Caption & Space(2) & "����ҽʦ��" & Label22.Caption & Space(5) & "����Ա��" & Space(10) & "�շ�Ա��"
Printer.CurrentX = 4
Printer.CurrentY = 22
Printer.Print "����ʱ�䣺" & Now
Printer.CurrentX = 1
Printer.CurrentY = 23
Printer.FontSize = 14
Printer.FontBold = True
Printer.Print " ע�����º���Ч��"

Printer.EndDoc
End Sub


Private Sub Command13_Click()
Adodc7.Recordset.Delete adAffectCurrent
End Sub

Private Sub Command14_Click()
Adodc7.Recordset.Update
End Sub

Private Sub Command15_Click()
If Text14.Text = "" Then
MsgBox " ������Ϣ����Ϊ�գ�����д�����ԣ�", vbExclamation, "����"
Exit Sub
Else

Dim RS As ADODB.Recordset
Set RS = Adodc6.Recordset
    Dim RSs As ADODB.Recordset
    Set RSs = Adodc7.Recordset
    On Error Resume Next
          RS.AddNew
          RS!���߱�� = Label19.Caption
          RS!�������� = Label2.Caption
          RS!��� = Text1.Text
          RS!���� = RichTextBox3.Text
          RS!ҽ������ = Combo1.Text
          RS!����ҽʦ = Label22.Caption
          RS!�������� = Date
          RS!����ʱ�� = Time
          RS!��ַ�� = Text14.Text
           RS.Update
           If Combo1.Text = "����" Then
           SSTab2.Tab = 0
          Call SSTab2_DblClick
           Text2.SetFocus
           End If
           If Combo1.Text = "���" Then
           SSTab2.Tab = 2
           DataCombo1.SetFocus
           End If
           If Combo1.Text = "סԺ" Then
           DataCombo4.SetFocus
           End If
 End If
End Sub

Private Sub Command16_Click()
Adodc2.Recordset.Delete
End Sub

Private Sub Command2_Click()
On Error Resume Next
cc = Format(Date, "YYYYMMDD")

          Adodc4.Recordset.AddNew
          Adodc4.Recordset.Fields("��ˮ��") = cc & Label19.Caption
          Adodc4.Recordset.Fields("��������") = Label2.Caption
          Adodc4.Recordset.Fields("�Ա�") = Label3.Caption
          Adodc4.Recordset.Fields("����") = Label5.Caption
          'Adodc4.Recordset.Fields("�����") = Label28.Caption
          Adodc4.Recordset.Fields("������") = DataCombo1.Text
          Adodc4.Recordset.Fields("�����Ŀ") = DataCombo3.Text + Text11.Text
          Adodc4.Recordset.Fields("�۸�") = Text7.Text
          Adodc4.Recordset.Fields("��λ") = Text6.Text
          Adodc4.Recordset.Fields("��������") = Date
          Adodc4.Recordset.Fields("����ʱ��") = Time
          Adodc4.Recordset.Fields("�������") = Label24.Caption
          Adodc4.Recordset.Fields("����ҽʦ") = Label22.Caption
          Adodc4.Recordset.Fields("״̬") = "���շ�"
          Adodc4.Recordset.Update
          DataGrid8.Refresh
End Sub

Private Sub Command3_Click()
Adodc4.Recordset.Delete
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Not (Combo2.Text = "ʹ�÷���" And Combo3.Text = "����" And Combo6.Text = "��λ") Then
cc = Format(Date, "YYYYMMDD")
dd = Adodc1.Recordset.RecordCount + 1
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("���") = dd
Adodc1.Recordset.Fields("��ˮ��") = Label19.Caption
Adodc1.Recordset.Fields("��������") = Label2.Caption
Adodc1.Recordset.Fields("�Ա�") = Label3.Caption
Adodc1.Recordset.Fields("����") = Label5.Caption
Adodc1.Recordset.Fields("ҩƷ����") = Text3.Text
Adodc1.Recordset.Fields("����") = Text5.Text
Adodc1.Recordset.Fields("��λ") = Combo6.Text

If Text12.Text = "" Then
Adodc1.Recordset.Fields("����") = "��˵������"
Else
Adodc1.Recordset.Fields("����") = Text12.Text & Combo4.Text
End If

Adodc1.Recordset.Fields("���") = TextBox1(12).Text
Adodc1.Recordset.Fields("����") = TextBox1(14).Text
Adodc1.Recordset.Fields("���") = Text4.Text
Adodc1.Recordset.Fields("״̬") = "���շ�"
Adodc1.Recordset.Fields("�÷�") = Combo2.Text & Combo3.Text
Adodc1.Recordset.Fields("����") = Label24.Caption
Adodc1.Recordset.Fields("ҽ��") = Label22.Caption
Adodc1.Recordset.Fields("����") = Date
Adodc1.Recordset.Fields("ʱ��") = Time
DataGrid2.Refresh
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Text8.Text = ""
TextBox1(14).Text = ""
TextBox1(12).Text = ""
Text12.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Text4.Text = ""
Command8_Click       '���洦��
Text2.SetFocus
Set DataGrid1.DataSource = Nothing
DataGrid1.Refresh
Else
Label26.Caption = "û��ѡ��ʹ�÷�����ʹ�ô�����"
End If
End Sub

Private Sub Command6_Click()

Adodc2.CursorLocation = adUseClient
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields("���߱��") = Label19.Caption
Adodc2.Recordset.Fields("����") = Label2.Caption
Adodc2.Recordset.Fields("�Ա�") = Label3.Caption
Adodc2.Recordset.Fields("סԺ��") = DataCombo4.Text
Adodc2.Recordset.Fields("���") = Text1.Text
Adodc2.Recordset.Fields("����") = Label5.Caption
Adodc2.Recordset.Fields("״̬") = "���Ŵ�"
Adodc2.Recordset.Update
DataGrid5.Refresh
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Text1.Text = "" Then
MsgBox "��ϲ���Ϊ�գ�����д�����ԣ�", vbInformation
Exit Sub
End If

If Text14.Text = "" Then
MsgBox "����鲻��Ϊ�գ�����д�����ԣ�", vbInformation
Exit Sub
End If

    Printer.ScaleMode = vbCentimeters
    Printer.Orientation = 2
    Printer.PaperSize = 13
    Printer.PaintPicture Image1.Picture, 0, 0, 13, 20
    Printer.PaintPicture Image1.Picture, 12, 0, 13, 20
    Printer.Line (12.5, 0)-(12.5, 20)
    
'For i = 0 To 17
  ' Printer.CurrentX = 0
'   Printer.CurrentY = i
  ' Printer.Print i
 ' Printer.CurrentX = i
 '   Printer.CurrentY = 0
 '   Printer.Print i
 '  Next i

            Printer.FontSize = 12
            Printer.CurrentX = 2.4
            Printer.CurrentY = 4.1
            Printer.Print Label2.Caption
             
            Printer.CurrentX = 8.5
            Printer.CurrentY = 4
            Printer.Print Label3.Caption
              
            Printer.CurrentX = 10
            Printer.CurrentY = 4
            Printer.Print Label5.Caption
            
            If Val(Label5.Caption) <= 14 Then
            Printer.Line (9, 1)-(11, 1)
            Printer.Line (9, 2)-(11, 2)
            Printer.Line (9, 1)-(9, 2)
            Printer.Line (11, 1)-(11, 2)
            Printer.CurrentX = 9.3
            Printer.CurrentY = 1.2
            Printer.FontSize = 16
            Printer.FontBold = True
            Printer.Print "����"
            Printer.FontBold = False
            Printer.FontSize = 12
            Else
            End If
            
            If Len(Label2.Caption) >= 3 Then
            Printer.CurrentX = 11.3
            Printer.CurrentY = 4
            Printer.Print "ά"
            End If
            
            Printer.CurrentX = 9
            Printer.CurrentY = 2.6
            Printer.Print Date
            
            Printer.CurrentX = 7
            Printer.CurrentY = 5
            Printer.Print Label7.Caption
           Printer.CurrentX = 8.4
           Printer.CurrentY = 5.9
           Printer.Print Label19.Caption
           Printer.CurrentX = 2
           Printer.CurrentY = 5.9
           Printer.Print Label24.Caption
           
           Printer.CurrentX = 4
           Printer.CurrentY = 6.5
           Printer.Print Text1.Text
           
           Printer.CurrentX = 9
           Printer.CurrentY = 6.5
           Printer.Print Text13.Text & "-" & Text14.Text
           
           Printer.CurrentX = 1
           Printer.CurrentY = 9
           Adodc1.Recordset.MoveFirst
           i = Adodc1.Recordset.RecordCount
           
            Printer.CurrentX = 8
           Printer.CurrentY = 15.7
           Printer.Print Label22.Caption
           '*****************************************************************
           Printer.FontSize = 12
            Printer.CurrentX = 2.4 + 12
            Printer.CurrentY = 4.1
            Printer.Print Label2.Caption
             
            Printer.CurrentX = 8.5 + 12
            Printer.CurrentY = 4
            Printer.Print Label3.Caption
              
            Printer.CurrentX = 10 + 12
            Printer.CurrentY = 4
            Printer.Print Label5.Caption
            
            If Val(Label5.Caption) <= 14 Then
            Printer.Line (9 + 12, 1)-(11 + 12, 1)
            Printer.Line (9 + 12, 2)-(11 + 12, 2)
            Printer.Line (9 + 12, 1)-(9 + 12, 2)
            Printer.Line (11 + 12, 1)-(11 + 12, 2)
            Printer.CurrentX = 9.3 + 12
            Printer.CurrentY = 1.2
            Printer.FontSize = 16
            Printer.FontBold = True
            Printer.Print "����"
            Printer.FontBold = False
            Printer.FontSize = 12
            Else
            End If
            
            If Len(Label2.Caption) >= 3 Then
            Printer.CurrentX = 11.3 + 12
            Printer.CurrentY = 4
            Printer.Print "ά"
            End If
            
            Printer.CurrentX = 9 + 12
            Printer.CurrentY = 2.6
            Printer.Print Date
            
            Printer.CurrentX = 7 + 12
            Printer.CurrentY = 5
            Printer.Print "  " & Label7.Caption
            
           Printer.CurrentX = 8.4 + 12
           Printer.CurrentY = 5.9
           Printer.Print Label19.Caption
           
           Printer.CurrentX = 9 + 12
           Printer.CurrentY = 6.5
           Printer.Print Text13.Text & "-" & Text14.Text
           Printer.CurrentX = 2 + 12
           Printer.CurrentY = 5.9
           Printer.Print Label24.Caption
           
           Printer.CurrentX = 4.5 + 12
           Printer.CurrentY = 6.5
           Printer.Print Text1.Text
           
           Printer.CurrentX = 1 + 12
           Printer.CurrentY = 9
           Adodc1.Recordset.MoveFirst
           i = Adodc1.Recordset.RecordCount
           
            Printer.CurrentX = 8 + 12
           Printer.CurrentY = 15.7
           Printer.Print Label22.Caption
           
           
           For cc = 0 To i Step 1.3
           dd = 9 + Val(cc)
           Printer.CurrentX = 1
           Printer.CurrentY = dd
           Printer.FontSize = 14
           Printer.FontBold = True
           Printer.Print Adodc1.Recordset.Fields("ҩƷ����") & Space(2) & Adodc1.Recordset.Fields("���") & "*" & Adodc1.Recordset.Fields("����") & Adodc1.Recordset.Fields("��λ")
           Printer.CurrentX = 1
           Printer.CurrentY = dd + 0.6
           Printer.FontBold = False
           Printer.FontSize = 12
           Printer.Print "   �÷���" & Adodc1.Recordset.Fields("�÷�") & Space(5) & "ÿ�μ�����" & Adodc1.Recordset.Fields("����")
           Printer.CurrentX = 14
           Printer.CurrentY = dd
            Printer.FontSize = 14
             Printer.FontBold = True
           Printer.Print Adodc1.Recordset.Fields("ҩƷ����") & Space(2) & Adodc1.Recordset.Fields("���") & "*" & Adodc1.Recordset.Fields("����") & Adodc1.Recordset.Fields("��λ")
           Printer.CurrentX = 14
           Printer.CurrentY = dd + 0.6
           Printer.FontSize = 12
           Printer.FontBold = False
           Printer.Print "   �÷���" & Adodc1.Recordset.Fields("�÷�") & Space(5) & "ÿ�μ�����" & Adodc1.Recordset.Fields("����")
           
           Adodc1.Recordset.MoveNext
           
           If Adodc1.Recordset.EOF = True Then
           Printer.EndDoc
           Exit Sub
           End If
           Next cc
           '################################################################
           
           
End Sub

Private Sub Command8_Click()
On Error Resume Next
Adodc1.Recordset.Update
End Sub

Private Sub Command9_Click()
On Error Resume Next
If (Text1.Text = "" And Text14.Text = "") Then
MsgBox "��Ϻ͵�ַ��Ϣ����Ϊ�գ�"
Exit Sub
End If
 Printer.PaperSize = 13   'vbPRPSB5 13 B5, 182 x 257 mm
 Printer.ScaleMode = vbCentimeters
 Printer.FontSize = 16
 Printer.Orientation = 1
 Printer.FontBold = True
Printer.CurrentX = 6
Printer.CurrentY = 1
Printer.Print "ɯ���ػĵ�������Ժ"
Printer.CurrentX = 7
Printer.CurrentY = 2
 Printer.FontSize = 12
 If DataCombo1.Text = "������" Then
Printer.Print "�����Ҽ�����뵥"

End If
If DataCombo1.Text = "B����" Then
Printer.Print "��ɫB��������뵥"

End If
If DataCombo1.Text = "ҽѧӰ����" Then
Printer.Print "X�߼�����뵥" '

End If
Printer.FontBold = False

Printer.Line (1, 2.8)-(17, 2.8)
Printer.Line (1, 2.85)-(17, 2.85)


Printer.CurrentX = 1
Printer.CurrentY = 3

Printer.Print "����ҽ�ƺţ�" & Label19.Caption & Space(2) & "����������" & Label2.Caption & Space(2) & "�Ա�" & Label3.Caption & Space(2) & "���䣺" & Label5.Caption
Printer.CurrentX = 1
Printer.CurrentY = 4
Printer.Print "���㷽ʽ��" & Label7.Caption & Space(10) & "��ַ��" & Text13.Text & "-" & Text14.Text & Space(10) & "��ϣ�" & Text1.Text
Printer.CurrentX = 1
Printer.CurrentY = 5
If DataCombo3.Text = "������Ŀ" Then

 Printer.Print "�����Ŀ��   " & Text11.Text & Space(3) & "���" & Text6.Text & Space(3) & "�۸񣺴�����     ���۽�________ "
Else
Printer.Print "�����Ŀ��" & DataCombo3.Text & Space(10) & "���" & Text6.Text & Space(10) & "�۸�" & Text7.Text & "Ԫ"
End If
Printer.Line (1, 5.5)-(17, 5.5)
Printer.CurrentX = 1
Printer.CurrentY = 6
Printer.Print "�ܽ�" & Space(15) & "����ҽ�Ʊ�����" & Space(15) & "�Էѽ�"
Printer.CurrentX = 1
Printer.CurrentY = 7
Printer.Print "������ң�" & Label24.Caption & Space(2) & "����ҽʦ��" & Label22.Caption & Space(5) & "����Ա��" & Space(10) & "�շ�Ա��"
Printer.CurrentX = 4
Printer.CurrentY = 8
Printer.Print "����ʱ�䣺" & Now
Printer.CurrentX = 4
Printer.CurrentY = 10
Printer.FontSize = 14
Printer.FontBold = True
Printer.Print " ע�����º���Ч��"

Printer.FontSize = 16
 Printer.Orientation = 1
Printer.CurrentX = 6
Printer.CurrentY = 15
Printer.Print "ɯ���ػĵ�������Ժ"
Printer.CurrentX = 7
Printer.CurrentY = 16
 Printer.FontSize = 12
 If DataCombo1.Text = "������" Then
Printer.Print "�����Ҽ�����뵥"
End If

If DataCombo1.Text = "B����" Then
Printer.Print "��ɫB��������뵥"
End If

If DataCombo1.Text = "ҽѧӰ����" Then
Printer.Print "X�߼�����뵥" '
End If
Printer.FontBold = False

Printer.Line (1, 16.8)-(17, 16.8)
Printer.Line (1, 16.85)-(17, 16.85)


Printer.CurrentX = 1
Printer.CurrentY = 17

Printer.Print "����ҽ�ƺţ�" & Label19.Caption & Space(2) & "����������" & Label2.Caption & Space(2) & "�Ա�" & Label3.Caption & Space(2) & "���䣺" & Label5.Caption
Printer.CurrentX = 1
Printer.CurrentY = 18
Printer.Print "���㷽ʽ��" & Label7.Caption & Space(10) & "��ַ��" & Text13.Text & "-" & Text14.Text & Space(10) & "��ϣ�" & Text1.Text
Printer.CurrentX = 1
Printer.CurrentY = 19
If DataCombo3.Text = "������Ŀ" Then
Printer.Print "�����Ŀ��    " & Text11.Text & Space(3) & "���" & Text6.Text & Space(3) & "�۸񣺴�����     ���۽�________ "
Else
Printer.Print "�����Ŀ��" & DataCombo3.Text & Space(10) & "���" & Text6.Text & Space(10) & "�۸�" & Text7.Text & "Ԫ"
End If

Printer.Line (1, 19.5)-(17, 19.5)
Printer.CurrentX = 1
Printer.CurrentY = 20
Printer.Print "�ܽ�" & Space(15) & "����ҽ�Ʊ�����" & Space(15) & "�Էѽ�"
Printer.CurrentX = 1
Printer.CurrentY = 21
Printer.Print "������ң�" & Label24.Caption & Space(2) & "����ҽʦ��" & Label22.Caption & Space(5) & "����Ա��" & Space(10) & "�շ�Ա��"
Printer.CurrentX = 4
Printer.CurrentY = 22
Printer.Print "����ʱ�䣺" & Now
Printer.CurrentX = 4
Printer.CurrentY = 23
Printer.FontSize = 14
Printer.FontBold = True
Printer.Print " ע�����º���Ч��"

Printer.EndDoc
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from �����Ŀ where ��������='" & DataCombo1.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set DataCombo3.RowSource = Mrc
DataCombo3.Text = "��ѡ������Ŀ"
End Sub

Private Sub DataCombo2_Change()
'DataCombo3.Text = DataCombo2.Text
'DataCombo3.ListField = DataCombo2.Text
End Sub



Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DataCombo3.SetFocus
End If
End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2.SetFocus
End If
End Sub

Private Sub DataCombo3_LostFocus()
If DataCombo3.Text = "������Ŀ" Then
Text11.Visible = True
Set Text7.DataSource = Nothing
Set Text6.DataSource = Nothing
Text11.SetFocus
Else
Text11.Visible = False
End If
On Error Resume Next
If Not DataCombo3.Text = "������Ŀ" Then
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from �����Ŀ where ��������='" & DataCombo1.Text & "'and ��Ŀ����='" & DataCombo3.Text & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc5.Recordset = Mrc
Set DataGrid6.DataSource = Mrc
DataGrid6.Refresh
Set Text6.DataSource = Mrc
Text6.DataField = "��λ"
Set Text7.DataSource = Mrc
Text7.DataField = "�۸�"
End If
End Sub

Private Sub DataGrid1_Click()
Text3.Text = DataGrid1.Columns("ҩƷ��").CellValue(DataGrid1.Bookmark)
TextBox1(12).Text = DataGrid1.Columns("���").CellValue(DataGrid1.Bookmark)
TextBox1(14).Text = DataGrid1.Columns("����").CellValue(DataGrid1.Bookmark)
'Combo6.Text = DataGrid1.Columns("��λ").CellValue(DataGrid1.Bookmark)

End Sub

Private Sub DataGrid3_Click()
If Text1.Text = "" Then

Text1.Text = DataGrid3.Columns("���").CellValue(DataGrid3.Bookmark)
Else
Text1.Text = Text1.Text & "," & DataGrid3.Columns("���").CellValue(DataGrid3.Bookmark)
End If
RichTextBox3.Text = DataGrid3.Columns("����").CellValue(DataGrid3.Bookmark)
End Sub


Private Sub DataGrid4_Click()
On Error Resume Next
Text1.Text = DataGrid4.Columns("���").CellText(DataGrid4.Bookmark)
RichTextBox3.Text = DataGrid4.Columns("����").CellText(DataGrid4.Bookmark)
Combo1.Text = DataGrid4.Columns("ҽ������").CellText(DataGrid4.Bookmark)
Text14.Text = DataGrid4.Columns("��").CellValue(DataGrid4.Bookmark)

End Sub

Private Sub DataGrid4_DblClick()
On Error Resume Next
��鵥��ӡ.Show
With ��鵥��ӡ
.Label4.Item(0).Caption = DataGrid4.Columns("���߱��").CellText(DataGrid4.Bookmark)
.Label4.Item(1).Caption = DataGrid4.Columns("���").CellText(DataGrid4.Bookmark)
.Label4.Item(2).Caption = DataGrid4.Columns("����").CellText(DataGrid4.Bookmark)
.Label4.Item(8).Caption = DataGrid4.Columns("ҽ������").CellText(DataGrid4.Bookmark)
.Label4.Item(9).Caption = DataGrid4.Columns("����ҽʦ").CellText(DataGrid4.Bookmark)
.Label4.Item(10).Caption = DataGrid4.Columns("��������").CellText(DataGrid4.Bookmark)
.Label4.Item(11).Caption = DataGrid4.Columns("����ʱ��").CellText(DataGrid4.Bookmark)
 End With
 
End Sub

Private Sub DataGrid8_DblClick()
On Error Resume Next
��鵥���.Show
With ��鵥���
.Label4.Item(0).Caption = DataGrid8.Columns("��ˮ��").CellText(DataGrid4.Bookmark)
.Label4.Item(1).Caption = DataGrid8.Columns("�����Ŀ").CellText(DataGrid4.Bookmark)
.Label4.Item(2).Caption = DataGrid8.Columns("�����").CellText(DataGrid4.Bookmark)
.Label4.Item(3).Caption = DataGrid8.Columns("������").CellText(DataGrid4.Bookmark)
.Label4.Item(4).Caption = DataGrid8.Columns("������").CellText(DataGrid4.Bookmark)
.Label4.Item(5).Caption = DataGrid8.Columns("�������").CellText(DataGrid4.Bookmark)
.Label4.Item(6).Caption = DataGrid8.Columns("���ʱ��").CellText(DataGrid4.Bookmark)
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set skinH_VB6 = Nothing
End Sub


Private Sub Label19_Change()
SSTab2.Visible = True
Text13.Text = Mid(Label19.Caption, 3, 2)
Text1.Text = ""
RichTextBox3.Text = ""
Combo1.Text = ""
End Sub

Private Sub RichTextBox1_GotFocus()
With RichTextBox1
.SelStart = 0
.SelLength = Len(.Text)
End With
Set DataList1.RowSource = Adodc2
   DataList1.ListField = "�ֲ�ʷ"
   DataList1.ReFill
   DataList1.Refresh
End Sub

Private Sub RichTextBox2_GotFocus()
Set DataList1.RowSource = Adodc2
   DataList1.ListField = "����ʷ"
   DataList1.ReFill
   DataList1.Refresh
End Sub
Private Sub RichTextBox5_GotFocus()
'Adodc2.Recordset.Close
DataList1.Visible = False
End Sub

Private Sub Label2_Change()
SSTab2.Visible = True
Text13.Text = Mid(Label19.Caption, 3, 2)
Text1.Text = ""
RichTextBox3.Text = ""
Combo1.Text = ""
End Sub


Private Sub Label7_Change()
If Label7.Caption = "�Է�" Then
Command12.Visible = True
Command9.Visible = False
End If
If Label7.Caption = "����ҽ��" Then
Command9.Visible = True
Command12.Visible = False
End If
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)

Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Dim conn As ADODB.Connection
Dim mrcc As ADODB.Recordset

'If SSTab2.Tab = 0 Then
'Text2.Text = ""
'Text3.Text = ""
'Text5.Text = ""
'Set Con = New ADODB.Connection
'Set Mrc = New ADODB.Recordset
'Con.CursorLocation = adUseClient
'Con.Open SQL
'Mrc.Open "select * from ���ﴦ�� where ��ˮ�� like'%" & Label19.Caption & "%'and ����='" & Date & "'", Con, adOpenKeyset, adLockOptimistic
'Set Adodc1.Recordset = Mrc
'Set DataGrid2.DataSource = Mrc
'Text9.Text = Adodc1.Recordset.RecordCount
'End If

If SSTab2.Tab = 1 Then
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from ���ﲡ�� where ���߱�� ='" & Label19.Caption & "'and ��������='" & Label2.Caption & "'", Con, adOpenKeyset, adOpenDynamic
Set Adodc6.Recordset = Mrc
Set DataGrid4.DataSource = Mrc

Set conn = New ADODB.Connection
Set mrcc = New ADODB.Recordset
Con.CursorLocation = adUseClient
conn.Open SQL
mrcc.Open "select * from סԺ�� where ���߱�� ='" & Label19.Caption & "'and ����='" & Label2.Caption & "'", conn, adOpenKeyset, adOpenDynamic
Set Adodc2.Recordset = mrcc
Set DataGrid5.DataSource = mrcc
End If

If SSTab2.Tab = 2 Then
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ��鵥 where ��ˮ�� like'%" & Label19.Caption & "%'and ��������='" & Label2.Caption & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc4.Recordset = Mrc
Set DataGrid8.DataSource = Mrc
Label27.Caption = "��¼����  " & Mrc.RecordCount
End If
End Sub

Private Sub SSTab2_DblClick()
On Error Resume Next
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset

If SSTab2.Tab = 2 Then
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ��鵥 where ��ˮ�� like'%" & Label19.Caption & "%'and ��������='" & Label2.Caption & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc4.Recordset = Mrc
Set DataGrid8.DataSource = Mrc
Label27.Caption = "��¼����  " & Mrc.RecordCount
End If

If SSTab2.Tab = 0 Then
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from ���ﴦ�� where ��ˮ�� like'%" & Label19.Caption & "%' and ��������='" & Label2.Caption & "'and ����='" & Date & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc1.Recordset = Mrc
Set DataGrid2.DataSource = Mrc
'Text9.Text = Adodc1.Recordset.RecordCount
End If

If SSTab2.Tab = 1 Then
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.Open SQL
Mrc.Open "select * from ���ﲡ�� where ���߱�� ='" & Label19.Caption & "'and ��������='" & Label2.Caption & "'", Con, adOpenKeyset, adOpenDynamic
Set Adodc6.Recordset = Mrc
Set DataGrid4.DataSource = Mrc
End If

End Sub

Private Sub SSTab2_GotFocus()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ��鵥 where ��ˮ��='" & Label19.Caption & "' and ��������='" & Label2.Caption & "' and ��������='" & Date & "'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid8.DataSource = Mrc
    DataGrid8.Refresh
End Sub

Private Sub Text11_Change()
Text7.Text = ""
Text6.Text = ""
End Sub

Private Sub Text12_GotFocus()
On Error Resume Next
If Val(Text5.Text) >= Val(Text8.Text) Then
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
Label26.Caption = "ҩƷ�������ܳ��������������ϵҩ����Ա��"
Else
Text4.Text = Val(Text5.Text) * Val(TextBox1(14).Text)

End If
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) >= 2 Then
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ҩƷ��� where ������ like'%" & Text2.Text & "%'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = Mrc
Set Text3.DataSource = Mrc
Set Text8.DataSource = Mrc
Set TextBox1(12).DataSource = Mrc
Set TextBox1(14).DataSource = Mrc
Else
End If
If Text2.Text = " " Then
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
SQL = "Provider=sqloledb.1;Data Source=NURA\SQLEXPRESS;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ҩ��", Con, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = Mrc
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
End Sub


Private Sub Text3_Change()
Text5.Text = "1"
End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub Text5_LostFocus()
On Error Resume Next
If Val(Text5.Text) >= Val(Text8.Text) Then
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
Label26.Caption = "ҩƷ�������ܳ��������������ϵҩ����Ա��"
Else
Text4.Text = Val(Text5.Text) * Val(TextBox1(14).Text)
Label26.Caption = ""
End If
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "ɾ��"
        Adodc1.Recordset.Delete
            Adodc1.Recordset.Delete
        Case "��ӡ����"
        
           
            MsgBox "��� 'ɾ����' ��ť���롣"
        Case "x28"
            'Ӧ��:��� 'x28' ��ť���롣
            MsgBox "��� 'x28' ��ť���롣"
        Case "refresh"
        DataGrid2.Refresh
        
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "application"
            'Ӧ��:��� 'application' ��ť���롣
            MsgBox "��� 'application' ��ť���롣"
        Case "����"
            'Ӧ��:��� '����' ��ť���롣
            MsgBox "��� '����' ��ť���롣"
        Case "��ϸ����"
            'Ӧ��:��� '��ϸ����' ��ť���롣
            MsgBox "��� '��ϸ����' ��ť���롣"
        Case "�����"
            'Ӧ��:��� '�����' ��ť���롣
            MsgBox "��� '�����' ��ť���롣"
    End Select
End Sub

Private Sub Command1_Click()
������ѯ.Show
End Sub
Private Sub Form_Load()
On Error Resume Next
Me.Width = 15000
Me.Height = 11800
Text1.Text = ""
RichTextBox3.Text = ""
Combo1.Text = ""

   Label24.Caption = MDIForm1.StatusBar1.Panels(4)
   Label22.Caption = MDIForm1.StatusBar1.Panels(3)
   SSTab2.Visible = False
    End Sub
