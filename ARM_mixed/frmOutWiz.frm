VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOutWiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   Icon            =   "frmOutWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameSave 
      Height          =   4455
      Left            =   120
      TabIndex        =   88
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label lblBlinker 
         Alignment       =   2  'Center
         Caption         =   "���� �����- ����� ������ � CORE IMS. �����."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4095
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "���6 -������� ����������"
      Height          =   7695
      Left            =   2400
      TabIndex        =   76
      Top             =   480
      Width           =   14175
      Begin GridEX20.GridEX gr2 
         Height          =   5835
         Left            =   240
         TabIndex        =   77
         Top             =   360
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   10292
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ItemCount       =   0
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   1
         Column(1)       =   "frmOutWiz.frx":030A
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmOutWiz.frx":036E
         FormatStyle(2)  =   "frmOutWiz.frx":044E
         FormatStyle(3)  =   "frmOutWiz.frx":05AA
         FormatStyle(4)  =   "frmOutWiz.frx":065A
         FormatStyle(5)  =   "frmOutWiz.frx":070E
         FormatStyle(6)  =   "frmOutWiz.frx":07E6
         FormatStyle(7)  =   "frmOutWiz.frx":089E
         ImageCount      =   0
         PrinterProperties=   "frmOutWiz.frx":08BE
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "��� 7 - ���� ����� � ������"
      Height          =   5535
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   21135
      Begin VB.CommandButton cmdPrnRas 
         Caption         =   "��� ������� �����������"
         Height          =   495
         Left            =   2880
         TabIndex        =   85
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CommandButton cmd6PrnKL 
         Caption         =   "������ ����������� �����"
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton cmd6PRNSRV 
         Caption         =   "������ ��������� �� ������"
         Height          =   495
         Left            =   5280
         TabIndex        =   36
         Top             =   4440
         Width           =   2775
      End
      Begin VSFlex8Ctl.VSFlexGrid srvGrid 
         Height          =   3975
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   7815
         _cx             =   13785
         _cy             =   7011
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   600
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOutWiz.frx":0A96
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   1
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "���4 - ����������� ����� ��� �������"
      Height          =   6255
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   10815
      Begin VB.CommandButton cmd6FindCell 
         Caption         =   "..."
         Height          =   375
         Left            =   5160
         TabIndex        =   81
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox txt4NewPlace 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   4560
         Width           =   4935
      End
      Begin VB.TextBox Txt4PackageWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox txt4Quantity 
         Height          =   375
         Left            =   2760
         TabIndex        =   47
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txt4InQry 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txt4FromQ 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txt4Good 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox txt4FullWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txt4GoodWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "��� ����� �����"
         Height          =   375
         Left            =   2760
         TabIndex        =   84
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "�������� ������ ��� ��������"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   4200
         Width           =   5055
      End
      Begin VB.Label Label21 
         Caption         =   "��� ����� ��������"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "���������� �������"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   46
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "���� ���������, ��������������"
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "�����"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "��� ����������� �������"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��� 1 - ����� ������"
      Height          =   6975
      Left            =   7920
      TabIndex        =   0
      Top             =   2880
      Width           =   7455
      Begin VB.TextBox txtSupplier 
         Height          =   300
         Left            =   120
         MaxLength       =   255
         TabIndex        =   59
         ToolTipText     =   "���������"
         Top             =   2220
         Width           =   3000
      End
      Begin VB.TextBox txtTTN 
         Height          =   300
         Left            =   120
         MaxLength       =   30
         TabIndex        =   58
         ToolTipText     =   "����� ���"
         Top             =   2925
         Width           =   3000
      End
      Begin VB.TextBox txtTranspNumber 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   56
         ToolTipText     =   "� ��"
         Top             =   4335
         Width           =   3000
      End
      Begin VB.TextBox txtContainer 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   55
         ToolTipText     =   "� ������� \ ����������"
         Top             =   5040
         Width           =   3000
      End
      Begin VB.TextBox txtStampNumber 
         Height          =   300
         Left            =   120
         MaxLength       =   20
         TabIndex        =   54
         ToolTipText     =   "����� ������"
         Top             =   5745
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.TextBox txtStampStatus 
         Height          =   300
         Left            =   120
         MaxLength       =   30
         TabIndex        =   53
         ToolTipText     =   "��������� ������"
         Top             =   6450
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.TextBox txtTheClient 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1440
         Width           =   6615
      End
      Begin VB.TextBox txtShipOrder 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "��� ������"
         Top             =   690
         Width           =   6015
      End
      Begin MTZ_PANEL.DropButton cmdShipOrder 
         Height          =   300
         Left            =   6240
         TabIndex        =   2
         Tag             =   "refopen.ico"
         ToolTipText     =   "��� ������"
         Top             =   690
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin MSMask.MaskEdBox txttemp_in_track 
         Height          =   300
         Left            =   3390
         TabIndex        =   50
         ToolTipText     =   "�����������"
         Top             =   3660
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   27
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtptrack_time_out 
         Height          =   300
         Left            =   3390
         TabIndex        =   51
         ToolTipText     =   "����� ������ ������"
         Top             =   2955
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd.MM.yyyy HH:mm:ss"
         Format          =   51970051
         CurrentDate     =   39006
      End
      Begin MSComCtl2.DTPicker dtpTrack_time_in 
         Height          =   300
         Left            =   3390
         TabIndex        =   52
         ToolTipText     =   "����� �������� ������"
         Top             =   2250
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd.MM.yyyy HH:mm:ss"
         Format          =   51970051
         CurrentDate     =   39006
      End
      Begin MSComCtl2.DTPicker dtpTTNDate 
         Height          =   300
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "���� ���"
         Top             =   3630
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd.MM.yyyy"
         Format          =   51970051
         CurrentDate     =   39006
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "����������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   69
         Top             =   1890
         Width           =   3000
      End
      Begin VB.Label lblTTN 
         BackStyle       =   0  'Transparent
         Caption         =   "����� ���:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   68
         Top             =   2595
         Width           =   3000
      End
      Begin VB.Label lblTTNDate 
         BackStyle       =   0  'Transparent
         Caption         =   "���� ���:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   67
         Top             =   3300
         Width           =   3000
      End
      Begin VB.Label lblTranspNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "� ��:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   66
         Top             =   4005
         Width           =   3000
      End
      Begin VB.Label lblContainer 
         BackStyle       =   0  'Transparent
         Caption         =   "� ������� \ ����������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   65
         Top             =   4710
         Width           =   3000
      End
      Begin VB.Label lblStampNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "����� ������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   64
         Top             =   5415
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label lblStampStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "��������� ������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   63
         Top             =   6120
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label lblTrack_time_in 
         BackStyle       =   0  'Transparent
         Caption         =   "����� �������� ������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3390
         TabIndex        =   62
         Top             =   1920
         Width           =   3000
      End
      Begin VB.Label lbltrack_time_out 
         BackStyle       =   0  'Transparent
         Caption         =   "����� ������ ������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3390
         TabIndex        =   61
         Top             =   2625
         Width           =   3000
      End
      Begin VB.Label lbltemp_in_track 
         BackStyle       =   0  'Transparent
         Caption         =   "�����������:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3390
         TabIndex        =   60
         Top             =   3330
         Width           =   3000
      End
      Begin VB.Label Label14 
         Caption         =   "������"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblQryCode 
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3000
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��� 3 - ������ � ������"
      Height          =   6375
      Left            =   1680
      TabIndex        =   16
      Top             =   240
      Width           =   7335
      Begin VB.Frame frameWait 
         Height          =   2775
         Left            =   1080
         TabIndex        =   86
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "���� �������� ����������� �������� ������ � �������� �������."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   240
            TabIndex        =   87
            Top             =   480
            Width           =   4575
         End
      End
      Begin VB.TextBox txtMainCell 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox txt3PackageWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox txt3Quantity 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox txt3FRomQ 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmd3ClearW 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         ToolTipText     =   "�������� ��� �  �����"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txt3InQry 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txt3GoodWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txt3FullWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txt3PWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txt3PNum 
         Height          =   405
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txt3Good 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   5535
      End
      Begin VB.CommandButton cmd3ClearNum 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         ToolTipText     =   "������ ����� ��� ���"
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "������ ��������� ��������"
         Height          =   375
         Left            =   2760
         TabIndex        =   82
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Label Label20 
         Caption         =   "��� ����� ��������"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label19 
         Caption         =   "��� �������"
         Height          =   375
         Left            =   2760
         TabIndex        =   71
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "���������� �������"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "���� ���������, ��������������"
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "��� ����� �����"
         Height          =   375
         Left            =   2760
         TabIndex        =   29
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "��� ����� � ��������"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "��� �������"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "������"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "�����"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "���5 - ������ ������� �� �������������� �����"
      Height          =   5055
      Left            =   2880
      TabIndex        =   15
      Top             =   1200
      Width           =   13575
      Begin VB.Label lbl5Out 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   3255
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��� 2 - ����� ������ ������"
      Height          =   6495
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   8535
      Begin VB.CommandButton cmdToClosePage 
         Caption         =   "������� � �������� ��������"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��������"
         Height          =   255
         Left            =   3120
         TabIndex        =   70
         Top             =   360
         Width           =   2055
      End
      Begin GridEX20.GridEX gr 
         Height          =   5355
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   9446
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ItemCount       =   0
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   1
         Column(1)       =   "frmOutWiz.frx":0AFB
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmOutWiz.frx":0B5F
         FormatStyle(2)  =   "frmOutWiz.frx":0C3F
         FormatStyle(3)  =   "frmOutWiz.frx":0D9B
         FormatStyle(4)  =   "frmOutWiz.frx":0E4B
         FormatStyle(5)  =   "frmOutWiz.frx":0EFF
         FormatStyle(6)  =   "frmOutWiz.frx":0FD7
         FormatStyle(7)  =   "frmOutWiz.frx":108F
         ImageCount      =   0
         PrinterProperties=   "frmOutWiz.frx":10AF
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "���������"
      Height          =   615
      Left            =   5160
      TabIndex        =   33
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddW 
      Caption         =   "��������� ������"
      Height          =   615
      Left            =   6960
      TabIndex        =   32
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "��������"
      Height          =   615
      Left            =   3360
      TabIndex        =   31
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "�����"
      Default         =   -1  'True
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      Handshaking     =   2
   End
   Begin VB.Image imgState 
      Height          =   8580
      Left            =   0
      Picture         =   "frmOutWiz.frx":1287
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   3015
   End
End
Attribute VB_Name = "frmOutWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 655
Option Explicit
' ���� ������� ��������





Dim StepNo As Integer
Dim XMLShipOrder As String
Dim XMLTheClient As String
Dim Item As ITTOUT.Application
Dim conn As ADODB.Connection
Private curQRow As ITTOUT.ITTOUT_LINES
Private LinePal As ITTOUT.ITTOUT_PALET
Private StopWeighting As Boolean
Private wave As MTZMCI.WavePlayer
Private emu As Boolean
Private port As String
Private psetup As String
Private poddon As ITTPL_DEF
Private isFull As Boolean
Private isCalibrated As Boolean
Private QFromStock As Double

' ��������� ��� ����:ITTPL �������
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" '��������
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" '�� ������ � ������
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" '���������� � ������
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" '������
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" '�������


' ��������� ������� ������ �������
Private Sub SetBtnPos(cmd As CommandButton, ByVal pos As Integer)
  On Error Resume Next
  cmd.Left = imgState.Width + (Me.ScaleWidth - imgState.Width) / 4 * (pos - 1)
End Sub



Private Sub cmd3ClearNum_Click()
On Error Resume Next
  txt3PNum = ""
End Sub



Private Sub cmd3ClearW_Click()
On Error Resume Next
  txt3FullWeight = 0
End Sub

Private Sub cmd6Close_Click()
On Error Resume Next
  If MsgBox("������� ����� ", vbYesNo) = vbYes Then
    'Item.StatusID = "{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}"
  End If
End Sub


'������ ������ �� ����������� �����
Private Sub cmd6PrnKL_Click()
On Error Resume Next
    
    Set repShowOL = Nothing
    Set repShowOL = New ReportShow
    repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
    repShowOL.ReportFilter = " instanceid='" & Item.id & "'"
    repShowOL.ReportPath = App.Path & "\out_OL.rpt"
    repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowOL.Run True
    Set repShowOL = Nothing
End Sub

'������ ������ �� �������
Private Sub cmd6PRNSRV_Click()
    
    On Error Resume Next
    Set repShowSRVOUT = Nothing
    Set repShowSRVOUT = New ReportShow
    repShowSRVOUT.ReportSource = "V_viewITTout_ITTout_SRV"
    repShowSRVOUT.ReportFilter = " instanceid='" & Item.id & "'"
    repShowSRVOUT.ReportPath = App.Path & "\out_srvq.rpt"
    repShowSRVOUT.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowSRVOUT.Run True
    Set repShowSRVOUT = Nothing
End Sub

' ��������� ������
Private Sub cmdAddW_Click()
On Error Resume Next
    If CheckAfter Then
      StepNo = 3
      ProcessStatus
    End If
End Sub

' �����
Private Sub cmdBack_Click()
On Error Resume Next
  If CheckAfter Then
      StepNo = 2
      ProcessStatus
  End If
End Sub

' ������
Private Sub cmdCancel_Click()
On Error Resume Next
  StepNo = 8
  ProcessStatus
End Sub

'�������� ������ ������
Private Sub cmdEdit_Click()
On Error Resume Next
 'gr_DblClick
End Sub

'�����
Private Sub cmdNext_Click()
On Error Resume Next
  If CheckAfter Then
    If StepNo = 3 Then
      If MsgBox("��������� ������ ������ ������� ?", vbYesNo, "��������") = vbYes Then
        'If MsgBox("���������������� �������� ������?", vbExclamation + vbYesNo, "��������") = vbYes Then
          StepNo = 6
          isFull = True
'        Else
'          log.message "����� �� �������� �������"
'          Exit Sub
'        End If
      Else
        StepNo = 4
        isFull = False
      End If
    Else
      StepNo = StepNo + 1
    End If
    
    ProcessStatus
  End If
End Sub






'������ ���� � ������������
Private Sub cmdPrnRas_Click()
On Error Resume Next
    
    Set repShowOL = Nothing
    Set repShowOL = New ReportShow
    repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
    repShowOL.ReportFilter = " instanceid='" & Item.id & "'"
    repShowOL.ReportPath = App.Path & "\out_ras.rpt"
    repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowOL.Run True
    Set repShowOL = Nothing

End Sub

'����� ������ �� ��������
Private Sub cmdShipOrder_Click()
On Error Resume Next
  On Error Resume Next
  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtShipOrder.Tag = "") Then
    ' call MsgBox("��� ������ ��� �������")
  Else
    txtShipOrder.Tag = AddSQLRefIds(txtShipOrder.Tag, "TheClient", txtTheClient.Tag)
    txtShipOrder.Tag = Replace(txtShipOrder.Tag, "%ID%", " 1=1 ")

    Call pars.Add("xml", txtShipOrder.Tag)
  End If
  If Manager.GetCustomObjects("cliFilter").Name <> "" Then
    Call pars.Add("filter", " and " & (Manager.GetCustomObjects("cliFilter").Name))
  End If
  Set res = Manager.GetSQLDataDialog(pars)
  If (Not res Is Nothing) Then
    Dim resStr As String
    resStr = res.Item("RESULT").Value
    If (resStr = "OK") Then
      txtShipOrder.Tag = res.Item("xml").Value
      If (txtShipOrder.Text <> res.Item("brief").Value) Then
        txtShipOrder.Text = res.Item("brief").Value
        Call txtShipOrder_Change
      End If
      MakeItem
      LoadHeader Item.ITTOUT_DEF.Item(1)
    Else
      Dim errStr As String
      errStr = res.Item("ErrorDescription").Value
      If (errStr <> vbNullString) Then
       Call MsgBox("������ ����������: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
  log.message "�������� " & txtShipOrder
End Sub

' ������� � �������� ������
Private Sub cmdToClosePage_Click()
StepNo = 7
ProcessStatus
End Sub

'����� ������ ��� ��������
Private Sub cmd6FindCell_Click()
Dim PTYPE As ITTD.ITTD_PLTYPE
  
  Set PTYPE = poddon.Pltype
  Dim f As frmGetCell
  Set f = New frmGetCell
  
  
  If PTYPE.TheCode = 0 Then
    f.PTYPE = 1
  Else
    f.PTYPE = 1.25
  End If
  
  f.itemid = Manager.GetIDFromXMLField(curQRow.good_id)
  On Error Resume Next
  f.country = ""
  f.country = curQRow.made_country.Name
  f.factory = ""
  f.factory = curQRow.factory.Name
  f.killplace = ""
  f.killplace = curQRow.KILL_NUMBER.Name
  err.Clear
  
  f.Show vbModal
  If f.OK Then
    txt4NewPlace = f.OutCode
    txt4NewPlace.Tag = f.OUtID
  End If
  Unload f
  Set f = Nothing
End Sub

'����� ����
Private Sub Command2_Click()
On Error Resume Next
  txt4FullWeight = "0"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
  If UnloadMode <> 1 Then
    Cancel = -1
  Else
  wave.StopPlaying
  Set wave = Nothing
  Timer1.Enabled = False
  End If
     
End Sub




'������ ������ �� ������ �����
Private Sub srvGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
  If Col = 0 Then Exit Sub
  Item.ITTOUT_SRV.Item(Row).Quantity = MyRound(srvGrid.TextMatrix(Row, Col))
  Item.ITTOUT_SRV.Item(Row).save
End Sub

Private Sub srvGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
If Col = 0 Then Cancel = True
End Sub


'���� ������ � �����
Private Sub Timer1_Timer()
On Error Resume Next
  Dim w As Double
  If StepNo = 3 Then
    If txt3PNum = "" Then
      txt3PNum.SetFocus
    End If
    If txt3FullWeight = "0" Or Not IsNumeric(txt3FullWeight) Then
      w = GetWeight
      If w > 0 Then
        txt3FullWeight = Round(w + 0.001, 1)
        MyBeep "Gruz"
      End If
    End If
  End If
  
  
  
  
  If StepNo = 4 Then
    
    If txt4FullWeight = "0" Or Not IsNumeric(txt4FullWeight) Then
      w = GetWeight
      If w > 0 Then
        txt4FullWeight = Round(w + 0.001, 1)
        MyBeep "Gruz"
      End If
    End If
  End If
  
End Sub

Private Sub txt3FullWeight_Change()
  On Error Resume Next
'  If GetSetting("RBH", "ITTSETTINGS", "RESTORE", "False") = "False" Then
'    txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
'  Else
    If isCalibrated Then Exit Sub
    txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
'  End If
  
End Sub

Private Sub txt3PackageWeight_Change()
On Error Resume Next
'  If GetSetting("RBH", "ITTSETTINGS", "RESTORE", "False") = "False" Then
'    txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
'  Else
    If isCalibrated Then Exit Sub
    txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
'  End If
End Sub



' ������� ����� �������
Private Sub txt3PNum_Change()
On Error Resume Next
  
  Dim cm As Boolean
  If GetSetting("RBH", "ITTSETTINGS", "MOROZ", "False") = "False" Then
    cm = False
  Else
    cm = True
  End If
   
  Dim conn As ADODB.Connection
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset
  
  'poddon.CurrentGood
  
   
  
  'log.Message "�������� ������� " & poddon.TheNumber
  
  If CheckPoddon(cm) Then
    
   'log.Message "�������� �������� ������� " & poddon.TheNumber
   
   Set rs = conn.Execute("select * from stock where PALLET_STATUS is null and pallet_id=" & poddon.CorePalette_ID)
  
   If GetSetting("RBH", "ITTSETTINGS", "RESTORE", "False") = "False" Then
    
    
    txt3PWeight = poddon.Weight
    txt3FullWeight = 0
    txt3Quantity = rs!custom_field1
    txt3PackageWeight = rs!custom_field3
    cmd3ClearW.Enabled = True
    If isCalibrated Then
      'log.Message "�������������� ���� �� ���� ��� �������������� ������!!!"
      txt3GoodWeight = rs!QTY_ON_HAND
      txt3FullWeight = rs!QTY_ON_HAND + Val(rs!custom_field1) * Val(rs!custom_field3) + poddon.Weight
      cmd3ClearW.Enabled = False
    End If
   Else
    'log.Message "�������������� ���� �� ����!!!"
    txt3PWeight = poddon.Weight
    txt3FullWeight = rs!QTY_ON_HAND + Val(rs!custom_field1) * Val(rs!custom_field3) + poddon.Weight
    txt3Quantity = rs!custom_field1
    txt3PackageWeight = rs!custom_field3
    If isCalibrated Then
      txt3GoodWeight = rs!QTY_ON_HAND
      txt3FullWeight = rs!QTY_ON_HAND + Val(rs!custom_field1) * Val(rs!custom_field3) + poddon.Weight
      cmd3ClearW.Enabled = False
    End If
    
   End If

    
  End If
End Sub


'�������� �������
Private Function CheckPoddon(Optional NoCheckShip As Boolean = True) As Boolean
On Error Resume Next
  Dim result As Boolean
  Dim i As Long
  result = True
  
  isCalibrated = False
  
  If txt3PNum <> "" Then
    If Len(txt3PNum) = 6 Then
      Set poddon = Nothing
      Set poddon = FindPoddon(txt3PNum)
      If Not poddon Is Nothing Then
        If poddon.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" Then
          Dim conn As ADODB.Connection
          Set conn = GetCoreConn
          If conn.State <> adStateOpen Then
            conn.open
          End If
          
          Dim rs As ADODB.Recordset
          Dim rs2 As ADODB.Recordset
          
          'poddon.CurrentGood
          
          Set rs = conn.Execute("select * from stock where PALLET_STATUS is null and pallet_id=" & poddon.CorePalette_ID)
          If rs.EOF Then
            MsgBox "����� �������: " & txt3PNum & "  �� ��������� � ���� CORE IMS"
            result = False
          Else
            If rs!item_id <> Manager.GetIDFromXMLField(curQRow.good_id) Then
              MsgBox "������� ����� �� ������� �� ��������� � ��������� ������"
              result = False
            Else
            
            
              If rs!status = 103 Then
                    MsgBox "������ ������������ ��� �������� (���������).", vbExclamation + vbOKOnly, "��������"
                    result = False
              Else
              
                If rs!custom_field2 & "" = "1" Then
                  isCalibrated = True
                  QFromStock = rs!QTY_ON_HAND
                End If
                
                Dim lid  As String
                lid = "" & rs!location_id
                If lid <> "" Then
                  Set rs = conn.Execute("select * from location where id=" & lid)
                  txtMainCell = rs!code
                  txtMainCell.Tag = rs!id
                End If
              
                If NoCheckShip = False Then
                  frameWait.Visible = True
                  DoEvents
                  ' ��������� ����� �� ��������� �� ������ ������ � ������ ���������
                  'Set rs2 = conn.Execute("select * from v_bami_vimorozka_rpt2 A  join stock B on  checksum(a.item_id, a.factory , a.country, a.Kill_place, a.IsBrak, a.made_date_to, a.vetsved) = " & _
                  '"checksum(b.item_id,b.custom_field4,b.custom_field6,b.custom_field11,b.custom_field12,b.custom_field9,b.custom_field7)  and b.PALLET_STATUS is null and b.pallet_id=" & poddon.CorePalette_ID)
                  Set rs2 = conn.Execute("exec  CheckPartiaMoroz " & poddon.CorePalette_ID & " ")
                  If Not rs2.EOF Then
                    If rs2!to_ship > 0 Then
                        If rs2!to_ship < rs!QTY_ON_HAND Then
                          MsgBox "C ������� ������� ����� ���� ��������� ������ " & rs2!to_ship & " ��. ������", vbExclamation + vbOKOnly, "��������"
                          
                          Dim mail As STDMail.Application
                          Dim idmail As String
                          idmail = CreateGUID2()
                          Manager.NewInstance idmail, "STDMail", "���������� " & Now
                          Set mail = Manager.GetInstanceObject(idmail)
                          If Not mail Is Nothing Then
                            For i = 1 To ITTDic.ITTD_EMAIL.Count
                              If ITTDic.ITTD_EMAIL.Item(i).IgnoreAddress = Boolean_Net Then
                                With mail.STDMail_To.Add
                                  .TheTo = ITTDic.ITTD_EMAIL.Item(i).EMail
                                  .TheType = MailSenderType_Komu
                                  .save
                                End With
                              End If
                            Next
                            
                            With mail.STDMail_Info.Add
                              .Subject = "���������� �� " & Now
                              .TheBody = "�������� ������ c ������� '" & poddon.code & "'  � ���������� (" & rs!QTY_ON_HAND & ") ��������� ����� '� ��������' (" & rs2!to_ship & ") "
                              .TheBody = .TheBody & " ��� ������:" & vbCrLf & rs2!item_code & " " & rs2!Description
                              .TheBody = .TheBody & " ������:" & rs2!country & " �����: " & rs2!factory & " �����:" & rs2!kill_place
                              .Sended = Boolean_Net
                              .save
                            End With
                            
                          End If
                       
                        End If

                      Else
                        
                            MsgBox "�������� ������ �� ������ ������ �������������.", vbExclamation + vbOKOnly, "��������"
                            result = False
                      End If
'                    Else
                        
'                          MsgBox "�������� ������ �� ������ ������ �������������.", vbExclamation + vbOKOnly, "��������"
'                          result = False
'
                  End If
                  frameWait.Visible = False
                End If
                  
              End If
            End If
          End If
        Else
          MsgBox "��������� �������: " & txt3PNum & "  ����������� ������� (" & poddon.Application.StatusName & ")"
          result = False
        End If
      Else
        MsgBox "����� �������: " & txt3PNum & "  �� ���������������"
        result = False
      End If
    End If
  End If
  If result = True Then
  
  End If
  CheckPoddon = result
End Function



Private Sub txt3PWeight_Change()
  On Error Resume Next
  If isCalibrated Then Exit Sub
  txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight)
End Sub

Private Sub txt3Quantity_Change()
On Error Resume Next
If isCalibrated Then Exit Sub
txt3GoodWeight = MyRound(txt3FullWeight) - MyRound(txt3PWeight) - (MyRound(txt3PackageWeight) * MyRound(txt3Quantity))
End Sub

Private Sub txt4FullWeight_Change()
  On Error Resume Next
  If isCalibrated Then Exit Sub
  txt4GoodWeight = MyRound(txt4FullWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
  
End Sub


  

Private Sub Form_Load()
On Error Resume Next
    StepNo = 0
  
  XMLShipOrder = "<SQLData>"
  XMLShipOrder = XMLShipOrder & "<connectionstring>ref</connectionstring>"
  XMLShipOrder = XMLShipOrder & "<connectionprovider>ref</connectionprovider>"
  XMLShipOrder = XMLShipOrder & "<query>select A.ID [���] , convert(varchar(30),A.NUMBER) +'  �� ' + convert(varchar(30),A.ORD_DATE,111)  [��������], PARTNER.Name [������]  from shipping_ORDER A left join PARTNER  on A.PARTNER_ID=PARTNER.ID where (a.STATUS = 1 or a.status =0) </query>"
  XMLShipOrder = XMLShipOrder & "<IDFieldName>���</IDFieldName>"
  XMLShipOrder = XMLShipOrder & "<BriefFields>��������</BriefFields>"
  XMLShipOrder = XMLShipOrder & "</SQLData>"
    
  
  XMLTheClient = "<SQLData>"
  XMLTheClient = XMLTheClient & "<connectionstring>ref</connectionstring>"
  XMLTheClient = XMLTheClient & "<connectionprovider>ref</connectionprovider>"
  XMLTheClient = XMLTheClient & "<query>select partner.ID, partner.Name from SHIPPING_ORDER join partner on SHIPPING_ORDER.partner_id=partner.id where SHIPPING_ORDER.ID='%ShipOrderID%'</query>"
  XMLTheClient = XMLTheClient & "<IDFieldName>ID</IDFieldName>"
  XMLTheClient = XMLTheClient & "<BriefFields>Name</BriefFields>"
  XMLTheClient = XMLTheClient & "</SQLData>"
    
    ProcessStatus
    Set conn = GetCoreConn
    If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
      Set wave = New MTZMCI.WavePlayer
      wave.OpenDevice
    End If
    
End Sub

'��������� ���������� ������ � ������� �������
Private Sub AdjFrame(f As Frame)
On Error Resume Next
  f.Top = 0
  f.Left = imgState.Width + 5 * Screen.TwipsPerPixelX
  f.Width = Me.ScaleWidth - imgState.Width - 10 * Screen.TwipsPerPixelX
  f.Height = Me.ScaleHeight - cmdNext.Height - 5 * Screen.TwipsPerPixelY
End Sub

'�� ������� ���� ������� - ��� 1 - ����� ������
Private Sub Before1()
On Error Resume Next
    txtShipOrder.Text = ""
    txtShipOrder.Tag = XMLShipOrder
    LoadBtnPictures cmdShipOrder, cmdShipOrder.Tag
    cmdShipOrder.RemoveAllMenu
    txtTheClient.Text = ""
    txtTheClient.Tag = XMLTheClient
End Sub

'������� �����
Private Sub MakeItem()
On Error Resume Next
'����� ����� � � ����� ����
  Dim rs As ADODB.Recordset
  Dim id As String
  Dim qID As String
  qID = Manager.GetIDFromXMLField(txtShipOrder.Tag)
  id = ""
  Set rs = Session.GetData("select instanceid from ITTOUT_DEF where ShipOrder like '%<ID>" & qID & "</ID>%'")
  If Not rs Is Nothing Then
    If Not rs.EOF Then
      id = rs!InstanceID
    End If
  End If
  rs.Close
  
  If conn.State <> ADODB.adStateOpen Then
    conn.open
  End If
    
  '���� ��� ������, �� ������������ �����
  If id = "" Then
    
    
    Dim errlines As Boolean
    ' ��������� ��� ��� ������ � ������ ������������ ���� �������������
    Set rs = conn.Execute("select B.code b_code,e.code e_code,d.code d_code from shipping_line A  join item B on A.item_id =B.id  join shipping_order C on a.order_id = C.id join partner D on c.partner_id= d.id join partner E on b.CLASS= e.CODE where e.code <>d.code and a.order_id='" & qID & "'")
    While Not rs.EOF
        MsgBox "��� �������� � ����� " & rs!b_code & " �� ������������� ������������� � ��������� � ������ ������" & vbCrLf & "��������� ������ � CORE IMS", vbOKOnly + vbExclamation, "��������"
        errlines = True
        rs.MoveNext
    Wend
    
    If errlines Then Exit Sub
    
    
    
    ' �������� �������� ������ �� core
    Set rs = conn.Execute("select * from shipping_order where id=" & Manager.GetIDFromXMLField(txtShipOrder.Tag))
    If rs.EOF Then Exit Sub
    
    
'    ������� ����� �����
    id = CreateGUID2
    Manager.NewInstance id, "ITTOUT", txtShipOrder
    Set Item = Manager.GetInstanceObject(id)
    
    
'    ��������� ��������
    With Item.ITTOUT_DEF.Add
      .ProcessDate = Date
      .ShipOrder = txtShipOrder.Tag
      .TheClient = txtTheClient.Tag
      
      .Supplier = "" & rs!street1
      '.TTN = rs!ACCOUNT_NUMBER
      '.TTNDate = Date
      '.TranspNumber = rs!Comment1
      .Container = "" & rs!TRACK_NUMBER2
      .Track_time_in = Now
      .track_time_out = DateAdd("h", 4, Now)
      .temp_in_track = -1

      
      .save
    End With
    rs.Close
    
    
'    ��������� ������ ������ � ���� ������ ��
    Dim XMLQRY_NUM As String
    Dim XMLLineAtQuery As String
    Dim XMLgood_ID As String
    
    Set rs = conn.Execute("select a.*, A.QTY_ORD QRY_NUMID, B.DESCRIPTION  BRIEF , B.code Articul from shipping_line A join item B on A.item_id =B.id where (a.PARENT_ID  is null or a.parent_id=0) and a.order_id='" & qID & "'")
    While Not rs.EOF
    Set curQRow = Item.ITTOUT_LINES.Add
    
      With curQRow
        
        XMLLineAtQuery = "<SQLData>"
        XMLLineAtQuery = XMLLineAtQuery & "<connectionstring>ref</connectionstring>"
        XMLLineAtQuery = XMLLineAtQuery & "<connectionprovider>ref</connectionprovider>"
        XMLLineAtQuery = XMLLineAtQuery & "<query>select A.ID [���], A.ORDER_ID [��� ������], A.QTY_ORD [����������] , B.DESCRIPTION [��������] from shipping_line A join item B on A.item_id =B.id </query>"
        XMLLineAtQuery = XMLLineAtQuery & "<IDFieldName>���</IDFieldName>"
        XMLLineAtQuery = XMLLineAtQuery & "<BriefFields>��������</BriefFields>"
        XMLLineAtQuery = XMLLineAtQuery & "<Brief>" & rs!brief & "</Brief>"
        XMLLineAtQuery = XMLLineAtQuery & "<ID>" & rs!id & "</ID>"
        XMLLineAtQuery = XMLLineAtQuery & "</SQLData>"
        
        .LineAtQuery = XMLLineAtQuery
              
        XMLQRY_NUM = "<SQLData>"
        XMLQRY_NUM = XMLQRY_NUM & "<connectionstring>ref</connectionstring>"
        XMLQRY_NUM = XMLQRY_NUM & "<connectionprovider>ref</connectionprovider>"
        XMLQRY_NUM = XMLQRY_NUM & "<query>select  QTY_ORD from shipping_line where ID='%LineAtQueryID%'</query>"
        XMLQRY_NUM = XMLQRY_NUM & "<IDFieldName>QTY_ORD</IDFieldName>"
        XMLQRY_NUM = XMLQRY_NUM & "<BriefFields>QTY_ORD</BriefFields>"
        XMLQRY_NUM = XMLQRY_NUM & "<ID>" & rs!QRY_NUMID & "</ID>"
        XMLQRY_NUM = XMLQRY_NUM & "<Brief>" & rs!QRY_NUMID & "</Brief>"
        XMLQRY_NUM = XMLQRY_NUM & "<LineAtQueryID>" & rs!id & "</LineAtQueryID>"
        XMLQRY_NUM = XMLQRY_NUM & "</SQLData>"
              
        .QRY_NUM = XMLQRY_NUM
       
        XMLgood_ID = "<SQLData>"
        XMLgood_ID = XMLgood_ID & "<connectionstring>ref</connectionstring>"
        XMLgood_ID = XMLgood_ID & "<connectionprovider>ref</connectionprovider>"
        XMLgood_ID = XMLgood_ID & "<query>select  item_id from shipping_line where ID='%LineAtQueryID%'</query>"
        XMLgood_ID = XMLgood_ID & "<IDFieldName>item_id</IDFieldName>"
        XMLgood_ID = XMLgood_ID & "<BriefFields>item_id</BriefFields>"
        XMLgood_ID = XMLgood_ID & "<ID>" & rs!item_id & "</ID>"
        XMLgood_ID = XMLgood_ID & "<Brief>" & rs!item_id & "</Brief>"
        XMLgood_ID = XMLgood_ID & "<LineAtQueryID>" & rs!id & "</LineAtQueryID>"
        XMLgood_ID = XMLgood_ID & "</SQLData>"
        
        .good_id = XMLgood_ID
        
        .edizm = "" & rs!UOM
        .articul = "" & rs!articul
        Set .made_country = FindCountry("" & rs!prod_country)
        
        On Error Resume Next
        
'        �������� ������������ ������ �� ����������� �� ������ ��������� ����� � CORE
        If Not .made_country Is Nothing Then
          Set .factory = FindFactory(.made_country.id, "" & rs!producer)
        End If
        
        If Not .factory Is Nothing Then
          Set .KILL_NUMBER = FindKill(.factory.id, "" & rs!KILL_NUMBER)
        End If
        
        Set .PartRef = FindPartia("" & rs!brief, "" & rs!LOT_SN)
        
        If Not IsNull(rs!made_date) Then .made_date = rs!made_date
        
        If Not IsNull(rs!exp_date) Then .exp_date = rs!exp_date
        
        If Not IsNull(rs!custom_field9) Then .made_date_to = CDate(rs!custom_field9)
        
        If Not IsNull(rs!custom_field7) Then .vetsved = rs!custom_field7
        
        .save
      End With
      ' �������� ������ ������ � �� ��
      Call GetNumValue(curQRow, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "OUT%P", "")
      rs.MoveNext
    Wend
    
'    ��������� ������ ����� �� ������ ����������� ��� �������� �������
     Set rs = Session.GetData("select * from ITTCS_DEF where clientcode like '%<ID>" & Manager.GetIDFromXMLField(txtTheClient.Tag) & "</ID>%'")
    Dim srvid As String
    Dim srvObj As ITTCS.Application
    Dim srv As ITTD_SRV
    If Not rs.EOF Then
      srvid = rs!InstanceID
      Set srvObj = Manager.GetInstanceObject(srvid)
      Dim i As Long
      For i = 1 To srvObj.ITTCS_LIN.Count
         Set srv = srvObj.ITTCS_LIN.Item(i).srv
         If srv.ForShipping = Boolean_Da Then
            If srvObj.ITTCS_LIN.Item(i).UseSrv = Boolean_Da Then
              With Item.ITTOUT_SRV.Add
                 Set .srv = srv
                 .Quantity = 0
                 .save
              End With
            End If
         End If
      Next
    End If
  Else
    Set Item = Manager.GetInstanceObject(id)
  End If
    
End Sub

'�������� ��������� ������ ��� �����������
Private Sub LoadHeader(Item As Object)
  txtSupplier = Item.Supplier
  txtTTN = Item.TTN
  dtpTTNDate = Date
  If Item.TTNDate <> 0 Then
   dtpTTNDate = Item.TTNDate
  Else
   dtpTTNDate.Value = Null
  End If
  txtTranspNumber = Item.TranspNumber
  txtContainer = Item.Container
  txtStampNumber = Item.StampNumber
  txtStampStatus = Item.StampStatus
  dtpTrack_time_in = Now
  If Item.Track_time_in <> 0 Then
   dtpTrack_time_in = Item.Track_time_in
  Else
   dtpTrack_time_in.Value = Null
  End If
  dtptrack_time_out = Now
  If Item.track_time_out <> 0 Then
   dtptrack_time_out = Item.track_time_out
  Else
   dtptrack_time_out.Value = Null
  End If
  txttemp_in_track = Item.temp_in_track

End Sub

'�� ������� ���� ������� - ��� 2 - ����� ������ ������
Private Sub Before2()
  If MsgBox("���������� �����������?", vbYesNo) = vbYes Then
    Set repShowSRVOUT = Nothing
    Set repShowSRVOUT = New ReportShow
    repShowSRVOUT.ReportSource = "V_viewITTOUT_ITTOUT_SRV"
    repShowSRVOUT.ReportFilter = " instanceid='" & Item.id & "'"
    repShowSRVOUT.ReportPath = App.Path & "\OUt_srv.rpt"
    repShowSRVOUT.PrinterName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowSRVOUT.Run True
  End If
  
' ��������� ��� ����:ITTOUT ��������
' "{70853C28-84B5-434E-8413-52DF8FBBB49B}" '���� ��������
' "{2CDDB562-63D7-483E-B95E-B579A9096CCC}" '��������� ���������
' "{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}" '�������� ���������
' "{CDCAFF7F-B013-40AF-BE61-1A27E35DB946}" '�����������
  
  Item.StatusID = "{70853C28-84B5-434E-8413-52DF8FBBB49B}" '���� ��������
  
  '��������������� ������� ����� ������
  gr.ItemCount = 0
  'Item.ITTOUT_LINES.Sort = "sequence"
  Item.ITTOUT_LINES.PrepareGrid gr
  Item.ITTOUT_LINES.Refresh
  gr.ItemCount = Item.ITTOUT_LINES.Count


End Sub

'�� �������� ���� ������� - ��� 3 - ������ � ������
Private Sub Before3()
  On Error Resume Next
  txt3PNum = ""
  txt3FullWeight = "0"
  txt3Good = 0
  txt3PWeight = 0
  txt3Quantity = 0
  txt3PackageWeight = 0
  
  If curQRow Is Nothing Then Exit Sub
  
  
'  ������� �����
  Dim XMLDocLineAtQuery As New DOMDocument
  Call XMLDocLineAtQuery.loadXML(curQRow.LineAtQuery)
  If (err.Number = 0 And XMLDocLineAtQuery.parseError.errorCode = 0) Then
    Dim nodeLineAtQuery As MSXML2.IXMLDOMNode
    
    For Each nodeLineAtQuery In XMLDocLineAtQuery.childNodes.Item(0).childNodes
      If (nodeLineAtQuery.baseName = "Brief") Then
       txt3Good.Text = nodeLineAtQuery.Text
       Exit For
      End If
    Next
  End If

'  �������� ���� �� ������
  Dim XMLDocQRY_NUM As New DOMDocument
  Dim plan As Double
On Error Resume Next
  If (curQRow.QRY_NUM <> "") Then
    Call XMLDocQRY_NUM.loadXML(curQRow.QRY_NUM)
    If (err.Number = 0 And XMLDocQRY_NUM.parseError.errorCode = 0) Then
      Dim nodeQRY_NUM As MSXML2.IXMLDOMNode
      
      For Each nodeQRY_NUM In XMLDocQRY_NUM.childNodes.Item(0).childNodes
        If (nodeQRY_NUM.baseName = "Brief") Then
          plan = MyRound("0" & nodeQRY_NUM.Text)
         Exit For
        End If
      Next
    End If
  End If
  
'  ��������� ���� ����������
  txt3FRomQ = plan
  txt3InQry.Text = plan - curQRow.CurValue


  
'  ������������� �����
  If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
    Set wave = New MTZMCI.WavePlayer
    wave.OpenDevice
  End If

'  ������������ COM  �����
  emu = Not (GetSetting("RBH", "ITTSETTINGS", "EMULATOR", "False") = "False")
  psetup = GetSetting("RBH", "ITTSETTINGS", "WSETUP", "4800,e,8,1")
  port = GetSetting("RBH", "ITTSETTINGS", "WPORT", 1)

  If Not emu Then
    If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
    End If
      
    MSComm1.Handshaking = comNone
    MSComm1.DTREnable = False
    MSComm1.EOFEnable = False
      
    MSComm1.Settings = psetup
    MSComm1.CommPort = port
    MSComm1.PortOpen = True
  End If
  
  
End Sub

'�� ���������� ���� ������� - ���4 - ����������� ����� ��� �������
Private Sub Before4()
  txt4FullWeight = 0
  txt4Good = txt3Good
  txt4FromQ = txt3FRomQ
  txt4InQry = txt3InQry
  txt4PackageWeight = txt3PackageWeight
  txt4NewPlace = txtMainCell
  txt4Quantity = 0
End Sub

'�� ������ ���� ������� - ���5 - ������ ������� �� �������������� �����
Private Sub Before5()

'
  
bye2:
  
  Exit Sub
  
bye:
  If err.Number <> 0 Then
    MsgBox err.Description, , "������ ���������� �� ������"
  End If

End Sub


'�� ������� ���� ������� - ���6 -������� ����������
Private Sub Before6()

  Dim strs As ADODB.Recordset
  Dim conn As ADODB.Connection
  
'  �������� ������ � core
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
'  ����� ������
  Set poddon = FindPoddon(txt3PNum)
     
  Dim netto As Double
  Dim korob As Integer
  Dim OK As Boolean
  
'  ������� ����
  netto = MyRound(txt3GoodWeight)
  korob = MyRound(txt3Quantity)
  
    
  Dim morosrs As ADODB.Recordset
  Dim delta As Double
  Dim protID As String
  Dim prot As ITTPR.Application
  
  
  If isFull Then
    ' ������� ���������
    Set morosrs = conn.Execute("select   min(LastRCV) LASTRCV  ,sum(in_quantity)  qin ,sum(in_boxes)  bin ,sum(out_quantity) qout  ,sum(out_boxes) bout  ,sum( dout_quantity) vimorozka  ,sum(stok_quantity) qstok from v_bami_vimorozka where pallet ='" & poddon.TheNumber & "' and rectype <>3")
  
    If Not morosrs Is Nothing Then
      delta = morosrs!qin - morosrs!qout - morosrs!qstok - morosrs!vimorozka * 0.0005 - morosrs!qout * 0.001 - morosrs!qin * 0.001
      delta = delta - netto - netto * 0.001 - netto / 30 * DateDiff("d", morosrs!lastrcv, Now) * 0.0005
      
      ' ���� ������ ��� ����� �� ����� �������������
      If delta > 0 Then
        ' ������� ��������
        protID = CreateGUID2
        Manager.NewInstance protID, "ITTPR", "�������� ����������� �� ������ �" & poddon.code
        Set prot = Manager.GetInstanceObject(protID)
        
'        ��������� ��������
        With prot.ITTPR_DEF.Add
          .TheDate = Date
          .poddon = poddon.code
          .InWeight = morosrs!qin
          .OutWeight = morosrs!qout + netto
          .Vesi = netto * 0.001 + morosrs!qout * 0.001 + morosrs!qin * 0.001
          .Moroz = netto / 30 * DateDiff("d", morosrs!lastrcv, Now) * 0.0005 + morosrs!vimorozka * 0.0005
          .WeightDelta = delta
          .InBoxes = morosrs!bin
          .OutBoxes = morosrs!bout + korob
          On Error Resume Next
          
          Set strs = conn.Execute("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & poddon.CorePalette_ID)
          .Good = curQRow.articul
          .Description = GetBRIEFFromXMLField(curQRow.LineAtQuery)
          .Client = GetBRIEFFromXMLField(Item.ITTOUT_DEF.Item(1).TheClient)
          .factory = strs!custom_field4
          .killplace = strs!custom_field11
          .country = strs!custom_field6
        
          If strs!status = 101 Then
            .brak = "����"
          Else
            .brak = " - ��� -"
          End If
          
'          ���������  ��������
          .save
          
        End With
        
    
        ' �������� ���
            
        Set RptActVes = New ReportShow
        RptActVes.ReportPath = App.Path & "\AktVes.rpt"
        RptActVes.ReportSource = "V_AUTOITTPR_DEF"
        RptActVes.ReportFilter = "instanceid ='" & protID & "'"
        Call RptActVes.Run(True)
        Set RptActVes = Nothing
        log.message "������ ��� � ����������� ������:" & poddon.code
        
        
'        �������� ����� �� ��������
        If MsgBox("��������� ������?", vbYesNo, "��������") = vbNo Then
            curQRow.ITTOUT_PALET.Refresh
            'poddon.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}"
            StepNo = 3
            ProcessStatus
            log.message "����� �� �������� ������:" & poddon.code
'            ����������
            Exit Sub
        End If
      
      End If
    End If
    
    
    ' ������ ��������� ������
    ' ��������� ��� ����:ITTPL �������
    ' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" '��������
    ' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" '�� ������ � ������
    ' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" '���������� � ������
    ' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" '������
    ' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" '�������
    
    curQRow.ITTOUT_PALET.Refresh
    Dim pweight As Double
    pweight = poddon.Weight
    
'    ������ ��������� ������� ��� ��������
    If MsgBox("������ ������ ������� ?", vbYesNo + vbDefaultButton2) = vbYes Then
      poddon.Application.StatusID = "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}"
      poddon.CurrentPosition = ""
      poddon.PackageWeight = 0
      poddon.CaliberQuantity = 0
      poddon.CurrentGood = ""
      poddon.save
    Else
      poddon.Application.StatusID = "{E9BFB749-A606-4DEF-A429-07D636F108C6}"
      poddon.CurrentGood = ""
      poddon.CurrentPosition = ""
      poddon.PackageWeight = 0
      poddon.CaliberQuantity = 0
      poddon.save
    End If
    
    curQRow.save
    
 
   
   
   ' �������� � ������ ��
    
'    ������� �������� �������
    Set LinePal = curQRow.ITTOUT_PALET.Add
    
    
'    ��������� ��������
    Set LinePal.TheNumber = FindPoddon(txt3PNum)
    LinePal.CaliberQuantity = MyRound(txt3Quantity)
    LinePal.GoodWithPaletWeight = MyRound(txt3FullWeight)
    LinePal.PackageWeight = MyRound(txt3PackageWeight)
    
    If isCalibrated Then
      LinePal.FullPackageWeight = MyRound(txt3FullWeight) - MyRound(txt3GoodWeight) - pweight
    Else
      LinePal.FullPackageWeight = MyRound(txt3PackageWeight) * MyRound(txt3Quantity)
    End If
    
    LinePal.IsEmpty = Boolean_Da
    LinePal.StoreCell = txtMainCell
    Set LinePal.made_country = curQRow.made_country
    Set LinePal.factory = curQRow.factory
    Set LinePal.KILL_NUMBER = curQRow.KILL_NUMBER
    Set LinePal.PartRef = curQRow.PartRef
    LinePal.made_date = curQRow.made_date
    LinePal.exp_date = curQRow.exp_date
    LinePal.made_date_to = curQRow.made_date_to
    LinePal.vetsved = curQRow.vetsved
    
    Set strs = conn.Execute("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & LinePal.TheNumber.CorePalette_ID)
    If strs!status = 101 Then
      LinePal.IsBrak = Boolean_Da
    Else
      LinePal.IsBrak = Boolean_Net
    End If
    
    If strs!custom_field2 & "" = "1" Then
      LinePal.isCalibrated = Boolean_Da
    Else
      LinePal.isCalibrated = Boolean_Net
    End If
    
'    ���������
    LinePal.save
     
    Call GetNumValue(LinePal, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "OUTPAL%P", "")
    
    log.message "�������� ������  " & txt3PNum & " ��� " & MyRound(txt3GoodWeight)
    
    ' �� ��������� ��� ������ �������!
    curQRow.CurValue = curQRow.CurValue + MyRound(txt3GoodWeight)
    curQRow.save
    
    frameSave.Visible = True
    DoEvents

    ' ���������� � CORE
    OK = False
    While Not OK
      OK = SaveShipRowToCore(txtShipOrder.Text, txt4NewPlace.Text, Item, poddon, curQRow, LinePal, True)
      If Not OK Then
         OK = Not MagicMessageBox("�� ������� ��������� ���������� � CORE. ������� � " & poddon.code & vbCrLf & "������ ������� ���������� ������")
        Dim conn2 As Object
        Set conn2 = GetCoreConn(True)
      End If
      
    Wend
    
    frameSave.Visible = False
    DoEvents
    
    ' ���������� ������ ������
    Set curQRow.made_country = LinePal.made_country
    Set curQRow.factory = LinePal.factory
    Set curQRow.KILL_NUMBER = LinePal.KILL_NUMBER
    Set curQRow.PartRef = LinePal.PartRef
    curQRow.save
    
  Else
  
'  �������� ��������
    curQRow.save
'    ������� ������ -� ��������
    Set LinePal = curQRow.ITTOUT_PALET.Add
     
     
'    ��������� ������ �� ������� (�������)
     Set poddon = FindPoddon(txt3PNum)
     poddon.CurrentWeightBrutto = MyRound(txt3FullWeight) - MyRound(txt4GoodWeight)
     poddon.CaliberQuantity = MyRound(txt3Quantity) - MyRound(txt4Quantity)
     poddon.PackageWeight = (MyRound(txt3Quantity) - MyRound(txt4Quantity)) * MyRound(txt3PackageWeight)
     err.Clear
     On Error Resume Next
     poddon.save
     
'     ��������� ������ �������
     With LinePal
      Set .TheNumber = poddon
      .IsEmpty = Boolean_Net
      
      Set strs = conn.Execute("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & LinePal.TheNumber.CorePalette_ID)
      
      ' ������� ���������
      
      .GoodWithPaletWeight = MyRound(txt4FullWeight) + MyRound(txt3PWeight)
      
      If isCalibrated Then
        .GoodWithPaletWeight = (MyRound(txt4Quantity)) * (MyRound(txt3PackageWeight) + (strs!QTY_ON_HAND / IIf(Val(strs!custom_field1) = 0, 1, Val(strs!custom_field1)))) + MyRound(txt3PWeight)
        .FullPackageWeight = MyRound(txt4Quantity) * MyRound(txt3PackageWeight)
      Else
        .FullPackageWeight = (MyRound(txt4Quantity)) * MyRound(txt3PackageWeight)
      End If
      
      
      .PackageWeight = MyRound(txt3PackageWeight)
      .CaliberQuantity = MyRound(txt4Quantity)
      
      ' ������� �������� �� �������
      .ReorgWeight = MyRound(txt3FullWeight) - MyRound(txt4GoodWeight)
      .ReorgPackageFullWeight = MyRound(txt4PackageWeight) * (MyRound(txt3Quantity) - MyRound(txt4Quantity))
      .ReorgCaliberQuantity = MyRound(txt4Quantity)
      .StoreCell = txtMainCell
      .BufferCell = txt4NewPlace
      
      Set LinePal.made_country = curQRow.made_country
      Set LinePal.factory = curQRow.factory
      Set LinePal.KILL_NUMBER = curQRow.KILL_NUMBER
      Set LinePal.PartRef = curQRow.PartRef
      LinePal.made_date = curQRow.made_date
      LinePal.exp_date = curQRow.exp_date
      LinePal.made_date_to = curQRow.made_date_to
      LinePal.vetsved = curQRow.vetsved
    
      

      If strs!status = 101 Then
        LinePal.IsBrak = Boolean_Da
      Else
        LinePal.IsBrak = Boolean_Net
      End If
      
       If strs!custom_field2 & "" = "1" Then
        LinePal.isCalibrated = Boolean_Da
       Else
          LinePal.isCalibrated = Boolean_Net
       End If
      err.Clear
'      ��������� ������� � ��
      .save
      
    End With
    
    Call GetNumValue(LinePal, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "OUTPAL%P", "")
    
    log.message "�������� ������  " & txt3PNum & " ��� " & MyRound(txt4GoodWeight)
    
    curQRow.CurValue = curQRow.CurValue + MyRound(txt4GoodWeight)
    curQRow.save
    
    

    
    
    frameSave.Visible = True
    DoEvents
    
    
   ' ��������� � CORE
    OK = False
    While Not OK
      OK = SaveShipRowToCore(txtShipOrder.Text, txt4NewPlace.Text, Item, poddon, curQRow, LinePal, False)
      If Not OK Then
         OK = Not MagicMessageBox("�� ������� ��������� ���������� � CORE. ������� � " & poddon.code & vbCrLf & "������ ������� ���������� ������")
        Dim conn3 As Object
        Set conn3 = GetCoreConn(True)
      End If
    Wend
  
    frameSave.Visible = False
    DoEvents
    
    Set curQRow.made_country = LinePal.made_country
    Set curQRow.factory = LinePal.factory
    Set curQRow.KILL_NUMBER = LinePal.KILL_NUMBER
    Set curQRow.PartRef = LinePal.PartRef
    curQRow.save
    
    
    ' �������� ������ �� �������
    PrintSticker LinePal.TheNumber
    
  End If

  On Error Resume Next
  MSComm1.PortOpen = False
  
  ' ��������� ������ ����� ������ � �����
  gr2.ItemCount = 0
  Item.ITTOUT_LINES.PrepareGrid gr2
  gr2.ItemCount = Item.ITTOUT_LINES.Count
  
End Sub

'�� �������� ���� ������� - ��� 7 - ���� ����� � ������
Private Sub before7()
 
  Dim i As Long
  If Item.ITTOUT_SRV.Count > 0 Then
    srvGrid.Rows = Item.ITTOUT_SRV.Count + 1
    
    
    Dim pcnt As Long
    pcnt = 0
    
    For i = 1 To Item.ITTOUT_LINES.Count
        pcnt = pcnt + Item.ITTOUT_LINES.Item(i).ITTOUT_PALET.Count
    Next
    
    Dim srv As ITTD.ITTD_SRV
 
    For i = 1 To Item.ITTOUT_SRV.Count
    
      Set srv = Item.ITTOUT_SRV.Item(i).srv
      srvGrid.TextMatrix(i, 0) = Item.ITTOUT_SRV.Item(i).srv.brief
      srvGrid.TextMatrix(i, 1) = Item.ITTOUT_SRV.Item(i).Quantity
      If srv.AutoSetPallet = Boolean_Da Then
         If Item.ITTOUT_SRV.Item(i).Quantity = 0 Then
                   srvGrid.TextMatrix(i, 1) = pcnt
         End If
      End If
    Next
        

  End If
End Sub

'�������� ������ � ������� ����� ������
Private Sub gr_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
  On Error Resume Next
  Item.ITTOUT_LINES.LoadRow gr, RowIndex, Bookmark, Values
End Sub

'����� ������� ������ ������
Private Sub gr_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
On Error Resume Next
  If gr.ItemCount = 0 Then
    Exit Sub
  End If
  
  If gr.Row > 0 Then
   If gr.RowIndex(gr.Row) > 0 Then
    If LastRow <> gr.Row Then
      
      Dim bm
      bm = gr.RowBookmark(gr.RowIndex(gr.Row))
      
      Set curQRow = Item.FindRowObject(Right(bm, Len(bm) - 38), Left(bm, 38))
      
    End If
   End If
  End If
End Sub

'������� ����� ���� �������
Private Sub ProcessStatus()
  Frame1.Visible = False
  Frame2.Visible = False
  Frame3.Visible = False
  Frame4.Visible = False
  Frame5.Visible = False
  Frame6.Visible = False
  Frame7.Visible = False
  
  cmdBack.Visible = False
  cmdNext.Caption = "�����"
  cmdAddW.Visible = False
  cmdCancel.Caption = "��������"
  cmdCancel.Visible = True

  Select Case StepNo
  Case 0
    cmdNext.Caption = "������ �������"
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  Case 1
  '��� 1 - ����� ������
    Before1
    Frame1.Visible = True
    AdjFrame Frame1
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 2
  '��� 2 - ����� ������ ������
    Before2
    Frame2.Visible = True
    AdjFrame Frame2
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 3
  '��� 3 - ������ � ������
    Before3
    Frame3.Visible = True
    AdjFrame Frame3
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  
  Case 4
  '���4 - ����������� ����� ��� �������
    Before4
    Frame4.Visible = True
    AdjFrame Frame4
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 5
  '���5 - ������ ������� �� �������������� �����
    Before5
    Frame5.Visible = True
    AdjFrame Frame5
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
   
  Case 6
  '���6 -������� ����������
    Before6
    If StepNo = 6 Then
    Frame6.Visible = True
    AdjFrame Frame6
    
    cmdBack.Visible = True
    cmdAddW.Visible = True
    cmdCancel.Visible = False
    cmdNext.Caption = "��������� ���������"
    cmdBack.Caption = "������ ������� ������"
    
    If curQRow.CurValue >= Manager.GetIDFromXMLField(curQRow.QRY_NUM) Then
      cmdAddW.Enabled = False
      SetBtnPos cmdCancel, 1
      SetBtnPos cmdBack, 3
      SetBtnPos cmdNext, 4
      SetBtnPos cmdAddW, 2
    Else
      cmdAddW.Enabled = True
      SetBtnPos cmdCancel, 1
      SetBtnPos cmdBack, 2
      SetBtnPos cmdNext, 3
      SetBtnPos cmdAddW, 4
    End If
    End If
    
    
   Case 7
   '��� 7 - ���� ����� � ������
    before7
    Frame7.Visible = True
    AdjFrame Frame7
    
  Case 8
    Unload Me
  End Select
   
  ' ������ �������� ��������������� ����
  If StepNo >= 0 And StepNo < 8 Then
    imgState.Picture = LoadPicture(App.Path & "\Design\LStep" & (StepNo) & ".bmp")
  Else
    imgState.Picture = LoadPicture(App.Path & "\Design\LStep0.bmp")
  End If
End Sub

'�������� ��������� ������ ����� ����
Private Function CheckAfter() As Boolean
  Dim result As Boolean
  
  Select Case StepNo
  Case 0
    ' do nothiing
    result = True
  Case 1
  
  
    ' ������ ����������� � ������
    If txtShipOrder = "" Then
      result = False
      MsgBox "������� ������� �����"
    Else
      result = True
    End If
    
    
  Case 2
    ' ������� ������ ������
    If curQRow Is Nothing Then
      result = False
      MsgBox "������� ������� ������ ������"
    Else
     result = True
    End If
    
    
  Case 3
    
    ' �������� ������ � ����� ��� �����
    '
     result = True
     
         
     If txt3FullWeight = "" Or Not IsNumeric(txt3FullWeight) Then
      MsgBox "��������� ��������� ���� ����� � �����"
      result = False
     End If
     
     
     If txt3GoodWeight = "" Then
      MsgBox "��� ����� �� �����"
      result = False
     End If
     
     If IsNumeric(txt3GoodWeight) And Val(txt3GoodWeight) <= 0 Then
      MsgBox "��� ����� �� ����� ���� �������"
      result = False
     End If
     
     ' ��������� ��������� ������� � ����
     
     result = CheckPoddon
     
     
  Case 4
    ' ���������� �������������  �����
    
    
    
      result = True
      If txt4FullWeight = "" Or Not IsNumeric(txt4FullWeight) Then
        MsgBox "��������� �������� ���� ����� � �����"
        result = False
      End If
      
      If MyRound(txt4GoodWeight) >= MyRound(txt3GoodWeight) Then
        MsgBox "��������� ������, ��� ���� �� �������, ��������������� ���"
        result = False
      End If

      If MyRound(txt4Quantity) >= MyRound(txt3Quantity) Then
        MsgBox "��������� ������ �������, ��� ���� �� �������, ��������������� ����������"
        result = False
      End If
      
      If txt4NewPlace = "" Then
        MsgBox "������� ������ ��� ��������"
        result = False
      End If
      
  Case 5
    result = True
    If result Then
      If MsgBox("���������������� �������� ������?", vbExclamation + vbYesNo, "��������") = vbYes Then
        result = True
      Else
        result = False
        log.message "����� �� �������� ������:" & poddon.code
      End If
    End If
    
    
    
  Case 6
   result = True
  
  Case 7
   
   ' ��������� �����
   If MsgBox("������� ����� ?", vbExclamation + vbYesNo) = vbYes Then
    CloseZakaz
    
' ��������� ��� ����:ITTOUT ��������
' "{70853C28-84B5-434E-8413-52DF8FBBB49B}" '���� ��������
' "{2CDDB562-63D7-483E-B95E-B579A9096CCC}" '��������� ���������
' "{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}" '�������� ���������
' "{CDCAFF7F-B013-40AF-BE61-1A27E35DB946}" '�����������
    
    Item.StatusID = "{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}" '�������� ���������
    
   End If
    result = True
    
    
  
  
  Case 8
  result = True
  
  Case 9
  result = True
  
  End Select
  CheckAfter = result
End Function

'������� �������� ���� � �����
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� Double
'Example:
' dim variable as Double
'  variable = me.GetWeight4()
Public Function GetWeight4() As Double
Attribute GetWeight4.VB_HelpID = 660
  On Error Resume Next
    Dim ws As String
    Dim ch As String
    Dim start As Single
    Dim ws1 As String
    Dim ws2 As String
    GetWeight4 = 0
    
    MSComm1.output = Chr(68)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_s1
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_s1
    Loop
    
    GetWeight4 = 0  ' �� ��������� ������
    Exit Function
    
answer_s1:
    
    ws = MSComm1.Input
    ' ������ ��� ��� ��������
    If Asc(Mid(ws, 1, 1)) >= 128 Then
    
      ''''''''''''''''''''''''''''''''''''
      '�������� !!!
      '
      ' ���� ����� ��������� �������
      start = Timer   ' Set start time.
      Do While Timer < start + 0.3
      Loop
      
      ' ���������� ��� ���
      MSComm1.output = Chr(68)
      
      
      start = Timer   ' Set start time.
      Do While Timer < start + 0.2
      Loop
      
      If MSComm1.InBufferCount > 0 Then GoTo answer_s2
      start = Timer   ' Set start time
      Do While Timer < start + 0.5
         If MSComm1.InBufferCount > 0 Then GoTo answer_s2
      Loop
      
    End If
    
    GetWeight4 = 0 ' ��� ������� ������
    Exit Function
    
answer_s2:

    ws = MSComm1.Input
    
    ' ������ ��� ��� ���� ��������
    If Asc(Mid(ws, 1, 1)) >= 128 Then
      MSComm1.output = Chr(69)
      start = Timer   ' Set start time.
      Do While Timer < start + 0.2
      Loop
      If MSComm1.InBufferCount > 0 Then GoTo answer_w1
      start = Timer   ' Set start time
      Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w1
      Loop
    End If
    
    GetWeight4 = 0 ' ��� �� ��������, ��� ��� ������
    Exit Function
    
answer_w1:

    ' ������ ��������� ����
    ws1 = MSComm1.Input
    
    
    ''''''''''''''''''''''''''''''''''''
    '�������� !!!
    '
    ' ���� ����� ��������� �������
    start = Timer   ' Set start time.
    Do While Timer < start + 0.3
    Loop
    
    ' ���������� ��� ��� ���
    MSComm1.output = Chr(69)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    Loop
    
    GetWeight4 = 0 '  ��� ������
    Exit Function
      
answer_w2:
    ws = MSComm1.Input
  
    If ws1 = ws Then
      GetWeight4 = (Asc(Mid(ws, 2, 1)) * 256 + Asc(Mid(ws, 1, 1))) / 10
    Else
      GetWeight4 = 0 ' ��� �� ��������, ���������� ���������
    End If
  
End Function

'�������� ��� ��� ������������
Private Function GetWeight() As Double
  If emu Then
    GetWeight = Rnd(Second(Now)) * 1000 + MyRound("0" & txt3PWeight)
  Else
    GetWeight = GetWeight4
  End If
End Function

'�������� ������
Private Sub MyBeep(ByVal BeepType As String)
      If Not wave Is Nothing Then
        On Error Resume Next
        wave.OpenFile App.Path & "\" & BeepType & ".wav"
        wave.Play
      End If
End Sub

'��������� ���� ������
Private Sub cmdTheClient_Click()
  On Error Resume Next
  
    
  
  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtTheClient.Tag = "") Then
    ' call MsgBox("��� ������ ��� �������")
  Else
    Call pars.Add("permanent", "true")
    txtTheClient.Tag = AddSQLRefIds(txtTheClient.Tag, "ShipOrder", txtShipOrder.Tag)
    txtTheClient.Tag = Replace(txtTheClient.Tag, "%ID%", " 1=1 ")
    Call pars.Add("xml", txtTheClient.Tag)
  End If
  Set res = Manager.GetSQLDataDialog(pars)
  If (Not res Is Nothing) Then
    Dim resStr As String
    resStr = res.Item("RESULT").Value
    If (resStr = "OK") Then
      txtTheClient.Tag = res.Item("xml").Value
      If (txtTheClient.Text <> res.Item("brief").Value) Then
        txtTheClient.Text = res.Item("brief").Value
        'mIDTheClient = res.Item("ID").Value
        'Call txtTheClient_Change
      End If
    Else
      Dim errStr As String
      errStr = res.Item("ErrorDescription").Value
      If (errStr <> vbNullString) Then
       Call MsgBox("������ ����������: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
End Sub

'������� ��� �������� ���4
Private Sub Txt4PackageWeight_Change()
If isCalibrated Then Exit Sub
txt4GoodWeight = MyRound(txt4FullWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
End Sub

'������� ���������� ��� 4
Private Sub txt4Quantity_Change()
  If isCalibrated Then
    txt4GoodWeight = MyRound(txt3GoodWeight) / MyRound(txt3Quantity) * MyRound(txt4Quantity)
  Else
    txt4GoodWeight = MyRound(txt4FullWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
  End If
End Sub

'������� �����
Private Sub txtShipOrder_Change()
  
If (txtShipOrder.Text = "") Then
  ' ������ Brief � ID
  If (txtShipOrder.Tag <> "") Then
    Dim XMLDoc As New DOMDocument
    Call XMLDoc.loadXML(txtShipOrder.Tag)
    Dim Node As MSXML2.IXMLDOMNode
    For Each Node In XMLDoc.childNodes.Item(0).childNodes
     If (Node.baseName = "ID") Then
       Node.Text = ""
     End If
     If (Node.baseName = "Brief") Then
       Node.Text = ""
     End If
    Next
    txtShipOrder.Tag = XMLDoc.XML
  End If
End If

Call cmdTheClient_Click
End Sub


'�������� ������
Private Sub CloseZakaz()
On Error Resume Next
  Dim conn As ADODB.Connection
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rlID As String
  
  
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
  Set cmd = New ADODB.Command
  
  Dim i As Long
  Dim oid As String
  oid = Manager.GetIDFromXMLField(Item.ITTOUT_DEF.Item(1).ShipOrder)
  
  cmd.CommandText = "update shipping_order set status=2 where id=" & oid
  Set cmd.ActiveConnection = conn
  err.Clear
  cmd.Execute

  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
  log.message "�������� ������ �� ��������" & txtShipOrder
  
  For i = 1 To Item.ITTOUT_LINES.Count
    Set curQRow = Item.ITTOUT_LINES.Item(i)
    
    rlID = Manager.GetIDFromXMLField(curQRow.good_id)

    cmd.CommandText = "update SHIPPING_LINE SET status=2 where order_id=" & oid & " and item_ID=" & rlID
    err.Clear
    Set cmd.ActiveConnection = conn
    cmd.Execute
    If err.Number <> 0 Then
      MsgBox err.Description
    End If
  Next
  
  cmd.CommandText = "delete from buf_loc where id=" & oid
  Set cmd.ActiveConnection = conn
  err.Clear
  cmd.Execute
  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
  
End Sub

'����������� ������� �����������
Private Sub gr2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
  On Error Resume Next
  Item.ITTOUT_LINES.LoadRow gr2, RowIndex, Bookmark, Values
End Sub
