VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Прием груза"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Шаг 1 - Выбор заказа"
      Height          =   6855
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtSupplier 
         Height          =   300
         Left            =   120
         MaxLength       =   255
         TabIndex        =   64
         ToolTipText     =   "Поставщик"
         Top             =   2130
         Width           =   3000
      End
      Begin VB.TextBox txtTTN 
         Height          =   300
         Left            =   120
         MaxLength       =   30
         TabIndex        =   63
         ToolTipText     =   "Номер ТТН"
         Top             =   2835
         Width           =   3000
      End
      Begin VB.TextBox txtTranspNumber 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   61
         ToolTipText     =   "№ ТС"
         Top             =   4245
         Width           =   3000
      End
      Begin VB.TextBox txtContainer 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   60
         ToolTipText     =   "№ прицепа \ контейнера"
         Top             =   4950
         Width           =   3000
      End
      Begin VB.TextBox txtStampNumber 
         Height          =   300
         Left            =   120
         MaxLength       =   20
         TabIndex        =   59
         ToolTipText     =   "Номер пломбы"
         Top             =   5655
         Width           =   3000
      End
      Begin VB.TextBox txtStampStatus 
         Height          =   300
         Left            =   120
         MaxLength       =   30
         TabIndex        =   58
         ToolTipText     =   "Состояние пломбы"
         Top             =   6360
         Width           =   3000
      End
      Begin VB.TextBox txtTheClient 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1440
         Width           =   6615
      End
      Begin VB.TextBox txtQryCode 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Код заказа"
         Top             =   690
         Width           =   6015
      End
      Begin MTZ_PANEL.DropButton cmdQryCode 
         Height          =   300
         Left            =   6240
         TabIndex        =   8
         Tag             =   "refopen.ico"
         ToolTipText     =   "Код заказа"
         Top             =   690
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin MSComCtl2.DTPicker dtpTTNDate 
         Height          =   300
         Left            =   120
         TabIndex        =   62
         ToolTipText     =   "Дата ТТН"
         Top             =   3540
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd.MM.yyyy"
         Format          =   53411843
         CurrentDate     =   39006
      End
      Begin MSMask.MaskEdBox txttemp_in_track 
         Height          =   300
         Left            =   3360
         TabIndex        =   72
         ToolTipText     =   "Темпиратура"
         Top             =   3540
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
         Left            =   3360
         TabIndex        =   73
         ToolTipText     =   "Время убытия машины"
         Top             =   2835
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd.MM.yyyy HH:mm:ss"
         Format          =   53411843
         CurrentDate     =   39006
      End
      Begin MSComCtl2.DTPicker dtpTrack_time_in 
         Height          =   300
         Left            =   3360
         TabIndex        =   74
         ToolTipText     =   "Время прибытия машины"
         Top             =   2130
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd.MM.yyyy HH:mm:ss"
         Format          =   53411843
         CurrentDate     =   39006
      End
      Begin VB.Label lblTrack_time_in 
         BackStyle       =   0  'Transparent
         Caption         =   "Время прибытия машины:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3360
         TabIndex        =   77
         Top             =   1800
         Width           =   3000
      End
      Begin VB.Label lbltrack_time_out 
         BackStyle       =   0  'Transparent
         Caption         =   "Время убытия машины:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3360
         TabIndex        =   76
         Top             =   2505
         Width           =   3000
      End
      Begin VB.Label lbltemp_in_track 
         BackStyle       =   0  'Transparent
         Caption         =   "Темпиратура:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3360
         TabIndex        =   75
         Top             =   3210
         Width           =   3000
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "Поставщик:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   71
         Top             =   1800
         Width           =   3000
      End
      Begin VB.Label lblTTN 
         BackStyle       =   0  'Transparent
         Caption         =   "Номер ТТН:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   70
         Top             =   2505
         Width           =   3000
      End
      Begin VB.Label lblTTNDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Дата ТТН:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   69
         Top             =   3210
         Width           =   3000
      End
      Begin VB.Label lblTranspNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "№ ТС:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   68
         Top             =   3915
         Width           =   3000
      End
      Begin VB.Label lblContainer 
         BackStyle       =   0  'Transparent
         Caption         =   "№ прицепа \ контейнера:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   67
         Top             =   4620
         Width           =   3000
      End
      Begin VB.Label lblStampNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Номер пломбы:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   66
         Top             =   5325
         Width           =   3000
      End
      Begin VB.Label lblStampStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Состояние пломбы:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   65
         Top             =   6030
         Width           =   3000
      End
      Begin VB.Label Label14 
         Caption         =   "Клиент"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblQryCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Код заказа:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3000
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Шаг 3 - Взвешивание поддона"
      Height          =   5775
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   8055
      Begin VB.TextBox txt3FromUser 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmd3ClearW 
         Caption         =   "x"
         Height          =   375
         Left            =   5280
         TabIndex        =   43
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton cmd3ClearNum 
         Caption         =   "x"
         Height          =   375
         Left            =   5280
         TabIndex        =   42
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txt3Good 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   840
         Width           =   5535
      End
      Begin VB.TextBox txt3InQry 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CommandButton cmdPNew 
         Caption         =   "Новый"
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdPPrint 
         Caption         =   "Печать стикера на поддон"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.TextBox txt3Weight 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5055
      End
      Begin VB.TextBox txt3Poddon 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   5055
      End
      Begin VB.Label Label15 
         Caption         =   "Количество в заказе"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label13 
         Caption         =   "Товар"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Осталось принять, планово"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Вес поддона"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Номер поддона"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Шаг 5 - Документ на поддон с грузом"
      Height          =   6615
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   10215
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
         Height          =   5175
         Left            =   480
         TabIndex        =   54
         Top             =   600
         Width           =   6255
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Шаг 4 - Вес поддона с грузом"
      Height          =   6975
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   7095
      Begin VB.TextBox txt4NewPlace 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   4920
         Width           =   2895
      End
      Begin VB.TextBox txt4FromUser 
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton cmd4ClearW 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   44
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txt4InQry 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txt4GoodWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txt4FullWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txt4CaliberQuantity 
         Height          =   405
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox txt4PWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txt4PNum 
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txt4CaliberWeight 
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   3960
         Width           =   2895
      End
      Begin VB.CheckBox chk4Caliber 
         Caption         =   "Калиброванный"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txt4Good 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label17 
         Caption         =   "Место в буферной зоне"
         Height          =   255
         Left            =   2760
         TabIndex        =   51
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "По заказу"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Осталось принять, планово"
         Height          =   255
         Left            =   2760
         TabIndex        =   35
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Вес груза на паддоне"
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Вес груза с поддоном"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Количество коробов"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Вес поддона"
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Поддон №"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Калиброванный вес"
         Height          =   615
         Left            =   2760
         TabIndex        =   22
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Товар"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Шаг 6 - Закрытие заказа"
      Height          =   6255
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   12135
      Begin VB.CommandButton cmd6Close 
         Caption         =   "Закрыть заказ"
         Height          =   615
         Left            =   360
         TabIndex        =   57
         Top             =   5280
         Width           =   3735
      End
      Begin VB.CommandButton cmd6PRNSRV 
         Caption         =   "Печать документа на услуги"
         Height          =   495
         Left            =   4320
         TabIndex        =   56
         Top             =   4560
         Width           =   3735
      End
      Begin VB.CommandButton cmd6PrnKL 
         Caption         =   "Печать контрольного листа"
         Height          =   495
         Left            =   360
         TabIndex        =   55
         Top             =   4560
         Width           =   3735
      End
      Begin VSFlex8Ctl.VSFlexGrid srvGrid 
         Height          =   3975
         Left            =   360
         TabIndex        =   53
         Top             =   360
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
         FormatString    =   $"frmInWiz.frx":0000
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
   Begin VB.Frame Frame2 
      Caption         =   "Шаг 2 - Выбор строки заказа"
      Height          =   6735
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Изменить"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   2055
      End
      Begin GridEX20.GridEX gr 
         Height          =   5835
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   6960
         _ExtentX        =   12277
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
         Column(1)       =   "frmInWiz.frx":0065
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmInWiz.frx":00C9
         FormatStyle(2)  =   "frmInWiz.frx":01A9
         FormatStyle(3)  =   "frmInWiz.frx":0305
         FormatStyle(4)  =   "frmInWiz.frx":03B5
         FormatStyle(5)  =   "frmInWiz.frx":0469
         FormatStyle(6)  =   "frmInWiz.frx":0541
         FormatStyle(7)  =   "frmInWiz.frx":05F9
         ImageCount      =   0
         PrinterProperties=   "frmInWiz.frx":0619
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отменить"
      Height          =   615
      Left            =   3000
      TabIndex        =   41
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin VB.CommandButton cmdAddW 
      Caption         =   "Следующий поддон"
      Height          =   615
      Left            =   6600
      TabIndex        =   34
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Следующий"
      Height          =   615
      Left            =   4800
      TabIndex        =   7
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Далее"
      Default         =   -1  'True
      Height          =   615
      Left            =   9720
      TabIndex        =   6
      Top             =   7680
      Width           =   1575
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Обработка заказа на прием груза"
      Height          =   855
      Left            =   3120
      TabIndex        =   11
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Image imgState 
      Height          =   8220
      Left            =   0
      Picture         =   "frmInWiz.frx":07F1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmInWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StepNo As Integer
Dim XMLQryCode As String
Dim XMLTheClient As String
Dim Item As ITTIN.Application
Dim conn As ADODB.Connection
Private curQRow As ITTIN.ITTIN_QLINE
Private StopWeighting As Boolean
Private wave As MTZMCI.WavePlayer
Private emu As Boolean
Private port As String
Private psetup As String
Private Poddon As ITTPL_DEF



' состояния для типа:ITTIN Приемка груза
' "{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}" 'Идет приемка
' "{49A919F7-94A6-49DE-9280-1EEAC973647B}" 'Оформляется
' "{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}" 'Приемка заершена
' "{E8BA9909-6680-4B2C-B446-F58EF91DCD17}" 'Приемка обработана


Private Sub SetBtnPos(cmd As CommandButton, ByVal pos As Integer)
  cmd.Left = imgState.Width + (Me.ScaleWidth - imgState.Width) / 4 * (pos - 1)
End Sub

Private Sub chk4Caliber_Click()
  If chk4Caliber.Value = vbChecked Then
    'txt4CaliberQuantity.Enabled = True
    txt4CaliberWeight.Enabled = True
  Else
    'txt4CaliberQuantity.Enabled = False
    txt4CaliberWeight.Enabled = False
  End If
End Sub


Private Sub cmd3ClearNum_Click()
txt3Poddon = ""
End Sub

Private Sub cmd3ClearW_Click()
txt3Weight = 0
End Sub

Private Sub cmd4ClearW_Click()
txt4FullWeight = 0
End Sub

Private Sub cmd6Close_Click()
  If MsgBox("Закрыть заказ ", vbYesNo) = vbYes Then
    Item.StatusID = "{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}"
  End If
End Sub

Private Sub cmd6PrnKL_Click()
    Dim repShow As ReportShow
    Set repShow = New ReportShow
    repShow.ReportSource = "V_viewITTIN_ITTIN_PALET"
    repShow.ReportFilter = " instanceid='" & Item.id & "'"
    repShow.ReportPath = App.Path & "\in_KL.rpt"
    repShow.PrinterName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShow.PrintOut
    Set repShow = Nothing
End Sub

Private Sub cmd6PRNSRV_Click()
    Dim repShow As ReportShow
    Set repShow = New ReportShow
    repShow.ReportSource = "V_viewITTIN_ITTIN_SRV"
    repShow.ReportFilter = " instanceid='" & Item.id & "'"
    repShow.ReportPath = App.Path & "\in_srvq.rpt"
    repShow.PrinterName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShow.PrintOut
    Set repShow = Nothing
End Sub

Private Sub cmdAddW_Click()
    If CheckAfter Then
      StepNo = 3
      ProcessStatus
    End If
End Sub

Private Sub cmdBack_Click()
  If CheckAfter Then
    If StepNo = 6 Then
      StepNo = 1
      ProcessStatus
    End If
    If StepNo = 5 Then
      StepNo = 2
      ProcessStatus
    End If
  End If
End Sub

Private Sub cmdCancel_Click()
  StepNo = 6
  ProcessStatus
End Sub

Private Sub cmdEdit_Click()
  gr_DblClick
End Sub

Private Sub cmdNext_Click()
  If CheckAfter Then
    StepNo = StepNo + 1
    ProcessStatus
  End If
End Sub






Private Sub cmdQryCode_Click()
  On Error Resume Next

    
'    txtQryCode.Tag = Replace(txtQryCode.Tag, "<ID>", "<IDOld>")
'    txtQryCode.Tag = Replace(txtQryCode.Tag, "</ID>", "</IDOld>")
'
 
  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtQryCode.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
  Else
    txtQryCode.Tag = Replace(txtQryCode.Tag, "%ID%", " 1=1 ")
    Call pars.Add("xml", txtQryCode.Tag)
  End If
  Set res = Manager.GetSQLDataDialog(pars)
  If (Not res Is Nothing) Then
    Dim resStr As String
    resStr = res.Item("RESULT").Value
    If (resStr = "OK") Then
      txtQryCode.Tag = res.Item("xml").Value
      If (txtQryCode.Text <> res.Item("brief").Value) Then
        txtQryCode.Text = res.Item("brief").Value
        'mIDQryCode = res.Item("ID").Value
        Call txtQryCode_Change
        MakeItem
        LoadHeader Item.ITTIN_DEF.Item(1)
      End If
    Else
      Dim errStr As String
      errStr = res.Item("ErrorDescription").Value
      If (errStr <> vbNullString) Then
       Call MsgBox("Ошибка исполнения: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
End Sub

Public Sub SaveHeader(Item As Object)
  Item.QryCode = txtQryCode.Tag
  Item.TheClient = txtTheClient.Tag
  Item.Supplier = txtSupplier
  Item.TTN = txtTTN
    If IsNull(dtpTTNDate) Then
      Item.TTNDate = 0
    Else
      Item.TTNDate = dtpTTNDate.Value
    End If
  Item.TranspNumber = txtTranspNumber
  Item.Container = txtContainer
  Item.StampNumber = txtStampNumber
  Item.StampStatus = txtStampStatus
    If IsNull(dtpTrack_time_in) Then
      Item.Track_time_in = 0
    Else
      Item.Track_time_in = dtpTrack_time_in.Value
    End If
    If IsNull(dtptrack_time_out) Then
      Item.track_time_out = 0
    Else
      Item.track_time_out = dtptrack_time_out.Value
    End If
  Item.temp_in_track = CDbl(txttemp_in_track)
  Item.save
End Sub

Private Sub MakeItem()
'Найти заказ у в нашей базе
  Dim rs As ADODB.Recordset
  Dim id As String
  Dim qID As String
  qID = Manager.GetIDFromXMLField(txtQryCode.Tag)
  id = ""
  Set rs = Session.GetData("select instanceid from ITTIN_DEF where QryCode like '%<ID>" & qID & "</ID>%'")
  If Not rs Is Nothing Then
    If Not rs.EOF Then
      id = rs!InstanceID
    End If
  End If
  rs.Close
  
  'Если нет заказа, то сформировать новый
  If id = "" Then
    id = CreateGUID2
    Manager.NewInstance id, "ITTIN", txtQryCode
    Set Item = Manager.GetInstanceObject(id)
    
    If conn.State <> ADODB.adStateOpen Then
      conn.Open
    End If
    
    Set rs = conn.Execute("select * from receiving_order where id=" & Manager.GetIDFromXMLField(txtQryCode.Tag))
    If rs.EOF Then Exit Sub
    
    
    With Item.ITTIN_DEF.Add
      .QryCode = txtQryCode.Tag
      .TheClient = txtTheClient.Tag
      .Supplier = rs!street1
      .TTN = rs!ACCOUNT_NUMBER
      .TTNDate = Date
      .TranspNumber = rs!Comment1
      .Container = rs!TRACK_NUMBER1
      .Track_time_in = Now
      .track_time_out = DateAdd("h", 4, Now)
      .temp_in_track = -1
      .save
    End With
    
    
    Dim XMLQRY_NUM As String
    Dim XMLLineAtQuery As String
    Dim XMLgood_ID As String
    
    Set rs = conn.Execute("select A.*, B.DESCRIPTION  BRIEF, B.code ARTICUL from receiving_line A join item B on A.item_id =B.id where a.order_id='" & qID & "'")
    While Not rs.EOF
      With Item.ITTIN_QLINE.Add
        
        .edizm = "" & rs!UOM
        .articul = "" & rs!articul
        .made_country = "" & rs!prod_country
        If Not IsNull(rs!Made_date) Then .Made_date = rs!Made_date
        If Not IsNull(rs!exp_date) Then .exp_date = rs!exp_date
        .KILL_NUMBER = "" & rs!KILL_NUMBER
        
        
        XMLLineAtQuery = "<SQLData>"
        XMLLineAtQuery = XMLLineAtQuery & "<connectionstring>ref</connectionstring>"
        XMLLineAtQuery = XMLLineAtQuery & "<connectionprovider>ref</connectionprovider>"
        XMLLineAtQuery = XMLLineAtQuery & "<query>select A.ID [Код], A.ORDER_ID [Код Заказа], A.QTY_ORD [Количество], B.DESCRIPTION [Наименование]  from receiving_line A join item B on A.item_id =B.id </query>"
        XMLLineAtQuery = XMLLineAtQuery & "<IDFieldName>Код</IDFieldName>"
        XMLLineAtQuery = XMLLineAtQuery & "<BriefFields>Наименование</BriefFields>"
        XMLLineAtQuery = XMLLineAtQuery & "<Brief>" & rs!brief & "</Brief>"
        XMLLineAtQuery = XMLLineAtQuery & "<ID>" & rs!id & "</ID>"
        XMLLineAtQuery = XMLLineAtQuery & "</SQLData>"
        
        .LineAtQuery = XMLLineAtQuery
        
        
        
        
        XMLQRY_NUM = "<SQLData>"
        XMLQRY_NUM = XMLQRY_NUM & "<connectionstring>ref</connectionstring>"
        XMLQRY_NUM = XMLQRY_NUM & "<connectionprovider>ref</connectionprovider>"
        XMLQRY_NUM = XMLQRY_NUM & "<query>select  QTY_ORD from receiving_line where ID='%LineAtQueryID%'</query>"
        XMLQRY_NUM = XMLQRY_NUM & "<IDFieldName>QTY_ORD</IDFieldName>"
        XMLQRY_NUM = XMLQRY_NUM & "<BriefFields>QTY_ORD</BriefFields>"
        XMLQRY_NUM = XMLQRY_NUM & "<ID>" & rs!QTY_ORD & "</ID>"
        XMLQRY_NUM = XMLQRY_NUM & "<Brief>" & rs!QTY_ORD & "</Brief>"
        XMLQRY_NUM = XMLQRY_NUM & "<LineAtQueryID>" & rs!id & "</LineAtQueryID>"
        XMLQRY_NUM = XMLQRY_NUM & "</SQLData>"
              
        .QRY_NUM = XMLQRY_NUM
         
        XMLgood_ID = "<SQLData>"
        XMLgood_ID = XMLgood_ID & "<connectionstring>ref</connectionstring>"
        XMLgood_ID = XMLgood_ID & "<connectionprovider>ref</connectionprovider>"
        XMLgood_ID = XMLgood_ID & "<query>select  item_id from RECEIVING_LINE where ID='%LineAtQueryID%'</query>"
        XMLgood_ID = XMLgood_ID & "<IDFieldName>ITEM_ID</IDFieldName>"
        XMLgood_ID = XMLgood_ID & "<BriefFields>ITEM_ID</BriefFields>"
        XMLgood_ID = XMLgood_ID & "<Brief>" & rs!item_id & "</Brief>"
        XMLgood_ID = XMLgood_ID & "<ID>" & rs!item_id & "</ID>"
        XMLgood_ID = XMLgood_ID & "<LineAtQueryID>" & rs!id & "</LineAtQueryID>"
        XMLgood_ID = XMLgood_ID & "</SQLData>"
        
        .good_id = XMLgood_ID
        
        .save
      End With
      rs.MoveNext
    Wend
    
    
    Set rs = Session.GetData("select * from ITTCS_DEF where clientcode like '%<ID>" & Manager.GetIDFromXMLField(txtTheClient.Tag) & "</ID>%'")
    Dim srvid As String
    Dim srvObj As ITTCS.Application
    Dim srv As ITTD_SRV
    srvid = rs!InstanceID
    Set srvObj = Manager.GetInstanceObject(srvid)
    Dim i As Long
    For i = 1 To srvObj.ITTCS_LIN.Count
       Set srv = srvObj.ITTCS_LIN.Item(i).srv
       If srv.ForReceiving = Boolean_Da Then
          If srvObj.ITTCS_LIN.Item(i).UseSrv = Boolean_Da Then
            With Item.ITTIN_SRV.Add
               Set .srv = srv
               .Quantity = 0
               .save
            End With
          End If
       End If
    Next
  Else
    Set Item = Manager.GetInstanceObject(id)
  End If
End Sub

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> 1 Then
    Cancel = -1
  Else
  wave.StopPlaying
  Set wave = Nothing
  End If
     
End Sub

Private Sub gr_DblClick()
If gr.ItemCount = 0 Then Exit Sub
    Dim u As Object
    Dim gui As Object
    Set gui = Manager.GetInstanceGUI(Item.id)
    
    Dim bm2
    bm2 = gr.RowBookmark(gr.RowIndex(gr.Row))
    Set u = Item.FindRowObject(Right(bm2, Len(bm2) - 38), Left(bm2, 38))
    If gui.ShowAddForm("", u) Then
      On Error Resume Next
      err.Clear
      u.save
      gr.RefreshRowBookmark bm2
    Else
        u.Refresh
    End If
    Set u = Nothing
Exit Sub
bye:
MsgBox err.Description, vbOKOnly + vbExclamation, "Изменение"
End Sub

Private Sub srvGrid_AfterEdit(ByVal Row As Long, ByVal col As Long)
  If col = 0 Then Exit Sub
  Item.ITTIN_SRV.Item(Row).Quantity = val(srvGrid.TextMatrix(Row, col))
  Item.ITTIN_SRV.Item(Row).save
End Sub

Private Sub srvGrid_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 0 Then Cancel = True
End Sub

Private Sub Timer1_Timer()
  Dim w As Double
  If StepNo = 3 Then
    
    If txt3Poddon = "" Then
      txt3Poddon.SetFocus
    End If
    
    If txt3Weight = "0" Or Not IsNumeric(txt3Weight) Then
      w = GetWeight
      If w > 0 Then
        txt3Weight = Round(w + 0.001, 1)
        MyBeep "Poddon"
      End If
    End If
    
  End If
  If StepNo = 4 Then
    If txt4FullWeight = "0" Or Not IsNumeric(txt4FullWeight) Then
      w = GetWeight
      If w > 0 And w > val(txt3Weight) + 5 Then
        txt4FullWeight = Round(w + 0.001, 1)
        MyBeep "Gruz"
      End If
  End If
  End If
End Sub

Private Sub txt3Poddon_Change()
  CheckPoddon
End Sub

Private Function CheckPoddon() As Boolean
  If txt3Poddon <> "" Then
    If Len(txt3Poddon) = 6 Then
      Set Poddon = Nothing
      Set Poddon = FindPoddon(txt3Poddon)
      If Not Poddon Is Nothing Then
        MyBeep "Nomer"
        txt3Weight = Poddon.Weight
      Else
        MsgBox "Номер паддона: " & txt3Poddon & "  не зарегистрирован"
      End If
    End If
  End If
End Function

Private Sub txt4CaliberWeight_Change()
  
  If val("0" & txt4GoodWeight) > 0 Then
    If val("0" & txt4CaliberWeight) > 0 Then
      txt4CaliberQuantity = txt4GoodWeight \ txt4CaliberWeight
    End If
  End If
End Sub

Private Sub txt4FullWeight_Change()
  On Error Resume Next
  txt4GoodWeight = Round(val(txt4FullWeight) - val(txt4PWeight) + 0.001, 1)
End Sub



Private Sub txtQryCode_Change()
If (txtQryCode.Text = "") Then
  ' Убрать Brief и ID
  If (txtQryCode.Tag <> "") Then
    Dim XMLDoc As New DOMDocument
    Call XMLDoc.loadXML(txtQryCode.Tag)
    Dim Node As MSXML2.IXMLDOMNode
    For Each Node In XMLDoc.childNodes.Item(0).childNodes
     If (Node.baseName = "ID") Then
       Node.Text = ""
     End If
     If (Node.baseName = "Brief") Then
       Node.Text = ""
     End If
    Next
    txtQryCode.Tag = XMLDoc.xml
  End If
End If

cmdTheClient_Click

End Sub
  

Private Sub Form_Load()
    StepNo = 0
    XMLQryCode = "<SQLData>"
    XMLQryCode = XMLQryCode & "<connectionstring>ref</connectionstring>"
    XMLQryCode = XMLQryCode & "<connectionprovider>ref</connectionprovider>"
    XMLQryCode = XMLQryCode & "<query>select A.ID [КОД] , convert(varchar(30),A.NUMBER) +'  от ' + convert(varchar(30),A.ORD_DATE,111)  [Название], B.Name [Клиент]  from RECEIVING_ORDER A left join PARTNER B on A.PARTNER_ID=B.ID</query>"
    XMLQryCode = XMLQryCode & "<IDFieldName>КОД</IDFieldName>"
    XMLQryCode = XMLQryCode & "<BriefFields>Название</BriefFields>"
    XMLQryCode = XMLQryCode & "</SQLData>"
    
  
    XMLTheClient = "<SQLData>"
    XMLTheClient = XMLTheClient & "<connectionstring>ref</connectionstring>"
    XMLTheClient = XMLTheClient & "<connectionprovider>ref</connectionprovider>"
    XMLTheClient = XMLTheClient & "<query>select partner.ID, partner.Name from RECEIVING_ORDER join partner on RECEIVING_ORDER.partner_id=partner.id where RECEIVING_ORDER.ID='%QryCodeID%'</query>"
    XMLTheClient = XMLTheClient & "<IDFieldName>ID</IDFieldName>"
    XMLTheClient = XMLTheClient & "<BriefFields>Name</BriefFields>"
    XMLTheClient = XMLTheClient & "</SQLData>"
    
    
    
    
    ProcessStatus
    Set conn = Manager.GetCustomObjects("refref")
    If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
      Set wave = New MTZMCI.WavePlayer
      wave.OpenDevice
    End If
    
End Sub

Private Sub AdjFrame(f As Frame)
  f.Top = 0
  f.Left = imgState.Width + 5 * Screen.TwipsPerPixelX
  f.Width = Me.ScaleWidth - imgState.Width - 10 * Screen.TwipsPerPixelX
  f.Height = Me.ScaleHeight - cmdNext.Height - 5 * Screen.TwipsPerPixelY
End Sub


Private Sub Before1()
    txtQryCode.Text = ""
    txtQryCode.Tag = XMLQryCode
    LoadBtnPictures cmdQryCode, cmdQryCode.Tag
    cmdQryCode.RemoveAllMenu
    txtTheClient.Text = ""
    txtTheClient.Tag = XMLTheClient
End Sub


Private Sub Before2()
  SaveHeader Item.ITTIN_DEF.Item(1)
  Dim repShow As ReportShow
  Set repShow = New ReportShow
  repShow.ReportSource = "V_viewITTIN_ITTIN_SRV"
  repShow.ReportFilter = " instanceid='" & Item.id & "'"
  repShow.ReportPath = App.Path & "\in_srv.rpt"
  repShow.PrinterName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
  repShow.PrintOut
  Set repShow = Nothing
  Item.StatusID = "{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}"
 
  'Инициализироать таблицу строк заказа
  gr.ItemCount = 0
  Item.ITTIN_QLINE.PrepareGrid gr
  gr.ItemCount = Item.ITTIN_QLINE.Count



End Sub

Private Sub Before3()
  On Error Resume Next
  txt3Poddon = ""
  txt3Weight = "0"
  
    If curQRow Is Nothing Then Exit Sub
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
  
      
      Dim XMLDocQRY_NUM As New DOMDocument
      Dim plan As Double
On Error Resume Next
  If (curQRow.QRY_NUM <> "") Then
    Call XMLDocQRY_NUM.loadXML(curQRow.QRY_NUM)
    If (err.Number = 0 And XMLDocQRY_NUM.parseError.errorCode = 0) Then
      Dim nodeQRY_NUM As MSXML2.IXMLDOMNode
      
      For Each nodeQRY_NUM In XMLDocQRY_NUM.childNodes.Item(0).childNodes
        If (nodeQRY_NUM.baseName = "Brief") Then
          plan = val("0" & nodeQRY_NUM.Text)
         Exit For
        End If
      Next
    End If
  End If
  
  txt3FromUser = plan
  txt3InQry.Text = plan - curQRow.CurValue

  
  
  If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
    Set wave = New MTZMCI.WavePlayer
    wave.OpenDevice
  End If

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


Private Sub Before4()
  If curQRow Is Nothing Then Exit Sub
  txt4NewPlace = ""
   Dim XMLDocLineAtQuery As New DOMDocument
   Call XMLDocLineAtQuery.loadXML(curQRow.LineAtQuery)
   If (err.Number = 0 And XMLDocLineAtQuery.parseError.errorCode = 0) Then
     Dim nodeLineAtQuery As MSXML2.IXMLDOMNode
     
     For Each nodeLineAtQuery In XMLDocLineAtQuery.childNodes.Item(0).childNodes
       If (nodeLineAtQuery.baseName = "Brief") Then
        txt4Good.Text = nodeLineAtQuery.Text
        Exit For
       End If
     Next
   End If
  
      
      Dim XMLDocQRY_NUM As New DOMDocument
      Dim plan As Double
On Error Resume Next
  If (curQRow.QRY_NUM <> "") Then
    Call XMLDocQRY_NUM.loadXML(curQRow.QRY_NUM)
    If (err.Number = 0 And XMLDocQRY_NUM.parseError.errorCode = 0) Then
      Dim nodeQRY_NUM As MSXML2.IXMLDOMNode
      
      For Each nodeQRY_NUM In XMLDocQRY_NUM.childNodes.Item(0).childNodes
        If (nodeQRY_NUM.baseName = "Brief") Then
          plan = val("0" & nodeQRY_NUM.Text)
         Exit For
        End If
      Next
    End If
  End If
  
  txt4FullWeight = 0
  txt4InQry.Text = plan - curQRow.CurValue
  
  txt4CaliberWeight = curQRow.CaliberWeight
  txt4CaliberQuantity = 0
  txt4PNum = txt3Poddon
  txt4PWeight = txt3Weight
  txt4FromUser = txt3FromUser
  
  ' состояния для типа:ITTPL Палетта
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана
  
  Dim P As ITTPL_DEF
  
  Set P = FindPoddon(txt3Poddon)
  If Not P Is Nothing Then
    P.Weight = val(txt3Weight)
    P.WDate = Date
    P.save
    P.Application.StatusID = "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}"
  End If
  
  If curQRow.IsCalibrated = Boolean_Da Then
    chk4Caliber.Value = vbChecked
    txt4CaliberWeight.Enabled = True
  Else
    chk4Caliber.Value = vbUnchecked
    txt4CaliberWeight.Enabled = False
  End If
  
  
  
End Sub

Private Sub Before5()
  Frame5.Visible = True
  AdjFrame Frame5
  lbl5Out = "Сохраняем информацию о поддоне"
  DoEvents
  Dim pal As ITTPL_DEF
  Dim LinePal As ITTIN_PALET
  Set LinePal = curQRow.ITTIN_PALET.Add
  With LinePal
    Set .TheNumber = FindPoddon(txt3Poddon)
    .PalWeight = val(txt3Weight)
    .GoodWithPaletWeight = val(txt4FullWeight)
    .CaliberQuantity = val(txt4CaliberQuantity)
    .BufferZonePlace = txt4NewPlace
    .save
  End With
  
  With curQRow
    If chk4Caliber.Value = vbChecked Then
     .IsCalibrated = Boolean_Da
     .CaliberWeight = val(txt4CaliberWeight)
    Else
     .IsCalibrated = Boolean_Net
     .CaliberWeight = 0
    End If
    .CurValue = .CurValue + val(txt4GoodWeight)
    .save
  End With
  If val(txt3InQry) < val(txt4GoodWeight) Then
    cmdAddW.Enabled = False
  Else
    cmdAddW.Enabled = True
  End If
  
  On Error Resume Next
  MSComm1.PortOpen = False
  
  SaveRCVRowToCore curQRow, LinePal
  
  Set pal = LinePal.TheNumber
  pal.CurrentPosition = txt4NewPlace
  pal.CurrentWeightBrutto = val(txt4FullWeight)
  pal.save
  
  pal.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}"
  
  
  
  lbl5Out = "Печатается документ на поддон"
  DoEvents
  
  Dim X As Printer
  For Each X In Printers
  If X.DeviceName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN") Then
  
  Set Printer = X
  Printer.Font = "Arial CYR"
  Printer.FontSize = 32

  Printer.Print "Поклажедатель: " & txtTheClient
  Printer.Print "Артикул: " & curQRow.articul & " Код: ";
  Printer.Font = "Code 128"
  Printer.FontSize = 48
  Printer.Print code128(curQRow.articul)
  
  Printer.Font = "Arial CYR"
  Printer.FontSize = 32
    
  Printer.Print "Заказ: " & txtQryCode
  Printer.Print "Партия: " & curQRow.PartRef.Name
  Printer.Print "Бойня №: " & curQRow.KILL_NUMBER.Name
  Printer.Print "Товар: " & txt4Good
  Printer.Print "Вес груза брутто (КГ.) : " & Round(val(txt4GoodWeight) - val(txt3Weight) + 0.001, 2)
  Printer.Print "Страна производитель: " & curQRow.made_country
  
  Printer.Print "Дата выпуска: " & curQRow.Made_date
  Printer.Print "Cрока годности: " & curQRow.exp_date
  
  
  If chk4Caliber.Value = vbChecked Then
    Printer.Print "Калиброванный товар"
    Printer.Print "Вес одного короба (КГ.): " & Round(txt4CaliberWeight + 0.001, 2)
  End If
  Printer.Print "Количество коробов: " & Round(txt4CaliberQuantity + 0.001, 0)
  Printer.NewPage
  
  lbl5Out = "Печатается документ напервичное размещение"
  DoEvents
  
  
  Printer.FontSize = 72
  Printer.Print "Поддон №"
  Printer.Print txt3Poddon
  Printer.Print "Буферная ячейка:"
  Printer.Print txt4NewPlace
  
  Printer.EndDoc
  
    
  lbl5Out = "Документы отправлены на принтер"
  DoEvents
  
   Exit For
  End If
  Next
bye2:
  
  Exit Sub
  
bye:
  If err.Number > 0 Then
    MsgBox err.Description, , "Печать документов на поддон"
  End If
  
End Sub

Private Sub Before6()
Dim i As Long
  srvGrid.Rows = Item.ITTIN_SRV.Count + 1
  For i = 1 To Item.ITTIN_SRV.Count
    srvGrid.TextMatrix(i, 0) = Item.ITTIN_SRV.Item(i).srv.brief
    srvGrid.TextMatrix(i, 1) = Item.ITTIN_SRV.Item(i).Quantity
  Next
End Sub

Private Sub Before7()

End Sub

Private Sub gr_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
  On Error Resume Next
  Item.ITTIN_QLINE.LoadRow gr, RowIndex, Bookmark, Values
End Sub

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

Private Sub ProcessStatus()
  Frame1.Visible = False
  Frame2.Visible = False
  Frame3.Visible = False
  Frame4.Visible = False
  Frame5.Visible = False
  Frame6.Visible = False
  cmdBack.Visible = False
  cmdNext.Caption = "Далее"
  cmdAddW.Visible = False
  cmdCancel.Visible = True

  Select Case StepNo
  Case 0
    cmdNext.Caption = "Начать процесс"
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  Case 1
  
    Before1
    Frame1.Visible = True
    AdjFrame Frame1
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 2
    Before2
    Frame2.Visible = True
    AdjFrame Frame2
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 3
    Before3
    Frame3.Visible = True
    AdjFrame Frame3
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  
  Case 4
    Before4
    Frame4.Visible = True
    AdjFrame Frame4

    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 5
    Before5
    Frame5.Visible = True
    AdjFrame Frame5
    cmdBack.Visible = True
    cmdAddW.Visible = True
    cmdBack.Caption = "Другая позиция заказа"
    cmdNext.Caption = "Закрыть заказ"
    
    If cmdAddW.Enabled Then
      SetBtnPos cmdCancel, 1
      SetBtnPos cmdNext, 2
      SetBtnPos cmdBack, 3
      SetBtnPos cmdAddW, 4
    Else
      SetBtnPos cmdCancel, 1
      SetBtnPos cmdBack, 4
      SetBtnPos cmdAddW, 2
      SetBtnPos cmdNext, 3
      
    End If
    
  Case 6
    Before6
    Frame6.Visible = True
    AdjFrame Frame6
    cmdBack.Visible = True
    cmdBack.Caption = "Следующий заказ"
    cmdNext.Caption = "Закрыть окно"
    cmdCancel.Visible = False
    
    SetBtnPos cmdCancel, 2
    SetBtnPos cmdBack, 3
    SetBtnPos cmdNext, 4
    
  Case 7
   Before7
   Unload Me
  End Select
  If StepNo >= 0 And StepNo < 7 Then
    imgState.Picture = LoadPicture(App.Path & "\Design\Step" & (StepNo) & ".bmp")
  Else
    imgState.Picture = LoadPicture(App.Path & "\Design\Step0.bmp")
  End If
End Sub


Private Function CheckAfter() As Boolean
  Dim result As Boolean
  
  Select Case StepNo
  Case 0
    ' do nothiing
    result = True
  Case 1
  
  
    ' Печать пустографки к заказу
    'After1
    result = True
    
  Case 2
    ' Выбрали строку заказа
    'After2
    result = True
    
  Case 3
    ' Взвесили поддон и ввели его номер
    '
     result = True
     
     If txt3Poddon = "" Then
      MsgBox "Считайте сканером, или введите на клавиатуре номер поддона"
      result = False
     End If
     
     If txt3Weight = "" Or Not IsNumeric(txt3Weight) Then
      MsgBox "Дождитесь полчения веса поддона с весов"
      result = False
     End If
     
     ' проверить состояние поддона в базе
     
     
  Case 4
    ' взвешиваем груз
      result = True
      If txt4FullWeight = "" Or txt4FullWeight = "0" Or Not IsNumeric(txt4FullWeight) Then
        MsgBox "Дождитесь полчения веса груза с весов"
        result = False
      ElseIf val(txt4FullWeight) > 1000 Then
        MsgBox "Вес поддона превышает 1000 кг."
        txt4FullWeight = 0
        result = False
      End If
    
  Case 5
     
    result = True
    
  Case 6
   ' сохраняем заказ
    result = True
  Case 7
  '
  result = True
  End Select
  CheckAfter = result
End Function



Public Function GetWeight4() As Double
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
    
    GetWeight4 = 0  ' не дождались ответа
    Exit Function
    
answer_s1:
    
    ws = MSComm1.Input
    ' первый раз вес стабилен
    If Asc(Mid(ws, 1, 1)) >= 128 Then
    
      ''''''''''''''''''''''''''''''''''''
      'ЗАДЕРЖКА !!!
      '
      ' ждем чтобы исключить дребезг
      start = Timer   ' Set start time.
      Do While Timer < start + 0.3
      Loop
      
      ' спрашиваем еще раз
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
    
    GetWeight4 = 0 ' нет второго ответа
    Exit Function
    
answer_s2:

    ws = MSComm1.Input
    
    ' второй раз вес тоже стабилен
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
    
    GetWeight4 = 0 ' вес не стабилен, или нет ответа
    Exit Function
    
answer_w1:

    ' прочли показания веса
    ws1 = MSComm1.Input
    
    
    ''''''''''''''''''''''''''''''''''''
    'ЗАДЕРЖКА !!!
    '
    ' ждем чтобы исключить дребезг
    start = Timer   ' Set start time.
    Do While Timer < start + 0.3
    Loop
    
    ' спрашиваем вес еще раз
    MSComm1.output = Chr(69)
    start = Timer   ' Set start time.
    Do While Timer < start + 0.2
    Loop
    
    If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    start = Timer   ' Set start time
    Do While Timer < start + 0.5
       If MSComm1.InBufferCount > 0 Then GoTo answer_w2
    Loop
    
    GetWeight4 = 0 '  нет ответа
    Exit Function
      
answer_w2:
    ws = MSComm1.Input
  
    If ws1 = ws Then
      GetWeight4 = (Asc(Mid(ws, 2, 1)) * 256 + Asc(Mid(ws, 1, 1))) / 10
    Else
      GetWeight4 = 0 ' вес не стабилен, отличаются показания
    End If
  
End Function

Private Function GetWeight() As Double
  If emu Then
    If StepNo = 4 Then
      GetWeight = Rnd(Second(Now)) * 1000 + val("0" & txt3Weight)
    Else
      GetWeight = Rnd(Second(Now)) * 40
    End If
  Else
    GetWeight = GetWeight4
  End If
End Function

Private Sub MyBeep(ByVal BeepType As String)
      If Not wave Is Nothing Then
        On Error Resume Next
        wave.OpenFile App.Path & "\" & BeepType & ".wav"
        wave.Play
      End If
End Sub

Private Sub cmdTheClient_Click()
  On Error Resume Next
  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtTheClient.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
  Else
    Call pars.Add("permanent", "true")
    txtTheClient.Tag = AddSQLRefIds(txtTheClient.Tag, "QryCode", txtQryCode.Tag)
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
'        mIDTheClient = res.Item("ID").Value
        'Call txtTheClient_Change
      End If
    Else
      Dim errStr As String
      errStr = res.Item("ErrorDescription").Value
      If (errStr <> vbNullString) Then
       Call MsgBox("Ошибка исполнения: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
End Sub


Private Sub SaveRCVRowToCore(ByVal CurRow As ITTIN_QLINE, LinePal As ITTIN_PALET)
  Dim conn As ADODB.Connection
  Set conn = Manager.GetCustomObjects("refref")
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rlID As String
  Dim palID As String
  
  rlID = Manager.GetIDFromXMLField(CurRow.LineAtQuery)
  palID = LinePal.TheNumber.CorePalette_ID
  
  
  ' запрашиваем свободное место в буферной зоне
  Dim bzrs As ADODB.Recordset
  Dim bzid As String
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  
  
  Set bzrs = conn.Execute( _
    "select  distinct location_id id from stock join location on location.id = stock.location_id " & _
    " Where stock.item_iD = " & Manager.GetIDFromXMLField(curQRow.good_id) & _
    " group by location_id,location.description " & _
    " having count(*) < convert(integer, substring(location.description, 0,charindex(';',location.description,0)))" _
  )
  
  
  Dim s As String

  If Not bzrs.EOF Then
    bzid = bzrs!id

      s = "insert into stock(SITE_ID,ITEM_ID,LOCATION_ID,ORDER_ID,QTY_ON_HAND,status,UNIT_COST,UOM,LOT_SN,REF_NUM,ORD_NUM,PALLET_ID,custom_field1,custom_field6,custom_field11,custom_field5,exp_date)" & _
      "values(1," & Manager.GetIDFromXMLField(curQRow.good_id) & "," & bzid & ",null," & txt4GoodWeight & ",0,0,'" & curQRow.edizm & "','" & curQRow.PartRef.Name & "','" & txtQryCode.Text & "','" & txtQryCode.Text & "'," & palID & "," & LinePal.CaliberQuantity & ",'" & CurRow.made_country.Name & "','" & CurRow.KILL_NUMBER.Name & "','" & CurRow.Made_date & "','" & CurRow.exp_date & "') "
      
      Set cmd = New ADODB.Command
      cmd.CommandType = adCmdText
      cmd.CommandText = s
      Set cmd.ActiveConnection = conn
      On Error Resume Next
      cmd.Execute
       If err.Number > 0 Then
        MsgBox err.Description
      End If
  Else
  
  
  Set bzrs = conn.Execute( _
  "select top 100 id from location where description like '%;B%' and  id not in ( " & _
  " select location_id from stock where location_id is not null )")
  
    If Not bzrs.EOF Then
      bzid = bzrs!id
      
       s = "insert into stock(SITE_ID,ITEM_ID,LOCATION_ID,ORDER_ID,QTY_ON_HAND,status,UNIT_COST,UOM,LOT_SN,REF_NUM,ORD_NUM,PALLET_ID,custom_field1,custom_field6,custom_field11,custom_field5,exp_date)" & _
      "        values(1," & Manager.GetIDFromXMLField(curQRow.good_id) & "," & bzid & ",null," & txt4GoodWeight & ",0,0,'" & curQRow.edizm & "','" & curQRow.PartRef.Name & "','" & txtQryCode.Text & "','" & txtQryCode.Text & "'," & palID & "," & LinePal.CaliberQuantity & ",'" & CurRow.made_country.Name & "','" & CurRow.KILL_NUMBER.Name & "','" & CurRow.Made_date & "'," & MakeMSSQLDate(CurRow.exp_date) & ") "
            
      Set cmd = New ADODB.Command
      cmd.CommandType = adCmdText
      cmd.CommandText = s
      Set cmd.ActiveConnection = conn
      On Error Resume Next
      err.Clear
      cmd.Execute
     
      If err.Number <> 0 Then
        MsgBox err.Description
      End If
    End If
  End If
  
  If bzid <> "" Then
    Set bzrs = conn.Execute("select code from location where id=" & bzid)
    txt4NewPlace = bzrs!Code
    LinePal.BufferZonePlace = bzrs!Code
  End If
  
  Set rs = conn.Execute("select * from RECEIVING_LINE where id=" & rlID)
  If rs.EOF Then Exit Sub
  
  conn.BeginTrans
  err.Clear
  cmd.CommandText = "INSERT INTO RECEIVING_HISTORY( [REF_NUMBER], [QTY_REC], [UOM], [LOT_SN], [EXP_DATE], [UNIT_PRICE], [COMMENTS], [REC_DATE], [TRACK_NUMBER2], [TRACK_NUMBER3], [LOCATION], [PALLET], [CONTAINER], [STATUS], [ORDER_ID], [ITEM_ID], [USER_ID], custom_field1)" & _
  "VALUES( '" & txtQryCode.Text & "', " & (LinePal.GoodWithPaletWeight - LinePal.PalWeight) & ",'" & rs!UOM & "', '" & CurRow.PartRef.Name & "'," & MakeMSSQLDate(CurRow.exp_date) & " , 0, ' ', getdate(), '" & Item.ITTIN_DEF.Item(1).TranspNumber & "', '" & Item.ITTIN_DEF.Item(1).TranspNumber & "','" & LinePal.BufferZonePlace & "','" & palID & "', '" & Item.ITTIN_DEF.Item(1).Container & "', 0, " & rs!ORDER_ID & "," & rs!item_id & ",1," & LinePal.CaliberQuantity & " )"
  Set cmd.ActiveConnection = conn
  err.Clear
  cmd.Execute
  
  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
   cmd.CommandText = "update RECEIVING_LINE SET QTY_PREV_REC =" & CurRow.CurValue & ", MADE_DATE=" & MakeMSSQLDate(CurRow.Made_date) & ", EXP_DATE=" & MakeMSSQLDate(CurRow.exp_date) & ",PROD_COUNTRY='" & CurRow.made_country.Name & "',KILL_NUMBER='" & CurRow.KILL_NUMBER.Name & "',LOT_SN='" & CurRow.PartRef.Name & "' where ID=" & rlID
   err.Clear
  Set cmd.ActiveConnection = conn
  cmd.Execute
  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  conn.CommitTrans
End Sub
