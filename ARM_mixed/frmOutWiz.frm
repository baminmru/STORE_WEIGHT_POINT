VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOutWiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Отгрузка"
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
         Caption         =   "Идет сохра- нение данных в CORE IMS. Ждите."
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
      Caption         =   "Шаг6 -текущие результаты"
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
      Caption         =   "Шаг 7 - Ввод услуг к заказу"
      Height          =   5535
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   21135
      Begin VB.CommandButton cmdPrnRas 
         Caption         =   "Акт весовых расхождений"
         Height          =   495
         Left            =   2880
         TabIndex        =   85
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CommandButton cmd6PrnKL 
         Caption         =   "Печать отборочного листа"
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton cmd6PRNSRV 
         Caption         =   "Печать документа на услуги"
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
      Caption         =   "Шаг4 - Отгружаемый товар без поддона"
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
         Caption         =   "Вес груза НЕТТО"
         Height          =   375
         Left            =   2760
         TabIndex        =   84
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "Буферная ячейка для остатков"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   4200
         Width           =   5055
      End
      Begin VB.Label Label21 
         Caption         =   "Вес одной упаковки"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "Количество коробов"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   46
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Надо отгрузить, ориентировочно"
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Заказано"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Товар"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Вес отгружаемых коробов"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Шаг 1 - Выбор заказа"
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
         ToolTipText     =   "Поставщик"
         Top             =   2220
         Width           =   3000
      End
      Begin VB.TextBox txtTTN 
         Height          =   300
         Left            =   120
         MaxLength       =   30
         TabIndex        =   58
         ToolTipText     =   "Номер ТТН"
         Top             =   2925
         Width           =   3000
      End
      Begin VB.TextBox txtTranspNumber 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   56
         ToolTipText     =   "№ ТС"
         Top             =   4335
         Width           =   3000
      End
      Begin VB.TextBox txtContainer 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   55
         ToolTipText     =   "№ прицепа \ контейнера"
         Top             =   5040
         Width           =   3000
      End
      Begin VB.TextBox txtStampNumber 
         Height          =   300
         Left            =   120
         MaxLength       =   20
         TabIndex        =   54
         ToolTipText     =   "Номер пломбы"
         Top             =   5745
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.TextBox txtStampStatus 
         Height          =   300
         Left            =   120
         MaxLength       =   30
         TabIndex        =   53
         ToolTipText     =   "Состояние пломбы"
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
         ToolTipText     =   "Код заказа"
         Top             =   690
         Width           =   6015
      End
      Begin MTZ_PANEL.DropButton cmdShipOrder 
         Height          =   300
         Left            =   6240
         TabIndex        =   2
         Tag             =   "refopen.ico"
         ToolTipText     =   "Код заказа"
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
         ToolTipText     =   "Темпиратура"
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
         ToolTipText     =   "Время убытия машины"
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
         ToolTipText     =   "Время прибытия машины"
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
         ToolTipText     =   "Дата ТТН"
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
         Caption         =   "Получатель:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   69
         Top             =   1890
         Width           =   3000
      End
      Begin VB.Label lblTTN 
         BackStyle       =   0  'Transparent
         Caption         =   "Номер ТТН:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   68
         Top             =   2595
         Width           =   3000
      End
      Begin VB.Label lblTTNDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Дата ТТН:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   67
         Top             =   3300
         Width           =   3000
      End
      Begin VB.Label lblTranspNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "№ ТС:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   66
         Top             =   4005
         Width           =   3000
      End
      Begin VB.Label lblContainer 
         BackStyle       =   0  'Transparent
         Caption         =   "№ прицепа \ контейнера:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   65
         Top             =   4710
         Width           =   3000
      End
      Begin VB.Label lblStampNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Номер пломбы:"
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
         Caption         =   "Состояние пломбы:"
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
         Caption         =   "Время прибытия машины:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3390
         TabIndex        =   62
         Top             =   1920
         Width           =   3000
      End
      Begin VB.Label lbltrack_time_out 
         BackStyle       =   0  'Transparent
         Caption         =   "Время убытия машины:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3390
         TabIndex        =   61
         Top             =   2625
         Width           =   3000
      End
      Begin VB.Label lbltemp_in_track 
         BackStyle       =   0  'Transparent
         Caption         =   "Темпиратура:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3390
         TabIndex        =   60
         Top             =   3330
         Width           =   3000
      End
      Begin VB.Label Label14 
         Caption         =   "Клиент"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblQryCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Код заказа:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3000
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Шаг 3 - Поддон с грузом"
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
            Caption         =   "Идет проверка возможности отгрузки товара с текущего поддона."
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
         ToolTipText     =   "Получить вес с  весов"
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
         ToolTipText     =   "ввести номер еще раз"
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Ячейка основного хранения"
         Height          =   375
         Left            =   2760
         TabIndex        =   82
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Label Label20 
         Caption         =   "Вес одной упаковки"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label19 
         Caption         =   "Вес поддона"
         Height          =   375
         Left            =   2760
         TabIndex        =   71
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "Количество коробов"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "Заказано"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Надо отгрузить, ориентировочно"
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Вес груза НЕТТО"
         Height          =   375
         Left            =   2760
         TabIndex        =   29
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Вес груза с поддоном"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Вес поддона"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Поддон"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Товар"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Шаг5 - Печать стикера на перепакованный товар"
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
      Caption         =   "Шаг 2 - Выбор строки заказа"
      Height          =   6495
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   8535
      Begin VB.CommandButton cmdToClosePage 
         Caption         =   "перейти к итоговой странице"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Изменить"
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
      Caption         =   "Следующий"
      Height          =   615
      Left            =   5160
      TabIndex        =   33
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddW 
      Caption         =   "Следующий поддон"
      Height          =   615
      Left            =   6960
      TabIndex        =   32
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отменить"
      Height          =   615
      Left            =   3360
      TabIndex        =   31
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Далее"
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
' Окно визарда отгрузки





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

' состояния для типа:ITTPL Палетта
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана


' уточнение позиций кнопок визарда
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
  If MsgBox("Закрыть заказ ", vbYesNo) = vbYes Then
    'Item.StatusID = "{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}"
  End If
End Sub


'запуск отчета по отборочному листу
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

'запуск отчета по услугам
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

' следующий поддон
Private Sub cmdAddW_Click()
On Error Resume Next
    If CheckAfter Then
      StepNo = 3
      ProcessStatus
    End If
End Sub

' назад
Private Sub cmdBack_Click()
On Error Resume Next
  If CheckAfter Then
      StepNo = 2
      ProcessStatus
  End If
End Sub

' отмена
Private Sub cmdCancel_Click()
On Error Resume Next
  StepNo = 8
  ProcessStatus
End Sub

'открытие строки заказа
Private Sub cmdEdit_Click()
On Error Resume Next
 'gr_DblClick
End Sub

'далее
Private Sub cmdNext_Click()
On Error Resume Next
  If CheckAfter Then
    If StepNo = 3 Then
      If MsgBox("Отгружаем текщую палету целиком ?", vbYesNo, "Уточните") = vbYes Then
        'If MsgBox("Зарегистрировать отгрузку палеты?", vbExclamation + vbYesNo, "Внимание") = vbYes Then
          StepNo = 6
          isFull = True
'        Else
'          log.message "Отказ от отгрузки паллеты"
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






'печать акта о расхождениях
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

'выбор заказа на отгрузку
Private Sub cmdShipOrder_Click()
On Error Resume Next
  On Error Resume Next
  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtShipOrder.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
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
       Call MsgBox("Ошибка исполнения: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
  log.message "Отгрузка " & txtShipOrder
End Sub

' переход к закрытию заказа
Private Sub cmdToClosePage_Click()
StepNo = 7
ProcessStatus
End Sub

'поиск ячейки для остатков
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

'сброс веса
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




'запись данных по объему услуг
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


'сбор данных с весов
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



' изменен номер поддона
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
  
   
  
  'log.Message "Проверка паддона " & poddon.TheNumber
  
  If CheckPoddon(cm) Then
    
   'log.Message "Возможна отгрузка паддона " & poddon.TheNumber
   
   Set rs = conn.Execute("select * from stock where PALLET_STATUS is null and pallet_id=" & poddon.CorePalette_ID)
  
   If GetSetting("RBH", "ITTSETTINGS", "RESTORE", "False") = "False" Then
    
    
    txt3PWeight = poddon.Weight
    txt3FullWeight = 0
    txt3Quantity = rs!custom_field1
    txt3PackageWeight = rs!custom_field3
    cmd3ClearW.Enabled = True
    If isCalibrated Then
      'log.Message "Восстановление веса из базы для калиброванного товара!!!"
      txt3GoodWeight = rs!QTY_ON_HAND
      txt3FullWeight = rs!QTY_ON_HAND + Val(rs!custom_field1) * Val(rs!custom_field3) + poddon.Weight
      cmd3ClearW.Enabled = False
    End If
   Else
    'log.Message "Восстановление веса из базы!!!"
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


'проверка поддона
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
            MsgBox "Номер паддона: " & txt3PNum & "  не обнаружен в базе CORE IMS"
            result = False
          Else
            If rs!item_id <> Manager.GetIDFromXMLField(curQRow.good_id) Then
              MsgBox "Артикул груза на палетте не совпадает с артикулом заказа"
              result = False
            Else
            
            
              If rs!status = 103 Then
                    MsgBox "Поддон заблокирован для отгрузки (выморозка).", vbExclamation + vbOKOnly, "Внимание"
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
                  ' проверяем можно ли отгружать из данной партии с учетом выморозки
                  'Set rs2 = conn.Execute("select * from v_bami_vimorozka_rpt2 A  join stock B on  checksum(a.item_id, a.factory , a.country, a.Kill_place, a.IsBrak, a.made_date_to, a.vetsved) = " & _
                  '"checksum(b.item_id,b.custom_field4,b.custom_field6,b.custom_field11,b.custom_field12,b.custom_field9,b.custom_field7)  and b.PALLET_STATUS is null and b.pallet_id=" & poddon.CorePalette_ID)
                  Set rs2 = conn.Execute("exec  CheckPartiaMoroz " & poddon.CorePalette_ID & " ")
                  If Not rs2.EOF Then
                    If rs2!to_ship > 0 Then
                        If rs2!to_ship < rs!QTY_ON_HAND Then
                          MsgBox "C данного поддона может быть отгружено только " & rs2!to_ship & " кг. товара", vbExclamation + vbOKOnly, "Внимание"
                          
                          Dim mail As STDMail.Application
                          Dim idmail As String
                          idmail = CreateGUID2()
                          Manager.NewInstance idmail, "STDMail", "Оповещение " & Now
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
                              .Subject = "Оповещение от " & Now
                              .TheBody = "Отгрузка товара c поддона '" & poddon.code & "'  в количестве (" & rs!QTY_ON_HAND & ") превышает объем 'к отргузке' (" & rs2!to_ship & ") "
                              .TheBody = .TheBody & " для товара:" & vbCrLf & rs2!item_code & " " & rs2!Description
                              .TheBody = .TheBody & " страна:" & rs2!country & " завод: " & rs2!factory & " бойня:" & rs2!kill_place
                              .Sended = Boolean_Net
                              .save
                            End With
                            
                          End If
                       
                        End If

                      Else
                        
                            MsgBox "Отгрузка товара по данной партии заблокирована.", vbExclamation + vbOKOnly, "Внимание"
                            result = False
                      End If
'                    Else
                        
'                          MsgBox "Отгрузка товара по данной партии заблокирована.", vbExclamation + vbOKOnly, "Внимание"
'                          result = False
'
                  End If
                  frameWait.Visible = False
                End If
                  
              End If
            End If
          End If
        Else
          MsgBox "Состояние паддона: " & txt3PNum & "  установлено неверно (" & poddon.Application.StatusName & ")"
          result = False
        End If
      Else
        MsgBox "Номер паддона: " & txt3PNum & "  не зарегистрирован"
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
  XMLShipOrder = XMLShipOrder & "<query>select A.ID [КОД] , convert(varchar(30),A.NUMBER) +'  от ' + convert(varchar(30),A.ORD_DATE,111)  [Название], PARTNER.Name [Клиент]  from shipping_ORDER A left join PARTNER  on A.PARTNER_ID=PARTNER.ID where (a.STATUS = 1 or a.status =0) </query>"
  XMLShipOrder = XMLShipOrder & "<IDFieldName>КОД</IDFieldName>"
  XMLShipOrder = XMLShipOrder & "<BriefFields>НАЗВАНИЕ</BriefFields>"
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

'установка очередного фрейма в видимую позицию
Private Sub AdjFrame(f As Frame)
On Error Resume Next
  f.Top = 0
  f.Left = imgState.Width + 5 * Screen.TwipsPerPixelX
  f.Width = Me.ScaleWidth - imgState.Width - 10 * Screen.TwipsPerPixelX
  f.Height = Me.ScaleHeight - cmdNext.Height - 5 * Screen.TwipsPerPixelY
End Sub

'до первого шага визарда - Шаг 1 - Выбор заказа
Private Sub Before1()
On Error Resume Next
    txtShipOrder.Text = ""
    txtShipOrder.Tag = XMLShipOrder
    LoadBtnPictures cmdShipOrder, cmdShipOrder.Tag
    cmdShipOrder.RemoveAllMenu
    txtTheClient.Text = ""
    txtTheClient.Tag = XMLTheClient
End Sub

'создать заказ
Private Sub MakeItem()
On Error Resume Next
'Найти заказ у в нашей базе
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
    
  'Если нет заказа, то сформировать новый
  If id = "" Then
    
    
    Dim errlines As Boolean
    ' Проверяем что код товара в заказе соответсвует коду поклажедателя
    Set rs = conn.Execute("select B.code b_code,e.code e_code,d.code d_code from shipping_line A  join item B on A.item_id =B.id  join shipping_order C on a.order_id = C.id join partner D on c.partner_id= d.id join partner E on b.CLASS= e.CODE where e.code <>d.code and a.order_id='" & qID & "'")
    While Not rs.EOF
        MsgBox "Для артикула с кодом " & rs!b_code & " не соответствует поклажедатель и выбранный в заказе клиент" & vbCrLf & "Исправьте ошибку в CORE IMS", vbOKOnly + vbExclamation, "Внимание"
        errlines = True
        rs.MoveNext
    Wend
    
    If errlines Then Exit Sub
    
    
    
    ' получаем описание заказа из core
    Set rs = conn.Execute("select * from shipping_order where id=" & Manager.GetIDFromXMLField(txtShipOrder.Tag))
    If rs.EOF Then Exit Sub
    
    
'    создаем новый заказ
    id = CreateGUID2
    Manager.NewInstance id, "ITTOUT", txtShipOrder
    Set Item = Manager.GetInstanceObject(id)
    
    
'    заполняем описание
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
    
    
'    формируем строки заказа в базе данных ВК
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
        XMLLineAtQuery = XMLLineAtQuery & "<query>select A.ID [Код], A.ORDER_ID [Код заказа], A.QTY_ORD [Количество] , B.DESCRIPTION [Название] from shipping_line A join item B on A.item_id =B.id </query>"
        XMLLineAtQuery = XMLLineAtQuery & "<IDFieldName>Код</IDFieldName>"
        XMLLineAtQuery = XMLLineAtQuery & "<BriefFields>Название</BriefFields>"
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
        
'        пытаемся восстановить ссылки на справочники по данным кастомных полей в CORE
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
      ' нумеруем строку заказа в БД ВК
      Call GetNumValue(curQRow, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "OUT%P", "")
      rs.MoveNext
    Wend
    
'    заполняем список услуг по данным справочника для текущего клиента
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

'загрузка заголовка заказа для отображения
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

'до второго шага визарда - Шаг 2 - Выбор строки заказа
Private Sub Before2()
  If MsgBox("Напечатать пустографку?", vbYesNo) = vbYes Then
    Set repShowSRVOUT = Nothing
    Set repShowSRVOUT = New ReportShow
    repShowSRVOUT.ReportSource = "V_viewITTOUT_ITTOUT_SRV"
    repShowSRVOUT.ReportFilter = " instanceid='" & Item.id & "'"
    repShowSRVOUT.ReportPath = App.Path & "\OUt_srv.rpt"
    repShowSRVOUT.PrinterName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowSRVOUT.Run True
  End If
  
' состояния для типа:ITTOUT Отгрузка
' "{70853C28-84B5-434E-8413-52DF8FBBB49B}" 'Идет отгрузка
' "{2CDDB562-63D7-483E-B95E-B579A9096CCC}" 'Обработка завершена
' "{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}" 'Отгрузка завершена
' "{CDCAFF7F-B013-40AF-BE61-1A27E35DB946}" 'Оформляется
  
  Item.StatusID = "{70853C28-84B5-434E-8413-52DF8FBBB49B}" 'Идет отгрузка
  
  'Инициализироать таблицу строк заказа
  gr.ItemCount = 0
  'Item.ITTOUT_LINES.Sort = "sequence"
  Item.ITTOUT_LINES.PrepareGrid gr
  Item.ITTOUT_LINES.Refresh
  gr.ItemCount = Item.ITTOUT_LINES.Count


End Sub

'до третьего шага визарда - Шаг 3 - Поддон с грузом
Private Sub Before3()
  On Error Resume Next
  txt3PNum = ""
  txt3FullWeight = "0"
  txt3Good = 0
  txt3PWeight = 0
  txt3Quantity = 0
  txt3PackageWeight = 0
  
  If curQRow Is Nothing Then Exit Sub
  
  
'  получае товар
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

'  получаем план по товару
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
  
'  заполняем поля количества
  txt3FRomQ = plan
  txt3InQry.Text = plan - curQRow.CurValue


  
'  инициализация звука
  If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
    Set wave = New MTZMCI.WavePlayer
    wave.OpenDevice
  End If

'  иициализация COM  порта
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

'до четвертого шага визарда - Шаг4 - Отгружаемый товар без поддона
Private Sub Before4()
  txt4FullWeight = 0
  txt4Good = txt3Good
  txt4FromQ = txt3FRomQ
  txt4InQry = txt3InQry
  txt4PackageWeight = txt3PackageWeight
  txt4NewPlace = txtMainCell
  txt4Quantity = 0
End Sub

'до пятого шага визарда - Шаг5 - Печать стикера на перепакованный товар
Private Sub Before5()

'
  
bye2:
  
  Exit Sub
  
bye:
  If err.Number <> 0 Then
    MsgBox err.Description, , "Печать документов на поддон"
  End If

End Sub


'до шестого шага визарда - Шаг6 -текущие результаты
Private Sub Before6()

  Dim strs As ADODB.Recordset
  Dim conn As ADODB.Connection
  
'  получили конект к core
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
'  нашли поддон
  Set poddon = FindPoddon(txt3PNum)
     
  Dim netto As Double
  Dim korob As Integer
  Dim OK As Boolean
  
'  считаем веса
  netto = MyRound(txt3GoodWeight)
  korob = MyRound(txt3Quantity)
  
    
  Dim morosrs As ADODB.Recordset
  Dim delta As Double
  Dim protID As String
  Dim prot As ITTPR.Application
  
  
  If isFull Then
    ' считаем выморозку
    Set morosrs = conn.Execute("select   min(LastRCV) LASTRCV  ,sum(in_quantity)  qin ,sum(in_boxes)  bin ,sum(out_quantity) qout  ,sum(out_boxes) bout  ,sum( dout_quantity) vimorozka  ,sum(stok_quantity) qstok from v_bami_vimorozka where pallet ='" & poddon.TheNumber & "' and rectype <>3")
  
    If Not morosrs Is Nothing Then
      delta = morosrs!qin - morosrs!qout - morosrs!qstok - morosrs!vimorozka * 0.0005 - morosrs!qout * 0.001 - morosrs!qin * 0.001
      delta = delta - netto - netto * 0.001 - netto / 30 * DateDiff("d", morosrs!lastrcv, Now) * 0.0005
      
      ' если больше чем можно со всеми погрешностями
      If delta > 0 Then
        ' создаем протокол
        protID = CreateGUID2
        Manager.NewInstance protID, "ITTPR", "Протокол расхождений на поддон №" & poddon.code
        Set prot = Manager.GetInstanceObject(protID)
        
'        заполняем протокол
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
            .brak = "БРАК"
          Else
            .brak = " - нет -"
          End If
          
'          сохраняем  протокол
          .save
          
        End With
        
    
        ' печатаем его
            
        Set RptActVes = New ReportShow
        RptActVes.ReportPath = App.Path & "\AktVes.rpt"
        RptActVes.ReportSource = "V_AUTOITTPR_DEF"
        RptActVes.ReportFilter = "instanceid ='" & protID & "'"
        Call RptActVes.Run(True)
        Set RptActVes = Nothing
        log.message "Создан акт о расхождении паддон:" & poddon.code
        
        
'        возможен отказ от отгрузки
        If MsgBox("Отгрузить поддон?", vbYesNo, "Уточните") = vbNo Then
            curQRow.ITTOUT_PALET.Refresh
            'poddon.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}"
            StepNo = 3
            ProcessStatus
            log.message "Отказ от отгрузки паддон:" & poddon.code
'            отказались
            Exit Sub
        End If
      
      End If
    End If
    
    
    ' меняем состояние подона
    ' состояния для типа:ITTPL Палетта
    ' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
    ' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
    ' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
    ' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
    ' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана
    
    curQRow.ITTOUT_PALET.Refresh
    Dim pweight As Double
    pweight = poddon.Weight
    
'    меняем состояние поддона при отгрузке
    If MsgBox("Отдаем палету клиенту ?", vbYesNo + vbDefaultButton2) = vbYes Then
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
    
 
   
   
   ' отражаем в заказе ВК
    
'    создаем описание паллеты
    Set LinePal = curQRow.ITTOUT_PALET.Add
    
    
'    заполняем описание
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
    
'    сохраняем
    LinePal.save
     
    Call GetNumValue(LinePal, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "OUTPAL%P", "")
    
    log.message "Отгрузка товара  " & txt3PNum & " вес " & MyRound(txt3GoodWeight)
    
    ' мы списываем всю палету целиком!
    curQRow.CurValue = curQRow.CurValue + MyRound(txt3GoodWeight)
    curQRow.save
    
    frameSave.Visible = True
    DoEvents

    ' записываем в CORE
    OK = False
    While Not OK
      OK = SaveShipRowToCore(txtShipOrder.Text, txt4NewPlace.Text, Item, poddon, curQRow, LinePal, True)
      If Not OK Then
         OK = Not MagicMessageBox("Не удалось сохранить информацию в CORE. Паллета № " & poddon.code & vbCrLf & "Повтор попытки сохранения данных")
        Dim conn2 As Object
        Set conn2 = GetCoreConn(True)
      End If
      
    Wend
    
    frameSave.Visible = False
    DoEvents
    
    ' записываем строку заказа
    Set curQRow.made_country = LinePal.made_country
    Set curQRow.factory = LinePal.factory
    Set curQRow.KILL_NUMBER = LinePal.KILL_NUMBER
    Set curQRow.PartRef = LinePal.PartRef
    curQRow.save
    
  Else
  
'  неполная отгрузка
    curQRow.save
'    создаем палету -в отгрузке
    Set LinePal = curQRow.ITTOUT_PALET.Add
     
     
'    уменьшаем данные на поддоне (остаток)
     Set poddon = FindPoddon(txt3PNum)
     poddon.CurrentWeightBrutto = MyRound(txt3FullWeight) - MyRound(txt4GoodWeight)
     poddon.CaliberQuantity = MyRound(txt3Quantity) - MyRound(txt4Quantity)
     poddon.PackageWeight = (MyRound(txt3Quantity) - MyRound(txt4Quantity)) * MyRound(txt3PackageWeight)
     err.Clear
     On Error Resume Next
     poddon.save
     
'     заполняем строку паллеты
     With LinePal
      Set .TheNumber = poddon
      .IsEmpty = Boolean_Net
      
      Set strs = conn.Execute("select * from STOCK where PALLET_STATUS is null and  PALLET_ID=" & LinePal.TheNumber.CorePalette_ID)
      
      ' столько отгрузили
      
      .GoodWithPaletWeight = MyRound(txt4FullWeight) + MyRound(txt3PWeight)
      
      If isCalibrated Then
        .GoodWithPaletWeight = (MyRound(txt4Quantity)) * (MyRound(txt3PackageWeight) + (strs!QTY_ON_HAND / IIf(Val(strs!custom_field1) = 0, 1, Val(strs!custom_field1)))) + MyRound(txt3PWeight)
        .FullPackageWeight = MyRound(txt4Quantity) * MyRound(txt3PackageWeight)
      Else
        .FullPackageWeight = (MyRound(txt4Quantity)) * MyRound(txt3PackageWeight)
      End If
      
      
      .PackageWeight = MyRound(txt3PackageWeight)
      .CaliberQuantity = MyRound(txt4Quantity)
      
      ' столько осталось на палетте
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
'      сохраняем паллету в вк
      .save
      
    End With
    
    Call GetNumValue(LinePal, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "OUTPAL%P", "")
    
    log.message "Отгрузка товара  " & txt3PNum & " вес " & MyRound(txt4GoodWeight)
    
    curQRow.CurValue = curQRow.CurValue + MyRound(txt4GoodWeight)
    curQRow.save
    
    

    
    
    frameSave.Visible = True
    DoEvents
    
    
   ' сохраняем в CORE
    OK = False
    While Not OK
      OK = SaveShipRowToCore(txtShipOrder.Text, txt4NewPlace.Text, Item, poddon, curQRow, LinePal, False)
      If Not OK Then
         OK = Not MagicMessageBox("Не удалось сохранить информацию в CORE. Паллета № " & poddon.code & vbCrLf & "Повтор попытки сохранения данных")
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
    
    
    ' печатаем стикер на остаток
    PrintSticker LinePal.TheNumber
    
  End If

  On Error Resume Next
  MSComm1.PortOpen = False
  
  ' обновляем список строк заказа в гриде
  gr2.ItemCount = 0
  Item.ITTOUT_LINES.PrepareGrid gr2
  gr2.ItemCount = Item.ITTOUT_LINES.Count
  
End Sub

'до седьмого шага визарда - Шаг 7 - Ввод услуг к заказу
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

'загрузка данных в таблицу строк заказа
Private Sub gr_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
  On Error Resume Next
  Item.ITTOUT_LINES.LoadRow gr, RowIndex, Bookmark, Values
End Sub

'выбор текущей строки заказа
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

'процесс смены шага визарда
Private Sub ProcessStatus()
  Frame1.Visible = False
  Frame2.Visible = False
  Frame3.Visible = False
  Frame4.Visible = False
  Frame5.Visible = False
  Frame6.Visible = False
  Frame7.Visible = False
  
  cmdBack.Visible = False
  cmdNext.Caption = "Далее"
  cmdAddW.Visible = False
  cmdCancel.Caption = "Отменить"
  cmdCancel.Visible = True

  Select Case StepNo
  Case 0
    cmdNext.Caption = "Начать процесс"
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  Case 1
  'Шаг 1 - Выбор заказа
    Before1
    Frame1.Visible = True
    AdjFrame Frame1
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 2
  'Шаг 2 - Выбор строки заказа
    Before2
    Frame2.Visible = True
    AdjFrame Frame2
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 3
  'Шаг 3 - Поддон с грузом
    Before3
    Frame3.Visible = True
    AdjFrame Frame3
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  
  Case 4
  'Шаг4 - Отгружаемый товар без поддона
    Before4
    Frame4.Visible = True
    AdjFrame Frame4
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 5
  'Шаг5 - Печать стикера на перепакованный товар
    Before5
    Frame5.Visible = True
    AdjFrame Frame5
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
   
  Case 6
  'Шаг6 -текущие результаты
    Before6
    If StepNo = 6 Then
    Frame6.Visible = True
    AdjFrame Frame6
    
    cmdBack.Visible = True
    cmdAddW.Visible = True
    cmdCancel.Visible = False
    cmdNext.Caption = "Закончить обработку"
    cmdBack.Caption = "Другая позиция заказа"
    
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
   'Шаг 7 - Ввод услуг к заказу
    before7
    Frame7.Visible = True
    AdjFrame Frame7
    
  Case 8
    Unload Me
  End Select
   
  ' грузим картинку соответствующую шагу
  If StepNo >= 0 And StepNo < 8 Then
    imgState.Picture = LoadPicture(App.Path & "\Design\LStep" & (StepNo) & ".bmp")
  Else
    imgState.Picture = LoadPicture(App.Path & "\Design\LStep0.bmp")
  End If
End Sub

'проверка состояния данных после шага
Private Function CheckAfter() As Boolean
  Dim result As Boolean
  
  Select Case StepNo
  Case 0
    ' do nothiing
    result = True
  Case 1
  
  
    ' Печать пустографки к заказу
    If txtShipOrder = "" Then
      result = False
      MsgBox "Следует выбрать заказ"
    Else
      result = True
    End If
    
    
  Case 2
    ' Выбрали строку заказа
    If curQRow Is Nothing Then
      result = False
      MsgBox "Следует выбрать строку заказа"
    Else
     result = True
    End If
    
    
  Case 3
    
    ' Взвесили поддон и ввели его номер
    '
     result = True
     
         
     If txt3FullWeight = "" Or Not IsNumeric(txt3FullWeight) Then
      MsgBox "Дождитесь получения веса груза с весов"
      result = False
     End If
     
     
     If txt3GoodWeight = "" Then
      MsgBox "Вес нетто не задан"
      result = False
     End If
     
     If IsNumeric(txt3GoodWeight) And Val(txt3GoodWeight) <= 0 Then
      MsgBox "Вес нетто не может быть нулевым"
      result = False
     End If
     
     ' проверить состояние поддона в базе
     
     result = CheckPoddon
     
     
  Case 4
    ' взвешиваем отругружаемый  товар
    
    
    
      result = True
      If txt4FullWeight = "" Or Not IsNumeric(txt4FullWeight) Then
        MsgBox "Дождитесь полчения веса груза с весов"
        result = False
      End If
      
      If MyRound(txt4GoodWeight) >= MyRound(txt3GoodWeight) Then
        MsgBox "Отгружаем больше, чем было на поддоне, откорректируйте вес"
        result = False
      End If

      If MyRound(txt4Quantity) >= MyRound(txt3Quantity) Then
        MsgBox "Отгружаем больше коробов, чем было на поддоне, откорректируйте количество"
        result = False
      End If
      
      If txt4NewPlace = "" Then
        MsgBox "Задайте ячейку для остатков"
        result = False
      End If
      
  Case 5
    result = True
    If result Then
      If MsgBox("Зарегистрировать отгрузку палеты?", vbExclamation + vbYesNo, "Внимание") = vbYes Then
        result = True
      Else
        result = False
        log.message "Отказ от отгрузки паддон:" & poddon.code
      End If
    End If
    
    
    
  Case 6
   result = True
  
  Case 7
   
   ' сохраняем заказ
   If MsgBox("Закрыть заказ ?", vbExclamation + vbYesNo) = vbYes Then
    CloseZakaz
    
' состояния для типа:ITTOUT Отгрузка
' "{70853C28-84B5-434E-8413-52DF8FBBB49B}" 'Идет отгрузка
' "{2CDDB562-63D7-483E-B95E-B579A9096CCC}" 'Обработка завершена
' "{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}" 'Отгрузка завершена
' "{CDCAFF7F-B013-40AF-BE61-1A27E35DB946}" 'Оформляется
    
    Item.StatusID = "{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}" 'Отгрузка завершена
    
   End If
    result = True
    
    
  
  
  Case 8
  result = True
  
  Case 9
  result = True
  
  End Select
  CheckAfter = result
End Function

'полчить значения веса с весов
'Parameters:
' параметров нет
'Returns:
'  значение типа Double
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

'получить вес или сэмулировать
Private Function GetWeight() As Double
  If emu Then
    GetWeight = Rnd(Second(Now)) * 1000 + MyRound("0" & txt3PWeight)
  Else
    GetWeight = GetWeight4
  End If
End Function

'звуковой сигнал
Private Sub MyBeep(ByVal BeepType As String)
      If Not wave Is Nothing Then
        On Error Resume Next
        wave.OpenFile App.Path & "\" & BeepType & ".wav"
        wave.Play
      End If
End Sub

'настройка поля клиент
Private Sub cmdTheClient_Click()
  On Error Resume Next
  
    
  
  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtTheClient.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
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
       Call MsgBox("Ошибка исполнения: " & errStr, vbOKOnly + vbCritical)
     End If
    End If
  End If
End Sub

'изменен вес упаковки шаг4
Private Sub Txt4PackageWeight_Change()
If isCalibrated Then Exit Sub
txt4GoodWeight = MyRound(txt4FullWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
End Sub

'Изменео количество шаг 4
Private Sub txt4Quantity_Change()
  If isCalibrated Then
    txt4GoodWeight = MyRound(txt3GoodWeight) / MyRound(txt3Quantity) * MyRound(txt4Quantity)
  Else
    txt4GoodWeight = MyRound(txt4FullWeight) - (MyRound(txt4PackageWeight) * MyRound(txt4Quantity))
  End If
End Sub

'изменен заказ
Private Sub txtShipOrder_Change()
  
If (txtShipOrder.Text = "") Then
  ' Убрать Brief и ID
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


'закрытие заказа
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
  
  log.message "Закрытие заказа на отгрузку" & txtShipOrder
  
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

'отображение текущих результатов
Private Sub gr2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
  On Error Resume Next
  Item.ITTOUT_LINES.LoadRow gr2, RowIndex, Bookmark, Values
End Sub
