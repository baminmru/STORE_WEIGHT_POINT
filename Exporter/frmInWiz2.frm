VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{1103DFFF-B8B4-437D-8D8D-4EA4A31D3424}#2.4#0"; "ITTINGUI.ocx"
Begin VB.Form frmInWiz2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Приемка"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14100
   Icon            =   "frmInWiz2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   14100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Шаг 3 - Выбор строки заказа"
      Height          =   6735
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   8655
      Begin VB.CommandButton cmdToClose 
         Caption         =   "На итоговую страницу"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   2415
      End
      Begin GridEX20.GridEX gr 
         Height          =   5835
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   8280
         _ExtentX        =   14605
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
         Column(1)       =   "frmInWiz2.frx":030A
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmInWiz2.frx":036E
         FormatStyle(2)  =   "frmInWiz2.frx":044E
         FormatStyle(3)  =   "frmInWiz2.frx":05AA
         FormatStyle(4)  =   "frmInWiz2.frx":065A
         FormatStyle(5)  =   "frmInWiz2.frx":070E
         FormatStyle(6)  =   "frmInWiz2.frx":07E6
         FormatStyle(7)  =   "frmInWiz2.frx":089E
         ImageCount      =   0
         PrinterProperties=   "frmInWiz2.frx":08BE
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Шаг4 - Параметры груза по умолчанию"
      Height          =   7695
      Left            =   3600
      TabIndex        =   8
      Top             =   840
      Width           =   12855
      Begin ITTINGUI.ITTIN_QLINE ITTIN_QLINE1 
         Height          =   6975
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   12303
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Шаг 6 - Вес поддона с грузом"
      Height          =   6735
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Width           =   10455
      Begin VB.TextBox txt4CaliberBrutto 
         Height          =   375
         Left            =   2760
         TabIndex        =   70
         Top             =   4680
         Width           =   2895
      End
      Begin VB.CommandButton cmd6FindCell 
         Caption         =   "..."
         Height          =   375
         Left            =   5160
         TabIndex        =   66
         ToolTipText     =   "Поиск ячейки"
         Top             =   5520
         Width           =   495
      End
      Begin VB.TextBox txt6Netto 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   6240
         Width           =   2895
      End
      Begin VB.TextBox txt6PackageWeight 
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   6240
         Width           =   2535
      End
      Begin VB.TextBox txt4Good 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   5535
      End
      Begin VB.CheckBox chk4Caliber 
         Caption         =   "Калиброванный"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txt4CaliberWeight 
         Height          =   375
         Left            =   2760
         TabIndex        =   20
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox txt4PNum 
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txt4PWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txt4CaliberQuantity 
         Height          =   405
         Left            =   120
         TabIndex        =   17
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox txt4FullWeight 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txt4GoodWeight 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txt4InQry 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton cmd4ClearW 
         Caption         =   "X"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txt4FromUser 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txt4NewPlace 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label19 
         Caption         =   "Калиброваный вес Брутто"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2760
         TabIndex        =   69
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "Вес товара НЕТТО"
         Height          =   255
         Left            =   2760
         TabIndex        =   64
         Top             =   6000
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Вес одной упаковки КГ."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   6000
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Товар"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Калиброванный вес НЕТТО кг"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Поддон №"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Вес поддона КГ."
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Количество коробов"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Вес груза с поддоном"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Вес груза БРУТТО КГ."
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Осталось принять, планово КГ."
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "По заказу КГ."
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Место в буферной зоне"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   5160
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Шаг 5 - Взвешивание поддона"
      Height          =   5895
      Left            =   3840
      TabIndex        =   33
      Top             =   1560
      Width           =   12615
      Begin VB.TextBox txt3Poddon 
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   2760
         Width           =   5055
      End
      Begin VB.TextBox txt3Weight 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5055
      End
      Begin VB.CommandButton cmdPPrint 
         Caption         =   "Печать стикера на поддон"
         Height          =   615
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.CommandButton cmdPNew 
         Caption         =   "Новый"
         Height          =   375
         Left            =   4320
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt3InQry 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txt3Good 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   840
         Width           =   5535
      End
      Begin VB.CommandButton cmd3ClearNum 
         Caption         =   "x"
         Height          =   375
         Left            =   5280
         TabIndex        =   36
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmd3ClearW 
         Caption         =   "x"
         Height          =   375
         Left            =   5280
         TabIndex        =   35
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox txt3FromUser 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Номер поддона"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Вес поддона КГ."
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Осталось принять, планово КГ."
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Товар"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Количество в заказе КГ."
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Шаг 1 - выбор заказа"
      Height          =   5535
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   14775
      Begin VB.TextBox txtTheClient 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1440
         Width           =   6615
      End
      Begin VB.TextBox txtQryCode 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Код заказа"
         Top             =   720
         Width           =   6015
      End
      Begin MTZ_PANEL.DropButton cmdQryCode 
         Height          =   300
         Left            =   6240
         TabIndex        =   4
         Tag             =   "refopen.ico"
         ToolTipText     =   "Код заказа"
         Top             =   720
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.Label Label14 
         Caption         =   "Клиент"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblQryCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Код заказа:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3000
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Шаг 9 - Закрытие заказа"
      Height          =   6135
      Left            =   3840
      TabIndex        =   52
      Top             =   1200
      Width           =   15615
      Begin VB.CommandButton cmdPrnRas 
         Caption         =   "Акт расхождений"
         Height          =   495
         Left            =   3960
         TabIndex        =   71
         Top             =   4560
         Width           =   2415
      End
      Begin VB.CommandButton cmdPrnEPL 
         Caption         =   "Акт о весе поддонов"
         Height          =   495
         Left            =   1920
         TabIndex        =   67
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton cmd6PrnKL 
         Caption         =   "Печать КЛП"
         Height          =   495
         Left            =   360
         TabIndex        =   54
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton cmd6PRNSRV 
         Caption         =   "Печать документа на услуги"
         Height          =   495
         Left            =   7920
         TabIndex        =   53
         Top             =   4440
         Width           =   2415
      End
      Begin VSFlex8Ctl.VSFlexGrid srvGrid 
         Height          =   3975
         Left            =   360
         TabIndex        =   55
         Top             =   360
         Width           =   9975
         _cx             =   17595
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
         FormatString    =   $"frmInWiz2.frx":0A96
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
   Begin VB.Frame Frame8 
      Caption         =   "Шаг 8 - Документ на поддон с грузом"
      Height          =   5415
      Left            =   3120
      TabIndex        =   50
      Top             =   1680
      Width           =   12375
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
         TabIndex        =   51
         Top             =   600
         Width           =   6255
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Шаг7 - Уточнение данных по поддону"
      Height          =   5295
      Left            =   2520
      TabIndex        =   48
      Top             =   1440
      Width           =   11655
      Begin ITTINGUI.ITTIN_PALET ITTIN_PALET1 
         Height          =   4695
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8281
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Далее"
      Default         =   -1  'True
      Height          =   615
      Left            =   9720
      TabIndex        =   59
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Следующий"
      Height          =   615
      Left            =   6120
      TabIndex        =   58
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddW 
      Caption         =   "Следующий поддон"
      Height          =   615
      Left            =   7920
      TabIndex        =   57
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отменить"
      Height          =   615
      Left            =   4320
      TabIndex        =   56
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Шаг 2 - Параметры заказа"
      Height          =   5535
      Left            =   4680
      TabIndex        =   1
      Top             =   480
      Width           =   7215
      Begin ITTINGUI.ITTIN_DEF ITTIN_DEF1 
         Height          =   4815
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8493
      End
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
      Height          =   8220
      Left            =   0
      Picture         =   "frmInWiz2.frx":0AFB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmInWiz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StepNo As Integer
Dim XMLQryCode As String
Dim XMLTheClient As String
Public Item As ITTIN.Application
Dim conn As ADODB.Connection
Private curQRow As ITTIN.ITTIN_QLINE
Dim LinePal As ITTIN_PALET
Dim pal As ITTPL_DEF
Private StopWeighting As Boolean
Private wave As MTZMCI.WavePlayer
Private emu As Boolean
Private port As String
Private psetup As String
Private poddon As ITTPL_DEF
Public NoMSG As Boolean
Public InPoddon As String
Public InWeight As String
Public SinglePoddon As Boolean



' состояния для типа:ITTIN Приемка груза
' "{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}" 'Идет приемка
' "{49A919F7-94A6-49DE-9280-1EEAC973647B}" 'Оформляется
' "{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}" 'Приемка заершена
' "{E8BA9909-6680-4B2C-B446-F58EF91DCD17}" 'Приемка обработана


' состояния для типа:ITTPL Палетта
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана


Public Sub Init()

    XMLQryCode = "<SQLData>"
    XMLQryCode = XMLQryCode & "<connectionstring>ref</connectionstring>"
    XMLQryCode = XMLQryCode & "<connectionprovider>ref</connectionprovider>"
    XMLQryCode = XMLQryCode & "<query>select A.ID [КОД] , convert(varchar(30),A.NUMBER) +'  от ' + convert(varchar(30),A.ORD_DATE,111)  [Название], partner.Name [Клиент]  from RECEIVING_ORDER A left join PARTNER  on A.PARTNER_ID=partner.ID where (a.STATUS = 1 or a.status =0)  </query>"  '
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
    
    Set conn = Manager.GetCustomObjects("refref")
    If GetSetting("RBH", "ITTSETTINGS", "SOUND", "False") <> "False" Then
      Set wave = New MTZMCI.WavePlayer
      wave.OpenDevice
    End If
End Sub

Private Sub SetBtnPos(cmd As CommandButton, ByVal pos As Integer)
On Error Resume Next
  cmd.Left = imgState.Width + (Me.ScaleWidth - imgState.Width) / 4 * (pos - 1)
End Sub

Private Sub chk4Caliber_Click()
On Error Resume Next
  If chk4Caliber.Value = vbChecked Then
    'txt4CaliberQuantity.Enabled = True
    txt4CaliberWeight.Enabled = True
    txt4CaliberBrutto.Enabled = True
    Label5.Enabled = True
    Label19.Enabled = True
    txt6PackageWeight.Locked = True
  Else
    txt4CaliberWeight = "0"
    txt4CaliberWeight.Enabled = False
    txt4CaliberBrutto = "0"
    txt4CaliberBrutto.Enabled = False
    Label5.Enabled = False
    Label19.Enabled = False
    txt6Netto = MyRound(txt4GoodWeight) - MyRound(txt4CaliberQuantity) * MyRound(txt6PackageWeight)
    txt6PackageWeight.Locked = False
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
 
End Sub

Private Sub cmd6FindCell_Click()
  Dim f As frmGetCell
  Set f = New frmGetCell
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

Private Sub cmd6PrnKL_Click()
On Error Resume Next
    Set repShowKL = Nothing
    Set repShowKL = New ReportShow
    repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
    repShowKL.ReportFilter = " instanceid='" & Item.id & "'"
    repShowKL.ReportPath = App.Path & "\in_KL.rpt"
    repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowKL.Run True
    Set repShowKL = Nothing
End Sub

Private Sub cmd6PRNSRV_Click()
On Error Resume Next
    Set repShowSRVIN = Nothing
    Set repShowSRVIN = New ReportShow
    repShowSRVIN.ReportSource = "V_viewITTIN_ITTIN_SRV"
    repShowSRVIN.ReportFilter = " instanceid='" & Item.id & "'"
    repShowSRVIN.ReportPath = App.Path & "\in_srvq.rpt"
    repShowSRVIN.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowSRVIN.Run True
    Set repShowSRVIN = Nothing
End Sub

Private Sub cmdAddW_Click()
On Error Resume Next
    If CheckAfter Then
      StepNo = 5
      ProcessStatus
    End If
End Sub

Private Sub cmdBack_Click()
On Error Resume Next
  If CheckAfter Then
      StepNo = 3
      ProcessStatus
  End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
  StepNo = 10
  ProcessStatus
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
  gr_DblClick
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
  If CheckAfter Then
    StepNo = StepNo + 1
    ProcessStatus
  End If
End Sub


Private Sub cmdPrnEPL_Click()
On Error Resume Next
    Set repShowINEPL = Nothing
    Set repShowINEPL = New ReportShow
    repShowINEPL.ReportSource = "V_viewITTIN_ITTIN_EPL"
    repShowINEPL.ReportFilter = " instanceid='" & Item.id & "'"
    repShowINEPL.ReportPath = App.Path & "\in_epl.rpt"
    repShowINEPL.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowINEPL.Run True
    Set repShowINEPL = Nothing
End Sub

Private Sub cmdPrnRas_Click()
    On Error Resume Next
    Set repShowKL = Nothing
    Set repShowKL = New ReportShow
    repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
    repShowKL.ReportFilter = " instanceid='" & Item.id & "'"
    repShowKL.ReportPath = App.Path & "\in_ras.rpt"
    repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
    repShowKL.Run True
    Set repShowKL = Nothing
End Sub

Private Sub cmdQryCode_Click()
  On Error Resume Next

  Dim pars As New NamedValues
  Dim res As NamedValues
  If (txtQryCode.Tag = "") Then
    ' call MsgBox("Нет данных для запроса")
  Else
    txtQryCode.Tag = Replace(txtQryCode.Tag, "%ID%", " 1=1 ")
    Call pars.Add("xml", txtQryCode.Tag)
  End If
  
  If Manager.GetCustomObjects("cliFilter").Name <> "" Then
    Call pars.Add("filter", " and " & (Manager.GetCustomObjects("cliFilter").Name))
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




Private Sub MakeItem()
On Error Resume Next
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
  
  Dim XMLQRY_NUM As String
  Dim XMLLineAtQuery As String
  Dim XMLgood_ID As String
  
  If conn.State <> ADODB.adStateOpen Then
     conn.Open
   End If
    
  'Если нет заказа, то сформировать новый
  If id = "" Then
    id = CreateGUID2
    Manager.NewInstance id, "ITTIN", txtQryCode
    Set Item = Manager.GetInstanceObject(id)
    
   
    
    Set rs = conn.Execute("select * from receiving_order where id=" & Manager.GetIDFromXMLField(txtQryCode.Tag))
    If rs.EOF Then Exit Sub
    
    
    With Item.ITTIN_DEF.Add
      .ProcessDate = Date
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
    
    
    
    Set rs = conn.Execute("select A.*, B.DESCRIPTION  BRIEF, B.code ARTICUL from receiving_line A join item B on A.item_id =B.id where (a.PARENT_ID  is null or a.parent_id=0) and a.order_id='" & qID & "'")
    While Not rs.EOF
      Set curQRow = Item.ITTIN_QLINE.Add
      With curQRow
        
        .edizm = "" & rs!UOM
        .articul = "" & rs!articul
        
        '.made_country = "" & rs!prod_country
        '.KILL_NUMBER = "" & rs!KILL_NUMBER
        
        If Not IsNull(rs!made_date) Then .made_date = rs!made_date
        If Not IsNull(rs!exp_date) Then .exp_date = rs!exp_date
        
        
        
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
      
      Call GetNumValue(curQRow, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "IN%P", "")
      
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
    ' проверяем не появилось ли новых строк к заказу
    
    Set rs = conn.Execute("select A.*, B.DESCRIPTION  BRIEF, B.code ARTICUL from receiving_line A join item B on A.item_id =B.id where (a.PARENT_ID  is null or a.parent_id=0) and a.order_id='" & qID & "'")
    
    While Not rs.EOF
      For i = 1 To Item.ITTIN_QLINE.Count
        Set curQRow = Item.ITTIN_QLINE.Item(i)
        If Manager.GetIDFromXMLField(curQRow.good_id) = rs!item_id Then
         GoTo next_in_core
        End If
      Next
    
    
      Set curQRow = Item.ITTIN_QLINE.Add
      With curQRow
        
        .edizm = "" & rs!UOM
        .articul = "" & rs!articul
        
        If Not IsNull(rs!made_date) Then .made_date = rs!made_date
        If Not IsNull(rs!exp_date) Then .exp_date = rs!exp_date
        
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
      
      Call GetNumValue(curQRow, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "IN%P", "")
next_in_core:
      rs.MoveNext
    Wend
    
    
    
  End If
End Sub

Private Sub LoadHeader(Item As Object)
On Error Resume Next
'  txtSupplier = Item.Supplier
'  txtTTN = Item.TTN
'  dtpTTNDate = Date
'  If Item.TTNDate <> 0 Then
'   dtpTTNDate = Item.TTNDate
'  Else
'   dtpTTNDate.Value = Null
'  End If
'  txtTranspNumber = Item.TranspNumber
'  txtContainer = Item.Container
'  txtStampNumber = Item.StampNumber
'  txtStampStatus = Item.StampStatus
'  dtpTrack_time_in = Now
'  If Item.Track_time_in <> 0 Then
'   dtpTrack_time_in = Item.Track_time_in
'  Else
'   dtpTrack_time_in.Value = Null
'  End If
'  dtptrack_time_out = Now
'  If Item.track_time_out <> 0 Then
'   dtptrack_time_out = Item.track_time_out
'  Else
'   dtptrack_time_out.Value = Null
'  End If
'  txttemp_in_track = Item.temp_in_track

End Sub

Private Sub cmdToClose_Click()
  StepNo = 9
  ProcessStatus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
  If UnloadMode <> 1 Then
    Cancel = -1
  Else
    wave.StopPlaying
    Set wave = Nothing
    Timer1.Enabled = False
    If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
    End If
  End If
     
End Sub

Private Sub gr_DblClick()
On Error Resume Next
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
On Error Resume Next
  If col = 0 Then Exit Sub
  Item.ITTIN_SRV.Item(Row).Quantity = MyRound(srvGrid.TextMatrix(Row, col))
  Item.ITTIN_SRV.Item(Row).save
End Sub

Private Sub srvGrid_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
On Error Resume Next
If col = 0 Then Cancel = True
End Sub

Private Sub Timer1_Timer()
  Dim w As Double
  On Error Resume Next
  If StepNo = 5 Then
    
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
  If StepNo = 6 Then
    If txt4FullWeight = "0" Or Not IsNumeric(txt4FullWeight) Then
      w = GetWeight
      If w > 0 And w > MyRound(txt3Weight) + 5 Then
        txt4FullWeight = Round(w + 0.001, 1)
        MyBeep "Gruz"
      End If
  End If
  End If
End Sub

Private Sub txt3Poddon_Change()
On Error Resume Next
  CheckPoddon
End Sub

Private Function CheckPoddon() As Boolean
On Error Resume Next
  If txt3Poddon <> "" Then
    If Len(txt3Poddon) = 6 Then
      Set poddon = Nothing
      Set poddon = FindPoddon(txt3Poddon)
      If Not poddon Is Nothing Then
        MyBeep "Nomer"
        txt3Weight = poddon.Weight
      Else
        MsgBox "Номер паддона: " & txt3Poddon & "  не зарегистрирован"
      End If
    End If
  End If
End Function

Private Sub txt4CaliberBrutto_Change()
  If chk4Caliber.Value = vbChecked Then
    If MyRound(txt4CaliberWeight) > 0 And MyRound(txt4CaliberBrutto) > 0 Then
      If MyRound(txt4CaliberWeight) < MyRound(txt4CaliberBrutto) Then
            txt6PackageWeight = MyRound(txt4CaliberBrutto) - MyRound(txt4CaliberWeight)
            'txt4CaliberQuantity = Round(MyRound(txt4GoodWeight) / MyRound(txt4CaliberBrutto) + 0.1)
      End If
    End If
  End If

End Sub

Private Sub txt4CaliberQuantity_Change()
  If chk4Caliber.Value = vbUnchecked Then
    txt6Netto = MyRound("0" & txt4GoodWeight) - MyRound("0" & txt6PackageWeight) * MyRound("0" & txt4CaliberQuantity)
  Else
    txt6Netto = MyRound("0" & txt4CaliberWeight) * MyRound("0" & txt4CaliberQuantity)
  End If
  
End Sub

Private Sub txt4CaliberWeight_Change()
'  On Error Resume Next
'  Static InCW As Boolean
'  If InCW Then Exit Sub
'  InCW = True
'
'  If chk4Caliber.Value = vbChecked Then
'    If MyRound(txt4CaliberWeight) > 0 And MyRound(txt4CaliberBrutto) > 0 Then
'      If MyRound(txt4CaliberWeight) < MyRound(txt4CaliberBrutto) Then
'            txt6PackageWeight = MyRound(txt4CaliberBrutto) - MyRound(txt4CaliberWeight)
'            txt4CaliberQuantity = txt4GoodWeight \ (MyRound(txt4CaliberWeight) + MyRound(txt6PackageWeight))
'      End If
'    End If
'  End If
'  InCW = False

 If chk4Caliber.Value = vbChecked Then
    If MyRound(txt4CaliberWeight) > 0 And MyRound(txt4CaliberBrutto) > 0 Then
      If MyRound(txt4CaliberWeight) < MyRound(txt4CaliberBrutto) Then
            txt6PackageWeight = MyRound(txt4CaliberBrutto) - MyRound(txt4CaliberWeight)
            'txt4CaliberQuantity = Round(MyRound(txt4GoodWeight) / MyRound(txt4CaliberBrutto) + 0.1)
      End If
    End If
  End If
End Sub

Private Sub txt4FullWeight_Change()
  On Error Resume Next
  txt4GoodWeight = Round(MyRound(txt4FullWeight) - MyRound(txt4PWeight) + 0.001, 1)
 
   
End Sub



Private Sub txt4GoodWeight_Change()
  txt4CaliberWeight_Change
  
  If chk4Caliber.Value = vbUnchecked Then
    txt6Netto = MyRound("0" & txt4GoodWeight) - MyRound("0" & txt6PackageWeight) * MyRound("0" & txt4CaliberQuantity)
  Else
    txt6Netto = MyRound("0" & txt4CaliberWeight) * MyRound("0" & txt4CaliberQuantity)
  End If
End Sub

Private Sub txt6PackageWeight_Change()
  If chk4Caliber.Value = vbChecked Then
    txt4CaliberQuantity_Change
  Else
    txt6Netto = MyRound("0" & txt4GoodWeight) - MyRound("0" & txt6PackageWeight) * MyRound("0" & txt4CaliberQuantity)
  End If

End Sub

Private Sub txtQryCode_Change()
On Error Resume Next
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
    txtQryCode.Tag = XMLDoc.XML
  End If
End If

cmdTheClient_Click

End Sub
  

Private Sub Form_Load()
On Error Resume Next
    StepNo = 0
    ProcessStatus
    Init
    
End Sub

Private Sub AdjFrame(f As Frame)
On Error Resume Next
  f.Top = 0
  f.Left = imgState.Width + 5 * Screen.TwipsPerPixelX
  f.Width = Me.ScaleWidth - imgState.Width - 10 * Screen.TwipsPerPixelX
  f.Height = Me.ScaleHeight - cmdNext.Height - 5 * Screen.TwipsPerPixelY
  
  If ITTIN_DEF1.Visible Then
    With ITTIN_DEF1
    .Left = 5 * Screen.TwipsPerPixelX
    .Top = 25 * Screen.TwipsPerPixelY
    .Width = f.Width - 10 * Screen.TwipsPerPixelX
    .Height = f.Height - 30 * Screen.TwipsPerPixelY
    End With
  End If
  
  If ITTIN_QLINE1.Visible Then
    With ITTIN_QLINE1
    .Left = 5 * Screen.TwipsPerPixelX
    .Top = 25 * Screen.TwipsPerPixelY
    .Width = f.Width - 10 * Screen.TwipsPerPixelX
    .Height = f.Height - 30 * Screen.TwipsPerPixelY
    End With
  End If
  
  If ITTIN_PALET1.Visible Then
    With ITTIN_PALET1
    .Left = 5 * Screen.TwipsPerPixelX
    .Top = 25 * Screen.TwipsPerPixelY
    .Width = f.Width - 10 * Screen.TwipsPerPixelX
    .Height = f.Height - 30 * Screen.TwipsPerPixelY
    End With
  End If
End Sub


Public Sub Before1()
On Error Resume Next
    txtQryCode.Text = ""
    txtQryCode.Tag = XMLQryCode
    LoadBtnPictures cmdQryCode, cmdQryCode.Tag
    cmdQryCode.RemoveAllMenu
    txtTheClient.Text = ""
    txtTheClient.Tag = XMLTheClient
End Sub


Public Sub Before2()
  On Error Resume Next
 
  Item.StatusID = "{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}"
 
  Item.ITTIN_DEF.Refresh
  
  Set ITTIN_DEF1.Item = Item.ITTIN_DEF.Item(1)
  ITTIN_DEF1.InitPanel
  
  Me.Caption = "Прием: " & txtQryCode
End Sub

Private Function After2() As Boolean
On Error Resume Next
  If ITTIN_DEF1.IsOK Then
    ITTIN_DEF1.save
    After2 = True
    Item.ITTIN_DEF.Item(1).save
    Else
  MsgBox "Не все обязательные поля заплолнены"
   
  End If
End Function
Private Function After4() As Boolean
On Error Resume Next
  If ITTIN_QLINE1.IsOK Then
    Dim rule As ITTD.ITTD_RULE
    Dim OK As Boolean
    OK = True
    Set rule = Item.ITTIN_DEF.Item(1).ThePartyRule
    ITTIN_QLINE1.save
    If rule.TheCountry = Boolean_Da Then
     If curQRow.made_country Is Nothing Then
      OK = False
      MsgBox "Не заплолнено поле: Страна производитель"
      Exit Function
     End If
    End If
    
    If rule.TheFactory = Boolean_Da Then
     If curQRow.factory Is Nothing Then
      OK = False
      MsgBox "Не заплолнено поле: Завод"
      Exit Function
     End If
    End If
    
    If rule.killplace = Boolean_Da Then
     If curQRow.KILL_NUMBER Is Nothing Then
      OK = False
      MsgBox "Не заплолнено поле: № Бойни"
      Exit Function
     End If
    End If
    
    
    
    
    'ITTIN_QLINE1.save
    After4 = True
    curQRow.save
  Else
  MsgBox "Не все обязательные поля заплолнены"
  End If
End Function

Private Function After6() As Boolean
On Error Resume Next
  After6 = True
  If txt4CaliberQuantity = "" Then
    MsgBox "Задайте количество коробов"
     After6 = False
     Exit Function
  End If
  
  If txt4NewPlace = "" Then
    MsgBox "Задайте буферную ячейку"
     After6 = False
     Exit Function
  End If

  If chk4Caliber.Value = vbChecked Then
     If MyRound(txt4CaliberWeight) > 0 And MyRound(txt4CaliberBrutto) > 0 Then
        If MyRound(txt4CaliberWeight) < MyRound(txt4CaliberBrutto) Then
            After6 = True
        Else
            MsgBox "Задайте параметры короба"
            After6 = False
        End If
      Else
            MsgBox "Задайте параметры короба"
            After6 = False
      End If
  End If
  
  
  ' проверяем  соразмерность коробов
  Dim rs As ADODB.Recordset
  Dim gid As String
  Dim tstqry As String
  gid = GetBRIEFFromXMLField(curQRow.good_id)
  
  tstqry = ""
  tstqry = tstqry & "select avg((ITTIN_PALET_GoodWithPaletWeight - poddonweight) / ITTIN_PALET_CaliberQuantity) test"
  tstqry = tstqry & " From dbo.V_viewITTIN_ITTIN_PALET"
  tstqry = tstqry & " Where ITTIN_PALET_IsBrak_VAL = 0"
  tstqry = tstqry & " and ITTIN_QLINE_good_id =" & gid
  
  If Not curQRow.made_country Is Nothing Then
    tstqry = tstqry & " and ITTIN_PALET_made_country_ID ='" & curQRow.made_country.id & "'"
  End If
  If Not curQRow.factory Is Nothing Then
    tstqry = tstqry & " and ITTIN_PALET_Factory_ID ='" & curQRow.factory.id & "'"
  End If
  
  If Not curQRow.KILL_NUMBER Is Nothing Then
    tstqry = tstqry & " and ITTIN_PALET_KILL_NUMBER_ID ='" & curQRow.KILL_NUMBER.id & "'"
  End If

  Dim corob As Double
  corob = MyRound(txt6Netto) / txt4CaliberQuantity
  Set rs = Session.GetData(tstqry)
  If Not rs Is Nothing Then
    If Not rs.EOF Then
        If corob > rs!test * 1.1 Or rs!test > corob * 1.1 Then
            If MsgBox("Вес одного короба (" & Round(corob, 2) & ") выходит за рамки средних (" & Round(rs!test, 2) & ") значений." & vbCrLf & "Коробов на подоне: <" & txt4CaliberQuantity & ">" & vbCrLf & "Вы уверены что задано верное количество коробов ? ", vbYesNo + vbExclamation, "Проверьте правильность ввода") = vbNo Then
              After6 = False
            End If
        End If
    End If
  End If
  
  
  
  

End Function

Private Function After7() As Boolean
On Error Resume Next
  After7 = False
  If ITTIN_PALET1.IsOK Then
    ITTIN_PALET1.save
    After7 = True
    On Error GoTo bye
    LinePal.save
  Else
    MsgBox "Не все обязательные поля заплолнены"
  End If
  Exit Function
bye:
  MsgBox err.Description
End Function

Private Function After9() As Boolean
  After9 = True
   If MsgBox("Закрыть заказ ", vbExclamation + vbYesNo) = vbYes Then
    Item.StatusID = "{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}"
    CloseZakaz
  End If
End Function

Private Sub CloseZakaz()
On Error Resume Next
  Dim conn As ADODB.Connection
  Dim cmd As ADODB.Command
  Dim rs As ADODB.Recordset
  Dim rlID As String
  Dim oid As String
  
  
  Set conn = Manager.GetCustomObjects("refref")
  If conn.State <> adStateOpen Then
    conn.Open
  End If
  oid = Manager.GetIDFromXMLField(Item.ITTIN_DEF.Item(1).QryCode)
  
  Dim i As Long
  Set cmd = New ADODB.Command
  cmd.CommandText = "update RECEIVING_order set status=2 where id=" & oid
  Set cmd.ActiveConnection = conn
  err.Clear
  cmd.Execute

  If err.Number <> 0 Then
    MsgBox err.Description
  End If
  
  
  For i = 1 To Item.ITTIN_QLINE.Count
    Set curQRow = Item.ITTIN_QLINE.Item(i)
    
    rlID = Manager.GetIDFromXMLField(curQRow.good_id)

    cmd.CommandText = "update RECEIVING_LINE SET status=2 where order_id = " & oid & " and item_ID=" & rlID
    err.Clear
    Set cmd.ActiveConnection = conn
    cmd.Execute
    If err.Number <> 0 Then
      MsgBox err.Description
    End If
  Next
'  Item.ITTIN_DEF.Item(1).track_time_out = Now
'  Item.ITTIN_DEF.Item(1).save
End Sub



Private Sub Before3()
On Error Resume Next
 'SaveHeader Item.ITTIN_DEF.Item(1)s
     If NoMSG = False Then
         If MsgBox("Напечатать пустографку?", vbYesNo) = vbYes Then
         
          Set repShowSRVIN = Nothing
          Set repShowSRVIN = New ReportShow
          repShowSRVIN.ReportSource = "V_viewITTIN_ITTIN_SRV"
          repShowSRVIN.ReportFilter = " instanceid='" & Item.id & "'"
          repShowSRVIN.ReportPath = App.Path & "\in_srv.rpt"
          repShowSRVIN.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
          repShowSRVIN.Run True
          Set repShowSRVIN = Nothing
          NoMSG = True
        End If
    End If

    'Инициализироать таблицу строк заказа
    gr.ItemCount = 0
    'Item.ITTIN_QLINE.Sort = "sequence"
    Item.ITTIN_QLINE.Refresh
    Item.ITTIN_QLINE.PrepareGrid gr
    
    gr.ItemCount = Item.ITTIN_QLINE.Count
End Sub

Private Sub Before5()
  On Error Resume Next
  txt3Poddon = InPoddon
  InPoddon = ""
  txt3Weight = InWeight
  InWeight = 0
  
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
          plan = MyRound("0" & nodeQRY_NUM.Text)
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
  Me.Caption = "Прием: " & txtQryCode & "\" & txt3Good & " (" & txt3InQry & ")"
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
On Error Resume Next

  Set ITTIN_QLINE1.Item = curQRow
  ITTIN_QLINE1.InitPanel

End Sub



Private Sub Before6()
On Error Resume Next
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
          plan = MyRound("0" & nodeQRY_NUM.Text)
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
  txt6PackageWeight = curQRow.PackageWeight
  
  txt4CaliberBrutto = curQRow.KorobBrutto
  txt4CaliberWeight = curQRow.KorobNetto

  If curQRow.isCalibrated = Boolean_Da Then
    chk4Caliber.Value = vbChecked
  Else
    chk4Caliber.Value = vbUnchecked
  End If
  
  
  
  ' состояния для типа:ITTPL Палетта
' "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" 'Взвешена
' "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}" 'На складе с грузом
' "{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}" 'Отправлена с грузом
' "{E9BFB749-A606-4DEF-A429-07D636F108C6}" 'Пустая
' "{588C5203-1E59-408E-92A1-B3DFED8C19FA}" 'Списана
  
  Dim P As ITTPL_DEF
  
  Set P = FindPoddon(txt3Poddon)
  If Not P Is Nothing Then
    P.Weight = MyRound(txt3Weight)
    P.WDate = Date
    P.save
    P.Application.StatusID = "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}"
  End If
  
  
  Dim i As Long
  Dim eplOK As Boolean
  eplOK = False
  For i = 1 To Item.ITTIN_EPL.Count
    If Item.ITTIN_EPL.Item(i).TheNumber.id = P.id Then
      If Item.ITTIN_EPL.Item(i).PalWeight > 0 Then
        eplOK = True
      Else
        Item.ITTIN_EPL.Item(i).PalWeight = P.Weight
        Item.ITTIN_EPL.Item(i).save
          eplOK = True
      End If
      
    End If
  Next
  If Not eplOK Then
    With Item.ITTIN_EPL.Add
      Set .TheNumber = P
      .PalWeight = P.Weight
      .save
    End With
  End If
  
  
  If curQRow.isCalibrated = Boolean_Da Then
    chk4Caliber.Value = vbChecked
    txt4CaliberWeight.Enabled = True
  Else
    chk4Caliber.Value = vbUnchecked
    txt4CaliberWeight.Enabled = False
  End If
End Sub

Private Sub before7()
On Error Resume Next
  curQRow.ITTIN_PALET.Refresh
  Set LinePal = curQRow.ITTIN_PALET.Add
  With LinePal
    Set poddon = FindPoddon(txt3Poddon)
    Set .TheNumber = poddon
    poddon.CurrentPosition = txt4NewPlace
    .PalWeight = MyRound(txt3Weight)
    .GoodWithPaletWeight = MyRound(txt4FullWeight)
    .CaliberQuantity = MyRound(txt4CaliberQuantity)
    .FullPackageWeight = MyRound("0" & txt6PackageWeight) * MyRound("0" & txt4CaliberQuantity)
    .made_date_to = curQRow.made_date_to
    .vetsved = curQRow.vetsved
    .BufferZonePlace = txt4NewPlace
    .PackageWeight = MyRound("0" & txt6PackageWeight)
    
    Set .made_country = curQRow.made_country
    Set .factory = curQRow.factory
    Set .KILL_NUMBER = curQRow.KILL_NUMBER
    Set .PartRef = curQRow.PartRef
    .made_date = curQRow.made_date
    .exp_date = curQRow.exp_date
    .VidOtruba = curQRow.VidOtruba
    .KorobBrutto = MyRound(txt4CaliberBrutto)
    .KorobNetto = MyRound(txt4CaliberWeight)
    
    
    If chk4Caliber.Value = vbChecked Then
      .isCalibrated = Boolean_Da
      
      If curQRow.KorobBrutto <= 0 Then
        curQRow.KorobBrutto = .KorobBrutto
      End If
      
      If curQRow.KorobNetto <= 0 Then
        curQRow.KorobNetto = .KorobNetto
      End If
      
      curQRow.save
      
    Else
      .isCalibrated = Boolean_Net
    End If
    
    On Error GoTo bye
    
    .save
    GoTo save_ok
    
bye:
  MsgBox err.Description
  curQRow.ITTIN_PALET.Refresh
  
save_ok:

  End With
  
  Call GetNumValue(LinePal, "sequence", "{E7F3EE01-4EC4-41D2-8657-BA22089DE0E5}", Now, "INPAL%P", "")
  
  
  
  Set ITTIN_PALET1.Item = LinePal
  ITTIN_PALET1.InitPanel
  
End Sub



Private Sub Before9()
On Error Resume Next
Dim i As Long
  srvGrid.Rows = Item.ITTIN_SRV.Count + 1
  For i = 1 To Item.ITTIN_SRV.Count
    srvGrid.TextMatrix(i, 0) = Item.ITTIN_SRV.Item(i).srv.brief
    srvGrid.TextMatrix(i, 1) = Item.ITTIN_SRV.Item(i).Quantity
  Next
End Sub


Private Sub Before8()
 On Error Resume Next
  
  With curQRow
    If chk4Caliber.Value = vbChecked Then
     .isCalibrated = Boolean_Da
     .CaliberWeight = MyRound(txt4CaliberWeight)
    Else
     .isCalibrated = Boolean_Net
     .CaliberWeight = 0
    End If
    
    .CurValue = .CurValue + MyRound(txt6Netto)
    
    .FullPackageWeight = .FullPackageWeight + MyRound(txt6PackageWeight) * MyRound(txt4CaliberQuantity)
    
    'If .PackageWeight = 0 Then
      .PackageWeight = MyRound(txt6PackageWeight)
    'End If
    
    .save
  End With
  
  
  If MyRound(txt3InQry) < MyRound(txt6Netto) Then
    cmdAddW.Enabled = False
  Else
    cmdAddW.Enabled = True
  End If
  
  On Error Resume Next
  If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
  End If
  
  
  SaveRCVRowToCore Item, curQRow, LinePal, txt4NewPlace.Tag, txtQryCode.Text
  
  Set pal = LinePal.TheNumber
  pal.CurrentPosition = txt4NewPlace
  pal.CurrentWeightBrutto = MyRound(txt4FullWeight)
  pal.CurrentGood = curQRow.good_id
  pal.CaliberQuantity = LinePal.CaliberQuantity
  pal.PackageWeight = MyRound(txt6PackageWeight) * MyRound(txt4CaliberQuantity)
  pal.save
    
  pal.Application.StatusID = "{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}"
  
  
  PrintSticker pal, MyRound(txt4CaliberWeight)
  
'  If GetSetting("RBH", "ITTSETTINGS", "PSTICKER", 1) Then
'
'  If MsgBox("Напечатать стикер на поддон?", vbYesNo) = vbYes Then
'
'    lbl5Out = "Печатается документ на поддон"
'    DoEvents
'
'    Dim X As Printer
'    For Each X In Printers
'    If X.DeviceName = GetSetting("RBH", "ITTSETTINGS", "DOCPRN") Then
'
'    Set Printer = X
'    Printer.Font = "Arial CYR"
'    Printer.FontSize = 32
'
'    Printer.FontBold = False
'    Printer.Print "Поклажедатель: ";
'    Printer.FontBold = True
'    Printer.Print txtTheClient
'
'    Printer.FontBold = False
'    Printer.Print "Заказ: ";
'    Printer.FontBold = True
'    Printer.Print txtQryCode
'
'    Printer.FontBold = False
'    Printer.Print "Поддон №";
'    Printer.FontBold = True
'    Printer.Print txt3Poddon & "  ";
'    Printer.Font = "Code 128"
'
'    Printer.FontBold = False
'    Printer.FontSize = 48
'    Printer.Print code128(txt3Poddon)
'
'    Printer.Font = "Arial CYR"
'    Printer.FontSize = 32
'
'    Printer.Print "Код: ";
'    Printer.FontBold = True
'    Printer.Print curQRow.articul & "";
'
'    Printer.Font = "Code 128"
'
'    Printer.FontBold = False
'    Printer.FontSize = 48
'    Printer.Print code128(curQRow.articul)
'
'    Printer.Font = "Arial CYR"
'    Printer.FontSize = 32
'    Printer.Print "Товар: ";
'    Printer.FontBold = True
'
'    Printer.Print Left(txt4Good & "", 30)
'    If Len(txt4Good & "") > 30 Then
'      Printer.Print Mid(txt4Good & "", 31, 36)
'    End If
'    If Len(txt4Good & "") > 30 + 36 Then
'      Printer.Print Mid(txt4Good & "", 31 + 36)
'    End If
'
'
'
'    If LinePal.IsBrak = Boolean_Da Then
'      Printer.Print "БРАК"
'    End If
'
'
'    Printer.FontBold = False
'
'    If Not curQRow.PartRef Is Nothing Then
'      Printer.Print "Партия: ";
'      Printer.FontBold = True
'      Printer.Print curQRow.PartRef.Name
'
'    End If
'
'    Printer.FontBold = False
'    Printer.Print "Страна производитель: ";
'    Printer.FontBold = True
'    Printer.Print curQRow.made_country.Name
'
'    Printer.FontBold = False
'    Printer.Print "Производитель: ";
'    Printer.FontBold = True
'    Printer.Print curQRow.factory.Name
'
'    Printer.FontBold = False
'    If Not curQRow.KILL_NUMBER Is Nothing Then
'      Printer.Print "Бойня: ";
'      Printer.FontBold = True
'      Printer.Print curQRow.KILL_NUMBER.Name
'    End If
'
'    Printer.FontBold = False
'    Printer.Print "Вес груза НЕТТО (КГ.) : ";
'    Printer.FontBold = True
'    Printer.Print MyRound(txt6Netto)
'
'    Printer.FontBold = False
'    Printer.Print "Вес груза Брутто (КГ.) : ";
'    Printer.FontBold = True
'    Printer.Print MyRound(txt4GoodWeight)
'
'    Printer.FontBold = False
'    Printer.Print "Вес поддона с грузом (КГ.) : ";
'    Printer.FontBold = True
'    Printer.Print MyRound(txt4FullWeight)
'
'    Printer.FontBold = False
'    Printer.Print "Вес упаковки (КГ.) : ";
'    Printer.FontBold = True
'    Printer.Print MyRound(txt4CaliberQuantity) * MyRound(txt6PackageWeight)
'
'
'    Printer.FontBold = False
'    Printer.Print "Дата выпуска: ";
'    Printer.FontBold = True
'    Printer.Print curQRow.Made_date
'
'    Printer.FontBold = False
'    Printer.Print "Cрок годности: ";
'    Printer.FontBold = True
'    Printer.Print curQRow.exp_date
'
'
'    If chk4Caliber.Value = vbChecked Then
'      Printer.FontBold = False
'      Printer.Print "Калиброванный товар"
'      Printer.Print "Вес одного короба (КГ.): ";
'      Printer.FontBold = True
'      Printer.Print Round(txt4CaliberWeight + 0.001, 2)
'    End If
'
'    Printer.FontBold = False
'    Printer.Print "Количество коробов: ";
'    Printer.FontBold = True
'    Printer.Print Round(txt4CaliberQuantity + 0.001, 0)
'
'    If GetSetting("RBH", "ITTSETTINGS", "PCELL", 0) = 1 Then
'      Printer.NewPage
'      lbl5Out = "Печатается документ напервичное размещение"
'      DoEvents
'      Printer.FontSize = 72
'      Printer.Print "Поддон №"
'      Printer.Print txt3Poddon
'      Printer.Print "Буферная яч.№"
'      Printer.Print txt4NewPlace
'    End If
'
'    Printer.EndDoc
'
'
'    lbl5Out = "Документы отправлены на принтер"
'    DoEvents
'
'
'
'   Exit For
'  End If
'  Next
'  End If
'  End If
'bye2:
'
'  Exit Sub
'
'bye:
'  If err.Number > 0 Then
'    MsgBox err.Description, , "Печать документов на поддон"
'  End If
  If SinglePoddon Then
    cmdCancel_Click
  End If
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

Public Sub ProcessStatus()
On Error Resume Next
  Frame1.Visible = False
  Frame2.Visible = False
  Frame3.Visible = False
  Frame4.Visible = False
  Frame5.Visible = False
  Frame6.Visible = False
  Frame7.Visible = False
  Frame8.Visible = False
  Frame9.Visible = False
  
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
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
'    cmdBack.Visible = True
'    cmdAddW.Visible = True
'    cmdBack.Caption = "Другая позиция заказа"
'    cmdNext.Caption = "Закрыть заказ"
'
'    If cmdAddW.Enabled Then
'      SetBtnPos cmdCancel, 1
'      SetBtnPos cmdNext, 2
'      SetBtnPos cmdBack, 3
'      SetBtnPos cmdAddW, 4
'    Else
'      SetBtnPos cmdCancel, 1
'      SetBtnPos cmdBack, 4
'      SetBtnPos cmdAddW, 2
'      SetBtnPos cmdNext, 3
'
'    End If
    
  Case 6
    Before6
    Frame6.Visible = True
    AdjFrame Frame6
    
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 7
    before7
    Frame7.Visible = True
    AdjFrame Frame7
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
  Case 8
    Before8
    cmdBack.Caption = "Другая позиция заказа"
    cmdAddW.Caption = "Следующий поддон"
    cmdBack.Visible = True
    cmdAddW.Visible = True
    Frame8.Visible = True
    AdjFrame Frame8
    
    If cmdAddW.Enabled Then
      SetBtnPos cmdCancel, 1
      SetBtnPos cmdNext, 2
      SetBtnPos cmdBack, 3
      SetBtnPos cmdAddW, 4
    Else
      SetBtnPos cmdAddW, 1
      SetBtnPos cmdCancel, 2
      SetBtnPos cmdBack, 4
      SetBtnPos cmdNext, 3
    End If

    
  Case 9
    Before9
    Frame9.Visible = True
    AdjFrame Frame9
    cmdNext.Caption = "Закрыть заказ"
    SetBtnPos cmdCancel, 3
    SetBtnPos cmdNext, 4
    
  Case 10
   Unload Me
   
  End Select
  
  
  If StepNo >= 0 And StepNo < 10 Then
    imgState.Picture = LoadPicture(App.Path & "\Design\Step" & (StepNo) & ".bmp")
  Else
    imgState.Picture = LoadPicture(App.Path & "\Design\Step0.bmp")
  End If
End Sub


Private Function CheckAfter() As Boolean
On Error Resume Next
  Dim result As Boolean
  
  Select Case StepNo
  Case 0
    ' do nothiing
    result = True
  Case 1
    ' Выбрали строку заказа
    '
  
    If txtQryCode.Text = "" Then
      result = False
      MsgBox "Следует выбрать заказ"
    Else
      result = True
    End If
  Case 2
    ' все поля заполнены
    result = After2
    
  Case 3
    If Not curQRow Is Nothing Then
      result = True
    Else
      MsgBox "Надо выбрать строку"
    End If
    
  Case 4
    result = After4
    
  Case 5
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
     
     Set poddon = FindPoddon(txt3Poddon)
     
     If poddon Is Nothing Then
         MsgBox "Поддон с таким номером не обнаружен в базе данных"
         result = False
     Else
      If poddon.Application.StatusID = "{6FDCC60F-8C10-47E3-BB36-110C49EF2144}" Or _
         poddon.Application.StatusID = "{E9BFB749-A606-4DEF-A429-07D636F108C6}" Then
         
         ' проверяем не взвешен ли он к другому заказу
         Dim checkrs As ADODB.Recordset
         Set checkrs = Session.GetData("select * from v_viewITTIN_ITTIN_EPL" & _
         " where  ITTIN_EPL_TheNumber_ID = '" & poddon.id & "' and " & _
         " instanceid <> '" & Item.id & "' and " & _
         " INTSANCEStatusID in ('{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}','{49A919F7-94A6-49DE-9280-1EEAC973647B}')")
         If checkrs Is Nothing Then
          MsgBox err.Description
         End If
         If Not checkrs.EOF Then
          MsgBox "Поддон с таким номером зарезервирован для заказа <" & checkrs!ITTIN_DEF_QryCode & "> и не можт быть использован"
          result = False
         End If
         
      Else
         MsgBox "Поддон с таким номером находится в состоянии <" & poddon.Application.StatusName & "> и не можт быть использован"
         result = False
      End If
     End If
     
     
     
  Case 6
    ' взвешиваем груз
    result = After6
    If txt4FullWeight = "" Or txt4FullWeight = "0" Or Not IsNumeric(txt4FullWeight) Then
      MsgBox "Дождитесь полчения веса груза с весов"
      result = False
    ElseIf MyRound(txt4FullWeight) > 1000 Then
      MsgBox "Вес поддона превышает 1000 кг."
      txt4FullWeight = 0
      result = False
    End If
    
    If MyRound(txt6Netto) <= 0 Then
      MsgBox "Значения веса НЕТТО должны быть больше 0"
      result = False
    End If
    
    If MyRound(txt6PackageWeight) <= 0 Then
      MsgBox "Значения веса упаковки должны быть больше 0"
      result = False
    End If
    
    If MyRound(txt6PackageWeight) > 10 Then
      MsgBox "Значения веса упаковки должны быть меньше 10 кг."
      result = False
    End If
    
    If chk4Caliber.Value = vbChecked Then
      If MyRound(txt4CaliberWeight) <= 0 Then
        MsgBox "Значения веса калиброванного товарадолжны быть больше 0"
        result = False
      End If
    End If
    
    If MyRound(txt4CaliberQuantity) <= 0 Then
      MsgBox "Количество коробов должны быть больше 0"
      result = False
    End If
   
    
    If result Then
      If MsgBox("Зарегистрировать прием палеты?", vbExclamation + vbYesNo, "Внимание") = vbYes Then
        result = True
      Else
        result = False
      End If
    End If
    
    
  Case 7
     
    result = After7
    
  Case 8
   ' сохраняем заказ
    result = True
  Case 9
  '
  result = After9
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
On Error Resume Next
  If emu Then
    If StepNo = 6 Then
      GetWeight = Rnd(Second(Now)) * 1000 + MyRound("0" & txt3Weight)
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





