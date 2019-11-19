VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Главное окно"
   ClientHeight    =   7170
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "out"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0624
            Key             =   "toQry"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":093E
            Key             =   "in"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C58
            Key             =   "auto"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F72
            Key             =   "div"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":128C
            Key             =   "add"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   2910
      ButtonWidth     =   2858
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Поддоны к заказу"
            Key             =   "toqry"
            Object.ToolTipText     =   "Взвешивание поддона к заказу"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Авто"
            Key             =   "auto"
            Object.ToolTipText     =   "Режим автовыбора заказа"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Приемка"
            Key             =   "in"
            Object.ToolTipText     =   "Приемка груза по заказу"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Отгрузка"
            Key             =   "out"
            Object.ToolTipText     =   "Отгрузка поддонов"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Разукомплектовать"
            Key             =   "div"
            Object.ToolTipText     =   "Перераспределить груз на 2 палеты"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Объединить"
            Key             =   "add"
            Object.ToolTipText     =   "Объединить груз с двух палет на одну"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Стикер"
            Key             =   "prn"
            Object.ToolTipText     =   "Печать стикера на поддон"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Timer MenuTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2355
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   1665
      Top             =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1080
      Top             =   840
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Файл"
      Begin VB.Menu mnuSetings 
         Caption         =   "Настройка"
         Begin VB.Menu mnuCoreSetup 
            Caption         =   "Настройка соединения"
         End
         Begin VB.Menu mnuPRNSetup 
            Caption         =   "Настройка принтеров"
         End
         Begin VB.Menu mnuWeightSetup 
            Caption         =   "Настройка весов"
         End
         Begin VB.Menu mnuNumbers 
            Caption         =   "Настройки системы"
         End
         Begin VB.Menu mnuSetup 
            Caption         =   "Общие настройки"
         End
      End
      Begin VB.Menu mnuServices 
         Caption         =   "Сервис"
         Begin VB.Menu mnuMakePoddon 
            Caption         =   "Регистрация поддонов"
         End
         Begin VB.Menu mnuSyncDict 
            Caption         =   "Синхронизировать справочники"
         End
         Begin VB.Menu mnuUpdateCore 
            Caption         =   "Обновить данные в Core"
         End
         Begin VB.Menu mnuLocationSize 
            Caption         =   "Управление объемом ячейки"
         End
         Begin VB.Menu mnuMakeOtbor 
            Caption         =   "Произвести отбор по выморозке"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu mnuOperations 
      Caption         =   "Операции"
      Begin VB.Menu mnuPWiz 
         Caption         =   "Поддоны к заказу"
      End
      Begin VB.Menu mnuAuto 
         Caption         =   "Авто"
      End
      Begin VB.Menu mnuProcessQuery 
         Caption         =   "Прием по заказу"
      End
      Begin VB.Menu mnuProcessShipping 
         Caption         =   "Отгрузка по закау"
      End
      Begin VB.Menu mnuSplitPoddon 
         Caption         =   "Разукомплектовать"
      End
      Begin VB.Menu mnuAssemble 
         Caption         =   "Объединить"
      End
   End
   Begin VB.Menu mnuJRNL 
      Caption         =   "Журналы"
      Begin VB.Menu mnuITTOPT 
         Caption         =   "Задание на перемещения"
         Begin VB.Menu mnuAllITTOPT 
            Caption         =   "Задание на перемещения - все состояния"
         End
         Begin VB.Menu mnuITTOPT_1 
            Caption         =   "Задание на перемещения :Выполнено"
         End
         Begin VB.Menu mnuITTOPT_6 
            Caption         =   "Задание на перемещения :К выполнению"
         End
         Begin VB.Menu mnuITTOPT_7 
            Caption         =   "Задание на перемещения :Создан"
         End
      End
      Begin VB.Menu mnuITTOUT 
         Caption         =   "Отгрузка"
         Begin VB.Menu mnuAllITTOUT 
            Caption         =   "Отгрузка - все состояния"
         End
         Begin VB.Menu mnuITTOUT_1 
            Caption         =   "Отгрузка :Оформляется"
         End
         Begin VB.Menu mnuITTOUT_2 
            Caption         =   "Отгрузка :Идет отгрузка"
         End
         Begin VB.Menu mnuITTOUT_3 
            Caption         =   "Отгрузка :Обработка завершена"
         End
         Begin VB.Menu mnuITTOUT_4 
            Caption         =   "Отгрузка :Отгрузка завершена"
         End
      End
      Begin VB.Menu mnuITTPL 
         Caption         =   "Палетта"
         Begin VB.Menu mnuAllITTPL 
            Caption         =   "Палетта - все состояния"
         End
         Begin VB.Menu mnuITTPL_1 
            Caption         =   "Палетта :Пустая"
         End
         Begin VB.Menu mnuITTPL_2 
            Caption         =   "Палетта :Взвешена"
         End
         Begin VB.Menu mnuITTPL_3 
            Caption         =   "Палетта :На складе с грузом"
         End
         Begin VB.Menu mnuITTPL_4 
            Caption         =   "Палетта :Списана"
         End
         Begin VB.Menu mnuITTPL_5 
            Caption         =   "Палетта :Отправлена с грузом"
         End
      End
      Begin VB.Menu mnuITTIN 
         Caption         =   "Приемка груза"
         Begin VB.Menu mnuAllITTIN 
            Caption         =   "Приемка груза - все состояния"
         End
         Begin VB.Menu mnuITTIN_1 
            Caption         =   "Приемка груза :Оформляется"
         End
         Begin VB.Menu mnuITTIN_2 
            Caption         =   "Приемка груза :Приемка заершена"
         End
         Begin VB.Menu mnuITTIN_3 
            Caption         =   "Приемка груза :Идет приемка"
         End
         Begin VB.Menu mnuITTIN_4 
            Caption         =   "Приемка груза :Приемка обработана"
         End
      End
      Begin VB.Menu mnuITTCS 
         Caption         =   "Услуги клиентов"
      End
      Begin VB.Menu mnuITTPR 
         Caption         =   "Протоколы расхождений"
      End
   End
   Begin VB.Menu mnuDictionary 
      Caption         =   "Справочники"
      Begin VB.Menu mnuITTFN 
         Caption         =   "Настройки системы"
      End
      Begin VB.Menu mnuITTOP 
         Caption         =   "Операторы и кладовщики"
      End
      Begin VB.Menu mnuITTD 
         Caption         =   "Справочник"
      End
      Begin VB.Menu mnuITTNO 
         Caption         =   "Настройки оптимизатора"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Отчеты"
      Begin VB.Menu mnuMySticker 
         Caption         =   "Стикер на поддон"
      End
      Begin VB.Menu mnuStickers 
         Caption         =   "Номера на поддон"
      End
      Begin VB.Menu mnuRptSrv 
         Caption         =   "Отчет по услугам"
      End
      Begin VB.Menu mnuExportSRV 
         Caption         =   "Выгрузка по услугам"
      End
      Begin VB.Menu mnuWrongLocation 
         Caption         =   "Проблемные ячейки"
      End
      Begin VB.Menu mnuVimorozka 
         Caption         =   "Отчет по выморозке"
      End
      Begin VB.Menu mnuRptHran 
         Caption         =   "Отчет по объему услуг хранения"
      End
      Begin VB.Menu mnuOtbor 
         Caption         =   "Отчет по отбору"
      End
      Begin VB.Menu mnuRptOtobrano 
         Caption         =   "Объемы отбора товара"
      End
      Begin VB.Menu mnuRpt103 
         Caption         =   "Заблокировано на выморозку"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "Окно"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Каскад"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Разложить вертикально"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Разложить горизонтально"
      End
      Begin VB.Menu mnuArrangeIcon 
         Caption         =   "Разложить иконки"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 570
Option Explicit
'Главное окно



Dim inTimer1 As Boolean
Dim inTimer2 As Boolean
Dim OnLoad As Boolean
Dim DelayedCommand As String

Private optEdit As Boolean


Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)


' формы журналов
' Задание на перемещения
Dim WithEvents jfmnuAllITTOPT As frmJournalShow2
Attribute jfmnuAllITTOPT.VB_VarHelpID = -1

Dim WithEvents jfmnuITTOPT_1 As frmJournalShow2
Attribute jfmnuITTOPT_1.VB_VarHelpID = -1

Dim WithEvents jfmnuITTOPT_6 As frmJournalShow2
Attribute jfmnuITTOPT_6.VB_VarHelpID = -1

Dim WithEvents jfmnuITTOPT_7 As frmJournalShow2
Attribute jfmnuITTOPT_7.VB_VarHelpID = -1

' отгрузка
Dim WithEvents jfmnuAllITTOUT As frmJournalShow2
Attribute jfmnuAllITTOUT.VB_VarHelpID = -1

Dim WithEvents jfmnuITTOUT_1 As frmJournalShow2
Attribute jfmnuITTOUT_1.VB_VarHelpID = -1

Dim WithEvents jfmnuITTOUT_2 As frmJournalShow2
Attribute jfmnuITTOUT_2.VB_VarHelpID = -1

Dim WithEvents jfmnuITTOUT_3 As frmJournalShow2
Attribute jfmnuITTOUT_3.VB_VarHelpID = -1

Dim WithEvents jfmnuITTOUT_4 As frmJournalShow2
Attribute jfmnuITTOUT_4.VB_VarHelpID = -1


' акты
Dim WithEvents jfmnuITTPR As frmJournalShow2
Attribute jfmnuITTPR.VB_VarHelpID = -1


' паллеты
Dim WithEvents jfmnuAllITTPL As frmJournalShow2
Attribute jfmnuAllITTPL.VB_VarHelpID = -1

Dim WithEvents jfmnuITTPL_1 As frmJournalShow2
Attribute jfmnuITTPL_1.VB_VarHelpID = -1

Dim WithEvents jfmnuITTPL_2 As frmJournalShow2
Attribute jfmnuITTPL_2.VB_VarHelpID = -1

Dim WithEvents jfmnuITTPL_3 As frmJournalShow2
Attribute jfmnuITTPL_3.VB_VarHelpID = -1

Dim WithEvents jfmnuITTPL_4 As frmJournalShow2
Attribute jfmnuITTPL_4.VB_VarHelpID = -1

Dim WithEvents jfmnuITTPL_5 As frmJournalShow2
Attribute jfmnuITTPL_5.VB_VarHelpID = -1


' приемка
Dim WithEvents jfmnuAllITTIN As frmJournalShow2
Attribute jfmnuAllITTIN.VB_VarHelpID = -1

Dim WithEvents jfmnuITTIN_1 As frmJournalShow2
Attribute jfmnuITTIN_1.VB_VarHelpID = -1

Dim WithEvents jfmnuITTIN_2 As frmJournalShow2
Attribute jfmnuITTIN_2.VB_VarHelpID = -1

Dim WithEvents jfmnuITTIN_3 As frmJournalShow2
Attribute jfmnuITTIN_3.VB_VarHelpID = -1

Dim WithEvents jfmnuITTIN_4 As frmJournalShow2
Attribute jfmnuITTIN_4.VB_VarHelpID = -1


' услуги
Dim WithEvents jfmnuITTCS As frmJournalShow2
Attribute jfmnuITTCS.VB_VarHelpID = -1



'обработка события инициализации журнала приемка
Private Sub jfmnuAllITTIN_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
  'bAdd = False
  'bDel = False
End Sub

' Обработка команды печати для окна журнала приемка
Private Sub jfmnuAllITTIN_OnPrint(usedefaut As Boolean)
 usedefaut = False
  On Error Resume Next
  Dim id As String
  id = jfmnuAllITTIN.jv.RowInstanceID(jfmnuAllITTIN.jv.Row)
  If id = "" Then Exit Sub
  
  Dim f As frmInPrint
  Set f = New frmInPrint
  f.Show vbModal
  If f.OK Then
    If f.optActRas.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_ras.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
  
    If f.optKLP.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_KL.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
    If f.optAct.Value Then
      On Error Resume Next
      Set repShowINEPL = Nothing
      Set repShowINEPL = New ReportShow
      repShowINEPL.ReportSource = "V_viewITTIN_ITTIN_EPL"
      repShowINEPL.ReportFilter = " instanceid='" & id & "'"
      repShowINEPL.ReportPath = App.Path & "\in_epl.rpt"
      repShowINEPL.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowINEPL.Run True
      Set repShowINEPL = Nothing
    End If
    
    If f.optSRV.Value Then
      On Error Resume Next
      Set repShowSRVIN = Nothing
      Set repShowSRVIN = New ReportShow
      repShowSRVIN.ReportSource = "V_viewITTIN_ITTIN_SRV"
      repShowSRVIN.ReportFilter = " instanceid='" & id & "'"
      repShowSRVIN.ReportPath = App.Path & "\in_srvq.rpt"
      repShowSRVIN.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowSRVIN.Run True
      Set repShowSRVIN = Nothing
    End If
  End If
  Unload f
  Set f = Nothing
End Sub

' управление действием при открытии журнала - задание на перемещения
Private Sub jfmnuAllITTOPT_OnEdit(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
usedefaut = True
optEdit = True
End Sub

' Обработка команды печати для окна журнала -задание на перемещение
Private Sub jfmnuAllITTOPT_OnPrint(usedefaut As Boolean)
  usedefaut = False
 
  On Error Resume Next
  Dim id As String
  id = jfmnuAllITTOPT.jv.RowInstanceID(jfmnuAllITTOPT.jv.Row)
  If id = "" Then Exit Sub
  
  Dim iop As ITTOPT.Application
  Dim def As ITTOPT.ITTOPT_DEF
  Dim rtype As ITTD.ITTD_RULE
  Set iop = Manager.GetInstanceObject(id)
  Set def = iop.ITTOPT_DEF.Item(1)
  Set rtype = iop.ITTOPT_DEF.Item(1).TheRule
  
  
  Dim csstr As String
  

  csstr = csstr & "Товар"

 If rtype.UseClient = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Клиент"
  End If
  
  If rtype.TheCountry = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Страна"
  End If
  
  If rtype.TheFactory = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ";"
   End If
   csstr = csstr & "Завод"
  End If
  
  If rtype.killplace = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Бойня"
  End If
  
 If rtype.UsePartia = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Партия"
  End If
  
 
  
  
   If rtype.UseVetsved = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Сертификат"
  End If
  
  
  If rtype.UsePalType = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Тип пал."
   
  End If
   
  If rtype.UseBrak = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Брак"
  End If
  
  If rtype.UseExpDate = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ";"
   End If
   csstr = csstr & "Срок годн. мес.; год"
  End If
 
 
  Set repShowMoves = Nothing
  Set repShowMoves = New ReportShow
  repShowMoves.Formulas.Add("Param").Expression = """" & csstr & """"
  repShowMoves.ReportSource = "V_viewITTOPT_ITTOPT_MOVE"
  repShowMoves.ReportFilter = " instanceid='" & id & "'"
  repShowMoves.ReportPath = App.Path & "\Moves.rpt"
  repShowMoves.Run True
  Set repShowMoves = Nothing
End Sub


' запуск оптимизатора или открытие формы ввода перемещений
Private Sub jfmnuAllITTOPT_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean) 'обработка события - Действие для окна журанла
  If optEdit Then
    usedefaut = True
    optEdit = False
  Else
    usedefaut = False
    On Error Resume Next
    Dim id As String
    id = jfmnuAllITTOPT.jv.RowInstanceID(jfmnuAllITTOPT.jv.Row)
    If id = "" Then Exit Sub
    Dim iop As ITTOPT.Application
    Set iop = Manager.GetInstanceObject(id)
    
    Manager.LockInstanceObject iop.id
    
    If iop.StatusID = "{300483B2-1D94-4A33-8ADF-ABF32E72E57B}" Then
    
      Dim f As frmMovings
      Set f = New frmMovings
      Set f.movetask = Manager.GetInstanceObject(id)
      f.Show vbModal
      Unload f
    Else
      If MsgBox("Запустить процесс оптимизации ?", vbYesNo, "Оптимизация") = vbYes Then
        Dim opt As ITT2OPTBST.BEFORESTATUS
        Set opt = New ITT2OPTBST.BEFORESTATUS
        opt.RunOptimization iop
      End If
    End If
    Manager.UnLockInstanceObject iop.id
  End If
End Sub

' настройка  журнала отгрузка
Private Sub jfmnuAllITTOUT_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
 'bAdd = False
 'bDel = False

End Sub

' Обработка команды печати для окна журнала - отгрузка
Private Sub jfmnuAllITTOUT_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
 On Error Resume Next
    Dim id As String
    id = jfmnuAllITTOUT.jv.RowInstanceID(jfmnuAllITTOUT.jv.Row)
    usedefaut = False
    If id = "" Then Exit Sub
    
    
    Dim f As frmOutPrint
    Set f = New frmOutPrint
    f.Show vbModal
    
    ' выбран вариант отчета
    If f.OK Then
    
      If f.optActRas.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_ras.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
      If f.optOTB.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_OL.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
      If f.optSRV.Value Then
        On Error Resume Next
        Set repShowSRVOUT = Nothing
        Set repShowSRVOUT = New ReportShow
        repShowSRVOUT.ReportSource = "V_viewITTout_ITTout_SRV"
        repShowSRVOUT.ReportFilter = " instanceid='" & id & "'"
        repShowSRVOUT.ReportPath = App.Path & "\out_srvq.rpt"
        repShowSRVOUT.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowSRVOUT.Run True
        Set repShowSRVOUT = Nothing
      End If
    End If
    Unload f
    Set f = Nothing
End Sub

' настройка журнала палет
Private Sub jfmnuAllITTPL_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
'bAdd = False
'bDel = False

End Sub

' копирование списка услуг в документ услуги клиентов
Private Sub jfmnuITTCS_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean) 'обработка события - Действие для окна журанла
  usedefaut = False
  
  Dim id As String
  id = jfmnuITTCS.jv.RowInstanceID(RowIndex)
  If id = "" Then Exit Sub
  Dim cs As ITTCS.Application
  Dim dic As ITTD.Application
  Dim rs As ADODB.Recordset
  Set rs = Manager.ListInstances("", "ITTD")
  Set dic = Manager.GetInstanceObject(rs!InstanceID)
  Set cs = Manager.GetInstanceObject(id)
  Dim i As Long
  If cs.ITTCS_LIN.Count = 0 Then
    If MsgBox("Добавить список услуг из справочника?", vbYesNo) = vbYes Then
      For i = 1 To dic.ITTD_SRV.Count
        With cs.ITTCS_LIN.Add
          Set .srv = dic.ITTD_SRV.Item(i)
          .UseSrv = Boolean_Da
          .save
        End With
      Next
    Else
      usedefaut = True
    End If
  Else
    usedefaut = True
  End If


End Sub

'обработка события инициализации журнала приемка
Private Sub jfmnuITTIN_1_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

'обработка события инициализации журнала приемка
Private Sub jfmnuITTIN_2_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
  bAdd = False
  bDel = False
End Sub

' Обработка команды печати для окна журнала - приемка
Private Sub jfmnuITTIN_2_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
 usedefaut = False
  On Error Resume Next
  Dim id As String
  id = jfmnuITTIN_2.jv.RowInstanceID(jfmnuITTIN_2.jv.Row)
  If id = "" Then Exit Sub
  
  ' выбор варианта отчета
  Dim f As frmInPrint
  Set f = New frmInPrint
  f.Show vbModal
  If f.OK Then
'    акт расхождений
    If f.optActRas.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_ras.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
  
'    КЛП
    If f.optKLP.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_KL.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
    
'    пустые палет к заказу
    If f.optAct.Value Then
      On Error Resume Next
      Set repShowINEPL = Nothing
      Set repShowINEPL = New ReportShow
      repShowINEPL.ReportSource = "V_viewITTIN_ITTIN_EPL"
      repShowINEPL.ReportFilter = " instanceid='" & id & "'"
      repShowINEPL.ReportPath = App.Path & "\in_epl.rpt"
      repShowINEPL.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowINEPL.Run True
      Set repShowINEPL = Nothing
    End If
    
'    услуги
    If f.optSRV.Value Then
      On Error Resume Next
      Set repShowSRVIN = Nothing
      Set repShowSRVIN = New ReportShow
      repShowSRVIN.ReportSource = "V_viewITTIN_ITTIN_SRV"
      repShowSRVIN.ReportFilter = " instanceid='" & id & "'"
      repShowSRVIN.ReportPath = App.Path & "\in_srvq.rpt"
      repShowSRVIN.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowSRVIN.Run True
      Set repShowSRVIN = Nothing
    End If
  End If
  Unload f
  Set f = Nothing
End Sub

'обработка события инициализации журнала приемка
Private Sub jfmnuITTIN_3_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

' Обработка команды печати для окна журнала - приемка
Private Sub jfmnuITTIN_3_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
  usedefaut = False
  On Error Resume Next
  Dim id As String
  id = jfmnuITTIN_3.jv.RowInstanceID(jfmnuITTIN_3.jv.Row)
  If id = "" Then Exit Sub
  
   ' выбор варианта отчета
  Dim f As frmInPrint
  Set f = New frmInPrint
  f.Show vbModal
  If f.OK Then
  '    акт расхождений
    If f.optActRas.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_ras.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
'  клп
    If f.optKLP.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_KL.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
'    пустые паллеты
    If f.optAct.Value Then
      On Error Resume Next
      Set repShowINEPL = Nothing
      Set repShowINEPL = New ReportShow
      repShowINEPL.ReportSource = "V_viewITTIN_ITTIN_EPL"
      repShowINEPL.ReportFilter = " instanceid='" & id & "'"
      repShowINEPL.ReportPath = App.Path & "\in_epl.rpt"
      repShowINEPL.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowINEPL.Run True
      Set repShowINEPL = Nothing
    End If
'    услуги
    If f.optSRV.Value Then
      On Error Resume Next
      Set repShowSRVIN = Nothing
      Set repShowSRVIN = New ReportShow
      repShowSRVIN.ReportSource = "V_viewITTIN_ITTIN_SRV"
      repShowSRVIN.ReportFilter = " instanceid='" & id & "'"
      repShowSRVIN.ReportPath = App.Path & "\in_srvq.rpt"
      repShowSRVIN.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowSRVIN.Run True
      Set repShowSRVIN = Nothing
    End If
  End If
  Unload f
  Set f = Nothing
End Sub

'обработка события инициализации журнала приемка
Private Sub jfmnuITTIN_4_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
  bAdd = False
  bDel = False
End Sub

' Обработка команды печати для окна журнала - приемка
Private Sub jfmnuITTIN_4_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
 usedefaut = False
  On Error Resume Next
  Dim id As String
  id = jfmnuITTIN_4.jv.RowInstanceID(jfmnuITTIN_4.jv.Row)
  If id = "" Then Exit Sub
  
   ' выбор варианта отчета
  Dim f As frmInPrint
  Set f = New frmInPrint
  f.Show vbModal
  If f.OK Then
  '    акт расхождений
    If f.optActRas.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_ras.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
'  клп
    If f.optKLP.Value Then
      Set repShowKL = Nothing
      Set repShowKL = New ReportShow
      repShowKL.ReportSource = "V_viewITTIN_ITTIN_PALET"
      repShowKL.ReportFilter = " instanceid='" & id & "'"
      repShowKL.ReportPath = App.Path & "\in_KL.rpt"
      repShowKL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowKL.Run True
      Set repShowKL = Nothing
    End If
'    пустые паллеты
    If f.optAct.Value Then
      On Error Resume Next
      Set repShowINEPL = Nothing
      Set repShowINEPL = New ReportShow
      repShowINEPL.ReportSource = "V_viewITTIN_ITTIN_EPL"
      repShowINEPL.ReportFilter = " instanceid='" & id & "'"
      repShowINEPL.ReportPath = App.Path & "\in_epl.rpt"
      repShowINEPL.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowINEPL.Run True
      Set repShowINEPL = Nothing
    End If
'    услуги
    If f.optSRV.Value Then
      On Error Resume Next
      Set repShowSRVIN = Nothing
      Set repShowSRVIN = New ReportShow
      repShowSRVIN.ReportSource = "V_viewITTIN_ITTIN_SRV"
      repShowSRVIN.ReportFilter = " instanceid='" & id & "'"
      repShowSRVIN.ReportPath = App.Path & "\in_srvq.rpt"
      repShowSRVIN.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
      repShowSRVIN.Run True
      Set repShowSRVIN = Nothing
    End If
  End If
  Unload f
  Set f = Nothing
End Sub

Private Sub jfmnuITTOPT_1_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
bAdd = False
bDel = False
End Sub

'обработка события редактирования журнала задания на перемещения
Private Sub jfmnuITTOPT_6_OnEdit(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
optEdit = True
End Sub

Private Sub jfmnuITTOPT_6_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
bAdd = False
bDel = False
End Sub

'обработка события печати журнала задания на перемещения
Private Sub jfmnuITTOPT_6_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
  usedefaut = False
 
  On Error Resume Next
  Dim id As String
  id = jfmnuITTOPT_6.jv.RowInstanceID(jfmnuITTOPT_6.jv.Row)
  If id = "" Then Exit Sub
  
  Dim iop As ITTOPT.Application
  Dim def As ITTOPT.ITTOPT_DEF
  Dim rtype As ITTD.ITTD_RULE
  Set iop = Manager.GetInstanceObject(id)
  Set def = iop.ITTOPT_DEF.Item(1)
  Set rtype = iop.ITTOPT_DEF.Item(1).TheRule
  
  
  Dim csstr As String
'  формируем название колонки параметров товара

  csstr = csstr & "Товар"

 If rtype.UseClient = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Клиент"
  End If
  
  If rtype.TheCountry = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Страна"
  End If
  
  If rtype.TheFactory = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ";"
   End If
   csstr = csstr & "Завод"
  End If
  
  If rtype.killplace = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Бойня"
  End If
  
 If rtype.UsePartia = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Партия"
  End If
  
 
  
  
   If rtype.UseVetsved = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Сертификат"
  End If
  
  
  If rtype.UsePalType = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Тип пал."
   
  End If
   
  If rtype.UseBrak = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & "; "
   End If
   csstr = csstr & "Брак"
  End If
  
  If rtype.UseExpDate = Boolean_Da Then
   If csstr <> "" Then
   csstr = csstr & ";"
   End If
   csstr = csstr & "Срок годн. мес.; год"
  End If
 
' выводим отчет
  Set repShowMoves = Nothing
  Set repShowMoves = New ReportShow
  repShowMoves.Formulas.Add("Param").Expression = """" & csstr & """"
  repShowMoves.ReportSource = "V_viewITTOPT_ITTOPT_MOVE"
  repShowMoves.ReportFilter = " instanceid='" & id & "'"
  repShowMoves.ReportPath = App.Path & "\Moves.rpt"
  repShowMoves.Run True
  Set repShowMoves = Nothing
  
  
End Sub

'обработка события действие журнала задания на перемещения
Private Sub jfmnuITTOPT_6_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean) 'обработка события - Действие для окна журанла
  If optEdit Then
    usedefaut = True
    optEdit = False
  Else
  
    usedefaut = False
    Dim f As frmMovings
    Set f = New frmMovings
    
    On Error Resume Next
    Dim id As String
    id = jfmnuITTOPT_6.jv.RowInstanceID(jfmnuITTOPT_6.jv.Row)
    If id = "" Then Exit Sub
    
    Set f.movetask = Manager.GetInstanceObject(id)
    f.Show vbModal
    Unload f
  End If
End Sub

'обработка события редактирования журнала задания на перемещения
Private Sub jfmnuITTOPT_7_OnEdit(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
optEdit = True
End Sub

Private Sub jfmnuITTOPT_7_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
bAdd = False
bDel = False
End Sub

'обработка события действие журнала задания на перемещения
Private Sub jfmnuITTOPT_7_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
  If optEdit Then
    usedefaut = True
    optEdit = False
  Else
     usedefaut = False
     On Error Resume Next
     Dim id As String
     id = jfmnuITTOPT_7.jv.RowInstanceID(jfmnuITTOPT_7.jv.Row)
     If id = "" Then Exit Sub
     Dim iop As ITTOPT.Application
     Set iop = Manager.GetInstanceObject(id)
     Manager.LockInstanceObject iop.id
     
'    в зависимости от состояния документе либо форма ввода перемещений, либо запуск оптимизации
    If iop.StatusID = "{300483B2-1D94-4A33-8ADF-ABF32E72E57B}" Then
     
       Dim f As frmMovings
       Set f = New frmMovings
       Set f.movetask = Manager.GetInstanceObject(id)
       f.Show vbModal
       Unload f
     Else
       If MsgBox("Запустить процесс оптимизации ?", vbYesNo, "Оптимизация") = vbYes Then
         Dim opt As ITT2OPTBST.BEFORESTATUS
         Set opt = New ITT2OPTBST.BEFORESTATUS
         opt.RunOptimization iop
       End If
     End If
     Manager.UnLockInstanceObject iop.id
  End If
End Sub

'обработка события инициализация журнала отгрузка
Private Sub jfmnuITTOUT_1_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub
'обработка события инициализация журнала отгрузка
Private Sub jfmnuITTOUT_2_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

'обработка события печать журнала отгрузка
Private Sub jfmnuITTOUT_2_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
 On Error Resume Next
    Dim id As String
    id = jfmnuITTOUT_2.jv.RowInstanceID(jfmnuITTOUT_2.jv.Row)
    usedefaut = False
    If id = "" Then Exit Sub
    
'    выбор формы отчета
    Dim f As frmOutPrint
    Set f = New frmOutPrint
    f.Show vbModal
    If f.OK Then
      
'      акт расхождений
      If f.optActRas.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_ras.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
'      отборочный лист
      If f.optOTB.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_OL.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
'      услуги
      If f.optSRV.Value Then
        On Error Resume Next
        Set repShowSRVOUT = Nothing
        Set repShowSRVOUT = New ReportShow
        repShowSRVOUT.ReportSource = "V_viewITTout_ITTout_SRV"
        repShowSRVOUT.ReportFilter = " instanceid='" & id & "'"
        repShowSRVOUT.ReportPath = App.Path & "\out_srvq.rpt"
        repShowSRVOUT.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowSRVOUT.Run True
        Set repShowSRVOUT = Nothing
      End If
    End If
    Unload f
    Set f = Nothing
End Sub

'обработка события инициализация журнала отгрузка
Private Sub jfmnuITTOUT_3_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

'обработка события печать журнала отгрузка
Private Sub jfmnuITTOUT_3_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
    On Error Resume Next
    Dim id As String
    id = jfmnuITTOUT_3.jv.RowInstanceID(jfmnuITTOUT_3.jv.Row)
    usedefaut = False
    If id = "" Then Exit Sub
    
    
    '    выбор формы отчета
    Dim f As frmOutPrint
    Set f = New frmOutPrint
    f.Show vbModal
    If f.OK Then
      
'      акт расхождений
      If f.optActRas.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_ras.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
'      отборочный лист
      If f.optOTB.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_OL.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
'      услуги
      If f.optSRV.Value Then
        On Error Resume Next
        Set repShowSRVOUT = Nothing
        Set repShowSRVOUT = New ReportShow
        repShowSRVOUT.ReportSource = "V_viewITTout_ITTout_SRV"
        repShowSRVOUT.ReportFilter = " instanceid='" & id & "'"
        repShowSRVOUT.ReportPath = App.Path & "\out_srvq.rpt"
        repShowSRVOUT.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowSRVOUT.Run True
        Set repShowSRVOUT = Nothing
      End If
    End If
    Unload f
    Set f = Nothing
    
End Sub

'обработка события инициализация журнала отгрузка
Private Sub jfmnuITTOUT_4_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
  'bAdd = False
  bDel = False
End Sub

'обработка события печать журнала отгрузка
Private Sub jfmnuITTOUT_4_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
 On Error Resume Next
    Dim id As String
    id = jfmnuITTOUT_4.jv.RowInstanceID(jfmnuITTOUT_4.jv.Row)
    usedefaut = False
    If id = "" Then Exit Sub
    
    
    '    выбор формы отчета
    Dim f As frmOutPrint
    Set f = New frmOutPrint
    f.Show vbModal
    If f.OK Then
      
'      акт расхождений
      If f.optActRas.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_ras.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
'      отборочный лист
      If f.optOTB.Value Then
        Set repShowOL = Nothing
        Set repShowOL = New ReportShow
        repShowOL.ReportSource = "V_viewITTOUT_ITTOUT_PALET"
        repShowOL.ReportFilter = " instanceid='" & id & "'"
        repShowOL.ReportPath = App.Path & "\out_OL.rpt"
        repShowOL.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowOL.Run True
        Set repShowOL = Nothing
      End If
      
'      услуги
      If f.optSRV.Value Then
        On Error Resume Next
        Set repShowSRVOUT = Nothing
        Set repShowSRVOUT = New ReportShow
        repShowSRVOUT.ReportSource = "V_viewITTout_ITTout_SRV"
        repShowSRVOUT.ReportFilter = " instanceid='" & id & "'"
        repShowSRVOUT.ReportPath = App.Path & "\out_srvq.rpt"
        repShowSRVOUT.PrinterName = "" 'GetSetting("RBH", "ITTSETTINGS", "DOCPRN", "")
        repShowSRVOUT.Run True
        Set repShowSRVOUT = Nothing
      End If
    End If
    Unload f
    Set f = Nothing
End Sub

'обработка события инициализация журнала паллеты
Private Sub jfmnuITTPL_1_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

'обработка события инициализация журнала паллеты
Private Sub jfmnuITTPL_2_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

'обработка события инициализация журнала паллеты
Private Sub jfmnuITTPL_3_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

'обработка события инициализация журнала паллеты
Private Sub jfmnuITTPL_4_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub

'обработка события инициализация журнала паллеты
Private Sub jfmnuITTPL_5_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
bAdd = False
bDel = False

End Sub


' создание документа  - акт расхождений
Private Sub jfmnuITTPR_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
usedefaut = False
 Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTPR", "Протокол" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False

End Sub

'обработка события инициализация журнала - акт расхождений
Private Sub jfmnuITTPR_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean) 'обработка события инициализации журнала
  'bAdd = False
  bFilter = False
End Sub

'обработка события печать журнала - акт расхождений
Private Sub jfmnuITTPR_OnPrint(usedefaut As Boolean)  ' Обработка команды печати для окна журнала
    usedefaut = False
    On Error Resume Next
    
    If MsgBox("Напечать только текущий акт?", vbYesNo, "Уточните") = vbYes Then
'      печать одного акта
      Dim id As String
      id = jfmnuITTPR.jv.RowInstanceID(jfmnuITTPR.jv.Row)
      If id = "" Then Exit Sub
      
      Set RptActVes = New ReportShow
      RptActVes.ReportPath = App.Path & "\AktVes.rpt"
      RptActVes.ReportSource = "V_AUTOITTPR_DEF"
      RptActVes.ReportFilter = "instanceid ='" & id & "'"
      Call RptActVes.Run(True)
      Set RptActVes = Nothing
    Else
'      печать всех актов
      Set RptActVes = New ReportShow
      RptActVes.ReportPath = App.Path & "\AktVesAll.rpt"
      RptActVes.ReportSource = "V_AUTOITTPR_DEF"
      Call RptActVes.Run(True)
      Set RptActVes = Nothing

    
    End If
End Sub

' загрузка формы
Private Sub MDIForm_Load()
  
  
  ' инициализация меню
  On_Load


 Dim o As ITTOP.Application
 Dim rs  As ADODB.Recordset
 Dim id As String
 Dim cliFilter As String
 Dim camFilter As String
 Dim i As Long, j As Long
  Set rs = Manager.ListInstances("", "ITTOP")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTOP", "Операторы и кладовщики"
End If
Set o = Manager.GetInstanceObject(id)

' формирование фильтров кладовщика  \ оператора для клиента и камеры
For i = 1 To o.ITTOP_KLNK.Count
  If o.ITTOP_KLNK.Item(i).TheUser Is MyUser Then
    For j = 1 To o.ITTOP_KLNK.Item(i).ITTOP_KCLI.Count
      If cliFilter <> "" Then
      cliFilter = cliFilter & ","
      End If
      cliFilter = cliFilter & "'" & GetBRIEFFromXMLField(o.ITTOP_KLNK.Item(i).ITTOP_KCLI.Item(j).TheClient) & "'"
    Next
    For j = 1 To o.ITTOP_KLNK.Item(i).ITTOP_KCAM.Count
     If camFilter <> "" Then
      camFilter = camFilter & " or "
      End If
      camFilter = camFilter & " location.code like '" & o.ITTOP_KLNK.Item(i).ITTOP_KCAM.Item(j).TheKamera.CameraMask & "' "
    Next
  
  End If
Next


For i = 1 To o.ITTOP_OPLNK.Count
  If o.ITTOP_OPLNK.Item(i).TheUser Is MyUser Then
   For j = 1 To o.ITTOP_OPLNK.Item(i).ITTOP_OPKAM.Count
     If camFilter <> "" Then
      camFilter = camFilter & " or "
      End If
      camFilter = camFilter & " location.code like '" & o.ITTOP_OPLNK.Item(i).ITTOP_OPKAM.Item(j).TheKamera.CameraMask & "' "
    Next
  End If
Next
 
 If camFilter <> "" Then
     camFilter = " and (" & camFilter & ") "
 End If

 If cliFilter <> "" Then
      cliFilter = " partner.name in ( " & cliFilter & ") "
 End If
 
 ' сохранение фильтров
 Dim Obj As DBuffer
 Set Obj = New DBuffer
 Obj.Name = camFilter
 Manager.AddCustomObjects Obj, "camFilter"
 
 Set Obj = New DBuffer
 Obj.Name = cliFilter
 Manager.AddCustomObjects Obj, "cliFilter"


  ' кеширование справочника
  Set rs = Manager.ListInstances("", "ITTD")
  If Not rs.EOF Then
    id = rs!InstanceID
    Set ITTDic = Manager.GetInstanceObject(id)
  End If
 
 

End Sub

'  выгрузка формы
Private Sub mdiForm_Unload(Cancel As Integer)
On Error Resume Next

' whait for finalize timer loops
inTimer1 = True
Me.Timer1.Enabled = False

inTimer2 = True
Me.Timer2.Enabled = False


Timer1.Enabled = False
Timer2.Enabled = False

On Error Resume Next

' unload all dynamically created journals and reports
UnloadObjects




Dim f As Form
For Each f In Forms
  If f.MDIChild = True Then
    On Error Resume Next
    'Call f.Controls.Item(0).object.Init(Nothing, Nothing, False, Nothing)
    Unload f
  End If
Next

  For Each f In Forms
      On Error Resume Next
      Debug.Print f.Name
  Next
  
  
  Set MyRole = Nothing
  Set MyUser = Nothing
  Set usr = Nothing

  log.message "Завершение сесии " & Session.sessionid
  Session.Logout
  Set Session = Nothing
  Manager.CloseClass
  Set Manager = Nothing

  If Command$ <> "DEBUG" Then
   TerminateProcess GetCurrentProcess, 0
  'Else
  ' End
  End If
End Sub

' о программе
Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub




'  сборка паддона
Private Sub mnuAssemble_Click()
Dim f As frmAssemblyWizard
Set f = New frmAssemblyWizard
f.Show vbModal
Unload f
Set f = Nothing
End Sub

' режим авто
Private Sub mnuAuto_Click()
    Dim f As frmAuto
    Set f = frmAuto
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub



' настройка соединения с CORE
Private Sub mnuCoreSetup_Click()
  Dim f As frmCoreSetup
  Set f = New frmCoreSetup
  f.Show vbModal
  Unload f
  Set f = Nothing
End Sub

' выгрузка услуг
Private Sub mnuExportSRV_Click()
Dim f As frmDates
  Set f = New frmDates
  f.Show vbModal
  If f.OK Then
  
'    подготовка параметров выгрузки
    Dim s As String
    If f.lbldfrom.Value = vbChecked Then
     s = "ProcessDate >=" & MakeMSSQLDate(f.dtpdfrom.Value)
    End If
    
    If f.lbldTo.Value = vbChecked Then
     If s <> "" Then
      s = s & " and "
     End If
     s = s & "ProcessDate <" & MakeMSSQLDate(f.dtpdTo.Value + 1)
    End If
   
    
    
    Dim fcol As Collection
    Set fcol = New Collection
    Dim rf As ObjHolder
    Set rf = New ObjHolder
    rf.id = "Client"
    rf.Name = "Client"
    fcol.Add rf
    Set rf = New ObjHolder
    rf.id = "ZType"
    rf.Name = "ZTYPE"
    fcol.Add rf
    
    Set rf = New ObjHolder
    rf.id = "ZAKAZ"
    rf.Name = "ZAKAZ"
    fcol.Add rf
    
    Set rf = New ObjHolder
    rf.id = "ProcessDate"
    rf.Name = "ProcessDate"
    fcol.Add rf
    
    Set rf = New ObjHolder
    rf.id = "srv"
    rf.Name = "srv"
    fcol.Add rf
    
    Set rf = New ObjHolder
    rf.id = "Quantity"
    rf.Name = "Quantity"
    fcol.Add rf
    
    Set rf = New ObjHolder
    rf.id = "CLIENT_ID"
    rf.Name = "CLIENT_ID"
    fcol.Add rf
    
    Set rf = New ObjHolder
    rf.id = "ZAKAZ_ID"
    rf.Name = "ZAKAZ_ID"
    fcol.Add rf
    
    Set rf = New ObjHolder
    rf.id = "SRV_ID"
    rf.Name = "SRV_ID"
    fcol.Add rf
    

'    выгрузка в форму
    Dim fxl As frmXL
    Set fxl = New frmXL
    ProcessQry "select * from V_SERVICE where Quantity <>0 and " & s & " order by ProcessDate,Client,ZType,ZAKAZ,srv", fcol, fxl, 0, "Выгрузка по услугам", True
    
'    открытие формы
    fxl.Show
    
  End If
End Sub


'выгрузка данных в форму для последующего сохранения в Excel
'параметры
'qry -запрос
'  fcol -  описание колонок
'  fxl -форма
'  ColPos - позиция колонок
'  ColsName - общий заголовок
'  skipEmpty - пропускать строки без данных
'результат -нет
Private Sub ProcessQry(ByVal qry As String, fcol As Object, fxl As frmXL, ColPos As Long, ByVal ColsName As String, skipEmpty As Boolean)
   'Debug.Print qry
   
   On Error GoTo bye
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData(qry)
  If rs Is Nothing Then Exit Sub
  If skipEmpty Then
      If rs.EOF Then Exit Sub
  End If
  Dim i As Long, j As Long, fp As Long
  Dim rf As ObjHolder
  i = 0
  fp = 0
  fxl.vfgr.Cols = 1 + ColPos
  fxl.vfgr.TextMatrix(0, ColPos) = ColsName
  For j = 1 To fcol.Count
    Set rf = fcol.Item(j)
    fxl.vfgr.Cols = 1 + ColPos + fp
    fxl.vfgr.TextMatrix(1, ColPos + fp) = rf.Name
    fp = fp + 1
  Next

  
  While Not rs.EOF
     On Error Resume Next
     ' Ищем строку для нашего ID
     'For i = 2 To fxl.vfgr.Rows
       'If rs!id = fxl.vfgr.RowData(i) Then
        fxl.vfgr.AddItem ""
        
        i = fxl.vfgr.Rows - 1
        
        fp = 0
        For j = 1 To fcol.Count
         Set rf = fcol.Item(j)
         
         
           If IsNull(rs.fields.Item(rf.id).Value) Then
            fxl.vfgr.TextMatrix(i, ColPos + fp) = ""
           Else
           
              fxl.vfgr.TextMatrix(i, ColPos + fp) = rs.fields.Item(rf.id).Value & ""
           
           End If
           fp = fp + 1
         
        Next
        'Exit For
       'End If
     'Next
   rs.MoveNext
  Wend
  Set rs = Nothing
'  On Error GoTo bye
'  For i = ColPos To ColPos + fp - 1
'    fxl.vfgr.TextMatrix(fxl.vfgr.Rows - 1, i) = fxl.vfgr.Aggregate(flexSTSum, 2, i, fxl.vfgr.Rows - 2, i)
'    fxl.vfgr.ColFormat(i) = "#,###.##"
'  Next
  ColPos = ColPos + fp
  fxl.vfgr.MergeRow(0) = True
  Exit Sub
bye:
  'Stop
  'Resume
 End Sub
 
' открытие журнала - протокол расхождений
Private Sub mnuITTPR_Click()

    Dim journal As Object
    On Error Resume Next
    If jfmnuITTPR Is Nothing Then
      Set jfmnuITTPR = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{D6B430E0-CCF2-4C4D-A0BB-A394492C05EA}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTPR.jv.journal = journal
      jfmnuITTPR.jv.OpenModal = False
      jfmnuITTPR.Caption = "Протокол расхождений"
      Me.MousePointer = vbHourglass
      DoEvents
      jfmnuITTPR.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTPR.Show
    jfmnuITTPR.WindowState = 0
    jfmnuITTPR.ZOrder 0
End Sub


' управление размером ячейки
Private Sub mnuLocationSize_Click()
    Dim f As frmLocSize
    Set f = New frmLocSize
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub

' произвести отбор по выморозке
Private Sub mnuMakeOtbor_Click()
    Dim f As frmMakeOtbor
    Set f = New frmMakeOtbor
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub

' создание поддонов
Private Sub mnuMakePoddon_Click()
  
    Dim f As frmAddPalette
    Set f = New frmAddPalette
    f.Show vbModal
    Manager.FreeAllInstanses
    Unload f
    Set f = Nothing
End Sub

' печать стикера
Private Sub mnuMySticker_Click()
  Dim f As frmPrintSticker
  Set f = New frmPrintSticker
  f.Show vbModal
  If f.OK Then
    PrintSticker f.poddon
  End If
  Unload f
  Set f = Nothing
End Sub


' открытие документа - настройки системы
Private Sub mnuNumbers_Click()
Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTFN")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTFN", "Нумератор"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub

' отчет по отбору
Private Sub mnuOtbor_Click()
  

      Dim conn As ADODB.Connection
      Dim rs As ADODB.Recordset
      
      Set conn = GetCoreConn
      
      Set rs = conn.Execute("select * from v_bami_vimorozka_rpt2 ")
      
      Set RptVimorozka2 = New ReportShow
      RptVimorozka2.ReportPath = App.Path & "\Otbor.rpt"
      Call RptVimorozka2.RunDirectRS(rs, False)
  
End Sub

' настройка принтеров
Private Sub mnuPRNSetup_Click()
  Dim f As frmPrnSetup
  Set f = New frmPrnSetup
  f.Show vbModal
  Unload f
  Set f = Nothing
End Sub

' выход
Private Sub mnuExit_Click()
  Unload Me
End Sub

' отгрузка
Private Sub mnuProcessShipping_Click()
  Dim f As frmOutWiz
  Set f = New frmOutWiz
  f.Show vbModal
  Unload f
  Set f = Nothing
End Sub

' взвешивание поддонов
Private Sub mnuPWiz_Click()
 Dim f As frmWizPoddons
 Set f = frmWizPoddons
 f.Show vbModal
 Unload f
 Set f = Nothing
End Sub

' отчет по заблокированным поддонам
Private Sub mnuRpt103_Click()
    Dim conn As ADODB.Connection
    
    Set RptStok103 = New ReportShow
    RptStok103.ReportPath = App.Path & "\stok103.rpt"
    RptStok103.ReportSource = "v_bami_stock103"
    Set conn = GetCoreConn
    
    Call RptStok103.Run(False, conn)
'    RptStok103.ExportPDF App.Path & "\" & "Stock103.pdf", conn
'
'    MailThisFile "Заблокированные поддоны", "Отчет по заблокированным поддонам на  " & Now & ".", App.Path & "\" & "Stock103.pdf"
    
End Sub

' отчет по объему хранения
Private Sub mnuRptHran_Click()
 Dim conn As ADODB.Connection
    
    Set RptHran = New ReportShow
    RptHran.ReportPath = App.Path & "\Objem.rpt"
    RptHran.ReportSource = "v_bami_hranenie"
    Set conn = GetCoreConn
    
    Call RptHran.Run(False, conn)
End Sub


' Отчет по отобранному товару
Private Sub mnuRptOtobrano_Click()
 Dim conn As ADODB.Connection
    
    Set RptOtobrano = New ReportShow
    RptOtobrano.ReportPath = App.Path & "\otobrano.rpt"
    RptOtobrano.ReportSource = "v_bami_stockblocked"
    Set conn = GetCoreConn
    
    Call RptOtobrano.Run(False, conn)
    

End Sub

' отчет по услугам
Private Sub mnuRptSrv_Click()
  Dim f As frmDates
  Set f = New frmDates
  f.Show vbModal
  If f.OK Then
  

    Set RptShowSRVALL = New ReportShow
    RptShowSRVALL.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "ZPRN", "")
    RptShowSRVALL.ReportPath = App.Path & "\srv_all.rpt"
    RptShowSRVALL.ReportSource = "V_SERVICE"
    
    Dim s As String
    If f.lbldfrom.Value = vbChecked Then
     s = "ProcessDate >=" & MakeMSSQLDate(f.dtpdfrom.Value)
    End If
    
    If f.lbldTo.Value = vbChecked Then
     If s <> "" Then
      s = s & " and "
     End If
     s = s & "ProcessDate <=" & MakeMSSQLDate(f.dtpdTo.Value + 1)
    End If
    RptShowSRVALL.ReportFilter = s
    RptShowSRVALL.Run
  End If
End Sub

' настройки
Private Sub mnuSetup_Click()
Dim f As frmSetup
Set f = New frmSetup
f.Show vbModal
Set f = Nothing
End Sub

' разборка поддона
Private Sub mnuSplitPoddon_Click()
Dim f As frmSplitWizard
Set f = New frmSplitWizard
f.Show vbModal
Unload f
Set f = Nothing
End Sub

' синхронизация справочников ( устаревший режим)
Private Sub mnuSyncDict_Click()
  Dim dic As ITTD.Application
  Dim conn As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim i As Long, j As Long
  Dim CountryOK As Boolean
  Dim FactoryOK As Boolean
  
  Set conn = GetCoreConn
  If conn.State <> adStateOpen Then
    conn.open
  End If
  
  Set rs = Manager.ListInstances("", "ITTD")
  Set dic = Manager.GetInstanceObject(rs!InstanceID)
'  синхронизация стран
  Set rs = conn.Execute("select * from COUNTRY_MEET")
  If rs Is Nothing Then Exit Sub
  While Not rs.EOF
    CountryOK = False
    For i = 1 To dic.ITTD_COUNTRY.Count
      If dic.ITTD_COUNTRY.Item(i).Code1 = rs!id Then
       CountryOK = True
       dic.ITTD_COUNTRY.Item(i).Name = rs!Name
       dic.ITTD_COUNTRY.Item(i).Code2 = rs!code
       dic.ITTD_COUNTRY.Item(i).save
       Exit For
      End If
     Next
     If CountryOK = False Then
      With dic.ITTD_COUNTRY.Add
        .Name = rs!Name
        .Code2 = rs!code
        .Code1 = rs!id
        .save
      End With
     End If
    
    rs.MoveNext
  Wend
  
'  синхронизация фабрик
   For i = 1 To dic.ITTD_COUNTRY.Count
     If IsNumeric(dic.ITTD_COUNTRY.Item(i).Code2) Then
      Set rs = conn.Execute("select * from PRODUCER_MEET where COUNTRY_MEET_CODE =" & dic.ITTD_COUNTRY.Item(i).Code2)
      If rs Is Nothing Then Exit Sub
      While Not rs.EOF
        CountryOK = False
         For j = 1 To dic.ITTD_FACTORY.Count
          If dic.ITTD_FACTORY.Item(j).Code1 = rs!id Then
           CountryOK = True
           dic.ITTD_FACTORY.Item(j).Name = rs!Name
           dic.ITTD_FACTORY.Item(j).Code2 = rs!code
           Set dic.ITTD_FACTORY.Item(j).country = dic.ITTD_COUNTRY.Item(i)
           dic.ITTD_FACTORY.Item(j).save
           Exit For
          End If
        Next
        If CountryOK = False Then
         With dic.ITTD_FACTORY.Add
           Set .country = dic.ITTD_COUNTRY.Item(i)
           .Name = rs!Name
           .Code2 = rs!code
           .Code1 = rs!id
           .save
         End With
        End If
        
        rs.MoveNext
      Wend
    End If
  Next
    
'    синхронизация боен
  For i = 1 To dic.ITTD_FACTORY.Count
   If IsNumeric(dic.ITTD_FACTORY.Item(i).Code2) Then
    Set rs = conn.Execute("select * from KILL_NUMBER_MEET where PRODUCER_MEET_CODE =" & dic.ITTD_FACTORY.Item(i).Code2)
    If rs Is Nothing Then Exit Sub
    While Not rs.EOF
      FactoryOK = False
       For j = 1 To dic.ITTD_KILLPLACE.Count
        If dic.ITTD_KILLPLACE.Item(j).Code1 = rs!id Then
         FactoryOK = True
         dic.ITTD_KILLPLACE.Item(j).Name = rs!Name
         dic.ITTD_KILLPLACE.Item(j).Code2 = rs!code
         Set dic.ITTD_KILLPLACE.Item(j).factory = dic.ITTD_FACTORY.Item(i)
         dic.ITTD_KILLPLACE.Item(j).save
         Exit For
        End If
      Next
      If FactoryOK = False Then
       With dic.ITTD_KILLPLACE.Add
         Set .factory = dic.ITTD_FACTORY.Item(i)
         .Name = rs!Name
         .Code2 = rs!code
         .Code1 = rs!id
         .save
       End With
      End If
      
      rs.MoveNext
    Wend
  End If
Next


   
MsgBox "Загрузка справочников завершена"

End Sub

' обновить данные в CORE
Private Sub mnuUpdateCore_Click()
  Dim f As frmUpdateCore
  Set f = New frmUpdateCore
  f.Show vbModal
  Set f = Nothing
End Sub

' отчет по выморозке
Private Sub mnuVimorozka_Click()
    Dim f As frmDate
    Set f = New frmDate
    f.Show vbModal
    If f.OK Then

      Dim conn As ADODB.Connection
      Dim rs As ADODB.Recordset
      
      Set conn = GetCoreConn
      Set rs = conn.Execute("select * from v_bami_vimorozka_rpt union all " & _
      "select partner_code ,item_id, item_code, description,  qin, qout, vimorozka * datediff(d,getdate()," & MakeMSSQLDate(f.dtpDate.Value) & "), pogreshnost,0, 0,0 from v_bami_stokmorozdayly")
      
      Set RptVimorozka = New ReportShow
      RptVimorozka.ReportPath = App.Path & "\Vimorozka.rpt"
      Call RptVimorozka.RunDirectRS(rs, False)
    End If
    Unload f
    Set f = Nothing
End Sub

' настройка весов
Private Sub mnuWeightSetup_Click()
  Dim f As frmWSetup
  Set f = New frmWSetup
  f.Show vbModal
  Unload f
  Set f = Nothing

End Sub

' печать стикеров с номерами поддонов
Private Sub mnuStickers_Click()
'  получить параметры
  Dim f As frmStickerRpt
  Set f = New frmStickerRpt
  f.Show vbModal
  If f.OK Then
  
'    вывод окна печати
    Set RptStickers = New ReportShow
    RptStickers.PrinterName = GetSetting("RBH", "ITTSETTINGS", "ZPRN", "")
    RptStickers.ReportPath = App.Path & "\palette.rpt"
    RptStickers.ReportSource = "V_AUTOITTPL_DEF"
    Dim s As String
    If f.txtFrom <> "" Then
     s = "ITTPL_DEF_TheNumber >=" & MyRound("0" & f.txtFrom)
    End If
    
    If f.txtFrom <> "" Then
     If s <> "" Then
      s = s & " and "
     End If
     s = s & "ITTPL_DEF_TheNumber <=" & MyRound("0" & f.txtTo)
    End If
    RptStickers.ReportFilter = s
    RptStickers.Run
  End If
  Unload f
  Set f = Nothing
End Sub

' приемка
Private Sub mnuProcessQuery_Click()
Dim f As frmInWiz2
Set f = New frmInWiz2
f.Show vbModal
Unload f
Set f = Nothing
End Sub

' отчет о ячейках со смешанным товаром
Private Sub mnuWrongLocation_Click()
    Dim conn As ADODB.Connection
    
    Set RptWrongLocation = New ReportShow
    RptWrongLocation.PrinterName = "" ' GetSetting("RBH", "ITTSETTINGS", "ZPRN", "")
    RptWrongLocation.ReportPath = App.Path & "\WrongLocation.rpt"
    RptWrongLocation.ReportSource = " v_bami_manual"
      Set conn = GetCoreConn
    
    Call RptWrongLocation.Run(False, conn)

End Sub



Private Sub Picture1_Click()

End Sub

' оповещение системы об активности сесии
Private Sub Timer2_Timer()
  If inTimer2 Then Exit Sub
  inTimer2 = True
  On Error Resume Next
  Call Session.Exec("SessionTouch", Nothing)
  inTimer2 = False
End Sub


'убрать лишние табуляции и переводы строк
Private Function NoTabs(ByVal s As String) As String
  NoTabs = Replace(Replace(Replace(Replace(s, vbTab, " "), vbCrLf, " "), vbCr, " "), vbLf, " ")
End Function

'сохранение струткуры меню в описании арм в базе данных
'Parameters:
' параметров нет
'Returns:
'  объект любого класса Visual Basic
'  ,или Nothing
'  ,или значение любого скалярного типа
'See Also:
'  On_Load
'Example:
' dim variable as Variant
'  variable = me.SynchronizeARMDescription()
' Set variable = me.SynchronizeARMDescription()
Public Function SynchronizeARMDescription()
Attribute SynchronizeARMDescription.VB_HelpID = 580
    Dim objARM As Object
    Dim objMenuItem As Menu
    Dim ObjItem As Object

    Set objARM = Manager.GetInstanceObject(ARMID)
    
    Dim i As Long
    Dim objRS As ADODB.Recordset
    Dim objEntryPoint As Object
    
    For i = 0 To Me.Controls.Count - 1
        Set ObjItem = Me.Controls(i)
        If UCase(TypeName(ObjItem)) = UCase("menu") Then
            If ObjItem.Caption <> "-" Then
              Debug.Print "Found menu " + ObjItem.Caption + "-" + ObjItem.Name
              
              Set objRS = Session.GetRowsEx("EntryPoints", ARMID, , "Caption='" + ObjItem.Caption + "' or Name='" & ObjItem.Name & "'")
              If objRS.EOF And objRS.BOF Then
                  Set objEntryPoint = objARM.EntryPoints.Add
                  objEntryPoint.Caption = ObjItem.Caption
                  objEntryPoint.Name = ObjItem.Name
                  objEntryPoint.AsToolbarItem = Boolean_Net
                  objEntryPoint.ActionType = 0 'MenuActionType_Nicego_ne_delat_
                  objEntryPoint.save
                  If err.Number <> 0 Then
                    MsgBox err.Description
                  End If
                  err.Clear
              Else
                  Set objEntryPoint = objARM.FindRowObject("EntryPoints", objRS!Entrypointsid)
                  If Not objEntryPoint Is Nothing Then
                    objEntryPoint.Caption = ObjItem.Caption
                    objEntryPoint.Name = ObjItem.Name
                    objEntryPoint.AsToolbarItem = Boolean_Net
                    objEntryPoint.save
                  End If
              End If
              objRS.Close
            End If
        End If
    Next
End Function



' обработка нажатия кнопок тулбара
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "in"
      If mnuProcessQuery.Visible And mnuProcessQuery.Enabled Then
        mnuProcessQuery_Click
      Else
        MsgBox "Действие не разрешено"
      End If
      
    Case "out"
      
      If mnuProcessShipping.Visible And mnuProcessShipping.Enabled Then
        mnuProcessShipping_Click
      Else
        MsgBox "Действие не разрешено"
      End If
    Case "toqry"
      If mnuPWiz.Visible And mnuPWiz.Enabled Then
        mnuPWiz_Click
      Else
        MsgBox "Действие не разрешено"
      End If
    Case "add"
      If mnuAssemble.Visible And mnuAssemble.Enabled Then
        mnuAssemble_Click
      Else
        MsgBox "Действие не разрешено"
      End If
    Case "div"
      If mnuSplitPoddon.Visible And mnuSplitPoddon.Enabled Then
        mnuSplitPoddon_Click
      Else
        MsgBox "Действие не разрешено"
      End If
    Case "auto"
      If mnuAuto.Visible And mnuAuto.Enabled Then
        mnuAuto_Click
      Else
        MsgBox "Действие не разрешено"
      End If
    Case "prn"
      If mnuMySticker.Visible And mnuMySticker.Enabled Then
        mnuMySticker_Click
      Else
        MsgBox "Действие не разрешено"
      End If
  End Select
End Sub

' действие при загрузке формы - управление меню
'Parameters:
' параметров нет
'See Also:
'  SynchronizeARMDescription
'Example:
'  call me.On_Load()
Public Sub On_Load()
Attribute On_Load.VB_HelpID = 575
   Me.Caption = App.FileDescription & " (" & Site & "\" & MyRole.Name & "\" & MyUser.brief & ")"
   log.message "Начало сессии " & Me.Caption & " " & Session.sessionid
   On Error Resume Next
   'If command$ <> "DEBUG" Then
     Dim c As Control
     For Each c In Me.Controls
      If TypeName(c) = "Menu" Then
         
        If CheckMenu(c.Name) = RoleMenuStatus_Hidden Then
          c.Visible = False
        Else
          frmSplash.lblWarning = "Инициализация меню: " & c.Caption
          DoEvents
        End If
      End If
     Next
  'End If
   Manager.FreeAllInstanses
End Sub




' управление расположением окон - иконки
Private Sub mnuArrangeIcon_Click()
  Me.Arrange vbArrangeIcons
End Sub

' управление расположением окон - каскад
Private Sub mnuCascade_Click()
  Me.Arrange vbCascade
End Sub

' управление расположением окон - горизонтально
Private Sub mnuTileHor_Click()
  Me.Arrange vbTileHorizontal
End Sub

' управление расположением окон - вертикально
Private Sub mnuTileVert_Click()
  Me.Arrange vbTileVertical
End Sub


'открыть форму редактирования объекта. или найти и показать уже открытую форму
Private Sub OpenForm(o As Object)
  Dim t As Form
  For Each t In Forms
    If t.Caption = o.Name Then
      t.WindowState = vbNormal
      t.ZOrder 0
      t.Show
      Me.MousePointer = vbNormal
      Exit Sub
    End If
  Next
  
  Dim f As frmObj
  Set f = New frmObj
  f.INIT o
  f.Show
  

End Sub










'открытие журанла отгрузки
Private Sub mnuAllITTOUT_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuAllITTOUT Is Nothing Then
      Set jfmnuAllITTOUT = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6EE193F7-B45F-4D5E-A6E0-391C006CB646}")
      Manager.LockInstanceObject journal.id
      Set jfmnuAllITTOUT.jv.journal = journal
      jfmnuAllITTOUT.jv.OpenModal = False
      jfmnuAllITTOUT.Caption = "Отгрузка - все состояния"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmITTOUT
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
    jfmnuAllITTOUT.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
      jfmnuAllITTOUT.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuAllITTOUT.Show
    jfmnuAllITTOUT.WindowState = 0
    jfmnuAllITTOUT.ZOrder 0
End Sub

'фильтр журанла отгрузки
Private Sub jfmnuAllITTOUT_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOUT
    Dim f As String
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
    jfmnuAllITTOUT.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'создание нового догкумента на отгрузку
Private Sub jfmnuAllITTOUT_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOUT", "Отгрузка" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла отгрузки - оформляется
Private Sub mnuITTOUT_1_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTOUT_1 Is Nothing Then
      Set jfmnuITTOUT_1 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6EE193F7-B45F-4D5E-A6E0-391C006CB646}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTOUT_1.jv.journal = journal
      jfmnuITTOUT_1.jv.OpenModal = False
      jfmnuITTOUT_1.Caption = "Отгрузка :Оформляется"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{CDCAFF7F-B013-40AF-BE61-1A27E35DB946}'"
    jfmnuITTOUT_1.jv.Filter.Add "AUTOITTOUT_DEF", f
    Dim fltr As frmITTOUT
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
    jfmnuITTOUT_1.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
      jfmnuITTOUT_1.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTOUT_1.Show
    jfmnuITTOUT_1.WindowState = 0
    jfmnuITTOUT_1.ZOrder 0
End Sub

'фильтр журанла отгрузки - оформляется
Private Sub jfmnuITTOUT_1_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOUT
    Dim f As String
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{CDCAFF7F-B013-40AF-BE61-1A27E35DB946}'"
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
    jfmnuITTOUT_1.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла отгрузки - оформляется
Private Sub jfmnuITTOUT_1_OnClearFilter()
   jfmnuITTOUT_1.jv.Filter.Add "AUTOITTOUT_DEF", " INTSANCEStatusID='{CDCAFF7F-B013-40AF-BE61-1A27E35DB946}'"
End Sub


'создание документа на отгрузку
Private Sub jfmnuITTOUT_1_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOUT", "Отгрузка" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла отгрузки - идет отгрузка
Private Sub mnuITTOUT_2_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTOUT_2 Is Nothing Then
      Set jfmnuITTOUT_2 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6EE193F7-B45F-4D5E-A6E0-391C006CB646}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTOUT_2.jv.journal = journal
      jfmnuITTOUT_2.jv.OpenModal = False
      jfmnuITTOUT_2.Caption = "Отгрузка :Идет отгрузка"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{70853C28-84B5-434E-8413-52DF8FBBB49B}'"
    jfmnuITTOUT_2.jv.Filter.Add "AUTOITTOUT_DEF", f
    Dim fltr As frmITTOUT
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
    jfmnuITTOUT_2.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
      jfmnuITTOUT_2.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTOUT_2.Show
    jfmnuITTOUT_2.WindowState = 0
    jfmnuITTOUT_2.ZOrder 0
End Sub

'фильтр журанла отгрузки - идет отгрузка
Private Sub jfmnuITTOUT_2_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOUT
    Dim f As String
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{70853C28-84B5-434E-8413-52DF8FBBB49B}'"
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
    jfmnuITTOUT_2.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

' сброс фильтра журанла отгрузки - идет отгрузка
Private Sub jfmnuITTOUT_2_OnClearFilter()
   jfmnuITTOUT_2.jv.Filter.Add "AUTOITTOUT_DEF", " INTSANCEStatusID='{70853C28-84B5-434E-8413-52DF8FBBB49B}'"
End Sub


'создавние документа - отгрузка
Private Sub jfmnuITTOUT_2_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOUT", "Отгрузка" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла отгрузки - обработка завершена
Private Sub mnuITTOUT_3_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTOUT_3 Is Nothing Then
      Set jfmnuITTOUT_3 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6EE193F7-B45F-4D5E-A6E0-391C006CB646}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTOUT_3.jv.journal = journal
      jfmnuITTOUT_3.jv.OpenModal = False
      jfmnuITTOUT_3.Caption = "Отгрузка :Обработка завершена"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{2CDDB562-63D7-483E-B95E-B579A9096CCC}'"
    jfmnuITTOUT_3.jv.Filter.Add "AUTOITTOUT_DEF", f
    Dim fltr As frmITTOUT
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
    jfmnuITTOUT_3.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
      jfmnuITTOUT_3.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTOUT_3.Show
    jfmnuITTOUT_3.WindowState = 0
    jfmnuITTOUT_3.ZOrder 0
End Sub

'фильтр журанла отгрузки - обработка завершена
Private Sub jfmnuITTOUT_3_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOUT
    Dim f As String
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{2CDDB562-63D7-483E-B95E-B579A9096CCC}'"
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
    jfmnuITTOUT_3.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла отгрузки - обработка завершена
Private Sub jfmnuITTOUT_3_OnClearFilter()
   jfmnuITTOUT_3.jv.Filter.Add "AUTOITTOUT_DEF", " INTSANCEStatusID='{2CDDB562-63D7-483E-B95E-B579A9096CCC}'"
End Sub

' создание документа на отгрузку
Private Sub jfmnuITTOUT_3_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOUT", "Отгрузка" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла отгрузки - Отгрузка завершена
Private Sub mnuITTOUT_4_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTOUT_4 Is Nothing Then
      Set jfmnuITTOUT_4 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6EE193F7-B45F-4D5E-A6E0-391C006CB646}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTOUT_4.jv.journal = journal
      jfmnuITTOUT_4.jv.OpenModal = False
      jfmnuITTOUT_4.Caption = "Отгрузка :Отгрузка завершена"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}'"
    jfmnuITTOUT_4.jv.Filter.Add "AUTOITTOUT_DEF", f
    Dim fltr As frmITTOUT
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
    jfmnuITTOUT_4.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
      jfmnuITTOUT_4.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTOUT_4.Show
    jfmnuITTOUT_4.WindowState = 0
    jfmnuITTOUT_4.ZOrder 0
End Sub

'фильтр журанла отгрузки - Отгрузка завершена
Private Sub jfmnuITTOUT_4_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOUT
    Dim f As String
    Set fltr = New frmITTOUT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}'"
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblShipOrder.Value = vbChecked Then
        f = f & " and ITTOUT_DEF_ShipOrder_ID='" & fltr.txtShipOrder.Tag & "'"
      End If
    jfmnuITTOUT_4.jv.Filter.Add "AUTOITTOUT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

' сброс фильтра журанла отгрузки - Отгрузка завершена
Private Sub jfmnuITTOUT_4_OnClearFilter()
   jfmnuITTOUT_4.jv.Filter.Add "AUTOITTOUT_DEF", " INTSANCEStatusID='{881CBAAC-BE9D-4216-AB25-ED3B2761F82F}'"
End Sub

' создание документа на отгрузку
Private Sub jfmnuITTOUT_4_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOUT", "Отгрузка" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub



'открытие журанла паллеты
Private Sub mnuAllITTPL_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuAllITTPL Is Nothing Then
      Set jfmnuAllITTPL = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6345F83E-3D6C-4782-B165-51AEADB4D040}")
      Manager.LockInstanceObject journal.id
      Set jfmnuAllITTPL.jv.journal = journal
      jfmnuAllITTPL.jv.OpenModal = False
      jfmnuAllITTPL.Caption = "Палетта - все состояния"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmITTPL
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
    jfmnuAllITTPL.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
      jfmnuAllITTPL.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuAllITTPL.Show
    jfmnuAllITTPL.WindowState = 0
    jfmnuAllITTPL.ZOrder 0
End Sub

'фильтр журанла паллеты
Private Sub jfmnuAllITTPL_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTPL
    Dim f As String
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
    jfmnuAllITTPL.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'создание документа - паллета
Private Sub jfmnuAllITTPL_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTPL", "Палетта" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла паллеты - пустая
Private Sub mnuITTPL_1_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTPL_1 Is Nothing Then
      Set jfmnuITTPL_1 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6345F83E-3D6C-4782-B165-51AEADB4D040}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTPL_1.jv.journal = journal
      jfmnuITTPL_1.jv.OpenModal = False
      jfmnuITTPL_1.Caption = "Палетта :Пустая"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{E9BFB749-A606-4DEF-A429-07D636F108C6}'"
    jfmnuITTPL_1.jv.Filter.Add "AUTOITTPL_DEF", f
    Dim fltr As frmITTPL
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
    jfmnuITTPL_1.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
      jfmnuITTPL_1.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTPL_1.Show
    jfmnuITTPL_1.WindowState = 0
    jfmnuITTPL_1.ZOrder 0
End Sub

'фильтр журанла паллеты - пустая
Private Sub jfmnuITTPL_1_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTPL
    Dim f As String
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{E9BFB749-A606-4DEF-A429-07D636F108C6}'"
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
    jfmnuITTPL_1.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла паллеты - пустая
Private Sub jfmnuITTPL_1_OnClearFilter()
   jfmnuITTPL_1.jv.Filter.Add "AUTOITTPL_DEF", " INTSANCEStatusID='{E9BFB749-A606-4DEF-A429-07D636F108C6}'"
End Sub

' создание документа - паллета
Private Sub jfmnuITTPL_1_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTPL", "Палетта" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла паллеты - Взвешена
Private Sub mnuITTPL_2_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTPL_2 Is Nothing Then
      Set jfmnuITTPL_2 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6345F83E-3D6C-4782-B165-51AEADB4D040}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTPL_2.jv.journal = journal
      jfmnuITTPL_2.jv.OpenModal = False
      jfmnuITTPL_2.Caption = "Палетта :Взвешена"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{6FDCC60F-8C10-47E3-BB36-110C49EF2144}'"
    jfmnuITTPL_2.jv.Filter.Add "AUTOITTPL_DEF", f
    Dim fltr As frmITTPL
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
    jfmnuITTPL_2.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
      jfmnuITTPL_2.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTPL_2.Show
    jfmnuITTPL_2.WindowState = 0
    jfmnuITTPL_2.ZOrder 0
End Sub

'фильтр журанла паллеты - Взвешена
Private Sub jfmnuITTPL_2_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTPL
    Dim f As String
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{6FDCC60F-8C10-47E3-BB36-110C49EF2144}'"
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
    jfmnuITTPL_2.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла паллеты - Взвешена
Private Sub jfmnuITTPL_2_OnClearFilter()
   jfmnuITTPL_2.jv.Filter.Add "AUTOITTPL_DEF", " INTSANCEStatusID='{6FDCC60F-8C10-47E3-BB36-110C49EF2144}'"
End Sub

' создание документа паллета
Private Sub jfmnuITTPL_2_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTPL", "Палетта" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла паллеты - На складе с грузом
Private Sub mnuITTPL_3_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTPL_3 Is Nothing Then
      Set jfmnuITTPL_3 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6345F83E-3D6C-4782-B165-51AEADB4D040}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTPL_3.jv.journal = journal
      jfmnuITTPL_3.jv.OpenModal = False
      jfmnuITTPL_3.Caption = "Палетта :На складе с грузом"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}'"
    jfmnuITTPL_3.jv.Filter.Add "AUTOITTPL_DEF", f
    Dim fltr As frmITTPL
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
    jfmnuITTPL_3.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
      jfmnuITTPL_3.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTPL_3.Show
    jfmnuITTPL_3.WindowState = 0
    jfmnuITTPL_3.ZOrder 0
End Sub

'фильтр журанла паллеты - На складе с грузом
Private Sub jfmnuITTPL_3_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTPL
    Dim f As String
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}'"
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
    jfmnuITTPL_3.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

' сброс фильтра журанла паллеты - На складе с грузом
Private Sub jfmnuITTPL_3_OnClearFilter()
   jfmnuITTPL_3.jv.Filter.Add "AUTOITTPL_DEF", " INTSANCEStatusID='{93E3DE6D-AB8D-48A6-84FD-152BF63FB14C}'"
End Sub

' создание документа - паллета
Private Sub jfmnuITTPL_3_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTPL", "Палетта" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла паллеты - Списана
Private Sub mnuITTPL_4_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTPL_4 Is Nothing Then
      Set jfmnuITTPL_4 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6345F83E-3D6C-4782-B165-51AEADB4D040}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTPL_4.jv.journal = journal
      jfmnuITTPL_4.jv.OpenModal = False
      jfmnuITTPL_4.Caption = "Палетта :Списана"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{588C5203-1E59-408E-92A1-B3DFED8C19FA}'"
    jfmnuITTPL_4.jv.Filter.Add "AUTOITTPL_DEF", f
    Dim fltr As frmITTPL
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
    jfmnuITTPL_4.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
      jfmnuITTPL_4.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTPL_4.Show
    jfmnuITTPL_4.WindowState = 0
    jfmnuITTPL_4.ZOrder 0
End Sub


'фильтр журанла паллеты - Списана
Private Sub jfmnuITTPL_4_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTPL
    Dim f As String
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{588C5203-1E59-408E-92A1-B3DFED8C19FA}'"
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
    jfmnuITTPL_4.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub


'сброс фильтра журанла паллеты - Списана
Private Sub jfmnuITTPL_4_OnClearFilter()
   jfmnuITTPL_4.jv.Filter.Add "AUTOITTPL_DEF", " INTSANCEStatusID='{588C5203-1E59-408E-92A1-B3DFED8C19FA}'"
End Sub

'создание документа - паллета
Private Sub jfmnuITTPL_4_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTPL", "Палетта" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


'открытие журанла паллеты - Отправлена с грузом
Private Sub mnuITTPL_5_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTPL_5 Is Nothing Then
      Set jfmnuITTPL_5 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{6345F83E-3D6C-4782-B165-51AEADB4D040}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTPL_5.jv.journal = journal
      jfmnuITTPL_5.jv.OpenModal = False
      jfmnuITTPL_5.Caption = "Палетта :Отправлена с грузом"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}'"
    jfmnuITTPL_5.jv.Filter.Add "AUTOITTPL_DEF", f
    Dim fltr As frmITTPL
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
    jfmnuITTPL_5.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
      jfmnuITTPL_5.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTPL_5.Show
    jfmnuITTPL_5.WindowState = 0
    jfmnuITTPL_5.ZOrder 0
End Sub

'фильтр журанла паллеты - Отправлена с грузом

Private Sub jfmnuITTPL_5_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTPL
    Dim f As String
    Set fltr = New frmITTPL
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}'"
      If fltr.lblCurrentWeightBrutto_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto<=" & Val(fltr.txtCurrentWeightBrutto_LE.Text)
      End If
      If fltr.lblPackageWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight>=" & Val(fltr.txtPackageWeight_GE.Text)
      End If
      If fltr.lblCode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Code like '%" & fltr.txtCode.Text & "%'"
      End If
      If fltr.lblTheNumber_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber<=" & Val(fltr.txtTheNumber_LE.Text)
      End If
      If fltr.lblCorePalette_ID_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID>=" & Val(fltr.txtCorePalette_ID_GE.Text)
      End If
      If fltr.lblWeight_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight>=" & Val(fltr.txtWeight_GE.Text)
      End If
      If fltr.lblPrivatePalet.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PrivatePalet='" & fltr.cmbPrivatePalet.Text & "'"
      End If
      If fltr.lblCaliberQuantity_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity>=" & Val(fltr.txtCaliberQuantity_GE.Text)
      End If
      If fltr.lblCaliberQuantity_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CaliberQuantity<=" & Val(fltr.txtCaliberQuantity_LE.Text)
      End If
      If fltr.lblPalKode.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PalKode like '%" & fltr.txtPalKode.Text & "%'"
      End If
      If fltr.lblCurrentPosition.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentPosition like '%" & fltr.txtCurrentPosition.Text & "%'"
      End If
      If fltr.lblCorePalette_ID_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CorePalette_ID<=" & Val(fltr.txtCorePalette_ID_LE.Text)
      End If
      If fltr.lblCurrentWeightBrutto_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentWeightBrutto>=" & Val(fltr.txtCurrentWeightBrutto_GE.Text)
      End If
      If fltr.lblPackageWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_PackageWeight<=" & Val(fltr.txtPackageWeight_LE.Text)
      End If
      If fltr.lblPltype.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Pltype_ID='" & fltr.txtPltype.Tag & "'"
      End If
      If fltr.lblWDate_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate>=" & MakeMSSQLDate(fltr.dtpWDate_GE.Value)
      End If
      If fltr.lblWDate_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_WDate<=" & MakeMSSQLDate(fltr.dtpWDate_LE.Value)
      End If
      If fltr.lblCurrentGood.Value = vbChecked Then
        f = f & " and ITTPL_DEF_CurrentGood_ID='" & fltr.txtCurrentGood.Tag & "'"
      End If
      If fltr.lblTheNumber_GE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_TheNumber>=" & Val(fltr.txtTheNumber_GE.Text)
      End If
      If fltr.lblWeight_LE.Value = vbChecked Then
        f = f & " and ITTPL_DEF_Weight<=" & Val(fltr.txtWeight_LE.Text)
      End If
    jfmnuITTPL_5.jv.Filter.Add "AUTOITTPL_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub


'сброс фильтра журанла паллеты - Отправлена с грузом
Private Sub jfmnuITTPL_5_OnClearFilter()
   jfmnuITTPL_5.jv.Filter.Add "AUTOITTPL_DEF", " INTSANCEStatusID='{7BD977D0-0EF9-4F0D-B047-E409BB1616CA}'"
End Sub

' создание документа паллета
Private Sub jfmnuITTPL_5_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTPL", "Палетта" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub



'открытие журанла приемка
Private Sub mnuAllITTIN_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuAllITTIN Is Nothing Then
      Set jfmnuAllITTIN = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{5AC03393-7686-4423-AD74-98673546FBA3}")
      Manager.LockInstanceObject journal.id
      Set jfmnuAllITTIN.jv.journal = journal
      jfmnuAllITTIN.jv.OpenModal = False
      jfmnuAllITTIN.Caption = "Приемка груза - все состояния"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmITTIN
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
    jfmnuAllITTIN.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
      jfmnuAllITTIN.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuAllITTIN.Show
    jfmnuAllITTIN.WindowState = 0
    jfmnuAllITTIN.ZOrder 0
End Sub

'фильтр журанла приемка
Private Sub jfmnuAllITTIN_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTIN
    Dim f As String
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
    jfmnuAllITTIN.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'создание документа - приемка
Private Sub jfmnuAllITTIN_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTIN", "Приемка груза" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла приемка - Оформляется
Private Sub mnuITTIN_1_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTIN_1 Is Nothing Then
      Set jfmnuITTIN_1 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{5AC03393-7686-4423-AD74-98673546FBA3}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTIN_1.jv.journal = journal
      jfmnuITTIN_1.jv.OpenModal = False
      jfmnuITTIN_1.Caption = "Приемка груза :Оформляется"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{49A919F7-94A6-49DE-9280-1EEAC973647B}'"
    jfmnuITTIN_1.jv.Filter.Add "AUTOITTIN_DEF", f
    Dim fltr As frmITTIN
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
    jfmnuITTIN_1.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
      jfmnuITTIN_1.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTIN_1.Show
    jfmnuITTIN_1.WindowState = 0
    jfmnuITTIN_1.ZOrder 0
End Sub

' фильтр журанла приемка - Оформляется
Private Sub jfmnuITTIN_1_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTIN
    Dim f As String
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{49A919F7-94A6-49DE-9280-1EEAC973647B}'"
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
    jfmnuITTIN_1.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'фильтр журанла приемка - Оформляется
Private Sub jfmnuITTIN_1_OnClearFilter()
   jfmnuITTIN_1.jv.Filter.Add "AUTOITTIN_DEF", " INTSANCEStatusID='{49A919F7-94A6-49DE-9280-1EEAC973647B}'"
End Sub

' создание документа приемка
Private Sub jfmnuITTIN_1_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTIN", "Приемка груза" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла приемка - Приемка завершена
Private Sub mnuITTIN_2_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTIN_2 Is Nothing Then
      Set jfmnuITTIN_2 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{5AC03393-7686-4423-AD74-98673546FBA3}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTIN_2.jv.journal = journal
      jfmnuITTIN_2.jv.OpenModal = False
      jfmnuITTIN_2.Caption = "Приемка груза :Приемка завершена"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}'"
    jfmnuITTIN_2.jv.Filter.Add "AUTOITTIN_DEF", f
    Dim fltr As frmITTIN
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
    jfmnuITTIN_2.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
      jfmnuITTIN_2.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTIN_2.Show
    jfmnuITTIN_2.WindowState = 0
    jfmnuITTIN_2.ZOrder 0
End Sub

'фильтр журанла приемка - Приемка завершена
Private Sub jfmnuITTIN_2_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTIN
    Dim f As String
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}'"
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
    jfmnuITTIN_2.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла приемка - Приемка завершена
Private Sub jfmnuITTIN_2_OnClearFilter()
   jfmnuITTIN_2.jv.Filter.Add "AUTOITTIN_DEF", " INTSANCEStatusID='{E3728A5B-6B62-48BF-9E5A-D4F0BCBFC75B}'"
End Sub

' создание документа - приемка
Private Sub jfmnuITTIN_2_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTIN", "Приемка груза" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла приемка -Идет приемка
Private Sub mnuITTIN_3_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTIN_3 Is Nothing Then
      Set jfmnuITTIN_3 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{5AC03393-7686-4423-AD74-98673546FBA3}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTIN_3.jv.journal = journal
      jfmnuITTIN_3.jv.OpenModal = False
      jfmnuITTIN_3.Caption = "Приемка груза :Идет приемка"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}'"
    jfmnuITTIN_3.jv.Filter.Add "AUTOITTIN_DEF", f
    Dim fltr As frmITTIN
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
    jfmnuITTIN_3.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
      jfmnuITTIN_3.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTIN_3.Show
    jfmnuITTIN_3.WindowState = 0
    jfmnuITTIN_3.ZOrder 0
End Sub

'фильтр журанла приемка -Идет приемка
Private Sub jfmnuITTIN_3_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTIN
    Dim f As String
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}'"
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
    jfmnuITTIN_3.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub
'сброс фильтра журанла приемка -Идет приемка
Private Sub jfmnuITTIN_3_OnClearFilter()
   jfmnuITTIN_3.jv.Filter.Add "AUTOITTIN_DEF", " INTSANCEStatusID='{EB3A7D03-EB3F-4541-AD93-D55C92BE02AC}'"
End Sub

' создание документа - приемка
Private Sub jfmnuITTIN_3_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTIN", "Приемка груза" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла приемка -Приемка обработана
Private Sub mnuITTIN_4_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTIN_4 Is Nothing Then
      Set jfmnuITTIN_4 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{5AC03393-7686-4423-AD74-98673546FBA3}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTIN_4.jv.journal = journal
      jfmnuITTIN_4.jv.OpenModal = False
      jfmnuITTIN_4.Caption = "Приемка груза :Приемка обработана"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{E8BA9909-6680-4B2C-B446-F58EF91DCD17}'"
    jfmnuITTIN_4.jv.Filter.Add "AUTOITTIN_DEF", f
    Dim fltr As frmITTIN
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
    jfmnuITTIN_4.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
      jfmnuITTIN_4.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTIN_4.Show
    jfmnuITTIN_4.WindowState = 0
    jfmnuITTIN_4.ZOrder 0
End Sub

'фильтр журанла приемка -Приемка обработана
Private Sub jfmnuITTIN_4_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTIN
    Dim f As String
    Set fltr = New frmITTIN
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{E8BA9909-6680-4B2C-B446-F58EF91DCD17}'"
      If fltr.lblThePartyRule.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ThePartyRule_ID='" & fltr.txtThePartyRule.Tag & "'"
      End If
      If fltr.lblProcessDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate<=" & MakeMSSQLDate(fltr.dtpProcessDate_LE.Value)
      End If
      If fltr.lblTrack_time_in_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in>=" & MakeMSSQLDate(fltr.dtpTrack_time_in_GE.Value)
      End If
      If fltr.lblProcessDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_ProcessDate>=" & MakeMSSQLDate(fltr.dtpProcessDate_GE.Value)
      End If
      If fltr.lblTranspNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TranspNumber like '%" & fltr.txtTranspNumber.Text & "%'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TheClient_ID='" & fltr.txtTheClient.Tag & "'"
      End If
      If fltr.lblTTN.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTN like '%" & fltr.txtTTN.Text & "%'"
      End If
      If fltr.lblSupplier.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Supplier like '%" & fltr.txtSupplier.Text & "%'"
      End If
      If fltr.lbltrack_time_out_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out>=" & MakeMSSQLDate(fltr.dtptrack_time_out_GE.Value)
      End If
      If fltr.lblQryCode.Value = vbChecked Then
        f = f & " and ITTIN_DEF_QryCode_ID='" & fltr.txtQryCode.Tag & "'"
      End If
      If fltr.lbltemp_in_track_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track<=" & Val(fltr.txttemp_in_track_LE.Text)
      End If
      If fltr.lblTTNDate_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate<=" & MakeMSSQLDate(fltr.dtpTTNDate_LE.Value)
      End If
      If fltr.lblStampNumber.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampNumber like '%" & fltr.txtStampNumber.Text & "%'"
      End If
      If fltr.lblContainer.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Container like '%" & fltr.txtContainer.Text & "%'"
      End If
      If fltr.lblTrack_time_in_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_Track_time_in<=" & MakeMSSQLDate(fltr.dtpTrack_time_in_LE.Value)
      End If
      If fltr.lblTTNDate_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_TTNDate>=" & MakeMSSQLDate(fltr.dtpTTNDate_GE.Value)
      End If
      If fltr.lblStampStatus.Value = vbChecked Then
        f = f & " and ITTIN_DEF_StampStatus like '%" & fltr.txtStampStatus.Text & "%'"
      End If
      If fltr.lbltemp_in_track_GE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_temp_in_track>=" & Val(fltr.txttemp_in_track_GE.Text)
      End If
      If fltr.lbltrack_time_out_LE.Value = vbChecked Then
        f = f & " and ITTIN_DEF_track_time_out<=" & MakeMSSQLDate(fltr.dtptrack_time_out_LE.Value)
      End If
    jfmnuITTIN_4.jv.Filter.Add "AUTOITTIN_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла приемка -Приемка обработана
Private Sub jfmnuITTIN_4_OnClearFilter()
   jfmnuITTIN_4.jv.Filter.Add "AUTOITTIN_DEF", " INTSANCEStatusID='{E8BA9909-6680-4B2C-B446-F58EF91DCD17}'"
End Sub

' создание документа - приемка
Private Sub jfmnuITTIN_4_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTIN", "Приемка груза" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub



'открытие журанла услуги клиентов
Private Sub mnuITTCS_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTCS Is Nothing Then
      Set jfmnuITTCS = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{D02217FD-2C39-46A2-B88D-011F9FAC08CA}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTCS.jv.journal = journal
      jfmnuITTCS.jv.OpenModal = False
      jfmnuITTCS.Caption = "Услуги клиентов"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmITTCS
    Set fltr = New frmITTCS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblCLIENTCODE.Value = vbChecked Then
        f = f & " and ITTCS_DEF_CLIENTCODE_ID='" & fltr.txtCLIENTCODE.Tag & "'"
      End If
    jfmnuITTCS.jv.Filter.Add "AUTOITTCS_DEF", f
    End If
      jfmnuITTCS.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTCS.Show
    jfmnuITTCS.WindowState = 0
    jfmnuITTCS.ZOrder 0
End Sub

'фильтр журанла услуги клиентов
Private Sub jfmnuITTCS_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTCS
    Dim f As String
    Set fltr = New frmITTCS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblCLIENTCODE.Value = vbChecked Then
        f = f & " and ITTCS_DEF_CLIENTCODE_ID='" & fltr.txtCLIENTCODE.Tag & "'"
      End If
    jfmnuITTCS.jv.Filter.Add "AUTOITTCS_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub


'создание документа - услуги клиентов
Private Sub jfmnuITTCS_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTCS", "Услуги клиентов" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub



' открытие документа  - настройки системы
Private Sub mnuITTFN_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTFN")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTFN", "Настройки системы"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub

' открытие документа  - Операторы и кладовщики
Private Sub mnuITTOP_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTOP")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTOP", "Операторы и кладовщики"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub

' открытие документа  - Справочник
Private Sub mnuITTD_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTD")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTD", "Справочник"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub


' открытие документа  - Настройки оптмизатора
Private Sub mnuITTNO_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "ITTNO")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "ITTNO", "Настройки оптмизатора"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub




'открытие журанла Задание на перемещения
Private Sub mnuAllITTOPT_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuAllITTOPT Is Nothing Then
      Set jfmnuAllITTOPT = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{70D005EF-420B-4A40-8CDF-B4800FCD0F1D}")
      Manager.LockInstanceObject journal.id
      Set jfmnuAllITTOPT.jv.journal = journal
      jfmnuAllITTOPT.jv.OpenModal = False
      jfmnuAllITTOPT.Caption = "Задание на перемещения - все состояния"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmITTOPT
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
    jfmnuAllITTOPT.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
      jfmnuAllITTOPT.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuAllITTOPT.Show
    jfmnuAllITTOPT.WindowState = 0
    jfmnuAllITTOPT.ZOrder 0
End Sub

' фильтр журанла Задание на перемещения
Private Sub jfmnuAllITTOPT_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOPT
    Dim f As String
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
    jfmnuAllITTOPT.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

' создание документа - Задание на перемещения
Private Sub jfmnuAllITTOPT_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOPT", "Задание на перемещения" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла Задание на перемещения -Задание исполнено
Private Sub mnuITTOPT_1_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTOPT_1 Is Nothing Then
      Set jfmnuITTOPT_1 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{70D005EF-420B-4A40-8CDF-B4800FCD0F1D}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTOPT_1.jv.journal = journal
      jfmnuITTOPT_1.jv.OpenModal = False
      jfmnuITTOPT_1.Caption = "Задание на перемещения :Выполнено"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}'"
    jfmnuITTOPT_1.jv.Filter.Add "AUTOITTOPT_DEF", f
    Dim fltr As frmITTOPT
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
    jfmnuITTOPT_1.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
      jfmnuITTOPT_1.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTOPT_1.Show
    jfmnuITTOPT_1.WindowState = 0
    jfmnuITTOPT_1.ZOrder 0
End Sub

'фильтр журанла Задание на перемещения -Задание исполнено
Private Sub jfmnuITTOPT_1_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOPT
    Dim f As String
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}'"
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
    jfmnuITTOPT_1.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub
'сброс фильтра журанла Задание на перемещения -Задание исполнено
Private Sub jfmnuITTOPT_1_OnClearFilter()
   jfmnuITTOPT_1.jv.Filter.Add "AUTOITTOPT_DEF", " INTSANCEStatusID='{0A7FC795-E787-4D17-9689-96EFFF8F0D9D}'"
End Sub

' создание документа -Задание на перемещения
Private Sub jfmnuITTOPT_1_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOPT", "Задание на перемещения" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub





'открытие журанла Задание на перемещения -К выполнению
Private Sub mnuITTOPT_6_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTOPT_6 Is Nothing Then
      Set jfmnuITTOPT_6 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{70D005EF-420B-4A40-8CDF-B4800FCD0F1D}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTOPT_6.jv.journal = journal
      jfmnuITTOPT_6.jv.OpenModal = False
      jfmnuITTOPT_6.Caption = "Задание на перемещения :К выполнению"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{300483B2-1D94-4A33-8ADF-ABF32E72E57B}'"
    jfmnuITTOPT_6.jv.Filter.Add "AUTOITTOPT_DEF", f
    Dim fltr As frmITTOPT
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
    jfmnuITTOPT_6.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
      jfmnuITTOPT_6.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTOPT_6.Show
    jfmnuITTOPT_6.WindowState = 0
    jfmnuITTOPT_6.ZOrder 0
End Sub

'фильтр журанла Задание на перемещения -Оформлено
Private Sub jfmnuITTOPT_6_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOPT
    Dim f As String
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{300483B2-1D94-4A33-8ADF-ABF32E72E57B}'"
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
    jfmnuITTOPT_6.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла Задание на перемещения -Оформлено
Private Sub jfmnuITTOPT_6_OnClearFilter()
   jfmnuITTOPT_6.jv.Filter.Add "AUTOITTOPT_DEF", " INTSANCEStatusID='{300483B2-1D94-4A33-8ADF-ABF32E72E57B}'"
End Sub

' создание документа -Задание на перемещения
Private Sub jfmnuITTOPT_6_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOPT", "Задание на перемещения" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

'открытие журанла Задание на перемещения -Создан
Private Sub mnuITTOPT_7_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuITTOPT_7 Is Nothing Then
      Set jfmnuITTOPT_7 = New frmJournalShow2
      Set journal = Manager.GetInstanceObject("{70D005EF-420B-4A40-8CDF-B4800FCD0F1D}")
      Manager.LockInstanceObject journal.id
      Set jfmnuITTOPT_7.jv.journal = journal
      jfmnuITTOPT_7.jv.OpenModal = False
      jfmnuITTOPT_7.Caption = "Задание на перемещения :Создан"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{C861FA15-0DF6-42D4-BCE9-2B38C3E6C0CB}'"
    jfmnuITTOPT_7.jv.Filter.Add "AUTOITTOPT_DEF", f
    Dim fltr As frmITTOPT
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
    jfmnuITTOPT_7.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
      jfmnuITTOPT_7.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuITTOPT_7.Show
    jfmnuITTOPT_7.WindowState = 0
    jfmnuITTOPT_7.ZOrder 0
End Sub

'фильтр журанла Задание на перемещения -Оформляется
Private Sub jfmnuITTOPT_7_OnFilter(UseDefault As Boolean)
    Dim fltr As frmITTOPT
    Dim f As String
    Set fltr = New frmITTOPT
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{C861FA15-0DF6-42D4-BCE9-2B38C3E6C0CB}'"
      If fltr.lblFactory.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_Factory like '%" & fltr.txtFactory.Text & "%'"
      End If
      If fltr.lblOPtDate_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate<=" & MakeMSSQLDate(fltr.dtpOPtDate_LE.Value)
      End If
      If fltr.lblgood.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_good like '%" & fltr.txtgood.Text & "%'"
      End If
      If fltr.lblmade_country.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_made_country like '%" & fltr.txtmade_country.Text & "%'"
      End If
      If fltr.lblOptType.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OptType_ID='" & fltr.txtOptType.Tag & "'"
      End If
      If fltr.lblTheClient.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_TheClient like '%" & fltr.txtTheClient.Text & "%'"
      End If
      If fltr.lblKILL_NUMBER.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_KILL_NUMBER like '%" & fltr.txtKILL_NUMBER.Text & "%'"
      End If
      If fltr.lblDateToOptimize_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize>=" & MakeMSSQLDate(fltr.dtpDateToOptimize_GE.Value)
      End If
      If fltr.lblIsBrak.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_IsBrak like '%" & fltr.txtIsBrak.Text & "%'"
      End If
      If fltr.lblOPtDate_GE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_OPtDate>=" & MakeMSSQLDate(fltr.dtpOPtDate_GE.Value)
      End If
      If fltr.lblDateToOptimize_LE.Value = vbChecked Then
        f = f & " and ITTOPT_DEF_DateToOptimize<=" & MakeMSSQLDate(fltr.dtpDateToOptimize_LE.Value)
      End If
    jfmnuITTOPT_7.jv.Filter.Add "AUTOITTOPT_DEF", f
    End If
    Unload fltr
    UseDefault = False
End Sub

'сброс фильтра журанла Задание на перемещения -Оформляется
Private Sub jfmnuITTOPT_7_OnClearFilter()
   jfmnuITTOPT_7.jv.Filter.Add "AUTOITTOPT_DEF", " INTSANCEStatusID='{C861FA15-0DF6-42D4-BCE9-2B38C3E6C0CB}'"
End Sub

' создание документа -Задание на перемещения
Private Sub jfmnuITTOPT_7_OnAdd(usedefaut As Boolean, Refesh As Boolean) ' Обработка события Добавить документ для окна журнала
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "ITTOPT", "Задание на перемещения" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub













'выгрузка объектов по завершению работы
Private Sub UnloadObjects()

On Error Resume Next

'  выгрузка форм журналов
Unload jfmnuAllITTOPT
Set jfmnuAllITTOPT = Nothing

Unload jfmnuITTOPT_1
Set jfmnuITTOPT_1 = Nothing


Unload jfmnuITTOPT_6
Set jfmnuITTOPT_6 = Nothing

Unload jfmnuITTOPT_7
Set jfmnuITTOPT_7 = Nothing



Unload jfmnuAllITTOUT
Set jfmnuAllITTOUT = Nothing

Unload jfmnuITTOUT_1
Set jfmnuITTOUT_1 = Nothing

Unload jfmnuITTOUT_2
Set jfmnuITTOUT_2 = Nothing

Unload jfmnuITTOUT_3
Set jfmnuITTOUT_3 = Nothing

Unload jfmnuITTOUT_4
Set jfmnuITTOUT_4 = Nothing


Unload jfmnuAllITTPL
Set jfmnuAllITTPL = Nothing

Unload jfmnuITTPL_1
Set jfmnuITTPL_1 = Nothing

Unload jfmnuITTPL_2
Set jfmnuITTPL_2 = Nothing

Unload jfmnuITTPL_3
Set jfmnuITTPL_3 = Nothing

Unload jfmnuITTPL_4
Set jfmnuITTPL_4 = Nothing

Unload jfmnuITTPL_5
Set jfmnuITTPL_5 = Nothing


Unload jfmnuAllITTIN
Set jfmnuAllITTIN = Nothing

Unload jfmnuITTIN_1
Set jfmnuITTIN_1 = Nothing

Unload jfmnuITTIN_2
Set jfmnuITTIN_2 = Nothing

Unload jfmnuITTIN_3
Set jfmnuITTIN_3 = Nothing

Unload jfmnuITTIN_4
Set jfmnuITTIN_4 = Nothing

Unload jfmnuITTCS
Set jfmnuITTCS = Nothing

Unload jfmnuITTPR
Set jfmnuITTPR = Nothing



Unload jfmnuAllITTOPT

Unload jfmnuITTOPT_1


Unload jfmnuITTOPT_6

Unload jfmnuITTOPT_7



Set jfmnuAllITTOPT = Nothing

Set jfmnuITTOPT_1 = Nothing


Set jfmnuITTOPT_6 = Nothing

Set jfmnuITTOPT_7 = Nothing

' выгрузка форм отчетов
Set repShowOL = Nothing
Set repShowSRVOUT = Nothing
Set repShowKL = Nothing
Set repShowSRVIN = Nothing
Set repShowINEPL = Nothing
Set RptShowSRVALL = Nothing
Set RptStickers = Nothing
Set RptWrongLocation = Nothing
Set repShowMoves = Nothing
Set RptNedostacha = Nothing
Set RptActVes = Nothing
Set RptVimorozka = Nothing
Set RptVimorozka2 = Nothing
Set RptHran = Nothing
Set RptStok103 = Nothing
Set RptOtobrano = Nothing



End Sub

