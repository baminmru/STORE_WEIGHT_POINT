VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Caption As String
Public ReportPath As String
Public ReportSource As String
Public ReportFilter As String
Public PrinterName As String

Private m_Formulas As Formulas
Private frm As frmReport


Public Function Formulas() As Formulas
  Set Formulas = m_Formulas
End Function

Private Sub Class_Initialize()
  Set m_Formulas = New Formulas
  Set frm = New frmReport
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Unload frm
  Set m_Formulas = Nothing
  Set frm = Nothing
End Sub



Public Sub Run(Optional ByVal Modal As Boolean = False, Optional ByVal conn As ADODB.Connection = Nothing)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
   If conn Is Nothing Then
  
    Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  Else
    If ReportFilter <> "" Then
      Set adoRs = conn.Execute("select * from " & ReportSource & " where  " & ReportFilter)
    Else
     Set adoRs = conn.Execute("select * from " & ReportSource)
    End If
  End If
  
  'Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  
  rpt.Database.SetDataSource adoRs
  
  If PrinterName <> "" Then
    Dim P As Printer
    For Each P In Printers
      If P.DeviceName = PrinterName Then
        rpt.SelectPrinter P.DriverName, P.DeviceName, P.port
        Exit For
      End If
    Next
  End If
  rpt.Database.Verify
  frm.Caption = Caption
  Set frm.rpt = rpt
  frm.CRViewer1.ReportSource = rpt
  frm.CRViewer1.ViewReport
  If Modal Then
    frm.Show vbModal
  Else
    frm.Show
  End If
  
  frm.WindowState = 0
  Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub


Public Sub RunDirectRS(ByVal rs As ADODB.Recordset, Optional ByVal Modal As Boolean = False)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
  
  rpt.Database.SetDataSource rs
  
  If PrinterName <> "" Then
    Dim P As Printer
    For Each P In Printers
      If P.DeviceName = PrinterName Then
        rpt.SelectPrinter P.DriverName, P.DeviceName, P.port
        Exit For
      End If
    Next
  End If
  rpt.Database.Verify
  frm.Caption = Caption
  Set frm.rpt = rpt
  frm.CRViewer1.ReportSource = rpt
  frm.CRViewer1.ViewReport
  If Modal Then
    frm.Show vbModal
  Else
    frm.Show
  End If
  
  frm.WindowState = 0
  Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub


Public Sub PrintOut(Optional ByVal conn As ADODB.Connection = Nothing)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
  If conn Is Nothing Then
  
    Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  Else
    If ReportFilter <> "" Then
      Set adoRs = conn.Execute("select * from " & ReportSource & " where  " & ReportFilter)
    Else
     Set adoRs = conn.Execute("select * from " & ReportSource)
    End If
  End If
  
  rpt.Database.SetDataSource adoRs
  
  
  rpt.Database.Verify
  
  If PrinterName <> "" Then
    Dim P As Printer
    For Each P In Printers
      If P.DeviceName = PrinterName Then
        
'        Set Printer = P
'        Printer.Orientation = 2
'        Dim w As Double, h As Double
'        Printer.PaperSize = vbPRPSA4
'        h = Printer.Height
'        w = Printer.Width
'
'        Printer.PaperSize = 256
'
'        Printer.Width = h
'        Printer.Height = w
 ''       rpt.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.port
        rpt.SelectPrinter P.DriverName, P.DeviceName, P.port
        
        Exit For
      End If
    Next
  End If
  
  rpt.PrintOut False, 1
  
  frm.Caption = Caption
  frm.Show
  frm.CRViewer1.ReportSource = rpt
  frm.CRViewer1.PrintReport
  frm.WindowState = 0
  Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub



Public Sub ExportPDF(ByVal FileName As String, Optional ByVal conn As ADODB.Connection = Nothing)
  Dim crX As CRAXDRT.Application
  Dim adoRs As ADODB.Recordset
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
  
  If conn Is Nothing Then
  
    Set adoRs = Session.GetRows(ReportSource, "", "", ReportFilter)
  Else
    If ReportFilter <> "" Then
      Set adoRs = conn.Execute("select * from " & ReportSource & " where  " & ReportFilter)
    Else
     Set adoRs = conn.Execute("select * from " & ReportSource)
    End If
  End If
  
  rpt.Database.SetDataSource adoRs
    
  rpt.Database.Verify
  
    
  rpt.ExportOptions.PDFExportAllPages = True
  rpt.ExportOptions.PDFFirstPageNumber = 1
  rpt.ExportOptions.DestinationType = crEDTDiskFile
  rpt.ExportOptions.DiskFileName = FileName
  rpt.ExportOptions.FormatType = crEFTPortableDocFormat
  'rpt.ExportOptions.PromptForExportOptions
  rpt.EnableParameterPrompting = False
  rpt.DisplayProgressDialog = False
  rpt.Export False
  
  Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "Отчет"
End Sub

Public Sub ExportPDFDirectRS(ByVal FileName As String, ByVal adoRs As ADODB.Recordset)
  Dim crX As CRAXDRT.Application
 
  On Error GoTo bye
  Set crX = New CRAXDRT.Application
  Dim rpt As CRAXDRT.Report
  Set rpt = crX.OpenReport(ReportPath)
  rpt.DiscardSavedData
  
  Dim ff As FormulaFieldDefinition
  Dim fd As FormulaData
  For Each fd In Formulas
    Set ff = rpt.FormulaFields.GetItemByName(fd.Name)
    ff.Text = fd.Expression
  Next
   
  rpt.Database.SetDataSource adoRs
  
  rpt.Database.Verify
  
    
  rpt.ExportOptions.PDFExportAllPages = True
  rpt.ExportOptions.PDFFirstPageNumber = 1
  rpt.ExportOptions.DestinationType = crEDTDiskFile
  rpt.ExportOptions.DiskFileName = FileName
  rpt.ExportOptions.FormatType = crEFTPortableDocFormat
  rpt.EnableParameterPrompting = False
  rpt.DisplayProgressDialog = False
  rpt.Export False
  
  Exit Sub
bye:
  MsgBox err.Description, vbOKOnly + vbCritical, "Экспорт"
End Sub

