VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITTD_COUNTRY_admi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 35
Option Explicit
'�������������� ���������� ������ ��� ����������� �����

'�������� ��� ������� �� ������ ���������  � ����������� �����
'Parameters:
'[IN][OUT]   TypeObject , ��� ���������: Object - ��������,
'[IN][OUT]   RowObject , ��� ���������: Object - ������,
'[IN][OUT]   GuiControlObject , ��� ���������: Object - �������,
'[IN][OUT]   UseDefault , ��� ���������: Boolean - ������������ ������� �� ���������,
'[IN][OUT]   MakeRefresh , ��� ���������: Boolean  - �������� �������
'Example:
'  call me.OnRun(...���������...)
Public Sub OnRun(ByRef TypeObject As Object, ByRef RowObject As Object, ByRef GuiControlObject As Object, ByRef UseDefault As Boolean, ByRef MakeRefresh As Boolean)
Attribute OnRun.VB_HelpID = 40

  UseDefault = False
  Dim f As frmCopyCountry
  Set f = New frmCopyCountry
  Set f.dic = TypeObject
  f.Show vbModal
  Set f = Nothing
  MakeRefresh = True
  
End Sub
