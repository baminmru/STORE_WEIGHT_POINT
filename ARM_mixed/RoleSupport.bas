Attribute VB_Name = "RoleSupport"
Attribute VB_HelpID = 1570
Option Explicit
' ��������� ���������������� �����


Public MyRole As ROLES.Application  'Object
Attribute MyRole.VB_VarHelpID = 1605

'�������� ����������� ������ ����
Public Enum RoleMenuStatus
  RoleMenuStatus_Unknown = 0
  RoleMenuStatus_Visible = 1
  RoleMenuStatus_Disabled = 2
  RoleMenuStatus_Hidden = 3
End Enum

'���������� ����� ��������� ���������
'Parameters:
'[IN][OUT]  Item , ��� ���������: Object - �������� ( Application),
'[IN][OUT]   NewStatus , ��� ���������: String  - ���������
'Returns:
' Boolean, ��������� ����������:
'   true  - �����
'   false - �����!
'See Also:
'  ARMID
'  CheckMenu
'  ChooseRole
'  GetDocumentMode
'  IsDocDenied
'  MyRole
'  RoleDocAllowDelete
'  RoleDocCanSwitchStatus
'Example:
' dim variable as Boolean
' variable = me.BeforeChangeStatus(...���������...)
Public Function BeforeChangeStatus(Item As Object, NewStatus As String) As Boolean
Attribute BeforeChangeStatus.VB_HelpID = 1580
  Dim logic As Object
  Dim result As Boolean
  result = True
  On Error Resume Next
  Set logic = CreateObject(Item.TypeName & "BST.BEFORESTATUS")
  If Not logic Is Nothing Then
    result = logic.Check(Item, NewStatus, MyUser, Item.TypeName)
    Set logic = Nothing
  End If
  BeforeChangeStatus = result
End Function

'����� ����, ���� ���� ����� ������� ������
'Parameters:
' ���������� ���
'Returns:
'  ������ ������ ������ Visual Basic
'  ,��� Nothing
'See Also:
'  ARMID
'  BeforeChangeStatus
'  CheckMenu
'  GetDocumentMode
'  IsDocDenied
'  MyRole
'  RoleDocAllowDelete
'  RoleDocCanSwitchStatus
'Example:
' dim variable as Object
' Set variable = me.ChooseRole()
Public Function ChooseRole() As Object
Attribute ChooseRole.VB_HelpID = 1590
Dim rs As ADODB.Recordset
Dim Q1 As String, Q2 As String, Q3 As String, Q4 As String
Dim res1 As String, res2 As String, resroles As String, armroles As String

    ' ���� ������  � ������� ������ �����������
    Q1 = CreateGUID2
    Call Session.TheFinder.FIND_IDS(Q1, "GROUPUSER", "TheUser", OpEQ, MyUser.id)
    Q2 = CreateGUID2
    Call Session.TheFinder.RowsToParents("GROUPUSER", Q1, Q2)
    Q3 = CreateGUID2
    Call Session.TheFinder.FIND_IDS(Q3, "ROLES_MAP", "TheGroup", OpIN_RESULT, Q2)
    res1 = CreateGUID2
    Call Session.TheFinder.RowsToInstances("ROLES_MAP", Q3, res1)
    Session.TheFinder.DropResults Q1
    Session.TheFinder.DropResults Q2
    Session.TheFinder.DropResults Q3
    
    Q1 = CreateGUID2
    Call Session.TheFinder.FIND_IDS(Q1, "ROLES_USER", "TheUser", OpEQ, MyUser.id)
    res2 = CreateGUID2
    Call Session.TheFinder.RowsToInstances("ROLES_USER", Q1, res2)
    Session.TheFinder.DropResults Q1
    
    
    
    resroles = CreateGUID2
    ' �������� ����� ����� ������������
    Session.TheFinder.QR_OR_QR res1, res2, resroles
    Session.TheFinder.DropResults res1
    Session.TheFinder.DropResults res2
    
    
    
    ' ��������� ����� �� ���� ��������� ��� ���
    Q1 = CreateGUID2
    Call Session.TheFinder.FIND_IDS(Q1, "ROLES_WP", "WP", OpEQ, ARMID)
    res1 = CreateGUID2
    Call Session.TheFinder.RowsToInstances("ROLES_WP", Q1, res1)
    Session.TheFinder.DropResults Q1
    armroles = CreateGUID2
    Session.TheFinder.QR_AND_QR resroles, res1, armroles
    
    Session.TheFinder.DropResults res1
    Session.TheFinder.DropResults resroles

    Set rs = Session.TheFinder.GetResults(armroles)
    If rs.EOF Then
        MsgBox "��� �� ��������� ������ � ���� �������", vbCritical + vbOKOnly, App.ProductName
        Set ChooseRole = Nothing
        Set rs = Nothing
        Session.TheFinder.DropResults armroles
        Exit Function
    End If
    
'    ��������� ������ ��������� �����
    Dim f As frmChooseRole
    Dim RoleObj As Object
    Set f = New frmChooseRole
    f.lstRole.Clear
    Dim i As Long
    Dim Col As Collection
    Set Col = New Collection
    i = 1
    While Not rs.EOF
        If Not IsNull(rs!result) Then
        Set RoleObj = Manager.GetInstanceObject(rs!result)
        Col.Add RoleObj, RoleObj.id
         f.lstRole.AddItem RoleObj.Name
        f.lstRole.ItemData(f.lstRole.NewIndex) = i
        i = i + 1
        End If
        rs.MoveNext
    Wend
    Set rs = Nothing
    Session.TheFinder.DropResults armroles
    If Col.Count = 1 Then
        Set ChooseRole = Col.Item(f.lstRole.ItemData(0))
        Unload f
        Set f = Nothing
        Set Col = Nothing
        Exit Function
    End If
    
    ' ���� ��������� ������ ����� ����, �� ���������� ����� ��� ��������� ����
    f.Show vbModal
    
    ' ��������� ����� �� ����� ����
    If f.OK Then
        Set ChooseRole = Col.Item(f.lstRole.ItemData(f.lstRole.ListIndex))
        Unload f
        Set f = Nothing
        Set Col = Nothing
        Exit Function
    Else
        Set ChooseRole = Nothing
        Unload f
        Set f = Nothing
        Set Col = Nothing
        Exit Function
    End If
End Function

'�������� ���� �� ���������� ���� ��� ������� ����
'Parameters:
'[IN]   menuName , ��� ���������: String  - ��� ����
'Returns:
'  ������ ������ RoleMenuStatus  - ��������� ����������� ����
'  ,��� Nothing
'See Also:
'  ARMID
'  BeforeChangeStatus
'  ChooseRole
'  GetDocumentMode
'  IsDocDenied
'  MyRole
'  RoleDocAllowDelete
'  RoleDocCanSwitchStatus
'Example:
' dim variable as RoleMenuStatus
' Set variable = me.CheckMenu(...���������...)
Public Function CheckMenu(ByVal menuName As String) As RoleMenuStatus
Attribute CheckMenu.VB_HelpID = 1585
  Dim ms As RoleMenuStatus
  ms = RoleMenuStatus_Unknown
  If MyRole Is Nothing Then
    Exit Function
  End If
  Dim i As Long, j As Long
  Dim rwp As ROLES_WP
  Dim ract As ROLES_ACT
  
  For i = 1 To MyRole.ROLES_WP.Count
    If MyRole.ROLES_WP.Item(i).WP.id = ARMID Then
          Set rwp = MyRole.ROLES_WP.Item(i)
      Exit For
    End If
  Next
  
  Set ract = FindRoleAct(rwp.ROLES_ACT, menuName)
  If Not ract Is Nothing Then
    If ract.Accesible = YesNo_Da Then
      ms = RoleMenuStatus_Visible
    End If
    If ract.Accesible = YesNo_Net Then
      ms = RoleMenuStatus_Hidden
    End If
  End If
  CheckMenu = ms
End Function

'����� ������ ��������������� ������ ����
Private Function FindRoleAct(ByVal Col As ROLES_ACT_COL, ByVal Name As String) As ROLES_ACT
  Dim i As Long, j As Long
  Dim ract As ROLES_ACT
  
  Set ract = Nothing
  For i = 1 To Col.Count
    If UCase(Col.Item(i).EntryPoints.Name) = UCase(Name) Then
      Set ract = Col.Item(i)
      Exit For
    End If
    If UCase(Col.Item(i).EntryPoints.Caption) = UCase(Name) Then
      Set ract = Col.Item(i)
      Exit For
    End If
    If ract Is Nothing Then
      Set ract = FindRoleAct(Col.Item(i).ROLES_ACT, Name)
      If Not ract Is Nothing Then Exit For
    End If
  Next
  Set FindRoleAct = ract
End Function

'�������� ����� ����������� ��������
'Parameters:
'[IN]   Obj , ��� ���������: Object  - ��������
'Returns:
'  �������� ���� String - �����
'See Also:
'  ARMID
'  BeforeChangeStatus
'  CheckMenu
'  ChooseRole
'  IsDocDenied
'  MyRole
'  RoleDocAllowDelete
'  RoleDocCanSwitchStatus
'Example:
' dim variable as String
' variable = me.GetDocumentMode(...���������...)
Public Function GetDocumentMode(ByVal Obj As Object) As String
Attribute GetDocumentMode.VB_HelpID = 1595
  Dim sid As String
  Dim tn As String
  Dim i As Long, j As Long
  GetDocumentMode = ""
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  For i = 1 To MyRole.ROLES_DOC.Count
    ' ����� ���
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
        ' ��� �������� � ������
        If MyRole.ROLES_DOC.Item(i).The_Denied = YesNo_Net Then
          For j = 1 To MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Count
            ' � ��������� �� ���������� ����������
            If sid = "" Then
              ' ���� ������ ��� ���������
              If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
                ' �������� ����������
                GetDocumentMode = MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_Mode.Name
                Exit Function
              End If
            Else
              ' ���� ���������  -  ���������� ������ � ������������� ����������
              If Not MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
                ' �����
                If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State.id = sid Then
                  If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_Mode Is Nothing Then
                     GetDocumentMode = ""
                  Else
                     ' �������� ����� ��������
                     GetDocumentMode = MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_Mode.Name
                  End If
                  Exit Function
                End If
              End If

            End If
          Next
        End If
      Exit For
    End If
  Next
  
End Function
'�������� ������� ��� ������ � ����������
'Parameters:
'[IN]   Obj , ��� ���������: Object  - ��������
'Returns:
' Boolean, ��������� ����������:
'   true  - ��������
'   false - ��������
'See Also:
'  ARMID
'  BeforeChangeStatus
'  CheckMenu
'  ChooseRole
'  GetDocumentMode
'  MyRole
'  RoleDocAllowDelete
'  RoleDocCanSwitchStatus
'Example:
' dim variable as Boolean
' variable = me.IsDocDenied(...���������...)
Public Function IsDocDenied(ByVal Obj As Object) As Boolean
Attribute IsDocDenied.VB_HelpID = 1600
  Dim sid As String
  Dim tn As String
  Dim mode As String
  Dim i As Long
  IsDocDenied = False
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  For i = 1 To MyRole.ROLES_DOC.Count
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
      If MyRole.ROLES_DOC.Item(i).The_Denied = YesNo_Da Then
        IsDocDenied = True
        Exit Function
      End If
    End If
  Next
End Function

'�������� ���������� �� ��������
'Parameters:
'[IN]   Obj , ��� ���������: Object  - ��������
'Returns:
' Boolean, ��������� ����������:
'   true  - ����� �������
'   false - ������
'See Also:
'  ARMID
'  BeforeChangeStatus
'  CheckMenu
'  ChooseRole
'  GetDocumentMode
'  IsDocDenied
'  MyRole
'  RoleDocCanSwitchStatus
'Example:
' dim variable as Boolean
' variable = me.RoleDocAllowDelete(...���������...)
Public Function RoleDocAllowDelete(ByVal Obj As Object) As Boolean
Attribute RoleDocAllowDelete.VB_HelpID = 1610
  Dim sid As String
  Dim tn As String
  Dim mode As String
  Dim i As Long, j As Long
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  RoleDocAllowDelete = True
  For i = 1 To MyRole.ROLES_DOC.Count
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
      If MyRole.ROLES_DOC.Item(i).AllowDeleteDoc = YesNo_Net Then
        RoleDocAllowDelete = False
        For j = 1 To MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Count
          If sid <> "" Then
            If Not MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
              If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State.id = sid Then
                If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).AllowDelete = Boolean_Net Then
                  RoleDocAllowDelete = False
                Else
                  RoleDocAllowDelete = True
                End If
                Exit For
              End If
            End If
          End If
        Next
        Exit Function
      End If
    End If
  Next
End Function

'����������  �� ����� ���������
'Parameters:
'[IN]   Obj , ��� ���������: Object  - ��������
'Returns:
' Boolean, ��������� ����������:
'   true  - ����� ������
'   false - ������
'See Also:
'  ARMID
'  BeforeChangeStatus
'  CheckMenu
'  ChooseRole
'  GetDocumentMode
'  IsDocDenied
'  MyRole
'  RoleDocAllowDelete
'Example:
' dim variable as Boolean
' variable = me.RoleDocCanSwitchStatus(...���������...)
Public Function RoleDocCanSwitchStatus(ByVal Obj As Object) As Boolean
Attribute RoleDocCanSwitchStatus.VB_HelpID = 1615
  Dim sid As String
  Dim tn As String
  Dim mode As String
  Dim i As Long, j As Long
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  RoleDocCanSwitchStatus = True
  For i = 1 To MyRole.ROLES_DOC.Count
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
        For j = 1 To MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Count
          If sid <> "" Then
            If Not MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
              If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State.id = sid Then
                If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).StateChangeDisabled = Boolean_Da Then
                  RoleDocCanSwitchStatus = False
                Else
                  RoleDocCanSwitchStatus = True
                End If
                Exit For
              End If
            End If
          End If
        Next
        Exit Function
    End If
  Next
End Function

'������������� �������� ���
'Parameters:
' ���������� ���
'Returns:
'  �������� ���� String - �������������
'See Also:
'  BeforeChangeStatus
'  CheckMenu
'  ChooseRole
'  GetDocumentMode
'  IsDocDenied
'  MyRole
'  RoleDocAllowDelete
'  RoleDocCanSwitchStatus
'Example:
' dim variable as String
'  variable = me.ARMID()
Public Function ARMID() As String
Attribute ARMID.VB_HelpID = 1575
   ARMID = "{FBDD7D58-A1D2-4326-9F89-477FD6C4CF97}"
End Function



