VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Optimizator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 185
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' основной класс оптимизатора

' список оптимизируемых ячеек
Public st As Boxes
Attribute st.VB_VarHelpID = 210

' результат обработки - перемещения
Public shifts As BoxShifts
Attribute shifts.VB_VarHelpID = 205

'расшифровка названия текущего отипмизируемого товара
Public CS_name As String

' пустые ячейки
Public EB As EmptyBoxes

' ячейки со смешанным товаром
Public Mixed As MixedBoxes


Private Const Stage1 As String = "Перемещение.Этап 1.Распределение"
Private Const Stage2 As String = "Перемещение.Этап 2.Исполнение"

'Оптимизация камеры
'Parameters:
'[IN][OUT]   KameraID , тип параметра: Integer - Номер камеры
'Returns:
'  значение типа Long - кол-во пустых ячеек
'See Also:
'  OptimizeAll
'  OptimizeRadius
'  shifts
'  st
'Example:
' dim variable as Long
' variable = me.OptimizeKamera(...параметры...)
Public Function OptimizeKamera(Optional KameraID As Integer = -1) As Long
Attribute OptimizeKamera.VB_HelpID = 195

  Set shifts = New BoxShifts
  
  Dim i As Long, j As Long
  Dim QMAX As Long, qcur As Long, qz As Long, cost As Long
  
  Dim b As Box, b2 As Box, bs As BoxShift
  
  
razbor:
  st.Base.Sort ("qcur")
   
  ' приоритетная разборка ячеек
  For i = 1 To st.Count
    Set b = st.Item(i)
    If b.Pallets.Count > 0 Then
      If b.qcur < b.QMAX / 2 Then
        b.GetFrom = False
        b.PutTo = True
      ElseIf b.qcur < b.QMAX Then
        b.PutTo = True
        b.GetFrom = False
      Else
        b.GetFrom = False
        b.PutTo = False
      End If
    Else
      b.PutTo = True
      b.GetFrom = False
    End If
    
    If b.HaveMixedPallet Then
     If Mixed.Item(b.Code).qcur > b.qcur Then
        b.GetFrom = True
        b.PutTo = False
        b.HavePut = False
      End If
    End If
  Next
  

  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If KameraID = -1 Or (b.T = KameraID) Then
         If b.GetFrom Then
nextpal0:
             If b.Pallets.Count > 0 Then
              For j = st.Count To 1 Step -1
                If i <> j Then
                  Set b2 = st.Item(j)
                  If b2.PutTo Then
                    If b2.Pallets.Count >= 0 Then
                      If b2.T = b.T Then
                        If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                            b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                            If b.HaveMixedPallet Then
                              Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                            End If
                            b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                            b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                            shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                            b.Pallets.Remove 1
                            
                            b2.GetFrom = False
                            b2.HavePut = True
                            
                            If b2.qcur + 1 > b2.QMAX Then
                              b2.PutTo = False
                            End If
                            GoTo nextpal0
                        End If
                      End If
                    End If
                  End If
                End If
              Next
            End If
         End If
      End If
  Next
  
  
  st.Base.Sort ("qcur")
   
  ' приоритетная разборка ячеек  - попытка использовать пустые ячейки для разбора
  For i = 1 To st.Count
    Set b = st.Item(i)
    If b.T = KameraID Or KameraID = -1 Then
      If b.HaveMixedPallet Then
        If Mixed.Item(b.Code).qcur > b.qcur Then
          If b.Pallets.Count > 0 Then
            If GetEmptyBox(b.T, b.qcur) Then
              GoTo razbor
            End If
          End If
        End If
      End If
    
    End If
  Next
  
  
  
  ' первый цикл  заполненные мене чем наполоину переносим в более полные
  st.Base.Sort ("qcur")
  
  For i = 1 To st.Count
    Set b = st.Item(i)
    If b.Pallets.Count > 0 Then
      If b.qcur < b.QMAX / 2 Then
        b.GetFrom = True
        b.PutTo = False
      ElseIf b.qcur < b.QMAX Then
        b.PutTo = True
        b.GetFrom = False
      Else
        b.GetFrom = False
        b.PutTo = False
      End If
    Else
    
      b.PutTo = False
      b.GetFrom = False
    End If
    
    If b.HaveMixedPallet Then
        If Mixed.Item(b.Code).qcur > b.qcur Then
          b.GetFrom = True
          b.PutTo = False
          b.HavePut = False
        End If
    End If
  Next
  

  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If KameraID = -1 Or (b.T = KameraID) Then
         If b.GetFrom Then
           If CanMoveFullCell(b, b.T) Then
nextpal:
             If b.Pallets.Count > 0 Then
              For j = st.Count To 1 Step -1
                If i <> j Then
                  Set b2 = st.Item(j)
                  If b2.PutTo Then
                    If b2.Pallets.Count > 0 Then
                      If b2.T = b.T Then
                        If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                            b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                            If b.HaveMixedPallet Then
                              Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                            End If
                            b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                            b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                            shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                            b.Pallets.Remove 1
                            
                            b2.GetFrom = False
                            b2.HavePut = True
                            
                            If b2.qcur + 1 > b2.QMAX Then
                              b2.PutTo = False
                            End If
                               ' DoEvents
                               ' ShowSklad
                            GoTo nextpal
                        End If
                      End If
                    End If
                  End If
                End If
              Next
            End If
           End If
         End If
      End If
  Next
'  ' ShowSklad
  
  ' второй цикл  - разрешаем класть в малозаполненные ячейки
   st.Base.Sort "qcur"
  
   For i = 1 To st.Count
    Set b = st.Item(i)
    If b.Pallets.Count > 0 Then
      If b.qcur < b.QMAX / 2 Then
        b.GetFrom = True
        b.PutTo = True
      ElseIf b.qcur < b.QMAX Then
        b.PutTo = True
        b.GetFrom = False
      Else
        b.GetFrom = False
        b.PutTo = False
      End If
    Else
    
      b.PutTo = True
      b.GetFrom = False
    End If
    
    If b.HaveMixedPallet Then
     If Mixed.Item(b.Code).qcur > b.qcur Then
        b.GetFrom = True
        b.PutTo = False
        b.HavePut = False
      End If
    End If
    
    If b.HavePut Then
        b.GetFrom = False
    End If
      
  Next

  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If KameraID = -1 Or (b.T = KameraID) Then
      If b.GetFrom Then
       If CanMoveFullCell(b, b.T) Then
nextpal_1:
       
        If b.Pallets.Count > 0 Then
          For j = st.Count To 1 Step -1
            If i <> j Then
                Set b2 = st.Item(j)
                If b2.PutTo Then
                  If b2.Pallets.Count > 0 Then
                    If b2.T = b.T Then
                     If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                          b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                           If b.HaveMixedPallet Then
                            Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                          End If
                          b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                          b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                          shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                          b.Pallets.Remove 1
                          b2.GetFrom = False
                          b2.HavePut = True
                          If b2.qcur + 1 > b2.QMAX Then
                            b2.PutTo = False
                          End If
                             ' DoEvents
                             ' ShowSklad
                          GoTo nextpal_1
                      End If
                    End If
                  End If
                End If
            End If
          Next
        End If
        
       End If
      End If
      End If
  Next


  ' цикл с заполненными на три четверти, разрешено класть в пустые ячейки
  st.Base.Sort ("qcur")
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.Pallets.Count > 0 Then
        If b.qcur < b.QMAX * 3 / 4 Then
          b.GetFrom = True
          b.PutTo = False
        ElseIf b.qcur < b.QMAX Then
          b.PutTo = True
          b.GetFrom = False
        Else
          b.GetFrom = False
          b.PutTo = False
        End If
      Else
        b.PutTo = True
        b.GetFrom = False
      End If
      If b.HaveMixedPallet Then
        If Mixed.Item(b.Code).qcur > b.qcur Then
          b.GetFrom = True
          b.PutTo = False
          b.HavePut = False
        End If
      End If
      If b.HavePut Then
        b.GetFrom = False
      End If
  Next
  ' ShowSklad

  For i = 1 To st.Count
      Set b = st.Item(i)
      If KameraID = -1 Or (b.T = KameraID) Then
      If b.GetFrom Then
       If CanMoveFullCell(b, b.T) Then
nextpal2:
        If b.Pallets.Count > 0 Then

          For j = st.Count To 1 Step -1
            Set b2 = st.Item(j)
            If i <> j And b2.PutTo Then
              If b2.Pallets.Count > 0 Then
                If b2.T = b.T Then
                   If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                      b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                      If b.HaveMixedPallet Then
                        Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                      End If
                      b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                      b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                      shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                      b.Pallets.Remove 1
                      b2.GetFrom = False
                      b2.HavePut = True
                      If b2.qcur + 1 > b2.QMAX Then
                        b2.PutTo = False
                      End If
                        ' DoEvents
                        ' ShowSklad
                      GoTo nextpal2
                  End If
                End If
              End If
            End If
          Next
        End If
       End If
      End If
      End If
  Next

  ' ShowSklad
  
  ' последний цикл разрешено разбирать любые ячейки куда еще не клали паллет
    st.Base.Sort ("qcur")
   For i = 1 To st.Count
    Set b = st.Item(i)
    If b.Pallets.Count > 0 Then
      If b.qcur < b.QMAX Then
        b.GetFrom = True
        b.PutTo = True
     Else
        b.GetFrom = False
        b.PutTo = False
      End If
    Else
      b.PutTo = False
      b.GetFrom = False
    End If
    If b.HaveMixedPallet Then
        If Mixed.Item(b.Code).qcur > b.qcur Then
          b.GetFrom = True
          b.PutTo = False
          b.HavePut = False
        End If
    End If
    If b.HavePut Then
        b.GetFrom = False
      End If
  Next

  For i = 1 To st.Count
      Set b = st.Item(i)
      If KameraID = -1 Or (b.T = KameraID) Then
        If b.GetFrom Then
           If CanMoveFullCell(b, b.T) Then
nextpal3:
            If b.Pallets.Count > 0 Then
              For j = st.Count To 1 Step -1
                If j <> i Then
                  Set b2 = st.Item(j)
                  If b2.PutTo Then
                    If b2.Pallets.Count > 0 Then
                      If b2.T = b.T Then
                         If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                            b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                            If b.HaveMixedPallet Then
                              Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                            End If
                            b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                            b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                            shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                            b.Pallets.Remove 1
                            b2.GetFrom = False
                            b2.HavePut = True
                            If b2.qcur + 1 > b2.QMAX Then
                              b2.PutTo = False
                            End If
                            ' DoEvents
                            ' ShowSklad
                            GoTo nextpal3
                        End If
                      End If
                    End If
                  End If
                End If
              Next
            End If
          End If
        End If
      End If
  Next
'
'  Me.Caption = ""
  OptimizeShifts
  
  QMAX = 0
  qcur = 0
  qz = 0
  cost = 0
  

'  подсчет пустых ячеек в оптимизируемом множестве
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.qcur = 0 Then
        If b.HaveMixedPallet Then
          If Mixed.Item(b.Code).qcur = 0 Then
            qz = qz + 1
          End If
        Else
          qz = qz + 1
        End If
      End If
     
  Next

  OptimizeKamera = qz
End Function

' Исключение лишних перемещений
Private Function OptimizeShifts() As Long
'  Dim sh As BoxShift
'  Dim sh2 As BoxShift
'  Dim goodShifts As BoxShifts
'
'  Dim i As Long
'  Dim j As Long
'  Dim OPt As Boolean
'  OPt = False
'  For i = 1 To shifts.Count
'    Set sh = shifts.Item(i)
'    If sh.ExcludeShift = False Then
'
'    For j = i + 1 To shifts.Count
'          Set sh2 = shifts.Item(j)
'          If sh2.ExcludeShift = False Then
'
'            If sh.ToCode = sh2.FromCode Then
'              sh.ToCode = sh2.ToCode
'              sh2.ExcludeShift = True
'              OPt = True
'            End If
'          End If
'    Next
'    End If
'  Next
'
'  If OPt Then
'    j = shifts.Count
'    Set goodShifts = New BoxShifts
'    For i = 1 To shifts.Count
'      Set sh = shifts.Item(i)
'      If sh.ExcludeShift = False Then
'        goodShifts.Add sh.ToCode, sh.FromCode, sh.pallet, sh.GoodInfo, sh.ptype
'
'      End If
'    Next
'    Set shifts = Nothing
'    Set shifts = goodShifts
'    OptimizeShifts
'  End If

End Function

'Оптимизация всего склада
'Parameters:
' параметров нет
'Returns:
'  значение типа Long кол-во пустых ячеек
'See Also:
'  OptimizeKamera
'  OptimizeRadius
'  shifts
'  st
'Example:
' dim variable as Long
'  variable = me.OptimizeAll()
Public Function OptimizeAll() As Long
Attribute OptimizeAll.VB_HelpID = 190


 Set shifts = New BoxShifts
  
  Dim i As Long, j As Long
  Dim QMAX As Long, qcur As Long, qz As Long, cost As Long
  
  Dim b As Box, b2 As Box, bs As BoxShift
  
razbor:
   ' разбор ячеек
   st.Base.Sort ("qcur")
   For i = 1 To st.Count
      Set b = st.Item(i)
      If b.Pallets.Count > 0 Then
        If b.qcur < b.QMAX / 2 Then
          b.GetFrom = False
          b.PutTo = True
        ElseIf b.qcur < b.QMAX Then
          b.PutTo = True
          b.GetFrom = False
        Else
          b.GetFrom = False
          b.PutTo = False
        End If
         If b.HaveMixedPallet Then
            If Mixed.Item(b.Code).qcur > b.qcur Then
              b.GetFrom = True
              b.PutTo = False
              b.HavePut = False
            End If
         End If
        
      Else
      
        b.PutTo = True
        b.GetFrom = False
      End If
  Next
   
   
   
     
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
       
nextpal0:
        If b.Pallets.Count > 0 Then
      
          For j = st.Count To 1 Step -1
           If i <> j Then
            Set b2 = st.Item(j)
            If b2.PutTo Then
             
                If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                    b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                    If b.HaveMixedPallet Then
                      Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                    End If
                    b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                    b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                    shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                    b.Pallets.Remove 1

                    b2.GetFrom = False
                    b2.HavePut = True
                    If b2.qcur + 1 > b2.QMAX Then
                      b2.PutTo = False
                    End If
                      
                    GoTo nextpal0
                End If
             
              End If
              
            End If
            
            ' DoEvents
          Next
        End If
       
      End If
  Next
   
   
    st.Base.Sort ("qcur")
   
  ' приоритетная разборка ячеек, пытаемся получить пустые ячейки, если не хватило занятых
  For i = 1 To st.Count
    Set b = st.Item(i)
    If b.HaveMixedPallet Then
      If Mixed.Item(b.Code).qcur > b.qcur Then
        If b.Pallets.Count > 0 Then
          If GetEmptyBox(b.T, b.qcur) Then GoTo razbor
        End If
      End If
    End If
  Next
   
   ' первый цикл. перемещение из заполненных менеее чем на половину в более полные
   st.Base.Sort ("qcur")
   For i = 1 To st.Count
      Set b = st.Item(i)
      If b.Pallets.Count > 0 Then
        If b.qcur < b.QMAX / 2 Then
          b.GetFrom = True
          b.PutTo = True
        ElseIf b.qcur < b.QMAX Then
          b.PutTo = True
          b.GetFrom = False
        Else
          b.GetFrom = False
          b.PutTo = False
        End If
         If b.HaveMixedPallet Then
            If Mixed.Item(b.Code).qcur > b.qcur Then
              b.GetFrom = True
              b.PutTo = False
              b.HavePut = False
            End If
         End If
        
      Else
      
        b.PutTo = False
        b.GetFrom = False
      End If
  Next
  
  
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
       If CanMoveFullCell(b, -1) Then
nextpal:
        If b.Pallets.Count > 0 Then
      
          For j = st.Count To 1 Step -1
           If i <> j Then
            Set b2 = st.Item(j)
            If b2.PutTo Then
             
                If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                    b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                    If b.HaveMixedPallet Then
                      Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                    End If
                    b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                    b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                    shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                    b.Pallets.Remove 1

                    b2.GetFrom = False
                    b2.HavePut = True
                    If b2.qcur + 1 > b2.QMAX Then
                      b2.PutTo = False
                    End If
                       ' DoEvents
                      ' ShowSklad
                    GoTo nextpal
                End If
             
              End If
              
            End If
            
            ' DoEvents
          Next
        End If
       End If
     
      End If
  Next
  
  ' второй цикл. перемещение из заполненных менеее чем на половину в любые
  st.Base.Sort ("qcur")
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.Pallets.Count > 0 Then
        If b.qcur < b.QMAX / 2 Then
          b.GetFrom = True
          b.PutTo = True
        ElseIf b.qcur < b.QMAX Then
          b.PutTo = True
          b.GetFrom = False
        Else
          b.GetFrom = False
          b.PutTo = False
        End If
      Else
      
        b.PutTo = True
        b.GetFrom = False
      End If
      If b.HaveMixedPallet Then
       If Mixed.Item(b.Code).qcur > b.qcur Then
          b.GetFrom = True
          b.PutTo = False
          b.HavePut = False
        End If
      End If
       If b.HavePut Then
        b.GetFrom = False
      End If
  Next
  
  st.Base.Sort ("qcur")
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
         If CanMoveFullCell(b, -1) Then
nextpal_1:
        If b.Pallets.Count > 0 Then
      
          For j = st.Count To 1 Step -1
            If i <> j Then
                Set b2 = st.Item(j)
                If b2.PutTo Then
             
                      If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                          b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                          If b.HaveMixedPallet Then
                             Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                          End If
                          b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                          b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                          shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                          b.Pallets.Remove 1

                          b2.GetFrom = False
                          b2.HavePut = True
                          If b2.qcur + 1 > b2.QMAX Then
                            b2.PutTo = False
                          End If
                            ' DoEvents
                            ' ShowSklad
                          GoTo nextpal_1
                      End If
              End If
            End If
          Next
        End If
        End If
        
      End If
  Next
  
  ' третий цикл.  разрешен разбор незаполненных ячеек
   st.Base.Sort ("qcur")
  For i = 1 To st.Count
    Set b = st.Item(i)
    If b.Pallets.Count > 0 Then
      If b.qcur < b.QMAX Then
        b.GetFrom = True
        b.PutTo = True
     Else
        b.GetFrom = False
        b.PutTo = False
      End If
    Else
      b.PutTo = False
      b.GetFrom = False
    End If
    If b.HaveMixedPallet Then
      If Mixed.Item(b.Code).qcur > b.qcur Then
        b.GetFrom = True
        b.PutTo = False
        b.HavePut = False
      End If
    End If
     If b.HavePut Then
        b.GetFrom = False
      End If
  Next

  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
         If CanMoveFullCell(b, -1) Then
nextpal3:
        If b.Pallets.Count > 0 Then
          For j = st.Count To 1 Step -1
            If j <> i Then
              Set b2 = st.Item(j)
              If b2.PutTo Then
            
                  If b2.T = b.T And (b2.X - b.X) * (b2.X - b.X) + (b2.Y - b.Y) * (b2.Y - b.Y) + (b2.Z - b.Z) * (b2.Z - b.Z) < 300 Then
                     If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                        b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                        If b.HaveMixedPallet Then
                          Mixed.Item(b.Code).qcur = Mixed.Item(b.Code).qcur - b.Pallets.Item(1).palletSize
                        End If
                        b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                        b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                        shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                        b.Pallets.Remove 1

                        b2.GetFrom = False
                        b2.HavePut = True
                        If b2.qcur + 1 > b2.QMAX Then
                          b2.PutTo = False
                        End If
                        GoTo nextpal3
                    End If
                  End If
               
              End If
            End If
          Next
        End If
        End If
      End If
  Next
  
  OptimizeShifts
   
  QMAX = 0
  qcur = 0
  qz = 0
  cost = 0
  
  
'  подсчет пустых ячеек в оптимизируемомо множестве
  For i = 1 To st.Count
      Set b = st.Item(i)
      
      If b.qcur = 0 Then
        If b.HaveMixedPallet Then
          If Mixed.Item(b.Code).qcur = 0 Then
            qz = qz + 1
          End If
        Else
          qz = qz + 1
        End If
      End If
     
  Next

  OptimizeAll = qz
End Function


'Оптимизация в пределах радиуса
'Parameters:
'[IN]   mx , тип параметра: Double  - центр х координата,
'[IN][OUT]   my , тип параметра: Double  - центр y координата,
'[IN][OUT]   mz , тип параметра: Double  - центр z координата,
'[IN]   Radius , тип параметра: Double  - радиус
'See Also:
'  OptimizeAll
'  OptimizeKamera
'  shifts
'  st
'Example:
'  call me.OptimizeRadius(...параметры...)
Public Sub OptimizeRadius(ByVal mx As Double, my As Double, mz As Double, ByVal Radius As Double)
Attribute OptimizeRadius.VB_HelpID = 200
  
  Set shifts = New BoxShifts
  
  Dim i As Long, j As Long
  Dim QMAX As Long, qcur As Long, qz As Long, cost As Long
  
  Dim b As Box, b2 As Box, bs As BoxShift
  
  If mx = 0 And my = 0 And mz = 0 Then
    mx = 0
    my = 0
    mz = 0
    qz = 0
    For i = 1 To st.Count
        Set b = st.Item(i)
        If b.Pallets.Count > 0 Then
          mx = mx + b.X * b.qcur
          my = my + b.Y * b.qcur
          mz = mz + b.Z * b.qcur
          qz = qz + b.qcur
        End If
        
    Next
  mx = mx / qz
  my = my / qz
  mz = mz / qz
  
    
  End If
  
  
  
   
  
   
  Dim movecnt As Long
    
  movecnt = 1
  While movecnt > 0
      movecnt = 0
      For i = 1 To st.Count
          Set b = st.Item(i)
          If (b.X - mx) * (b.X - mx) + (b.Y - my) * (b.Y - my) + (b.Z - mz) * (b.Z - mz) > Radius And b.Pallets.Count > 0 Then
nextpal3:
            If b.Pallets.Count > 0 Then
              For j = st.Count To 1 Step -1
                If j <> i Then
                  Set b2 = st.Item(j)
                   If (b2.X - mx) * (b2.X - mx) + (b2.Y - my) * (b2.Y - my) + (b2.Z - mz) * (b2.Z - mz) <= Radius Then
                     If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                            b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                            b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                            b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                            shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                            b.Pallets.Remove 1
                            movecnt = movecnt + 1
                            b2.HavePut = True
                            ' DoEvents
                            ' ShowSklad
                            GoTo nextpal3
                      End If
                  End If
                End If
              Next
            
            End If
        
          End If
      Next
      ' ShowSklad
  Wend
  
 
  
   For i = 1 To st.Count
    Set b = st.Item(i)
    If b.Pallets.Count > 0 Then
      If b.qcur < b.QMAX Then
        b.GetFrom = True
        b.PutTo = True
     Else
        b.GetFrom = False
        b.PutTo = False
      End If
    Else
      b.PutTo = True
      b.GetFrom = False
    End If
     If b.HaveMixedPallet Then
       b.GetFrom = True
      b.PutTo = False
      b.HavePut = False
    End If
     If b.HavePut Then
        b.GetFrom = False
      End If
  Next

  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
nextpal4:
        If b.Pallets.Count > 0 Then
          For j = st.Count To 1 Step -1
            If j <> i Then
              Set b2 = st.Item(j)
              If b2.PutTo Then
                If b2.Pallets.Count > 0 Then
                     If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                        b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                        b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                        b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                        shifts.Add b2.Code, b.Code, b.Pallets.Item(1).PalCode, CS_name, b.Pallets.Item(1).PalType, IIf(b.HaveMixedPallet, Stage1, Stage2)
                        b.Pallets.Remove 1
                        b2.HavePut = True
                        b2.GetFrom = False
                        If b2.qcur = b2.QMAX Then
                          b2.PutTo = False
                        End If
                        ' DoEvents
                        ' ShowSklad
                        GoTo nextpal4
                    End If
                End If
              End If
              
            End If
          Next
         
        End If
      End If
  Next
  
  OptimizeShifts
  
End Sub





' Проверяем сможем ли мы разгрузить полную ячейку
' параметры
' T - номер камеры, или -1 для исключения привязки к камере
' CellToMove - ячейка подлежащая разгрузке
' Результат - true - можем разгрузить
Private Function CanMoveFullCell(CellToMove As Box, ByVal T As Integer) As Boolean
  Dim cnt As Long
  Dim OK As Boolean
  
  Dim b As Box, b2 As Box
  
  If CellToMove.HavePut Then
    CanMoveFullCell = False
    Exit Function
  End If
  
  If CellToMove.GetFrom = False Then
    CanMoveFullCell = False
    Exit Function
  End If
  
 
    
    
  
  Dim i As Long, j As Long
  Dim st_test As Boxes
  Set st_test = New Boxes
  For i = 1 To st.Count
    Set b2 = st_test.Add(st.Item(i).qcur, st.Item(i).QMAX, st.Item(i).T, st.Item(i).Z, st.Item(i).Y, st.Item(i).X, st.Item(i).Code, "")
    b2.GetFrom = st.Item(i).GetFrom
    b2.PutTo = st.Item(i).PutTo
    b2.HavePut = st.Item(i).HavePut
    For j = 1 To st.Item(i).Pallets.Count
      b2.Pallets.Add st.Item(i).Pallets.Item(j).PalCode, st.Item(i).Pallets.Item(j).GoodParam, st.Item(i).Pallets.Item(j).palletSize, st.Item(i).Pallets.Item(j).PalType
    Next
  Next
  
  st_test.Base.Sort ("qcur")


  Set b = New Box
  With CellToMove
    b.Code = .Code
    b.qcur = .qcur
    b.QMAX = .QMAX
    b.X = .X
    b.Y = .Y
    b.Z = .Z
    b.T = .T
    b.GetFrom = .GetFrom
    b.PutTo = .PutTo
    b.HavePut = .HavePut
    For j = 1 To .Pallets.Count
      b.Pallets.Add .Pallets.Item(j).PalCode, .Pallets.Item(j).GoodParam, .Pallets.Item(j).palletSize, .Pallets.Item(j).PalType
    Next
  End With
  
nextpal_1:
  If b.Pallets.Count > 0 Then
    
        For j = st_test.Count To 1 Step -1
          Set b2 = st_test.Item(j)
            If b.Code <> b2.Code Then
              If T > 0 Then
                If b.T = b2.T Then
                  OK = True
                Else
                  OK = False
                End If
              Else
                OK = True
              End If
                
                If OK Then
                  If b2.PutTo Then
                     If b2.qcur + b.Pallets.Item(1).palletSize <= b2.QMAX Then
                          b.qcur = b.qcur - b.Pallets.Item(1).palletSize
                          b2.qcur = b2.qcur + b.Pallets.Item(1).palletSize
                          b2.Pallets.Add b.Pallets.Item(1).PalCode, b.Pallets.Item(1).GoodParam, b.Pallets.Item(1).palletSize, b.Pallets.Item(1).PalType
                          b.Pallets.Remove 1
                          b2.GetFrom = False
                          b2.HavePut = True
                          If b2.qcur = b2.QMAX Then
                            b2.PutTo = False
                          End If
                          GoTo nextpal_1
                      End If
                  End If
                End If
          End If
        Next
      
    End If
  
    If b.Pallets.Count > 0 Then
      CanMoveFullCell = False
    Else
      CanMoveFullCell = True
    End If

End Function


' запрос пустой ячейки из пула
' параметры
'[IN][OUT]   Kamera , тип параметра: Integer - Номер камеры
'Returns:
'  нет
Private Function GetEmptyBox(ByVal Kamera As Integer, Optional ByVal MinSZ As Double = 1)
  Dim i As Long
  Dim OK As Boolean
  GetEmptyBox = False
  If EB Is Nothing Then Exit Function
  If EB.Count = 0 Then Exit Function
  
  
  
  For i = 1 To EB.Count
    OK = True
    
    If EB.Item(i).Used Then
    OK = False
    End If
    If Kamera > 0 Then
      If Kamera <> EB.Item(i).Kamera Then
        OK = False
      End If
    End If
    If OK Then
        If EB.Item(i).QMAX < MinSZ Then
            OK = False
        End If
    End If
    
    If OK Then
      st.Add 0, EB.Item(i).QMAX, Kamera, 1, 1, 1, EB.Item(i).Code, ""
      EB.Item(i).Used = True
      GetEmptyBox = True
      Exit Function
    End If
  Next

End Function
