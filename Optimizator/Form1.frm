VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtZ 
      Height          =   375
      Left            =   11760
      TabIndex        =   13
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Left            =   10800
      TabIndex        =   12
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox pOut 
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   10995
      TabIndex        =   10
      Top             =   2520
      Width           =   11055
   End
   Begin VB.TextBox txtlevel 
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Text            =   "17"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Компактнее"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Оптимизировать СКЛАД"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Оптимизировать камеру"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtCount 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Text            =   "1000"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Init"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Z"
      Height          =   375
      Left            =   11520
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   375
      Left            =   10560
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "R"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   375
      Left            =   9480
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim st As Boxes
Dim fin As Boxes
Dim shifts As BoxShifts
Dim Radius As Double




Private Sub cmdSave_Click()
 Dim i As Long
 Dim ff As Integer
 ff = FreeFile
 Open App.Path & "\shifts.txt" For Output As #ff
 
 shifts.Base.Sort ("FromCode")
 For i = 1 To shifts.Count
 Print #ff, shifts.Item(i).FromCode, shifts.Item(i).ToCode, st.Item(shifts.Item(i).ToCode).qmax
    
  Next
  Close #ff
End Sub

Private Sub Command1_Click()
  Set st = New Boxes
 
  Dim i As Long
  Dim qmax As Long, qcur As Long, qz As Long
  
  Dim b As Box
  
  For i = 1 To Val(txtCount)
again:
   Set b = New Box
   b.X = 0 + Rnd * 60
   b.Y = 0 + Rnd * 20
   b.Z = 0 + Rnd * 6
   b.T = 1 + b.X / 15
   
   If b.X Mod 2 = 0 Then
    b.qmax = 1 '+ Rnd * 6
    If Rnd > 0.5 Then
      b.qcur = 1
    Else
      b.qcur = 0
    End If
    
   Else
    b.qmax = 7
    b.qcur = Rnd * 7
   End If
   
   
   
   
   b.Code = Right("00" & b.X, 2) & "." & Right("00" & b.Y, 2) & "." & Right("00" & b.Z, 2) & "." & Right("00" & b.T, 2)
   If st.Item(b.Code) Is Nothing Then
       st.Add b.qcur, b.qmax, b.T, b.Z, b.Y, b.X, b.Code
      Else
      GoTo again
   End If
  
  Next
  qmax = 0
  qcur = 0
  qz = 0
  For i = 1 To st.Count
      Set b = st.Item(i)
      qmax = qmax + b.qmax
      qcur = qcur + b.qcur
      If b.qcur = 0 Then
        qz = qz + 1
      End If
  Next
  Text1.Text = qmax & ":" & qcur & ":" & qz
  ShowSklad
End Sub

Private Sub Command2_Click()
  'Set fin = New Boxes
  Set shifts = New BoxShifts
  
  Dim i As Long, j As Long
  Dim qmax As Long, qcur As Long, qz As Long, cost As Long
  
  Dim b As Box, b2 As Box, bs As BoxShift
  
  For i = 1 To st.Count
    Set b = st.Item(i)
    If b.qcur > 0 Then
      If b.qcur < b.qmax / 2 Then
        b.GetFrom = True
        b.PutTo = False
      ElseIf b.qcur < b.qmax Then
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
  Next
  
  st.Base.Sort ("T")
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
      
nextpal:
       
        If b.qcur > 0 Then
      
          For j = 1 To st.Count
            Set b2 = st.Item(j)
            If b2.PutTo Then
              If b2.qcur > 0 Then
                If b2.T = b.T Then
                  If b2.qcur < b2.qmax Then
                  
                      b.qcur = b.qcur - 1
                      b2.qcur = b2.qcur + 1
                      shifts.Add b2.Code, b.Code
                      b2.GetFrom = False
                      b2.HavePut = True
                      
                      If b2.qcur = b2.qmax Then
                        b2.PutTo = False
                      End If
                         DoEvents
                         ShowSklad
                      GoTo nextpal
                  End If
                End If
              End If
            End If
          Next
        End If
        
     
      End If
  Next
'  ShowSklad
  
  
  
   For i = 1 To st.Count
    Set b = st.Item(i)
    If b.qcur > 0 Then
      If b.qcur < b.qmax / 2 Then
        b.GetFrom = True
        b.PutTo = True
      ElseIf b.qcur < b.qmax Then
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
    If b.HavePut Then
        b.GetFrom = False
      End If
  Next
  
  st.Base.Sort ("T")
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
      
nextpal_1:
       
        If b.qcur > 0 Then
          For j = 1 To st.Count
            If i <> j Then
                Set b2 = st.Item(j)
                If b2.PutTo Then
                  If b2.qcur > 0 Then
                    If b2.T = b.T Then
                      If b2.qcur < b2.qmax Then
                      
                          b.qcur = b.qcur - 1
                          b2.qcur = b2.qcur + 1
                          shifts.Add b2.Code, b.Code
                          b2.GetFrom = False
                          b2.HavePut = True
                          If b2.qcur = b2.qmax Then
                            b2.PutTo = False
                          End If
                             DoEvents
                             ShowSklad
                          GoTo nextpal_1
                      End If
                    End If
                  End If
                End If
            End If
          Next
        End If
        
     
      End If
  Next

  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.qcur > 0 Then
        If b.qcur < b.qmax * 3 / 4 Then
          b.GetFrom = True
          b.PutTo = False
        ElseIf b.qcur < b.qmax Then
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
      If b.HavePut Then
        b.GetFrom = False
      End If
  Next
  ShowSklad

  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
nextpal2:
        If b.qcur > 0 Then

          For j = 1 To st.Count
            Set b2 = st.Item(j)
            If i <> j And b2.PutTo Then
              If b2.qcur > 0 Then
                If b2.T = b.T Then
                  If b2.qcur < b2.qmax Then
                      b.qcur = b.qcur - 1
                      b2.qcur = b2.qcur + 1
                      shifts.Add b2.Code, b.Code
                      b2.GetFrom = False
                      b2.HavePut = True
                      If b2.qcur = b2.qmax Then
                        b2.PutTo = False
                      End If
                        DoEvents
                        ShowSklad
                      GoTo nextpal2
                  End If
                End If
              End If
            End If
          Next
        End If

      End If
  Next

  ShowSklad
  
   For i = 1 To st.Count
    Set b = st.Item(i)
    If b.qcur > 0 Then
      If b.qcur < b.qmax Then
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
    If b.HavePut Then
        b.GetFrom = False
      End If
  Next

  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
nextpal3:
        If b.qcur > 0 Then
          For j = 1 To st.Count
            If j <> i Then
              Set b2 = st.Item(j)
              If b2.PutTo Then
                If b2.qcur > 0 Then
                  If b2.T = b.T Then

                    If b2.qcur < b2.qmax Then
                        b.qcur = b.qcur - 1
                        b2.qcur = b2.qcur + 1
                        shifts.Add b2.Code, b.Code
                        b2.GetFrom = False
                        b2.HavePut = True
                        If b2.qcur = b2.qmax Then
                          b2.PutTo = False
                        End If
                        DoEvents
                        ShowSklad
                        GoTo nextpal3
                    End If
                  End If
                End If
              End If
            End If
          Next
        End If

        DoEvents
        ShowSklad
      End If
  Next
  
  Me.Caption = ""
  OptimizeShifts
  
  qmax = 0
  qcur = 0
  qz = 0
  cost = 0
  
  For i = 1 To shifts.Count
    cost = cost + st.Item(shifts.Item(i).FromCode).GetCost(st.Item(shifts.Item(i).FromCode).T = st.Item(shifts.Item(i).ToCode).T)
    cost = cost + st.Item(shifts.Item(i).ToCode).PutCost(st.Item(shifts.Item(i).FromCode).T = st.Item(shifts.Item(i).ToCode).T)
    'Debug.Print shifts.Item(i).FromCode & "->" & shifts.Item(i).ToCode
  Next
  For i = 1 To st.Count
      Set b = st.Item(i)
      qmax = qmax + b.qmax
      qcur = qcur + b.qcur
      
      If b.qcur = 0 Then
        qz = qz + 1
      End If
      If b.qcur > 0 And b.qcur < b.qmax Then
        Debug.Print b.Code, b.qmax, b.qcur, b.GetFrom
      End If
      
  Next
  Text2.Text = qmax & ":" & qcur & ":" & qz & ":" & CInt(cost / 60) & " h:" & shifts.Count & " shifts"
  ShowSklad
  
End Sub

Private Sub OptimizeShifts()
  Dim sh As BoxShift
  Dim sh2 As BoxShift
  Dim goodShifts As BoxShifts
  
  Dim i As Long
  Dim j As Long
  Dim opt As Boolean
  opt = False
  For i = 1 To shifts.Count
    Set sh = shifts.Item(i)
    If sh.ExcludeShift = False Then
  
        For j = i + 1 To shifts.Count
              Set sh2 = shifts.Item(j)
              If sh2.ExcludeShift = False Then
                If sh.ToCode = sh2.FromCode Then
                  sh.ToCode = sh2.ToCode
                  sh2.ExcludeShift = True
                  opt = True
                End If
              End If
        Next
    End If
  Next
  
  If opt Then
  j = shifts.Count
    Set goodShifts = New BoxShifts
    For i = 1 To shifts.Count
      Set sh = shifts.Item(i)
      If sh.ExcludeShift = False Then
        goodShifts.Add sh.ToCode, sh.FromCode
      
      End If
    Next
    Set shifts = Nothing
    Set shifts = goodShifts
    Me.Caption = Me.Caption + "||" & j & "->" & shifts.Count
    DoEvents
    OptimizeShifts
  End If

End Sub

Private Sub Command3_Click()
 'Set fin = New Boxes
  Set shifts = New BoxShifts
  
  Dim i As Long, j As Long
  Dim qmax As Long, qcur As Long, qz As Long, cost As Long
  
  Dim b As Box, b2 As Box, bs As BoxShift
  
   For i = 1 To st.Count
      Set b = st.Item(i)
      If b.qcur > 0 Then
        If b.qcur < b.qmax / 2 Then
          b.GetFrom = True
          b.PutTo = True
        ElseIf b.qcur < b.qmax Then
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
  Next
  
  st.Base.Sort ("T")
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
nextpal:
        If b.qcur > 0 Then
      
          For j = 1 To st.Count
           If i <> j Then
            Set b2 = st.Item(j)
            If b2.PutTo Then
              'If b2.qcur > 0 Then
                'If True Then
                  If b2.qcur < b2.qmax Then
                  
                      b.qcur = b.qcur - 1
                      b2.qcur = b2.qcur + 1
                      shifts.Add b2.Code, b.Code
                      b2.GetFrom = False
                      b2.HavePut = True
                      If b2.qcur = b2.qmax Then
                        b2.PutTo = False
                      End If
                         DoEvents
                        ShowSklad
                      GoTo nextpal
                  End If
                'End If
              'End If
              End If
              
            End If
            Me.Caption = i & " : " & j
            DoEvents
          Next
        End If
     
     
      End If
  Next
  
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.qcur > 0 Then
        If b.qcur < b.qmax / 2 Then
          b.GetFrom = True
          b.PutTo = True
        ElseIf b.qcur < b.qmax Then
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
       If b.HavePut Then
        b.GetFrom = False
      End If
  Next
  
  st.Base.Sort ("T")
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
nextpal_1:
        If b.qcur > 0 Then
      
          For j = 1 To st.Count
            If i <> j Then
                Set b2 = st.Item(j)
                If b2.PutTo Then
                  'If b2.qcur > 0 Then
                    'If True Then
                      If b2.qcur < b2.qmax Then
                      
                          b.qcur = b.qcur - 1
                          b2.qcur = b2.qcur + 1
                          shifts.Add b2.Code, b.Code
                          b2.GetFrom = False
                          b2.HavePut = True
                          If b2.qcur = b2.qmax Then
                            b2.PutTo = False
                          End If
                            DoEvents
                            ShowSklad
                          GoTo nextpal_1
                      End If
                    'End If
                  'End If
                End If
            End If
          Next
        End If
        
      End If
  Next
  
 ' ShowSklad
  
'     For i = 1 To st.Count
'      Set b = st.Item(i)
'      If b.qcur > 0 Then
'        If b.qcur < b.qmax * 3 / 4 Then
'          b.GetFrom = True
'          b.PutTo = False
'        ElseIf b.qcur < b.qmax Then
'          b.PutTo = True
'          b.GetFrom = False
'        Else
'          b.GetFrom = False
'          b.PutTo = False
'        End If
'      Else
'        b.PutTo = False
'        b.GetFrom = False
'      End If
'  Next
'
'  For i = 1 To st.Count
'      Set b = st.Item(i)
'      If b.GetFrom Then
'nextpal2:
'        If b.qcur > 0 Then
'
'          For j = 1 To st.Count
'            Set b2 = st.Item(j)
'            If b2.PutTo Then
'              If b2.qcur > 0 Then
'                If True Then
'                  If b2.qcur < b2.qmax Then
'                      b.qcur = b.qcur - 1
'                      b2.qcur = b2.qcur + 1
'                      shifts.Add b2.Code, b.Code
'                      b2.GetFrom = False
'                      b2.HavePut = True
'                      If b2.qcur = b2.qmax Then
'                        b2.PutTo = False
'                      End If
'                      GoTo nextpal2
'                  End If
'                End If
'              End If
'            End If
'          Next
'        End If
'
'      End If
'  Next
'
'  ShowSklad
  
  For i = 1 To st.Count
    Set b = st.Item(i)
    If b.qcur > 0 Then
      If b.qcur < b.qmax Then
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
     If b.HavePut Then
        b.GetFrom = False
      End If
  Next

  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
nextpal3:
        If b.qcur > 0 Then
          For j = 1 To st.Count
            If j <> i Then
              Set b2 = st.Item(j)
              If b2.PutTo Then
                'If b2.qcur > 0 Then
                  If b2.T = b.T And (b2.X - b.X) * (b2.X - b.X) + (b2.Y - b.Y) * (b2.Y - b.Y) + (b2.Z - b.Z) * (b2.Z - b.Z) < 300 Then

                    If b2.qcur < b2.qmax Then
                        b.qcur = b.qcur - 1
                        b2.qcur = b2.qcur + 1
                        shifts.Add b2.Code, b.Code
                        b2.GetFrom = False
                        b2.HavePut = True
                        If b2.qcur = b2.qmax Then
                          b2.PutTo = False
                        End If
                        GoTo nextpal3
                    End If
                  End If
                'End If
              End If
            End If
          Next
        End If
        DoEvents
        ShowSklad
      End If
  Next
  
  OptimizeShifts
   
  qmax = 0
  qcur = 0
  qz = 0
  cost = 0
  
  For i = 1 To shifts.Count
    cost = cost + st.Item(shifts.Item(i).FromCode).GetCost(st.Item(shifts.Item(i).FromCode).T = st.Item(shifts.Item(i).ToCode).T)
    cost = cost + st.Item(shifts.Item(i).ToCode).PutCost(st.Item(shifts.Item(i).FromCode).T = st.Item(shifts.Item(i).ToCode).T)
    'Debug.Print shifts.Item(i).FromCode & "->" & shifts.Item(i).ToCode
  Next
  Debug.Print "Test after"
  For i = 1 To st.Count
      Set b = st.Item(i)
      qmax = qmax + b.qmax
      qcur = qcur + b.qcur
      
      If b.qcur = 0 Then
        qz = qz + 1
      End If
      If b.qcur > 0 And b.qcur < b.qmax Then
        Debug.Print b.Code, b.qmax, b.qcur, b.GetFrom
      End If
      
  Next
  Text3.Text = qmax & ":" & qcur & ":" & qz & ":" & CInt(cost / 60) & " h:" & shifts.Count & " shifts"
  ShowSklad
End Sub

Private Sub Command4_Click()
 
  Radius = Val(txtlevel) * Val(txtlevel)
  Dim mx As Double, my As Double, mz As Double
  
  Set shifts = New BoxShifts
  
  Dim i As Long, j As Long
  Dim qmax As Long, qcur As Long, qz As Long, cost As Long
  
  Dim b As Box, b2 As Box, bs As BoxShift
  
  If Val("0" & txtX) = 0 And Val("0" & txtY) = 0 And Val("0" & txtZ) = 0 Then
    mx = 0
    my = 0
    mz = 0
    qz = 0
    For i = 1 To st.Count
        Set b = st.Item(i)
        If b.qcur > 0 Then
          mx = mx + b.X * b.qcur
          my = my + b.Y * b.qcur
          mz = mz + b.Z * b.qcur
          qz = qz + b.qcur
        End If
        
    Next
    mx = mx / qz
    my = my / qz
    mz = mz / qz
  Else
    mx = Val("0" & txtX)
    my = Val("0" & txtY)
    mz = Val("0" & txtZ)
    If mx < 0 Or mx > 61 Then
      mx = 31
    End If
    If my < 0 Or my > 21 Then
      mx = 11
    End If
    If mz < 0 Or mx > 11 Then
      mz = 5
    End If
    
  End If
  
  
  
   
  Dim movecnt As Long
    
  movecnt = 1
  While movecnt > 0
      movecnt = 0
      For i = 1 To st.Count
          Set b = st.Item(i)
          If (b.X - mx) * (b.X - mx) + (b.Y - my) * (b.Y - my) + (b.Z - mz) * (b.Z - mz) > Radius And b.qcur > 0 Then
nextpal3:
            If b.qcur > 0 Then
              For j = 1 To st.Count
                If j <> i Then
                  Set b2 = st.Item(j)
                  If b2.qcur < b2.qmax Then
                      If (b2.X - mx) * (b2.X - mx) + (b2.Y - my) * (b2.Y - my) + (b2.Z - mz) * (b2.Z - mz) <= Radius Then
                            b.qcur = b.qcur - 1
                            b2.qcur = b2.qcur + 1
                            shifts.Add b2.Code, b.Code
                            movecnt = movecnt + 1
                            b2.HavePut = True
                            DoEvents
                            ShowSklad
                            GoTo nextpal3
                      End If
                  End If
                End If
              Next
            
            End If
        
          End If
      Next
      ShowSklad
  Wend
  
 
  
   For i = 1 To st.Count
    Set b = st.Item(i)
    If b.qcur > 0 Then
      If b.qcur < b.qmax Then
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
     If b.HavePut Then
        b.GetFrom = False
      End If
  Next

  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.GetFrom Then
nextpal4:
        If b.qcur > 0 Then
          For j = 1 To st.Count
            If j <> i Then
              Set b2 = st.Item(j)
              If b2.PutTo Then
                If b2.qcur > 0 Then
                  'If b2.T = b.T Then

                    If b2.qcur < b2.qmax Then
                        b.qcur = b.qcur - 1
                        b2.qcur = b2.qcur + 1
                        shifts.Add b2.Code, b.Code
                        b2.HavePut = True
                        b2.GetFrom = False
                        If b2.qcur = b2.qmax Then
                          b2.PutTo = False
                        End If
                        DoEvents
                        ShowSklad
                        GoTo nextpal4
                    End If
                  'End If
                End If
              End If
              
            End If
          Next
         
        End If
      End If
  Next
  
  OptimizeShifts
  
  qmax = 0
  qcur = 0
  qz = 0
  cost = 0
  
  For i = 1 To shifts.Count
    cost = cost + st.Item(shifts.Item(i).FromCode).GetCost(st.Item(shifts.Item(i).FromCode).T = st.Item(shifts.Item(i).ToCode).T)
    cost = cost + st.Item(shifts.Item(i).ToCode).PutCost(st.Item(shifts.Item(i).FromCode).T = st.Item(shifts.Item(i).ToCode).T)
    'Debug.Print shifts.Item(i).FromCode & "->" & shifts.Item(i).ToCode
  Next
  For i = 1 To st.Count
      Set b = st.Item(i)
      qmax = qmax + b.qmax
      qcur = qcur + b.qcur
      
      If b.qcur = 0 Then
        qz = qz + 1
      End If
      If b.qcur > 0 And b.qcur < b.qmax Then
        Debug.Print b.Code, b.qmax, b.qcur, b.GetFrom
      End If
      
  Next
  Text4.Text = qmax & ":" & qcur & ":" & qz & ":" & CInt(cost / 60) & " h:" & shifts.Count & " shifts"
  ShowSklad
End Sub



Private Sub ShowSklad()
  Dim i As Integer, j As Long, k As Long

  Dim qmax As Long, qcur As Long, qz As Long, cost As Long
  
  Dim b As Box, b2 As Box, bs As BoxShift
  'pOut.Cls
  If st Is Nothing Then Exit Sub
  
  Dim cv(0 To 61, 0 To 21, 0 To 7) As Double
  Dim mv(0 To 61, 0 To 21, 0 To 7) As Double
  
  For i = 1 To st.Count
      Set b = st.Item(i)
      If b.qmax > 0 Then
        cv(b.X, b.Y, b.Z) = cv(b.X, b.Y, b.Z) + b.qcur
        mv(b.X, b.Y, b.Z) = mv(b.X, b.Y, b.Z) + b.qmax
     
      End If

  Next
  
  For i = 0 To 60
    For j = 0 To 20
    
           
      pOut.Line (i * 15 * Screen.TwipsPerPixelX, j * 15 * Screen.TwipsPerPixelY)-((i + 1) * 15 * Screen.TwipsPerPixelX, (j + 1) * 15 * Screen.TwipsPerPixelY), RGB(255, 255, 255), BF
    
      For k = 0 To 6
        If mv(i, j, k) <> 0 Then
          If cv(i, j, k) > 0 Then
            pOut.Line _
            (i * 15 * Screen.TwipsPerPixelX + 2 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY + k * 2 * Screen.TwipsPerPixelY + j * 15 * Screen.TwipsPerPixelY)- _
            ((cv(i, j, k) / mv(i, j, k) * 13 * Screen.TwipsPerPixelX) + (i) * 15 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY + (j) * 15 * Screen.TwipsPerPixelY + k * 2 * Screen.TwipsPerPixelY), vbRed
          Else
                 pOut.Line _
          (i * 15 * Screen.TwipsPerPixelX + 2 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY + k * 2 * Screen.TwipsPerPixelY + j * 15 * Screen.TwipsPerPixelY)- _
          (13 * Screen.TwipsPerPixelX + i * 15 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY + (j) * 15 * Screen.TwipsPerPixelY + k * 2 * Screen.TwipsPerPixelY), vbBlack

          End If
          
        Else
          pOut.Line _
          (i * 15 * Screen.TwipsPerPixelX + 2 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY + k * 2 * Screen.TwipsPerPixelY + j * 15 * Screen.TwipsPerPixelY)- _
          (13 * Screen.TwipsPerPixelX + i * 15 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY + (j) * 15 * Screen.TwipsPerPixelY + k * 2 * Screen.TwipsPerPixelY), RGB(192, 192, 192)
        End If
        

      Next
    Next
  Next
End Sub

Private Sub Command5_Click()
ShowSklad
End Sub





Private Sub Form_Resize()
On Error Resume Next
pOut.Left = 0
pOut.Height = Me.ScaleHeight - pOut.Top
pOut.Width = Me.ScaleWidth
End Sub
