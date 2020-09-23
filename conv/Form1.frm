VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Convolution Of Two Signals"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Output1 
      Height          =   1935
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   120
      Width           =   4815
   End
   Begin VB.ComboBox output 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox accum 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox SignalB 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox SignalA 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox MatB 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "1,2,3"
      Top             =   2760
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONVOLUTE THE TWO SIGNALS"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox MatA 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "1,2,3"
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "Vote Here"
      BeginProperty Font 
         Name            =   "Script"
         Size            =   24
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      MouseIcon       =   "Form1.frx":0E42
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":114C
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "Signal B"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Signal A"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim counter As Integer
Dim Index_A As Integer
Dim Index_B As Integer
Dim MatrixA As String
Dim MatrixB As String
Dim prev_semi_colon As Integer
Dim next_semi_colon As Integer
Dim diff As Integer
Dim Element As String
Dim counter1 As Double
Dim counter2 As Double
Dim LengthA As Integer
Dim LengthB As Integer
Dim OP_element As String
Dim Mat_Counter As Integer
Dim ml As String
Dim mr As String
Dim u As Integer
Dim Ocounter As Integer
Dim INcounter As Integer
Dim LENGTH_COUNTER As Integer
Dim no_of_semi As Integer
Dim counter_ckeck_semi As Integer
Dim Matrix As String
Dim occurance As Integer
Dim count_semi_colons As Integer
Dim Ref As Integer
Dim InvRef As Integer
Dim Middle As String
'Dim AconvB As String
Dim counter3 As Integer
Dim AconvlB As String
Dim vote As String



Private Sub Command1_Click()
    Call Execute
End Sub



Public Function Add_to_list_A(Element)
    Call SignalA.AddItem(Element, Index_A)
    Index_A = Index_A + 1
End Function

Public Function Add_to_list_B(Element)
    Call SignalB.AddItem(Element, Index_B)
    Index_B = Index_B + 1
End Function



Function get_element(Matrix, Length, Choice)
    count_semi_colons = 0
    For Mat_Counter = 1 To Length
    
        prev_semi_colon = 0
        ml = Left(Matrix, Mat_Counter)
        mr = Right(ml, 1)
    
        If mr = Chr(44) Then
            count_semi_colons = count_semi_colons + 1
            prev_semi_colon = next_semi_colon
            next_semi_colon = Mat_Counter
            diff = Abs(next_semi_colon - prev_semi_colon - 1)
            Element = Right(Left(Matrix, next_semi_colon - 1), diff)
        
            Select Case Choice
                Case "A"
                Call Add_to_list_A(Element)
                Case "B"
                Call Add_to_list_B(Element)
            End Select
             
        End If
        
        If Mat_Counter = Length Then
            If count_semi_colons = 0 Then
              Element = Matrix

                 Select Case Choice
                     Case "A"
                     Call Add_to_list_A(Element)
                     Case "B"
                     Call Add_to_list_B(Element)
                 End Select
            Else
            
          
                next_semi_colon = Length
                Element = Right(Left(Matrix, next_semi_colon), diff)

                 Select Case Choice
                     Case "A"
                     Call Add_to_list_A(Element)
                     Case "B"
                     Call Add_to_list_B(Element)
                 End Select

            End If
    End If
    Next Mat_Counter
'Print count_semi_colons
End Function





Function get_SignalA(i)
    get_SignalA = Val(SignalA.List(i - 1))
End Function

Function AconvB(i, value)
    Call output.AddItem(value, i - 1)
End Function


Function accumulator(i, value)
    Call accum.AddItem(value, i - 1)
End Function


Function get_SignalB(i)
    get_SignalB = Val(SignalB.List(i - 1))
End Function


Function get_accumulator(i)
    get_accumulator = Val(accum.List(i - 1))
End Function


Public Sub Execute()
    SignalA.Clear
    SignalB.Clear
    output.Clear
    accum.Clear
    Index_A = 0
    Index_B = 0
    
    MatrixA = MatA.Text
    Call get_element(MatrixA, Len(MatrixA), "A")

    MatrixB = MatB.Text
    Call get_element(MatrixB, Len(MatrixB), "B")


    LengthA = SignalA.ListCount
    LengthB = SignalB.ListCount
    
    
    
    Call check_length(LengthA, LengthB)
    
    For counter1 = 0 To LengthA - 1
        
         For counter2 = 0 To counter1 - 1
            Call accumulator(counter2 + 1, 0)
         Next counter2


        For counter2 = counter1 To LengthA + LengthB - 2
            u = u + 1
            Call accumulator((counter2 + 1), get_SignalA(counter1 + 1) * get_SignalB(u))
        Next counter2
        u = 0

    Next counter1

    For Ocounter = 1 To LengthA + LengthB - 1
         OP_element = 0
            
         For INcounter = 1 To LengthA
             OP_element = OP_element + get_accumulator(Ocounter + (INcounter - 1) * (LengthA + LengthB - 1))
             'Print Ocounter + (INcounter - 1) * 5
         Next INcounter

         Call AconvB((Ocounter), OP_element)

    Next Ocounter
    Call AddToOutput
End Sub


Function check_length(LengthA, LengthB)
If LengthA <> LengthB Then
    
   If LengthA > LengthB Then
    
            For LENGTH_COUNTER = 1 To Abs(LengthA - LengthB)
                Call Add_to_list_B(0)
            Next LENGTH_COUNTER
   Else
            For LENGTH_COUNTER = 1 To Abs(LengthA - LengthB)
                Call Add_to_list_A(0)
            Next LENGTH_COUNTER
   End If

End If
End Function


Function AddToOutput()
    Call GetOutput
    Output1.Text = "A=" & vbCrLf & "[ " & MatA.Text & " ]" & vbCrLf & vbCrLf _
                   & "B=" & vbCrLf & "[ " & MatB.Text & " ]" & vbCrLf & vbCrLf _
                   & "A conv B = " & "[" & AconvlB & " ]"
    AconvlB = ""
End Function

Private Sub GetOutput()
    For counter3 = 0 To output.ListCount - 1
        AconvlB = AconvlB & "  " & output.List(counter3)
    Next counter3
End Sub

Private Sub Label4_Click()
vote = "http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=42795&lngWId=1"
Call RunBrowser(vote, 10, 1)
End Sub
