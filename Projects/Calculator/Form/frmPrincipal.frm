VERSION 5.00
Begin VB.Form frmPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..:: Calculator ::.."
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDisplayCalc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   3200
   End
   Begin VB.CommandButton cmdPoint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "."
      Height          =   720
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   750
   End
   Begin VB.CommandButton cmdZero 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "0"
      Height          =   720
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   750
   End
   Begin VB.CommandButton cmdEqual 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "="
      Height          =   720
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   750
   End
   Begin VB.CommandButton cmdDivision 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "/"
      Height          =   720
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   510
   End
   Begin VB.CommandButton cmdSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "-"
      Height          =   720
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   510
   End
   Begin VB.CommandButton cmdSum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "+"
      Height          =   720
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   510
   End
   Begin VB.CommandButton cmdOne 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "1"
      Height          =   720
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   750
   End
   Begin VB.CommandButton cmdEight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "8"
      Height          =   720
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   750
   End
   Begin VB.CommandButton cmdFive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "5"
      Height          =   720
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   750
   End
   Begin VB.CommandButton cmdTwo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "2"
      Height          =   720
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   750
   End
   Begin VB.CommandButton cmdNine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "9"
      Height          =   720
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   750
   End
   Begin VB.CommandButton cmdSix 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "6"
      Height          =   720
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   750
   End
   Begin VB.CommandButton cmdThree 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "3"
      Height          =   720
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   750
   End
   Begin VB.CommandButton cmdMulti 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "*"
      Height          =   720
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   510
   End
   Begin VB.CommandButton cmdFour 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "4"
      Height          =   720
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   750
   End
   Begin VB.CommandButton cmdSeven 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "7"
      Height          =   720
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   750
   End
   Begin VB.CommandButton cmdBackspace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "<"
      Height          =   720
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   3150
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   240
      Width           =   3200
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const numbers As String = "1,2,3,4,5,6,7,8,9"
Private Const operators As String = "+,-,*,/"
Private Const INITIAL_VALUE As String = "0"
Private rebootDisplay As Boolean

Private Sub cmdBackspace_Click()
    txtDisplay.Text = Right(txtDisplay.Text, Len(txtDisplay.Text) - 1)
    If (txtDisplay.Text = vbNullString) Then txtDisplay.Text = INITIAL_VALUE
End Sub

Private Sub cmdDivision_Click()
    PrintOperator
End Sub

Private Sub cmdEight_Click()
    PrintNumber
End Sub

Private Sub cmdEqual_Click()
    If (txtDisplayCalc.Text = vbNullString) Then Exit Sub
    ExecuteCalc
    txtDisplayCalc.Text = vbNullString
    rebootDisplay = True
End Sub

Private Sub cmdFive_Click()
    PrintNumber
End Sub

Private Sub cmdFour_Click()
    PrintNumber
End Sub

Private Sub cmdMulti_Click()
    PrintOperator
End Sub

Private Sub cmdNine_Click()
    PrintNumber
End Sub

Private Sub cmdOne_Click()
    PrintNumber
End Sub

Private Sub cmdPoint_Click()
    PrintPoint
End Sub

Private Sub cmdSeven_Click()
    PrintNumber
End Sub

Private Sub cmdSix_Click()
    PrintNumber
End Sub

Private Sub cmdSub_Click()
    PrintOperator
End Sub

Private Sub cmdSum_Click()
    PrintOperator
End Sub

Private Sub cmdThree_Click()
    PrintNumber
End Sub

Private Sub cmdTwo_Click()
    PrintNumber
End Sub

Private Sub cmdZero_Click()
    PrintNumber
End Sub

Private Sub PrintNumber()
    Dim number As String
    number = Screen.ActiveControl.Caption
    If (InStr(1, numbers, number) < 0) Then Exit Sub
    If (txtDisplay.Text = INITIAL_VALUE And number = INITIAL_VALUE) Then Exit Sub
    If (txtDisplay.Text = INITIAL_VALUE And number <> INITIAL_VALUE) Then txtDisplay.Text = vbNullString
    If (rebootDisplay) Then txtDisplay.Text = vbNullString: rebootDisplay = False
    txtDisplay.Text = txtDisplay.Text & number
End Sub

Private Sub PrintPoint()
    If (txtDisplay.MaxLength = Len(txtDisplay.Text) + 1) Then Exit Sub
    If (InStr(1, txtDisplay.Text, ".") > 0) Then Exit Sub
    txtDisplay.Text = txtDisplay.Text & "."
End Sub

Private Sub PrintOperator()
    Dim operator As String
    operator = Screen.ActiveControl.Caption
    If (txtDisplay.Text = INITIAL_VALUE And txtDisplayCalc.Text = vbNullString) Then Exit Sub
    If (txtDisplay.Text = INITIAL_VALUE And txtDisplayCalc.Text <> vbNullString) Then
        txtDisplayCalc.Text = Left(txtDisplayCalc.Text, Len(txtDisplayCalc.Text) - 1)
        txtDisplayCalc.Text = txtDisplayCalc.Text & operator
        Exit Sub
    End If
    txtDisplayCalc.Text = txtDisplayCalc.Text & " " & txtDisplay.Text & " " & operator
    txtDisplay.Text = INITIAL_VALUE
End Sub

Private Function IsOperator(value As String) As Boolean
    If (value = vbNullString) Then IsOperator = False: Exit Function
    IsOperator = (InStr(1, operators, value) > 0)
End Function

Private Sub ExecuteCalc()
    Dim values As Variant
    Dim value As Variant
    
    Dim result As Double
    
    Dim numbers As Collection
    Dim operators As Collection
    
    Set numbers = New Collection
    Set operators = New Collection
    
    values = Split(txtDisplayCalc.Text & " " & txtDisplay.Text, " ")
    For Each value In values
        If (IsNumeric(CStr(value))) Then
            numbers.Add CDbl(Replace(value, ".", ","))
        End If
    Next
    
    result = numbers(1)
    Call numbers.Remove(1)
    
    For Each value In values
        If (IsOperator(CStr(value))) Then
            Select Case value
                Case "+"
                    result = result + numbers(1)
                Case "-"
                    result = result - numbers(1)
                Case "/"
                    If (numbers(1) = 0) Then
                        txtDisplay.Text = "err_div_zero"
                        txtDisplayCalc.Text = ""
                        rebootDisplay = True
                        Exit Sub
                    Else
                        result = result / numbers(1)
                    End If
                Case "*"
                    result = result * numbers(1)
            End Select
            Call numbers.Remove(1)
        End If
    Next
    
    txtDisplay.Text = Replace(result, ",", ".")
    
    Set numbers = Nothing
    Set operators = Nothing
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 'zero
            cmdZero.SetFocus
            cmdZero_Click
        Case 49 'one
            cmdOne.SetFocus
            cmdOne_Click
        Case 50 'two
            cmdTwo.SetFocus
            cmdTwo_Click
        Case 51 'three
            cmdThree.SetFocus
            cmdThree_Click
        Case 52 'four
            cmdFour.SetFocus
            cmdFour_Click
        Case 53 'five
            cmdFive.SetFocus
            cmdFive_Click
        Case 54 'six
            cmdSix.SetFocus
            cmdSix_Click
        Case 55 'seven
            cmdSeven.SetFocus
            cmdSeven_Click
        Case 56 'eight
            cmdEight.SetFocus
            cmdEight_Click
        Case 57 'nine
            cmdNine.SetFocus
            cmdNine_Click
        Case 43 'sum
            cmdSum.SetFocus
            cmdSum_Click
        Case 45 'sub
            cmdSub.SetFocus
            cmdSub_Click
        Case 47 'div
            cmdDivision.SetFocus
            cmdDivision_Click
        Case 42 'mult
            cmdMulti.SetFocus
            cmdMulti_Click
        Case 61, 13 'equal
            cmdEqual.SetFocus
            cmdEqual_Click
        Case 46 'point
            cmdPoint.SetFocus
            cmdPoint_Click
        Case 8
            cmdBackspace.SetFocus
            cmdBackspace_Click
        Case 27
            End
    End Select
End Sub

