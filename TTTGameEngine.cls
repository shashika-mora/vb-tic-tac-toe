VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event ComputerWins(x As Byte, y As Byte)
Event UserWins(x As Byte, y As Byte)
Private TicTacToe(3, 3) As Byte
Private isWon As Boolean
Private Type pos
 x As Byte
 y As Byte
End Type
Public Sub Initialize()
 For i = 0 To 2
 For j = 0 To 2
 TicTacToe(i, j) = 0
 Next
 Next
 isWon = False
End Sub
Function FindXwin() As Boolean
 Dim win As Boolean
 win = TicTacToe(0, 0) = 1 And TicTacToe(1, 0) = 1 And TicTacToe(2, 0) = 1
 win = win Or TicTacToe(0, 1) = 1 And TicTacToe(1, 1) = 1 And TicTacToe(2, 1) = 1
 win = win Or TicTacToe(0, 2) = 1 And TicTacToe(1, 2) = 1 And TicTacToe(2, 2) = 1
 win = win Or TicTacToe(1, 0) = 1 And TicTacToe(1, 1) = 1 And TicTacToe(1, 2) = 1
 win = win Or TicTacToe(2, 0) = 1 And TicTacToe(2, 1) = 1 And TicTacToe(2, 2) = 1
win = win Or TicTacToe(0, 0) = 1 And TicTacToe(0, 1) = 1 And TicTacToe(0, 2) = 1
 win = win Or TicTacToe(0, 0) = 1 And TicTacToe(1, 1) = 1 And TicTacToe(2, 2) = 1
 win = win Or TicTacToe(0, 2) = 1 And TicTacToe(1, 1) = 1 And TicTacToe(2, 0) = 1
 If win Then
 FindXwin = True
 isWon = True
 End If
End Function
Function FindOwin() As Boolean
 Dim win As Boolean
 win = TicTacToe(0, 0) = 2 And TicTacToe(1, 0) = 2 And TicTacToe(2, 0) = 2
 win = win Or TicTacToe(0, 1) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(2, 1) = 2
 win = win Or TicTacToe(0, 2) = 2 And TicTacToe(1, 2) = 2 And TicTacToe(2, 2) = 2
 win = win Or TicTacToe(1, 0) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(1, 2) = 2
 win = win Or TicTacToe(2, 0) = 2 And TicTacToe(2, 1) = 2 And TicTacToe(2, 2) = 2
 win = win Or TicTacToe(0, 0) = 2 And TicTacToe(0, 1) = 2 And TicTacToe(0, 2) = 2
 win = win Or TicTacToe(0, 0) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(2, 2) = 2
 win = win Or TicTacToe(0, 2) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(2, 0) = 2
 If win Then
 FindOwin = True
 isWon = True
 End If
End Function
Function ComputerMove(ByRef x, ByRef y) As Boolean
 Dim gpos As pos

'This Function Is Responsible For the Logic Of the
'Game It Finds out Whether There is a move so
'computer can win the game else It finds is there is a
'way to stop user from wining, else if will do a
'random Move
 If isWon Then
 GameLogic = False
 Exit Function
 End If
 If TicTacToe(0, 0) = 2 And TicTacToe(1, 0) = 2 And TicTacToe(2, 0) = 0 Then
 gpos.x = 2
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 0) = 2 And TicTacToe(2, 0) = 2 And TicTacToe(0, 0) = 0 Then
 gpos.x = 0
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 2 And TicTacToe(2, 0) = 2 And TicTacToe(1, 0) = 0 Then
 gpos.x = 1
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 2 And TicTacToe(0, 1) = 2 And TicTacToe(0, 2) = 0 Then
 gpos.x = 0
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 1) = 2 And TicTacToe(0, 2) = 2 And TicTacToe(0, 0) = 0 Then
 gpos.x = 0
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 2 And TicTacToe(0, 2) = 2 And TicTacToe(0, 1) = 0 Then
 gpos.x = 0
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 1) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(2, 1) = 0 Then
 gpos.x = 2
gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 1) = 2 And TicTacToe(2, 1) = 2 And TicTacToe(0, 1) = 0 Then
 gpos.x = 0
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 1) = 2 And TicTacToe(2, 1) = 2 And TicTacToe(1, 1) = 0 Then
 gpos.x = 1
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 0) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(1, 2) = 0 Then
 gpos.x = 1
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 0) = 2 And TicTacToe(1, 2) = 2 And TicTacToe(1, 1) = 0 Then
 gpos.x = 1
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 1) = 2 And TicTacToe(1, 2) = 2 And TicTacToe(1, 0) = 0 Then
 gpos.x = 1
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 2) = 2 And TicTacToe(1, 2) = 2 And TicTacToe(2, 2) = 0 Then
 gpos.x = 2
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 2) = 2 And TicTacToe(2, 2) = 2 And TicTacToe(0, 2) = 0 Then
 gpos.x = 0
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 2) = 2 And TicTacToe(2, 2) = 2 And TicTacToe(1, 2) = 0 Then
 gpos.x = 1
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 0) = 2 And TicTacToe(2, 1) = 2 And TicTacToe(2, 2) = 0 Then
 gpos.x = 2
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 1) = 2 And TicTacToe(2, 2) = 2 And TicTacToe(2, 0) = 0 Then
 gpos.x = 2
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 0) = 2 And TicTacToe(2, 2) = 2 And TicTacToe(2, 1) = 0 Then
 gpos.x = 2
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(2, 2) = 0 Then
 gpos.x = 2
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 2 And TicTacToe(2, 2) = 2 And TicTacToe(1, 1) = 0 Then
 gpos.x = 1
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 1) = 2 And TicTacToe(2, 2) = 2 And TicTacToe(0, 0) = 0 Then
 gpos.x = 0
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 0) = 2 And TicTacToe(1, 1) = 2 And TicTacToe(0, 2) = 0 Then
 gpos.x = 0
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 1) = 2 And TicTacToe(0, 2) = 2 And TicTacToe(2, 0) = 0 Then
 gpos.x = 2
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 0) = 2 And TicTacToe(0, 2) = 2 And TicTacToe(1, 1) = 0 Then
 gpos.x = 1
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 1 And TicTacToe(1, 0) = 1 And TicTacToe(2, 0) = 0 Then
 gpos.x = 2
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 0) = 1 And TicTacToe(2, 0) = 1 And TicTacToe(0, 0) = 0 Then
 gpos.x = 0
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 1 And TicTacToe(2, 0) = 1 And TicTacToe(1, 0) = 0 Then
 gpos.x = 1
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 1 And TicTacToe(0, 1) = 1 And TicTacToe(0, 2) = 0 Then
 gpos.x = 0
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 1) = 1 And TicTacToe(0, 2) = 1 And TicTacToe(0, 0) = 0 Then
 gpos.x = 0
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 1 And TicTacToe(0, 2) = 1 And TicTacToe(0, 1) = 0 Then
 gpos.x = 0
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 1) = 1 And TicTacToe(1, 1) = 1 And TicTacToe(2, 1) = 0 Then
 gpos.x = 2
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
ElseIf TicTacToe(2, 0) = 1 And TicTacToe(2, 1) = 1 And TicTacToe(2, 2) = 0 Then
 gpos.x = 2
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 1) = 1 And TicTacToe(2, 2) = 1 And TicTacToe(2, 0) = 0 Then
 gpos.x = 2
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 0) = 1 And TicTacToe(2, 2) = 1 And TicTacToe(2, 1) = 0 Then
 gpos.x = 2
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 1 And TicTacToe(1, 1) = 1 And TicTacToe(2, 2) = 0 Then
 gpos.x = 2
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(0, 0) = 1 And TicTacToe(2, 2) = 1 And TicTacToe(1, 1) = 0 Then
 gpos.x = 1
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 1) = 1 And TicTacToe(2, 2) = 1 And TicTacToe(0, 0) = 0 Then
 gpos.x = 0
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(2, 0) = 1 And TicTacToe(1, 1) = 1 And TicTacToe(0, 2) = 0 Then
 gpos.x = 0
 gpos.y = 2
 TicTacToe(gpos.x, gpos.y) = 2
 ElseIf TicTacToe(1, 1) = 1 And TicTacToe(0, 2) = 1 And TicTacToe(2, 0) = 0 Then
 gpos.x = 2
 gpos.y = 0
 TicTacToe(gpos.x, gpos.y) = 2
ElseIf TicTacToe(2, 0) = 1 And TicTacToe(0, 2) = 1 And TicTacToe(1, 1) = 0 Then
 gpos.x = 1
 gpos.y = 1
 TicTacToe(gpos.x, gpos.y) = 2
 Else

 Dim isNotFill As Boolean
 For i = 0 To 2
 For j = 0 To 2
 If TicTacToe(i, j) = 0 Then
 isNotFill = True
 GoTo ExFor
 End If
 Next j
 Next i
ExFor:
 If Not isNotFill Then
 gpos.x = 4
 gpos.y = 4

 x = gpos.x
 y = gpos.y
 GameLogic = False
 Exit Function
 End If

 Do
 Randomize Now
 gpos.x = Round(Rnd * 2, 0)
 gpos.y = Round(Rnd * 2, 0)
 x = gpos.x
 y = gpos.y
 DoEvents
 Loop While TicTacToe(gpos.x, gpos.y) <> 0

 TicTacToe(gpos.x, gpos.y) = 2
 End If
 If FindOwin Then RaiseEvent ComputerWins(gpos.x, gpos.y)
 x = gpos.x
 y = gpos.y
 GameLogic = True

End Function
Function UserMove(x As Byte, y As Byte) As Boolean
 If isWon Then Exit Function

 If TicTacToe(x, y) = 0 Then
 TicTacToe(x, y) = 1

 UserMove = True
 If FindXwin Then RaiseEvent UserWins(x, y)
 Exit Function
 Else
 UserMove = False
 Exit Function
 End If
End Function
Sub PrintTicTacToe()
 Debug.Print "*********************** Tic Tac Toe" & "********************"
 Debug.Print TicTacToe(0, 0) & " " & TicTacToe(0, 1) & " " & TicTacToe(0, 2)
 Debug.Print TicTacToe(1, 0) & " " & TicTacToe(1, 1) & " " & TicTacToe(1, 2)
 Debug.Print TicTacToe(2, 0) & " " & TicTacToe(2, 1) & " " & TicTacToe(2, 2)
 Debug.Print "***********************************************" & "**********"
End Sub
