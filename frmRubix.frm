VERSION 5.00
Begin VB.Form frmRubix 
   Caption         =   "Rubix Cube"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5565
   Icon            =   "frmRubix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTwistLeft3 
      Height          =   495
      Left            =   1560
      Picture         =   "frmRubix.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Rotate Down"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistLeft2 
      Height          =   495
      Left            =   1560
      Picture         =   "frmRubix.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Rotate Down"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistLeft1 
      Height          =   495
      Left            =   1560
      Picture         =   "frmRubix.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Rotate Down"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistup3 
      Height          =   495
      Left            =   3000
      Picture         =   "frmRubix.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Rotate Down"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistup2 
      Height          =   495
      Left            =   2520
      Picture         =   "frmRubix.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Rotate Down"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistup1 
      Height          =   495
      Left            =   2040
      Picture         =   "frmRubix.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Rotate Down"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistAcross3 
      Height          =   495
      Left            =   3480
      Picture         =   "frmRubix.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Rotate Down"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistAcross2 
      Height          =   495
      Left            =   3480
      Picture         =   "frmRubix.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Rotate Down"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistAcross1 
      Height          =   495
      Left            =   3480
      Picture         =   "frmRubix.frx":2652
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Rotate Down"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistDown3 
      Height          =   495
      Left            =   3000
      Picture         =   "frmRubix.frx":2A94
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Rotate Down"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistDown2 
      Height          =   495
      Left            =   2520
      Picture         =   "frmRubix.frx":2ED6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Rotate Down"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdTwistDown1 
      Height          =   495
      Left            =   2040
      Picture         =   "frmRubix.frx":3318
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Rotate Down"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   3000
      Top             =   6480
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   2520
      Top             =   6480
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   2040
      Top             =   6480
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   3000
      Top             =   6000
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   2520
      Top             =   6000
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   2040
      Top             =   6000
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   3000
      Top             =   5520
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   2520
      Top             =   5520
      Width           =   495
   End
   Begin VB.Shape side6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   2040
      Top             =   5520
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   3000
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   2520
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   2040
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   3000
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   2520
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   2040
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   3000
      Top             =   3960
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   2520
      Top             =   3960
      Width           =   495
   End
   Begin VB.Shape side5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   2040
      Top             =   3960
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   1080
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   600
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   120
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   1080
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   600
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   120
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   1080
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   600
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   4920
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   4440
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   3960
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   4920
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   4440
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   3960
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   4920
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   4440
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   3960
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   3000
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   2520
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   2040
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   3000
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   2520
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   2040
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   3000
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   2520
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   2040
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   3000
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   2520
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   2040
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   3000
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   2520
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   2040
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   3000
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   2520
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape side1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   2040
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFileScramble 
         Caption         =   "&Scramble"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmRubix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'+----------------------------------+
'|@@@@ A very simple rubix cube @@@@|
'+----------------------------------+
'|By Ian McCarthy                   |
'|    oeb@cotse.com                 |
'|       http://oeb.fwaggle.net     |
'| Steal it, Break it, and by all   |
'| Means if you find someone stupid |
'| Enough to buy it ..... SELL IT   |
'+----------------------------------+

Dim Face1(1 To 3, 1 To 3) As Integer
Dim Face2(1 To 3, 1 To 3) As Integer
Dim Face3(1 To 3, 1 To 3) As Integer
Dim Face4(1 To 3, 1 To 3) As Integer
Dim Face5(1 To 3, 1 To 3) As Integer
Dim Face6(1 To 3, 1 To 3) As Integer
Dim i As Integer
Dim j As Integer
Dim ShiftHold(1 To 3) As Integer
Dim unf As Integer
Dim k As Integer
Dim l As Integer
Dim z As Integer

Dim RandVar1 As Integer
Dim RandVar2 As Integer
Dim randVar3 As Integer

Private Sub SortCube()
'Sorts the cube to fil each array with a single colour
'Side 1 is Red
'Side 2 is Green
'Side 3 is Blue
'Side 4 is Black
'Side 5 is White
'And side 6 is Yellow
'These are kept as integers for ease of use later
For i = 1 To 3
    For j = 1 To 3
        Face1(i, j) = 1
    Next j
Next i
    
For i = 1 To 3
    For j = 1 To 3
        Face2(i, j) = 2
    Next j
Next i

For i = 1 To 3
    For j = 1 To 3
        Face3(i, j) = 3
    Next j
Next i

For i = 1 To 3
    For j = 1 To 3
        Face4(i, j) = 4
    Next j
Next i

For i = 1 To 3
    For j = 1 To 3
        Face5(i, j) = 5
    Next j
Next i

For i = 1 To 3
    For j = 1 To 3
        Face6(i, j) = 6
    Next j
Next i

End Sub

Private Sub ShiftDown(ByVal side As Integer)
'Rotates the indecated row Downwards
For i = 1 To 3
    ShiftHold(i) = Face6(i, side)
    Face6(i, side) = Face5(i, side)
    Face5(i, side) = Face3(i, side)
    Face3(i, side) = Face1(i, side)
    Face1(i, side) = ShiftHold(i)
Next i
End Sub

Private Sub ShiftUp(ByVal side As Integer)
'Rotates the indecated row Upwards
For i = 1 To 3
    ShiftHold(i) = Face1(i, side)
    Face1(i, side) = Face3(i, side)
    Face3(i, side) = Face5(i, side)
    Face5(i, side) = Face6(i, side)
    Face6(i, side) = ShiftHold(i)
Next i
End Sub

Private Sub ShiftAcrossRight(ByVal side As Integer)
'Rotates the indecated row to the right (Note the bottom array (6) infact
'Rotates Backwards
If side = 1 Then
    k = 3
    ElseIf side = 2 Then
        k = 2
    ElseIf side = 3 Then
        k = 1
End If
For i = 1 To 3
    If i = 1 Then
        l = 3
        ElseIf i = 2 Then
            l = 2
        ElseIf i = 3 Then
            l = 1
    End If
    
    ShiftHold(i) = Face2(side, i)
    Face2(side, i) = Face6(k, l)
    Face6(k, l) = Face4(side, i)
    Face4(side, i) = Face3(side, i)
    Face3(side, i) = ShiftHold(i)
Next i
End Sub

Private Sub ShiftAcrossLeft(ByVal side As Integer)
'Rotates the indecated row to the Left (Note the bottom array (6) infact
'Rotates Backwards
If side = 1 Then
    k = 3
    ElseIf side = 2 Then
        k = 2
    ElseIf side = 3 Then
        k = 1
End If
For i = 1 To 3
    If i = 1 Then
        l = 3
        ElseIf i = 2 Then
            l = 2
        ElseIf i = 3 Then
            l = 1
    End If
    
    ShiftHold(i) = Face3(side, i)
    Face3(side, i) = Face4(side, i)
    Face4(side, i) = Face6(k, l)
    Face6(k, l) = Face2(side, i)
    Face2(side, i) = ShiftHold(i)
Next i
End Sub

'All this is to deal with the movement buttons, and just calls the
'needed sub routines

Private Sub cmdTwistAcross1_Click()
unf = 1
Call ShiftAcrossRight(unf)
Call ShowCube
End Sub

Private Sub cmdTwistAcross2_Click()
unf = 2
Call ShiftAcrossRight(unf)
Call ShowCube
End Sub

Private Sub cmdTwistAcross3_Click()
unf = 3
Call ShiftAcrossRight(unf)
Call ShowCube
End Sub

Private Sub cmdTwistDown1_Click()
unf = 1
Call ShiftDown(unf)
Call ShowCube
End Sub
Private Sub cmdTwistDown2_Click()
unf = 2
Call ShiftDown(unf)
Call ShowCube
End Sub
Private Sub cmdTwistDown3_Click()
unf = 3
Call ShiftDown(unf)
Call ShowCube
End Sub

Private Sub cmdTwistLeft1_Click()
unf = 1
Call ShiftAcrossLeft(unf)
Call ShowCube
End Sub

Private Sub cmdTwistLeft2_Click()
unf = 2
Call ShiftAcrossLeft(unf)
Call ShowCube
End Sub

Private Sub cmdTwistLeft3_Click()
unf = 3
Call ShiftAcrossLeft(unf)
Call ShowCube
End Sub

Private Sub cmdTwistup1_Click()
unf = 1
Call ShiftUp(unf)
Call ShowCube
End Sub

Private Sub cmdTwistup2_Click()
unf = 2
Call ShiftUp(unf)
Call ShowCube
End Sub

Private Sub cmdTwistup3_Click()
unf = 3
Call ShiftUp(unf)
Call ShowCube
End Sub

Private Sub Form_Load()
'Here it is just used to fill the Array
Call SortCube
End Sub

Private Sub ShowCube()
'This updates all the colours and is used when the buttons are clicked etc
z = 0
For i = 1 To 3
    For j = 1 To 3
        side1(z).BackColor = GetSideColour(Face1(i, j))
        'Get the number from the array, and translate it into a colour
        z = z + 1
        'z is just used for the index of the control arrays (the shapes)
    Next j
Next i

z = 0
For i = 1 To 3
    For j = 1 To 3
        side2(z).BackColor = GetSideColour(Face2(i, j))
        z = z + 1
    Next j
Next i

z = 0
For i = 1 To 3
    For j = 1 To 3
        side3(z).BackColor = GetSideColour(Face3(i, j))
        z = z + 1
    Next j
Next i

z = 0
For i = 1 To 3
    For j = 1 To 3
        side4(z).BackColor = GetSideColour(Face4(i, j))
        z = z + 1
    Next j
Next i

z = 0
For i = 1 To 3
    For j = 1 To 3
        side5(z).BackColor = GetSideColour(Face5(i, j))
        z = z + 1
    Next j
Next i

z = 0
For i = 1 To 3
    For j = 1 To 3
        side6(z).BackColor = GetSideColour(Face6(i, j))
        z = z + 1
    Next j
Next i
End Sub

Private Function GetSideColour(ByVal num As Integer) As Long
'This function just translates the numbers stored in the array into a
'colour, It would also have been possible to store the colour codes
'Directy in the array
If num = 1 Then
    GetSideColour = vbRed
    ElseIf num = 2 Then
        GetSideColour = vbGreen
    ElseIf num = 3 Then
        GetSideColour = vbBlue
    ElseIf num = 4 Then
        GetSideColour = vbBlack
    ElseIf num = 5 Then
        GetSideColour = vbWhite
    ElseIf num = 6 Then
        GetSideColour = vbYellow
End If
End Function

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileNew_Click()
Call SortCube
Call ShowCube
End Sub

Private Sub Scramble()
Dim n As Integer
'This sub is incomplete and is causing problems
Randomize
RandVar1 = Int((Rnd * 100) + 1)
n = 1
For n = 1 To 10
    Randomize
    RandVar2 = Int((Rnd * 100) + 1)
    randVar3 = Int((Rnd * 100) + 1)
    If RandVar2 > 66 Then
        unf = 3
        ElseIf RandVar2 > 33 Then
            unf = 2
        ElseIf RandVar2 > 0 Then
            unf = 1
    End If
    
    If randVar3 > 75 Then
        Call ShiftUp(unf)
        ElseIf randVar3 > 50 Then
            Call ShiftDown(unf)
        ElseIf randVar3 > 25 Then
            Call ShiftAcrossLeft(unf)
        ElseIf randVar3 > 0 Then
            Call ShiftAcrossRight(unf)
    End If
    'txthmmm.Text = i
    DoEvents
Next n
Call ShowCube
End Sub

Private Sub mnuFileScramble_Click()
Call Scramble
End Sub
