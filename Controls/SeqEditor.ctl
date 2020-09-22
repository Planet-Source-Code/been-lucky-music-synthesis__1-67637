VERSION 5.00
Begin VB.UserControl SeqEditor 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   LockControls    =   -1  'True
   ScaleHeight     =   14.7
   ScaleMode       =   0  'User
   ScaleWidth      =   10.148
   ToolboxBitmap   =   "SeqEditor.ctx":0000
   Begin VB.Line Cursor 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   2.038
      X2              =   7.473
      Y1              =   5.6
      Y2              =   5.6
   End
End
Attribute VB_Name = "SeqEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' NOTE: PROPERTY BAG CODE IS STILL TO BE WRITTEN
'       SO IF YOU USE THE CONTROL, YOU WILL HAVE
'       TO SET THE CONTROL PROPERTY AT RUNTIME.


'-- WORKAROUND TO PRESERVE CASE OF ENUMS BELOW

    #If False Then
    Private bgPlainWhite, bgDarkBlue, bgChocolateBrown, bgLeafGreen
    #End If
'--
    #If False Then
    Private LeftButton, RightButton
    #End If

'
'-- COLOR CHOICE AVAILABLE TO VB USER

Public Enum bgColorChoice
    bgPlainWhite
    bgDarkBlue
    bgChocolateBrown
    bgLeafGreen
    End Enum
Public Enum MouseButton
    LeftButton = 1
    RightButton = 2
    End Enum

'
'-- INTERNAL VARIABLES USED
Dim cpNumbersOn                 As Boolean
Dim cpCursorOn                  As Boolean
Dim cpTopRow                    As Integer
Dim cpLeftColumn                As Integer
Dim cpRowCount                  As Integer
Dim cpColumnCount               As Integer
Dim Row                         As Integer
Dim Col                         As Integer
Dim cpMouseDownColumn           As Integer
Dim cpMouseUpColumn             As Integer
Dim CellDifference              As Single

'
'-- COLORS
Dim clrNOTEFRONT                As Long
Dim clrNOTEBACK                 As Long
Dim clrNOTEBORDER               As Long
Dim clrCURSOR                   As Long
Dim clrGRID                     As Long
Dim clrEDITORBACK               As Long
Dim clrEDITORFRONT              As Long
Dim clrTEXT                     As Long
'--

Public Event RightClick(Column As Integer, Row As Integer, CellDifference As Single)
Public Event LeftClick(Column As Integer, Row As Integer, CellDifference As Single)
Public Event MouseMove(Button As Integer, Column As Integer, Row As Integer, CellDifference As Single)
'--

Private Sub UserControl_Initialize()
    
    On Error Resume Next
    
'   SET THE DEFAULT PROPERTIES

    cpRowCount = 24
    cpColumnCount = 20
    cpTopRow = 1
    cpLeftColumn = 1:
    ColorChoice = bgDefault
    Cursor.X1 = 0
    Cursor.X2 = ScaleWidth
    ShowCursor = True
       
End Sub

Public Property Let ColorChoice(C As bgColorChoice)
    
    Select Case C
    
    Case bgDefault:
         Cursor.BorderColor = vbBlack
         clrNOTEFRONT = &HC0C0A0
         clrNOTEBACK = &HFFF0F0
         clrNOTEBORDER = &HA6A6A6
         clrGRID = &HE6E6E6
         clrEDITORBACK = vbWhite
         clrEDITORFRONT = vbRed
         clrTEXT = vbBlack
    Case bgDarkBlue:
         Cursor.BorderColor = vbWhite
         clrEDITORBACK = &H4000&
         clrEDITORFRONT = vbYellow
         clrNOTEFRONT = &HFFC0C0
         clrNOTEBACK = &H808000
         clrNOTEBORDER = &HA6A6A6
         clrGRID = &H8000&
         clrTEXT = vbYellow
    Case bgChocolateBrown
         Cursor.BorderColor = vbWhite
         clrEDITORBACK = &H8080&
         clrEDITORFRONT = vbYellow
         clrNOTEFRONT = &HFF&
         clrNOTEBACK = &HC0C0&
         clrNOTEBORDER = &HE0E0E0
         clrGRID = &H4040&
         clrTEXT = vbYellow
    Case bgLeafGreen
         Cursor.BorderColor = vbWhite
         clrEDITORBACK = &H8000&          '
         clrEDITORFRONT = vbYellow
         clrNOTEFRONT = &HFFFF&
         clrNOTEBACK = &H0&
         clrNOTEBORDER = &HFFFFFF
         clrGRID = &H0&
         clrTEXT = vbWhite
    End Select

    DrawGrid    ' REDRAW WITH NEW SETTINGS

End Property

'-- PROPERTIES
'-- GETS
Public Property Get NumbersOn() As Boolean
    NumbersOn = cpNumbersOn
End Property
Public Property Get CursorOn() As Boolean
    CursorOn = cpCursorOn
End Property
Public Property Get ColorChoice() As bgColorChoice
    ColorChoice = 1
End Property
Public Property Get Rows() As Integer
    Rows = cpRowCount
End Property
Public Property Get Columns() As Integer
    Columns = cpColumnCount
End Property
Public Property Get LeftCol() As Integer        '
    LeftCol = cpLeftColumn                      '
End Property
Public Property Get TopRow() As Integer         '
    TopRow = cpTopRow
End Property
Public Property Get MouseUpCol() As Integer     ' READONLY ...
    MouseUpCol = cpMouseUpColumn                ' THERE IS NO LET
End Property                                    '
Public Property Get MouseDownCol() As Integer   ' READONLY ...
    MouseDownCol = cpMouseDownColumn            ' THERE IS NO LET
End Property                                    '
'-- LETS
Public Property Let NumbersOn(B As Boolean)
    cpNumbersOn = B
End Property
Public Property Let CursorOn(B As Boolean)
    cpCursorOn = B
    Cursor.Visible = Not B
End Property
Public Property Let Rows(R As Integer)
    cpRowCount = R
    DrawGrid
End Property
Public Property Let Columns(C As Integer)
    cpColumnCount = C
    DrawGrid
End Property
Public Property Let LeftCol(C As Integer)
    cpLeftColumn = C
    DrawGrid
End Property
Public Property Let TopRow(R As Integer)
    cpTopRow = R
    DrawGrid
End Property

'-- METHODS

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If CursorOn Then
        Cursor.Visible = False
        CursorY = Int(Y) + 0.5   ' ADD 0.5 TO CENTER CURSOR
        Cursor.Y1 = CursorY
        Cursor.Y2 = CursorY
        Cursor.Visible = True
        End If
    
    Row = Int(Y) + cpTopRow
    Col = Int(X) + cpLeftColumn
    CellDifference = X - Int(X)
    
'   BOUNDCHECK -- BETTER DO IT

    If Row < LeftCol Then Row = LeftCol
    If Col < TopRow Then Col = TopRow
    If Row > Rows Then Row = Rows
    If Col > Columns Then Col = Columns
    
    RaiseEvent MouseMove(Button, Col, Row, CellDifference)
   
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cpMouseUpColumn = Int(X) + cpLeftColumn
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            
    Row = Int(Y) + cpTopRow
    Col = Int(X) + cpLeftColumn
    CellDifference = X - Int(X)
    
    cpMouseDownColumn = Int(X) + cpLeftColumn
    
    Select Case Button
    Case 2: RaiseEvent RightClick(Col, Row, CellDifference)
    Case 1: RaiseEvent LeftClick(Col, Row, CellDifference)
    End Select
     
End Sub

Sub DrawGrid()

    Cls
    
    BackColor = clrEDITORBACK
    ForeColor = clrTEXT
    ScaleWidth = cpColumnCount
    ScaleHeight = cpRowCount
    
    ' HORIZONTAL GRID LINES
    For i = 1 To cpRowCount
    Line (0, i)-(cpColumnCount, i), clrGRID
    Next

    ' VERTICAL GRID LINES AND NUMBER AT TOP
    For i = 1 To cpColumnCount
    Line (i, 0)-(i, cpRowCount), clrGRID
    If cpNumbersOn Then
        CurrentX = i - 0.9: CurrentY = 0.1:
        Print i + cpLeftColumn - 1
        End If
    Next
    
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
    DrawGrid
    Cursor.X2 = ScaleWidth

End Sub

Public Sub EraseNote(Column As Integer, Row As Integer)

    DrawNote Column, Row, 0   ' Period = 0 = Erase

End Sub

Public Sub DrawNote(Column As Integer, Row As Integer, Period As Single)

    BackColor = clrEDITORBACK
    ForeColor = clrTEXT
    
    X1 = Int(Column) - LeftCol
    Y1 = Int(Row) - TopRow  ' BASE ADJUST

    '-- ERASE TOP2BOTTOM
    Line (X1, 0)-(X1 + 1, ScaleHeight), clrEDITORBACK, BF
    
    '-- REDRAW ERASED PART OF GRID
    For i = 0 To ScaleHeight
    Line (X1, i)-(X1 + 1, i + 1), clrGRID, B
    Next
    
    If Period <> 0 Then
        '-- DRAW EVENTBOX
        Line (X1, Y1)-(X1 + 1, Y1 + 1), clrNOTEBACK, BF
        Line (X1, Y1)-(X1 + Period, Y1 + 1), clrNOTEFRONT, BF
        '-- ADD SOME SIMPLE 3D EFFECT
        DrawWidth = 2
        Line (X1 + 0.1, Y1 + 1)-(X1 + Period, Y1 + 1), vbBlack      ' BOTTOM SHADOW
        Line (X1 + Period, Y1)-(X1 + Period, Y1 + 1), vbBlack       ' RIGHT SHADOW
        DrawWidth = 1
        Line (X1, Y1)-(X1 + Period, Y1 + 1), clrNOTEBORDER, B
    End If

    '-- EVENTBOX NUMBER..PRINT NUMBER ON TOP
    If cpNumbersOn Then
    CurrentX = X1 + 0.1: CurrentY = 0.1:
    Print Int(Column)
    End If
    
End Sub

