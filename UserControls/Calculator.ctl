VERSION 5.00
Begin VB.UserControl Calculator 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4656
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   10.2
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   ToolboxBitmap   =   "Calculator.ctx":0000
   Begin VB.PictureBox picCalculator 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4608
      Left            =   24
      Picture         =   "Calculator.ctx":0312
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   24
      Width           =   3228
      Begin VB.PictureBox picBuffer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   372
         Left            =   120
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4056
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.PictureBox picMemory 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   264
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   146
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1752
         Begin VB.Label lblMemory 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   1680
            TabIndex        =   5
            Top             =   24
            Width           =   48
         End
      End
      Begin VB.PictureBox picDisplay 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   156
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   242
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   2904
         Begin VB.Label lblDisplay 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   360
            TabIndex        =   3
            Top             =   36
            Width           =   2520
         End
      End
   End
   Begin VB.ListBox lstReceipt 
      Appearance      =   0  'Flat
      Height          =   4512
      ItemData        =   "Calculator.ctx":0995
      Left            =   3276
      List            =   "Calculator.ctx":0997
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   72
      Width           =   2772
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Calculator Control
'
'Author Ben Vonk
'30-08-2008 First version

Option Explicit

' Public Events
Public Event Closing(Value As String)
Public Event ExitFocus()
Public Event ReceiptState(Showed As Boolean)

' Private Constants
Private Const KEY_ADD             As Integer = 13
Private Const KEY_BACKSPACE       As Integer = 25
Private Const KEY_CLEAR           As Integer = 23
Private Const KEY_CLEARENTRY      As Integer = 24
Private Const KEY_CLOSE           As Integer = 33
Private Const KEY_COPY            As Integer = 32
Private Const KEY_DECIMAL         As Integer = 11
Private Const KEY_DECIMALCOUNT    As Integer = 22
Private Const KEY_DIVIDE          As Integer = 16
Private Const KEY_DOUBLENULL      As Integer = 10
Private Const KEY_POWERTO         As Integer = 17
Private Const KEY_INVERT          As Integer = 19
Private Const KEY_MEMORY_ADD      As Integer = 28
Private Const KEY_MEMORY_CLEAR    As Integer = 26
Private Const KEY_MEMORY_RECALL   As Integer = 27
Private Const KEY_MEMORY_SUBTRACT As Integer = 29
Private Const KEY_MULTIPLY        As Integer = 15
Private Const KEY_PERCENT         As Integer = 20
Private Const KEY_RECEIPT         As Integer = 30
Private Const KEY_RECEIPTCLEAR    As Integer = 31
Private Const KEY_RETURN          As Integer = 12
Private Const KEY_INVERSE         As Integer = 18
Private Const KEY_SQUAREROOT      As Integer = 21
Private Const KEY_SUBTRACT        As Integer = 14
Private Const FLOODFILLBORDER     As Long = 0

' Public Enumreration
Public Enum BorderStyles
   NoBorder
   Raised
   Sunken
   Edged
   Colored
End Enum

' Private Enumerations
Private Enum ButtonStates
   ButtonUp
   ButtonDown
   ButtonPressed
End Enum

Private Enum InputTypes
   NoInput
   Numbers
   DecimalSign
   Inputs
   Operands
   Calculation
End Enum

' Private Types
Private Type PointAPI
   X                              As Long
   Y                              As Long
End Type

Private Type Rect
   Left                           As Long
   Top                            As Long
   Right                          As Long
   Bottom                         As Long
End Type

Private Type CalcButtons
   Caption                        As String
   BackColor                      As Long
   ForeColor                      As Long
   Tag                            As String
   Rect                           As Rect
End Type

' Private Variables
Private CalculationError          As Boolean
Private DecimalFlag               As Boolean
Private KeyDown                   As Boolean
Private m_ReceiptVisible          As Boolean
Private m_RepeatKey               As Boolean
Private m_ShowPressedKey          As Boolean
Private MemoryError               As Boolean
Private MouseOut                  As Boolean
Private m_BorderStyle             As BorderStyles
Private CalcButton()              As CalcButtons
Private MemoryValue               As Double
Private Operand(2)                As Double
Private LastInput                 As InputTypes
Private ButtonIndex               As Integer
Private m_DecimalCount            As Integer
Private OperandIndex              As Integer
Private PressedKey                As Integer
Private m_BackColor               As Long
Private m_BorderColor             As Long
Private m_CalculatorColor         As Long
Private m_CancelKeysBackColor     As Long
Private m_CancelKeysForeColor     As Long
Private m_DisplaysBackColor       As Long
Private m_DisplaysOffColor        As Long
Private m_CalculatorOnColor       As Long
Private m_EnterKeyBackColor       As Long
Private m_EnterKeyForeColor       As Long
Private m_ErrorColor              As Long
Private m_FunctionKeysBackColor   As Long
Private m_FunctionKeysForeColor   As Long
Private m_MemoryKeysBackColor     As Long
Private m_MemoryKeysForeColor     As Long
Private m_MemoryOnColor           As Long
Private m_NumberKeysBackColor     As Long
Private m_NumberKeysForeColor     As Long
Private m_OperandKeysBackColor    As Long
Private m_OperandKeysForeColor    As Long
Private m_PressedKeyColor         As Long
Private m_ReceiptBackColor        As Long
Private m_ReceiptForeColor        As Long
Private DecimalPoint              As String
Private OperationFlag             As String

' Private API's
Private Declare Function CreatePolygonRgn Lib "GDI32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function ExtFloodFill Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function FillRgn Lib "GDI32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

   BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   m_BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call Refresh

End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."

   BorderColor = m_BorderColor

End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)

   m_BorderColor = NewBorderColor
   PropertyChanged "BorderColor"
   
   Call Refresh

End Property

Public Property Get BorderStyle() As BorderStyles
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."

   BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As BorderStyles)

   m_BorderStyle = NewBorderStyle
   PropertyChanged "BorderStyle"
   
   Call Refresh

End Property

Public Property Get CalculatorColor() As OLE_COLOR
Attribute CalculatorColor.VB_Description = "Returns/sets the calculator background color."

   CalculatorColor = m_CalculatorColor

End Property

Public Property Let CalculatorColor(ByVal NewCalculatorColor As OLE_COLOR)

   m_CalculatorColor = NewCalculatorColor
   PropertyChanged "CalculatorColor"
   
   Call Refresh

End Property

Public Property Get CalculatorOnColor() As OLE_COLOR
Attribute CalculatorOnColor.VB_Description = "Returns/sets the calculator display foreground color when calculator is on."

   CalculatorOnColor = m_CalculatorOnColor

End Property

Public Property Let CalculatorOnColor(ByVal NewDisplayOnColor As OLE_COLOR)

   m_CalculatorOnColor = NewDisplayOnColor
   PropertyChanged "CalculatorOnColor"
   
   If LastInput > NoInput Then Call Refresh

End Property

Public Property Get CancelKeysBackColor() As OLE_COLOR
Attribute CancelKeysBackColor.VB_Description = "Returns/sets the cancel Keys background color."

   CancelKeysBackColor = m_CancelKeysBackColor

End Property

Public Property Let CancelKeysBackColor(ByVal NewCancelKeysBackColor As OLE_COLOR)

   m_CancelKeysBackColor = NewCancelKeysBackColor
   PropertyChanged "CancelKeysBackColor"
   
   Call Refresh

End Property

Public Property Get CancelKeysForeColor() As OLE_COLOR
Attribute CancelKeysForeColor.VB_Description = "Returns/sets the cancel Keys foreground color."

   CancelKeysForeColor = m_CancelKeysForeColor

End Property

Public Property Let CancelKeysForeColor(ByVal NewCancelKeysForeColor As OLE_COLOR)

   m_CancelKeysForeColor = NewCancelKeysForeColor
   PropertyChanged "CancelKeysForeColor"
   
   Call Refresh

End Property

Public Property Get DecimalCount() As Integer
Attribute DecimalCount.VB_Description = "Returns/sets the number of decimals that must be showed. (-1 will show all)"

   DecimalCount = m_DecimalCount

End Property

Public Property Let DecimalCount(ByVal NewDecimalCount As Integer)

   If NewDecimalCount < -1 Then NewDecimalCount = -1
   If NewDecimalCount > 4 Then NewDecimalCount = 4
   
   m_DecimalCount = NewDecimalCount
   PropertyChanged "DecimalCount"
   ButtonIndex = KEY_DECIMALCOUNT
   
   Call SetDecimals

End Property

Public Property Get DisplaysBackColor() As OLE_COLOR
Attribute DisplaysBackColor.VB_Description = "Returns/sets the displays background color."

   DisplaysBackColor = m_DisplaysBackColor

End Property

Public Property Let DisplaysBackColor(ByVal NewDisplayBackColor As OLE_COLOR)

   m_DisplaysBackColor = NewDisplayBackColor
   PropertyChanged "DisplaysBackColor"
   
   Call Refresh

End Property

Public Property Get DisplaysOffColor() As OLE_COLOR
Attribute DisplaysOffColor.VB_Description = "Returns/sets the display foreground color when display is off."

   DisplaysOffColor = m_DisplaysOffColor

End Property

Public Property Let DisplaysOffColor(ByVal NewDisplayOffColor As OLE_COLOR)

   m_DisplaysOffColor = NewDisplayOffColor
   PropertyChanged "DisplaysOffColor"
   
   Call Refresh

End Property

Public Property Get EnterKeyBackColor() As OLE_COLOR
Attribute EnterKeyBackColor.VB_Description = "Returns/sets the enter Key background color."

   EnterKeyBackColor = m_EnterKeyBackColor

End Property

Public Property Let EnterKeyBackColor(ByVal NewEnterKeyBackColor As OLE_COLOR)

   m_EnterKeyBackColor = NewEnterKeyBackColor
   PropertyChanged "EnterKeyBackColor"
   
   Call Refresh

End Property

Public Property Get EnterKeyForeColor() As OLE_COLOR
Attribute EnterKeyForeColor.VB_Description = "Returns/sets the enter Key foreground color."

   EnterKeyForeColor = m_EnterKeyForeColor

End Property

Public Property Let EnterKeyForeColor(ByVal NewEnterKeyForeColor As OLE_COLOR)

   m_EnterKeyForeColor = NewEnterKeyForeColor
   PropertyChanged "EnterKeyForeColor"
   
   Call Refresh

End Property

Public Property Get ErrorColor() As OLE_COLOR
Attribute ErrorColor.VB_Description = "Returns/sets the error sign foreground color."

   ErrorColor = m_ErrorColor

End Property

Public Property Let ErrorColor(ByVal NewErrorColor As OLE_COLOR)

   m_ErrorColor = NewErrorColor
   PropertyChanged "ErrorColor"
   
   If CalculationError Then Call ShowKeyPressed

End Property

Public Property Get FunctionKeysBackColor() As OLE_COLOR
Attribute FunctionKeysBackColor.VB_Description = "Returns/sets the function Keys background color."

   FunctionKeysBackColor = m_FunctionKeysBackColor

End Property

Public Property Let FunctionKeysBackColor(ByVal NewFunctionKeysBackColor As OLE_COLOR)

   m_FunctionKeysBackColor = NewFunctionKeysBackColor
   PropertyChanged "FunctionKeysBackColor"
   
   Call Refresh

End Property

Public Property Get FunctionKeysForeColor() As OLE_COLOR
Attribute FunctionKeysForeColor.VB_Description = "Returns/sets the function Keys foreground color."

   FunctionKeysForeColor = m_FunctionKeysForeColor

End Property

Public Property Let FunctionKeysForeColor(ByVal NewFunctionKeysForeColor As OLE_COLOR)

   m_FunctionKeysForeColor = NewFunctionKeysForeColor
   PropertyChanged "FunctionKeysForeColor"
   
   Call Refresh

End Property

Public Property Get MemoryKeysBackColor() As OLE_COLOR
Attribute MemoryKeysBackColor.VB_Description = "Returns/sets the memory Keys background color."

   MemoryKeysBackColor = m_MemoryKeysBackColor

End Property

Public Property Let MemoryKeysBackColor(ByVal NewMemoryKeysBackColor As OLE_COLOR)

   m_MemoryKeysBackColor = NewMemoryKeysBackColor
   PropertyChanged "MemoryKeysBackColor"
   
   Call Refresh

End Property

Public Property Get MemoryKeysForeColor() As OLE_COLOR
Attribute MemoryKeysForeColor.VB_Description = "Returns/sets the memory Keys foreground color."

   MemoryKeysForeColor = m_MemoryKeysForeColor

End Property

Public Property Let MemoryKeysForeColor(ByVal NewMemoryKeysForeColor As OLE_COLOR)

   m_MemoryKeysForeColor = NewMemoryKeysForeColor
   PropertyChanged "MemoryKeysForeColor"
   
   Call Refresh

End Property

Public Property Get MemoryOnColor() As OLE_COLOR
Attribute MemoryOnColor.VB_Description = "Returns/sets the memory display foreground color when memory is set."

   MemoryOnColor = m_MemoryOnColor

End Property

Public Property Let MemoryOnColor(ByVal NewMemoryColor As OLE_COLOR)

   m_MemoryOnColor = NewMemoryColor
   PropertyChanged "MemoryOnColor"
   
   If MemoryValue Then lblMemory.ForeColor = m_MemoryOnColor

End Property

Public Property Get NumberKeysBackColor() As OLE_COLOR
Attribute NumberKeysBackColor.VB_Description = "Returns/sets the number Keys background color."

   NumberKeysBackColor = m_NumberKeysBackColor

End Property

Public Property Let NumberKeysBackColor(ByVal NewNumberKeysBackColor As OLE_COLOR)

   m_NumberKeysBackColor = NewNumberKeysBackColor
   PropertyChanged "NumberKeysBackColor"
   
   Call Refresh

End Property

Public Property Get NumberKeysForeColor() As OLE_COLOR
Attribute NumberKeysForeColor.VB_Description = "Returns/sets the number Keys foreground color."

   NumberKeysForeColor = m_NumberKeysForeColor

End Property

Public Property Let NumberKeysForeColor(ByVal NewNumberKeysForeColor As OLE_COLOR)

   m_NumberKeysForeColor = NewNumberKeysForeColor
   PropertyChanged "NumberKeysForeColor"
   
   Call Refresh

End Property

Public Property Get OperandKeysBackColor() As OLE_COLOR
Attribute OperandKeysBackColor.VB_Description = "Returns/sets the operand Keys background color."

   OperandKeysBackColor = m_OperandKeysBackColor

End Property

Public Property Let OperandKeysBackColor(ByVal NewOperandKeysBackColor As OLE_COLOR)

   m_OperandKeysBackColor = NewOperandKeysBackColor
   PropertyChanged "OperandKeysBackColor"
   
   Call Refresh

End Property

Public Property Get OperandKeysForeColor() As OLE_COLOR
Attribute OperandKeysForeColor.VB_Description = "Returns/sets the operand Keys foreground color."

   OperandKeysForeColor = m_OperandKeysForeColor

End Property

Public Property Let OperandKeysForeColor(ByVal NewOperandKeysForeColor As OLE_COLOR)

   m_OperandKeysForeColor = NewOperandKeysForeColor
   PropertyChanged "OperandKeysForeColor"
   
   Call Refresh

End Property

Public Property Get PressedKeyColor() As OLE_COLOR
Attribute PressedKeyColor.VB_Description = "Returns/sets the pressed Key foreground color."

   PressedKeyColor = m_PressedKeyColor

End Property

Public Property Let PressedKeyColor(ByVal NewPressedKeyColor As OLE_COLOR)

   m_PressedKeyColor = NewPressedKeyColor
   PropertyChanged "PressedKeyColor"
   
   If ShowPressedKey Then Call ShowKeyPressed

End Property

Public Property Get ReceiptBackColor() As OLE_COLOR
Attribute ReceiptBackColor.VB_Description = "Returns/sets the receipt background color."

   ReceiptBackColor = m_ReceiptBackColor

End Property

Public Property Let ReceiptBackColor(ByVal NewReceiptBackColor As OLE_COLOR)

   m_ReceiptBackColor = NewReceiptBackColor
   PropertyChanged "ReceiptBackColor"
   lstReceipt.BackColor = m_ReceiptBackColor

End Property

Public Property Get ReceiptForeColor() As OLE_COLOR
Attribute ReceiptForeColor.VB_Description = "Returns/sets the receipt foreground color."

   ReceiptForeColor = m_ReceiptForeColor

End Property

Public Property Let ReceiptForeColor(ByVal NewReceiptForeColor As OLE_COLOR)

   m_ReceiptForeColor = NewReceiptForeColor
   PropertyChanged "ReceiptForeColor"
   lstReceipt.ForeColor = m_ReceiptForeColor

End Property

Public Property Get ReceiptVisible() As Boolean
Attribute ReceiptVisible.VB_Description = "Returns/sets a value that determines whether the receipt is visible or hidden."

   ReceiptVisible = m_ReceiptVisible

End Property

Public Property Let ReceiptVisible(ByVal NewReceiptVisible As Boolean)

   m_ReceiptVisible = NewReceiptVisible
   PropertyChanged "ReceiptVisible"
   
   Call ShowReceipt
   Call DrawButton(KEY_RECEIPT, ButtonUp)

End Property

Public Property Get RepeatKey() As Boolean
Attribute RepeatKey.VB_Description = "Returns/sets a value that determines whether the pressed key can be repeat or not."

   RepeatKey = m_RepeatKey

End Property

Public Property Let RepeatKey(ByVal NewRepeatKey As Boolean)

   m_RepeatKey = NewRepeatKey
   PropertyChanged "RepeatKey"

End Property

Public Property Get ShowPressedKey() As Boolean
Attribute ShowPressedKey.VB_Description = "Returns/sets a value that determines whether the pressed key will be showed in the display."

   ShowPressedKey = m_ShowPressedKey

End Property

Public Property Let ShowPressedKey(ByVal NewShowPressedKey As Boolean)

   m_ShowPressedKey = NewShowPressedKey
   PropertyChanged "ShowPressedKey"

End Property

Public Sub Refresh()

Dim intCount As Integer
Dim intX     As Integer
Dim intY     As Integer

   With picCalculator
      UserControl.BackColor = m_BackColor
      .BackColor = m_BackColor
      .FillStyle = vbFSSolid
      
      For intCount = 1 To 4
         .FillColor = Choose(intCount, m_BackColor, m_CalculatorColor, m_DisplaysBackColor, m_DisplaysBackColor)
         intX = Choose(intCount, 1, 10, 30, 30)
         intY = Choose(intCount, 1, 10, 30, 110)
         ExtFloodFill .hDC, intX, intY, &H404040, FLOODFILLBORDER
      Next 'intCount
      
      For intCount = 236 To 251 Step 3
         ExtFloodFill .hDC, intCount, 80, &H404040, FLOODFILLBORDER
      Next 'intCount
      
      .FillStyle = vbFSTransparent
      picDisplay.BackColor = m_DisplaysBackColor
      picMemory.BackColor = m_DisplaysBackColor
   End With
   
   Call DrawBorder
   Call FillButtons
   Call ShowButtons

End Sub

Private Function CheckDecimalCount(ByVal InputValue As Double) As Double

Dim intCount As Integer
Dim dblValue As Double

   If m_DecimalCount > -1 Then
      For intCount = 4 To m_DecimalCount Step -1
         dblValue = InputValue * (10 ^ intCount)
         
         If dblValue - Val(dblValue) > 0.4 Then dblValue = dblValue + 1
         
         dblValue = Val(dblValue) / (10 ^ intCount)
      Next 'intCount
      
      CheckDecimalCount = dblValue
      
   Else
      CheckDecimalCount = InputValue
   End If

End Function

Private Function CheckDigitsCount(ByVal InputValue As String) As Double

Dim intPointer As Integer

   If (Len(InputValue) > 15) Or InStr(InputValue, "E") Then
      intPointer = InStr(InputValue, "E")
      
      If intPointer Then
         CheckDigitsCount = CDbl(Left(InputValue, intPointer - 1))
         
         If Val(Abs(Mid(InputValue, intPointer + 1))) > 14 Then CalculationError = True
         
      Else
         CheckDigitsCount = CDbl(Left(InputValue, 16))
      End If
      
   Else
      CheckDigitsCount = InputValue
   End If
   
   If Not CalculationError Then Call ShowKeyPressed

End Function

Private Function IsInButton(ByRef ButtonRect As Rect, ByVal X As Long, ByVal Y As Long) As Boolean

   IsInButton = PtInRect(ButtonRect, X, Y)

End Function

Private Sub Calculate(ByVal Index As Integer)

Dim blnShowValue As Boolean
Dim intPointer   As Integer
Dim strTag       As String

   ReDim strOperand(2) As String
   
   If (LastInput = Numbers) Or (LastInput = Inputs) Or (LastInput = Calculation) Then OperandIndex = OperandIndex + 1
   
   Select Case OperandIndex
      Case 1
         Operand(1) = CDbl(lblDisplay.Caption)
         
      Case 2
         Operand(2) = CDbl(lblDisplay.Caption)
         strOperand(0) = Int(Operand(1))
         strOperand(1) = Int(Operand(1)) + Int(Operand(2))
         strOperand(2) = Int(Operand(2))
         
         If (Len(CStr(Operand(1))) > 15) And (Len(strOperand(1))) > Len(strOperand(0)) Then
            strOperand(1) = Operand(1)
            Operand(1) = Left(strOperand(1), Len(strOperand(1)) - (Len(strOperand(2)) - Len(strOperand(0))))
         End If
         
         Select Case OperationFlag
            Case "+"
               Operand(1) = Operand(1) + Operand(2)
               
            Case "-"
               Operand(1) = Operand(1) - Operand(2)
               
            Case "X"
               Operand(1) = Operand(1) * Operand(2)
               
            Case "/"
               If Operand(2) = 0 Then
                  CalculationError = True
                  Operand(1) = 0
                  
                  Call ShowKeyPressed
                  
               Else
                  Operand(1) = Operand(1) / Operand(2)
               End If
               
            Case "^"
               Operand(1) = Operand(1) ^ Operand(2)
               
            Case "="
               Operand(1) = Operand(2)
         End Select
         
         strOperand(1) = Operand(1)
         Operand(1) = CheckDigitsCount(strOperand(1))
         Operand(1) = CheckDecimalCount(Operand(1))
         
         If Operand(1) = Int(Operand(1)) Then
            lblDisplay.Caption = Format(Operand(1), "0.")
            
         Else
            lblDisplay.Caption = Operand(1)
         End If
         
         blnShowValue = True
   End Select
   
   With lstReceipt
      strTag = CalcButton(Index).Tag
      
      If CalculationError And (OperationFlag = "/") And (Operand(2) = 0) Then
         ' Devide by 0 error!
         ' Hold the / in the OperationFlag to check the CE button
         strTag = OperationFlag
         
      ElseIf InStr("+-*/^", strTag) Or (LastInput = Calculation) Or (OperationFlag = strTag) And (Left(.List(.ListCount - 1), 1) = strTag) Then
         If strTag <> "=" Then
            If OperandIndex = 2 Then
               Call FillReceipt("=")
               Call FillReceipt(lblDisplay.Caption)
               Call FillReceipt(strTag)
               
            Else
               Call FillReceipt(strTag)
            End If
         End If
         
      ElseIf strTag = "=" Then
         Call FillReceipt(strTag)
         
         If blnShowValue Then Call FillReceipt(lblDisplay.Caption)
         
      Else
         If blnShowValue And (Left(.List(.ListCount - 1), 1) <> " ") Then
            Call FillReceipt(vbCrLf)
            Call FillReceipt(lblDisplay.Caption)
         End If
         
         Call FillReceipt(strTag)
      End If
   End With
   
   If OperandIndex > 1 Then OperandIndex = 1
   
   LastInput = Operands
   OperationFlag = strTag
   Erase strOperand

End Sub

Private Sub CalculateMemory(ByVal InputValue As String, ByVal Add As Boolean)

Dim dblMemoryValue As Double

   dblMemoryValue = MemoryValue
   
   If Add Then
      MemoryValue = MemoryValue + CDbl(InputValue)
      
   Else
      MemoryValue = MemoryValue - CDbl(InputValue)
   End If
   
   MemoryValue = CheckDigitsCount(CStr(MemoryValue))
   
   If CalculationError Then
      MemoryError = True
      CalculationError = False
      MemoryValue = dblMemoryValue
   End If

End Sub

Private Sub CheckAllButtons(ByVal X As Long, ByVal Y As Long)

Dim intIndex As Integer

   For intIndex = 0 To UBound(CalcButton)
      If IsInButton(CalcButton(intIndex).Rect, X, Y) Then
         ButtonIndex = intIndex
         Exit For
      End If
   Next 'intIndex

End Sub

Private Sub ClearCalculator(ByVal Flag As String)

Dim intCount As Integer

   DecimalFlag = False
   CalculationError = False
   
   If Flag = "CE" Then Exit Sub
   
   For intCount = 0 To 2
      Operand(intCount) = vbDefault
   Next 'intCount
   
   lblDisplay.ForeColor = m_DisplaysOffColor
   lblDisplay.Caption = Format(0, "0.")
   OperandIndex = vbDefault
   OperationFlag = Flag
   LastInput = NoInput
   lstReceipt.Clear

End Sub

Private Sub ClearEntry()

Dim blnError As Boolean

   If CalculationError And (OperationFlag = "/") And (Operand(2) = 0) Then
      Exit Sub
      
   ElseIf Not CalculationError Then
      lblDisplay.Caption = Format(0, "0.")
   End If
   
   blnError = CalculationError
   
   Call ClearCalculator("CE")
   
   If blnError Then Exit Sub
   
   If (Left(lstReceipt.List(lstReceipt.ListCount - 1), 1) = "=") Or (LastInput = Calculation) Then
      Call FillReceipt(vbCrLf)
      Call FillReceipt(lblDisplay.Caption)
      
      LastInput = Operands
      
   ElseIf Len(Trim(lstReceipt.List(lstReceipt.ListCount - 1))) Then
      Call FillReceipt(lblDisplay.Caption)
   End If

End Sub

Private Sub DoBackSpace()

Dim blnNegative As Boolean

   With lblDisplay
      If Left(.Caption, 1) = "-" Then
         blnNegative = True
         .Caption = Mid(.Caption, 2)
      End If
      
      If Right(.Caption, 1) = DecimalPoint Then
         If DecimalFlag Then
            DecimalFlag = False
            
         Else
            .Caption = Mid(.Caption, 1, Len(.Caption) - 2) & Right(.Caption, 1)
            
            If Len(.Caption) = 1 Then .Caption = "0" & .Caption
         End If
         
      Else
         .Caption = Left(.Caption, Len(.Caption) - 1)
      End If
      
      If .Caption = DecimalPoint Then .Caption = Format(0, "0.")
      If blnNegative And CDbl(.Caption) <> 0 Then .Caption = "-" & .Caption
      
      If Len(Trim(lstReceipt.List(lstReceipt.ListCount - 1))) > 1 Then
         If LastInput = Operands Then Call FillReceipt(vbCrLf)
         
         Call FillReceipt(.Caption)
      End If
   End With
   
   LastInput = Numbers

End Sub

Private Sub DoFunction(ByVal KeyIndex As Integer)

Dim blnError   As Boolean
Dim intPointer As Integer

   Select Case KeyIndex
      Case KEY_MEMORY_RECALL
         If MemoryValue Then
            Operand(0) = MemoryValue
            DecimalFlag = False
            LastInput = Inputs
         End If
         
      Case KEY_INVERT
         If CDbl(lblDisplay.Caption) = 0 Then
            CalculationError = True
            
            Call ShowKeyPressed
            
         Else
            Operand(0) = 1 / CDbl(lblDisplay.Caption)
         End If
         
         LastInput = Calculation
         Operand(0) = CheckDigitsCount(CStr(Operand(0)))
         
         If m_ReceiptVisible Then Call FillReceipt(vbCrLf)
         
      Case KEY_INVERSE
         Operand(0) = -CDbl(lblDisplay.Caption)
         DecimalFlag = False
         LastInput = Numbers
         
      ' KEY_PERCENT and KEY_SQUAREROOT
      Case Else
         If KeyIndex = KEY_PERCENT Then
            Operand(0) = Operand(1) / 100 * CDbl(lblDisplay.Caption)
            
         Else
            Operand(0) = CDbl(lblDisplay.Caption)
            
            If Operand(0) < 0 Then
               blnError = True
               Operand(0) = Abs(Operand(0))
            End If
            
            Operand(0) = Sqr(Operand(0))
         End If
         
         LastInput = Calculation
         Operand(0) = CheckDigitsCount(CStr(Operand(0)))
         Operand(0) = CheckDecimalCount(Operand(0))
         
         If Not CalculationError Then
            If m_ReceiptVisible Then Call FillReceipt(vbCrLf)
            
            DecimalFlag = True
         End If
         
         If blnError Then CalculationError = True
   End Select

End Sub

Private Sub DrawBorder()

Const BDR_EDGED  As Long = &H16
Const BDR_RAISED As Long = &H5
Const BDR_SUNKEN As Long = &HA
Const BF_BOTTOM  As Long = &H8
Const BF_LEFT    As Long = &H1
Const BF_RIGHT   As Long = &H4
Const BF_TOP     As Long = &H2
Const BF_RECT    As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Dim rctFrame     As Rect

   Cls
   
   If m_BorderStyle = Colored Then
      Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), m_BorderColor, B
      
   ElseIf m_BorderStyle Then
      With rctFrame
         .Top = 0
         .Left = 0
         .Right = ScaleWidth
         .Bottom = ScaleHeight
      End With
      
      DrawEdge hDC, rctFrame, Choose(m_BorderStyle, BDR_RAISED, BDR_SUNKEN, BDR_EDGED), BF_RECT
   End If

End Sub

Private Sub DrawButton(ByVal Index As Integer, ByVal ButtonState As ButtonStates)

Const ALTERNATE  As Long = 1

Dim lngBrush     As Long
Dim lngRegion    As Long
Dim strCaption   As String

   ReDim lngColor(1) As Long
   
   If ButtonState = ButtonUp Then
      lngColor(0) = &HE0E0E0
      lngColor(1) = &H575757
      
   Else
      lngColor(0) = &H707070
      lngColor(1) = &HE7E7E7
   End If
   
   ReDim ptaRegion(1 To 4) As PointAPI
   
   With CalcButton(Index)
      strCaption = .Caption
      
      With .Rect
         picCalculator.Line (.Left + 1, .Top)-(.Right, .Top), lngColor(0)
         picCalculator.Line (.Left, .Top + 1)-(.Left, .Bottom), lngColor(0)
         picCalculator.Line (.Left + 1, .Bottom)-(.Right, .Bottom), lngColor(1)
         picCalculator.Line (.Right, .Top + 1)-(.Right, .Bottom), lngColor(1)
         DoEvents
         ptaRegion(1).X = .Left + 1
         ptaRegion(1).Y = .Top + 1
         ptaRegion(2).X = .Right
         ptaRegion(2).Y = .Top + 1
         ptaRegion(3).X = .Right
         ptaRegion(3).Y = .Bottom
         ptaRegion(4).X = .Left + 1
         ptaRegion(4).Y = .Bottom
         picCalculator.CurrentX = .Left + (.Right - .Left - picCalculator.TextWidth(strCaption)) \ 2 + ButtonState
         picCalculator.CurrentY = .Top + ((.Bottom - .Top) - picCalculator.TextHeight(strCaption)) \ 2 + ButtonState
      End With
      
      lngBrush = CreateSolidBrush(.BackColor)
      lngRegion = CreatePolygonRgn(ptaRegion(1), 4, ALTERNATE)
      
      If lngRegion Then FillRgn picCalculator.hDC, lngRegion, lngBrush
      
      If Len(.Caption) Then
         picCalculator.ForeColor = .ForeColor
         picCalculator.Print .Caption
         
      Else
         Call DrawButtonImage(Index, ButtonState, .ForeColor, .Rect.Left, .Rect.Top, picCalculator)
      End If
   End With
   
   DoEvents
   DeleteObject lngRegion
   DeleteObject lngBrush
   Erase ptaRegion, lngColor

End Sub

Private Sub DrawButtonImage(ByVal Index As Integer, ByVal ButtonState As ButtonStates, ByVal Color As Long, ByVal X As Long, ByVal Y As Long, ByRef Box As PictureBox)

Dim intCount As Integer
Dim lngX     As Long
Dim lngY     As Long

   With CalcButton(Index).Rect
      Box.DrawWidth = 2
      Box.Line (.Left + 2, .Top + 2)-(.Right - 2, .Bottom - 2), CalcButton(Index).BackColor, BF
   End With
   
   If Index = KEY_SUBTRACT Then
      If Box.Name = "picDisplay" Then lngY = (4 And (Screen.TwipsPerPixelY <> 12))
      
      lngX = ButtonUp + 12
      lngY = ButtonUp + 16 - lngY
      Box.Line (X + lngX, Y + lngY)-(X + lngX + 7, Y + lngY), Color, BF
      
   ElseIf Index = KEY_SQUAREROOT Then
      If Box.Name = "picDisplay" Then lngY = (5 And (Screen.TwipsPerPixelY <> 12))
      
      lngX = ButtonState + 8
      lngY = ButtonState - lngY
      Box.Line (X + lngX, Y + lngY + 14)-(X + lngX + 4, Y + lngY + 22), Color
      Box.Line -(X + lngX + 9, Y + lngY + 10), Color
      Box.Line -(X + lngX + 16, Y + lngY + 10), Color
      
   ElseIf Index = KEY_BACKSPACE Then
      If Box.Name = "picDisplay" Then lngY = (5 And (Screen.TwipsPerPixelY <> 12))
      
      lngX = ButtonState + 8 + (10 And (Box.Name = "picCalculator"))
      lngY = ButtonState - lngY
      Box.Line (X + lngX, Y + lngY + 16)-(X + lngX + 16, Y + lngY + 16), Color
      Box.Line (X + lngX + 1, Y + lngY + 15)-(X + lngX + 7, Y + lngY + 12), Color
      Box.Line (X + lngX + 1, Y + lngY + 16)-(X + lngX + 7, Y + lngY + 19), Color
      Box.Line -(X + lngX + 7, Y + lngY + 12), Color
      Box.PSet (X + lngX, Y + lngY + 15), Color
      
   ElseIf Index = KEY_COPY Then
      lngX = ButtonState + 9
      lngY = ButtonState + 7
      Box.DrawWidth = 1
      Box.Line (X + lngX + 9, Y + lngY + 2)-(X + lngX + 9, Y + lngY), Color
      Box.Line -(X + lngX, Y + lngY), Color
      Box.Line -(X + lngX, Y + lngY + 12), Color
      Box.Line -(X + lngX + 6, Y + lngY + 12), Color
      Box.Line -(X + lngX + 6, Y + lngY + 3), Color
      Box.Line -(X + lngX + 14, Y + lngY + 3), Color
      Box.Line -(X + lngX + 14, Y + lngY + 6), Color
      Box.Line -(X + lngX + 17, Y + lngY + 6), Color
      Box.Line -(X + lngX + 17, Y + lngY + 15), Color
      Box.Line -(X + lngX + 6, Y + lngY + 15), Color
      Box.Line -(X + lngX + 6, Y + lngY + 12), Color
      Box.PSet (X + lngX + 10, Y + lngY + 1), Color
      Box.PSet (X + lngX + 11, Y + lngY + 2), Color
      Box.PSet (X + lngX + 15, Y + lngY + 4), Color
      Box.PSet (X + lngX + 16, Y + lngY + 5), Color
      
      For intCount = 1 To 3
         Box.Line (X + lngX + 2, Y + lngY + 3 + intCount * 2)-(X + lngX + 5, Y + lngY + 3 + intCount * 2), Color
         Box.Line (X + lngX + 8, Y + lngY + 6 + intCount * 2)-(X + lngX + 16, Y + lngY + 6 + intCount * 2), Color
      Next 'intCount
      
   ElseIf Index = KEY_CLOSE Then
      lngX = ButtonState + 18
      lngY = ButtonState + 10
      Box.Circle (X + lngX, Y + lngY + 4), 9, Color
      Box.Line (X + lngX, Y + lngY)-(X + lngX + 1, Y + lngY + 9), Color, BF
      Box.PSet (X + lngX - 8, Y + lngY + 3), Color
      Box.PSet (X + lngX - 10, Y + lngY + 3), CalcButton(KEY_CLOSE).BackColor
      
   ElseIf Index = KEY_RECEIPT Then
      If m_ReceiptVisible Then
         lngX = ButtonState + 10
         lngY = ButtonState + 8
         Box.DrawWidth = 3
         Box.Line (X + lngX, Y + lngY)-(X + lngX + 15, Y + lngY + 16), Color, B
         Box.DrawWidth = 2
         Box.Line (X + lngX, Y + lngY + 6)-(X + lngX + 13, Y + lngY + 7), Color, BF
         Box.Line (X + lngX, Y + lngY + 10)-(X + lngX + 14, Y + lngY + 10), Color
         Box.Line (X + lngX, Y + lngY + 13)-(X + lngX + 15, Y + lngY + 13), Color
         Box.Line (X + lngX + 4, Y + lngY + 6)-(X + lngX + 4, Y + lngY + 17), Color
         Box.Line (X + lngX + 7, Y + lngY + 6)-(X + lngX + 7, Y + lngY + 17), Color
         Box.Line (X + lngX + 10, Y + lngY + 6)-(X + lngX + 10, Y + lngY + 17), Color
         
      Else
         Call DrawReceiptImage(Index, ButtonState, X, Y)
      End If
      
   ElseIf Index = KEY_RECEIPTCLEAR Then
      If (ButtonState = ButtonDown) And lstReceipt.ListCount Then
         Call DrawReceiptImage(Index, ButtonState, X, Y)
         
      Else
         lngX = ButtonState + 9
         lngY = ButtonState + 7
         Box.DrawWidth = 1
         Box.Line (X + lngX, Y + lngY)-(X + lngX + 17, Y + lngY + 18), Color, B
      End If
   End If
   
   DoEvents
   Box.DrawWidth = 1

End Sub

Private Sub DrawReceiptImage(ByVal Index As Integer, ByVal ButtonState As ButtonStates, ByVal X As Long, ByVal Y As Long)

Dim lngX As Long
Dim lngY As Long

   With CalcButton(Index)
      picCalculator.DrawWidth = 1
      lngX = 9 + ButtonState
      lngY = 7 + ButtonState
      ' Border
      picCalculator.Line (X + lngX, Y + lngY)-(X + lngX + 17, Y + lngY + 18), .ForeColor, B
      ' +
      picCalculator.Line (X + lngX + 4, Y + lngY + 5)-(X + lngX + 7, Y + lngY + 5), .ForeColor
      picCalculator.Line (X + lngX + 5, Y + lngY + 4)-(X + lngX + 5, Y + lngY + 7), .ForeColor
      ' 4
      picCalculator.Line (X + lngX + 10, Y + lngY + 2)-(X + lngX + 10, Y + lngY + 5), .ForeColor
      picCalculator.Line -(X + lngX + 14, Y + lngY + 5), Color
      picCalculator.Line (X + lngX + 13, Y + lngY + 2)-(X + lngX + 13, Y + lngY + 9), .ForeColor
      ' =
      picCalculator.Line (X + lngX + 4, Y + lngY + 13)-(X + lngX + 7, Y + lngY + 13), .ForeColor
      picCalculator.Line (X + lngX + 4, Y + lngY + 15)-(X + lngX + 7, Y + lngY + 15), .ForeColor
      ' 5
      picCalculator.Line (X + lngX + 13, Y + lngY + 11)-(X + lngX + 10, Y + lngY + 11), .ForeColor
      picCalculator.Line -(X + lngX + 10, Y + lngY + 13), .ForeColor
      picCalculator.Line -(X + lngX + 13, Y + lngY + 13), .ForeColor
      picCalculator.Line -(X + lngX + 13, Y + lngY + 16), .ForeColor
      picCalculator.Line -(X + lngX + 9, Y + lngY + 16), .ForeColor
      
      If Index = KEY_RECEIPTCLEAR Then
         picCalculator.DrawWidth = 2
         picCalculator.Line (X + lngX + 2, Y + lngY + 16)-(X + lngX + 15, Y + lngY + 2), Color
         picCalculator.Line (X + lngX + 2, Y + lngY + 2)-(X + lngX + 15, Y + lngY + 16), Color
      End If
      
      picCalculator.DrawWidth = 1
      DoEvents
   End With

End Sub

Private Sub FillButtons()

Dim intIndex As Integer

   lblDisplay.ForeColor = m_DisplaysOffColor
   lstReceipt.BackColor = m_ReceiptBackColor
   lstReceipt.ForeColor = m_ReceiptForeColor
   picDisplay.ForeColor = m_DisplaysOffColor
   picMemory.ForeColor = m_DisplaysOffColor
   
   ReDim CalcButton(33) As CalcButtons
   
   For intIndex = 0 To UBound(CalcButton)
      With CalcButton(intIndex)
         If intIndex < KEY_RETURN Then
            .Caption = Choose(intIndex + 1, "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "00", DecimalPoint)
            .BackColor = m_NumberKeysBackColor
            .ForeColor = m_NumberKeysForeColor
            
         ElseIf intIndex = KEY_RETURN Then
            .Caption = "Enter"
            .BackColor = m_EnterKeyBackColor
            .ForeColor = m_EnterKeyForeColor
            .Tag = "="
            
         ElseIf intIndex < KEY_CLEAR Then
            .Caption = Choose(intIndex - KEY_RETURN, "+", "", "X", "/", "x^y", "±", "1/x", "%", "", DecimalPoint & "x")
            .BackColor = m_OperandKeysBackColor
            .ForeColor = m_OperandKeysForeColor
            .Tag = Choose(intIndex - KEY_RETURN, "+", "-", "X", "/", "^", "", "", "", "Sqr", "")
            
         ElseIf intIndex < KEY_MEMORY_CLEAR Then
            .Caption = Choose(intIndex - KEY_CLEAR + 1, "C", "CE", "")
            .BackColor = m_CancelKeysBackColor
            .ForeColor = m_CancelKeysForeColor
            
         ElseIf intIndex < KEY_RECEIPT Then
            .Caption = Choose(intIndex - KEY_MEMORY_CLEAR + 1, "MC", "MR", "M+", "M-")
            .BackColor = m_MemoryKeysBackColor
            .ForeColor = m_MemoryKeysForeColor
            
         ' Function key colors
         Else
            .BackColor = m_FunctionKeysBackColor
            .ForeColor = m_FunctionKeysForeColor
         End If
         
         With .Rect
            .Top = Choose(intIndex + 1, 330, 290, 290, 290, 250, 250, 250, 210, 210, 210, 330, 330, 330, 290, 290, 250, 250, 170, 290, 250, 210, 210, 210, 170, 170, 170, 130, 130, 130, 130, 130, 130, 100, 100)
            .Left = Choose(intIndex + 1, 98, 18, 58, 98, 18, 58, 98, 18, 58, 98, 18, 138, 178, 138, 178, 138, 178, 198, 218, 218, 138, 178, 218, 18, 78, 138, 18, 58, 98, 138, 178, 214, 178, 214)
            
            If (intIndex = KEY_DOUBLENULL) Or (intIndex = KEY_RETURN) Then
               .Right = .Left + 71
               
            ElseIf (intIndex = KEY_POWERTO) Or (intIndex = KEY_CLEAR) Or (intIndex = KEY_CLEARENTRY) Or (intIndex = KEY_BACKSPACE) Then
               .Right = .Left + 51
               
            ElseIf (intIndex = KEY_RECEIPT) Or (intIndex = KEY_RECEIPTCLEAR) Or (intIndex = KEY_CLOSE) Or (intIndex = KEY_COPY) Then
               .Right = .Left + 35
               
            Else
               .Right = .Left + 31
            End If
            
            If (intIndex = KEY_CLOSE) Or (intIndex = KEY_COPY) Then
               .Bottom = .Top + 29
               
            Else
               .Bottom = .Top + 31
            End If
         End With
      End With
   Next 'intIndex
   
   Call ShowMemory
   Call SetDecimals
   Call ToggleReceipt

End Sub

Private Sub FillReceipt(ByVal Value As String)

Dim intSpaces  As Integer
Dim strOperand As String

   If Not m_ReceiptVisible Then Exit Sub
   
   Call DrawReceiptImage(KEY_RECEIPTCLEAR, ButtonUp, CalcButton(KEY_RECEIPTCLEAR).Rect.Left, CalcButton(KEY_RECEIPTCLEAR).Rect.Top)
   
   Value = Trim(Value)
   
   With lstReceipt
      If InStr("+-/X^=" & vbCrLf, Value) Then
         .AddItem ""
         
         If Value <> vbCrLf Then
            .List(.NewIndex) = Value
            .ListIndex = .NewIndex
         End If
         
         .ListIndex = .ListCount - 1
         
      Else
         If .ListCount = 0 Then .AddItem ""
         
         .ListIndex = .ListCount - 1
         
         If Left(.List(.ListIndex), 1) <> " " Then
            strOperand = Left(.List(.ListIndex), 1)
            
            If (strOperand = "") And ((ButtonIndex > 11) And (ButtonIndex < 23)) Then
               strOperand = CalcButton(ButtonIndex).Caption
               
               If strOperand = "" Then strOperand = CalcButton(ButtonIndex).Tag
               
               .List(.ListIndex) = strOperand
            End If
         End If
         
         If (CDbl(Value) = Int(CDbl(Value))) And (InStr(Value, DecimalPoint) = Len(Value)) Then Value = Format(Value, "0.")
         
         With picBuffer
            .Width = lstReceipt.Width
            .Cls
            picBuffer.Print Value;
            intSpaces = (.ScaleWidth - .CurrentX) / .TextWidth(" ") - 2 - Len(strOperand)
         End With
         
         .List(.ListIndex) = strOperand & Space(intSpaces) & Value
      End If
   End With

End Sub

Private Sub HandlePressedKey(ByVal Index As Integer)

Dim blnNegative  As Boolean
Dim blnShowValue As Boolean
Dim intPointer   As Integer
Dim strKeyValue  As String

   If MemoryError Then
      MemoryError = False
      
      Call ShowMemory
   End If
   
   If CalculationError And ((Index <> KEY_CLEAR) And (Index <> KEY_CLEARENTRY)) Then Exit Sub
   
   ' Handle ExtendedKeys
   If Index > KEY_MEMORY_SUBTRACT Then
      Select Case Index
         Case KEY_COPY
            Call SetToClipboard
            
         Case KEY_CLOSE
            MemoryValue = 0
            lblDisplay.ForeColor = m_DisplaysOffColor
            lblMemory.ForeColor = m_DisplaysOffColor
            strKeyValue = lblDisplay.Caption
            
            Call ShowMemory
            Call ClearCalculator("C")
            Call DrawButton(KEY_RECEIPTCLEAR, ButtonUp)
            
            RaiseEvent Closing(strKeyValue)
            
         Case KEY_RECEIPT
            m_ReceiptVisible = Not m_ReceiptVisible
            
            Call ToggleReceipt
            
         Case KEY_RECEIPTCLEAR
            If lstReceipt.ListCount Then lstReceipt.Clear
      End Select
      
      Exit Sub
      
   ' Handle MemoryKeys
   ElseIf Index > KEY_BACKSPACE Then
      Select Case Index
         Case KEY_MEMORY_CLEAR
            MemoryValue = 0
            
         Case KEY_MEMORY_RECALL
            Call DoFunction(Index)
            
            blnShowValue = True
            
         Case KEY_MEMORY_ADD
            Call CalculateMemory(lblDisplay.Caption, True)
            
         Case KEY_MEMORY_SUBTRACT
            Call CalculateMemory(lblDisplay.Caption, False)
      End Select
      
      Call ShowMemory
      
   ' Handle FunctionKeys
   ElseIf (Index > KEY_POWERTO) Then
      Select Case Index
         Case KEY_INVERT, KEY_PERCENT, KEY_SQUAREROOT, KEY_INVERSE
            Call DoFunction(Index)
            
            blnShowValue = True
            
         Case KEY_DECIMALCOUNT
            m_DecimalCount = m_DecimalCount + 1
            
            If m_DecimalCount > 4 Then m_DecimalCount = -1
            
            Call SetDecimals
            
         Case KEY_CLEAR
            Call ClearCalculator("C")
            Call DrawButton(KEY_RECEIPTCLEAR, ButtonUp)
            
         Case KEY_CLEARENTRY
            Call ClearEntry
            
         Case KEY_BACKSPACE
            If LastInput = Calculation Then Exit Sub
            
            Call DoBackSpace
      End Select
      
   ' Handle OperandKeys
   ElseIf Index > KEY_DECIMAL Then
      Call Calculate(Index)
      
   ElseIf Index < 0 Then
      Exit Sub
      
   ' Handle DecimalKey and NumberKeys
   Else
      With lblDisplay
         If Index = KEY_DECIMAL Then
            DecimalFlag = True
            
         Else
            If LastInput = Numbers Then
               If Left(.Caption, 1) = "-" Then
                  blnNegative = True
                  lblDisplay.Caption = Mid(lblDisplay.Caption, 2)
               End If
               
               If Len(lblDisplay.Caption) > 15 Then
                  If blnNegative Then .Caption = "-" & .Caption
                  
                  Exit Sub
               End If
            End If
            
            strKeyValue = CalcButton(Index).Caption
            
            If Len(.Caption & strKeyValue) > 16 Then strKeyValue = Mid(strKeyValue, 2)
            
            If LastInput <> Numbers Then
               .Caption = Format(0, "0.")
               
               If LastInput <> DecimalSign Then DecimalFlag = False
               
            ElseIf (strKeyValue = "0") And (Val(.Caption) = 0) Then
               strKeyValue = ""
            End If
            
            If DecimalFlag Then
               If m_DecimalCount > -1 Then
                  intPointer = Len(.Caption) - InStr(.Caption, DecimalPoint) + 1
                  
                  If intPointer > m_DecimalCount Then strKeyValue = ""
               End If
               
               .Caption = .Caption & strKeyValue
               
            Else
               .Caption = Format(CStr(CDbl("0" & .Caption)) & strKeyValue, "0.")
            End If
            
            If blnNegative Then .Caption = "-" & .Caption
         End If
      End With
      
      If (LastInput = Operands) And (OperationFlag = "=") Then Call FillReceipt(vbCrLf)
      
      If DecimalFlag And (LastInput = Operands) Then
         LastInput = DecimalSign
         
      Else
         LastInput = Numbers
      End If
      
      Call FillReceipt(lblDisplay.Caption)
   End If
   
   lblDisplay.ForeColor = m_CalculatorOnColor
   
   Call ShowKeyPressed(Index)
   
   If blnShowValue Then
      If Operand(0) = Int(Operand(0)) Then
         lblDisplay.Caption = Format(Operand(0), "0.")
         
      Else
         lblDisplay.Caption = Operand(0)
      End If
      
      Call FillReceipt(lblDisplay.Caption)
   End If

End Sub

Private Sub SetDecimals()

Dim strCode As String

   If m_DecimalCount = -1 Then
      strCode = "x"
      
   Else
      strCode = m_DecimalCount
   End If
   
   CalcButton(KEY_DECIMALCOUNT).Caption = CalcButton(KEY_DECIMAL).Caption & strCode
   
   If ButtonIndex > -1 Then Call DrawButton(ButtonIndex, ButtonUp)

End Sub

Private Sub SetDefaults()

   DecimalPoint = Format(0, ".")
   ButtonIndex = -1
   
   Call ClearCalculator("C")

End Sub

Private Sub SetDisplays()

   With CalcButton(KEY_MEMORY_CLEAR).Rect
      picDisplay.Left = .Left - 6
      picMemory.Left = .Left + 4
      picMemory.Width = CalcButton(KEY_MEMORY_SUBTRACT).Rect.Right - picMemory.Left - 2
   End With
   
   With CalcButton(KEY_CLOSE).Rect
      picDisplay.Top = (.Top - picDisplay.Height) \ 2 - 2
      picDisplay.Width = .Right + 8 - picDisplay.Left
      picMemory.Top = .Top + 3
      picMemory.Height = .Bottom - .Top - 10
   End With
   
   With lblMemory
      .Font.Size = .Font.Size + (2 And (Screen.TwipsPerPixelY <> 12))
      .AutoSize = True
      .Top = (picMemory.ScaleHeight - .Height) \ 2
      .Left = picMemory.ScaleWidth - .Width - 3
   End With
   
   With lblDisplay
      .Font.Size = .Font.Size + (2 And (Screen.TwipsPerPixelY <> 12))
      .AutoSize = True
      .Top = (picDisplay.ScaleHeight - .Height) \ 2 + 1
      .Left = picDisplay.ScaleWidth - .Width - 5
   End With
   
End Sub

Private Sub SetPressedKey(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal ButtonState As ButtonStates)

Const vbKeyPowerTo As Integer = 154

Static blnMemory As Boolean

Dim intIndex     As Integer

   intIndex = KeyCode
   
   If Shift And vbShiftMask Then
      ' +
      If KeyCode = 187 Then
         intIndex = vbKeyAdd
         
      ' *
      ElseIf KeyCode = 56 Then
         intIndex = vbKeyMultiply
         
      ' *
      ElseIf KeyCode = 54 Then
         intIndex = vbKeyPowerTo
      End If
      
   ' =
   ElseIf KeyCode = 187 Then
      intIndex = vbKeyReturn
      
   ' , or .
   ElseIf (KeyCode = 188) Or (KeyCode = 190) Then
      intIndex = vbKeyDecimal
      
   ' -
   ElseIf KeyCode = 189 Then
      intIndex = vbKeySubtract
      
   ' /
   ElseIf KeyCode = 191 Then
      intIndex = vbKeyDivide
   End If
   
   Select Case intIndex
      Case vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
         intIndex = KeyCode - vbKey0 - (vbKey0 And (KeyCode >= vbKeyNumpad0))
         
      Case vbKeyDecimal
         intIndex = KEY_DECIMAL
         
      ' Add or Add to Memory
      Case vbKeyAdd
         If blnMemory Then
            intIndex = KEY_MEMORY_ADD
            
            If ButtonState = ButtonUp Then blnMemory = False
            
         Else
            intIndex = KEY_ADD
         End If
         
      Case vbKeyMultiply
         intIndex = KEY_MULTIPLY
         
      ' Subtract or Subtract from  Memory
      Case vbKeySubtract
         If blnMemory Then
            intIndex = KEY_MEMORY_SUBTRACT
            
            If ButtonState = ButtonUp Then blnMemory = False
            
         Else
            intIndex = KEY_SUBTRACT
         End If
         
      Case vbKeyDivide
         intIndex = KEY_DIVIDE
         
      Case vbKeyReturn
         intIndex = KEY_RETURN
         
      Case vbKeyBack
         intIndex = KEY_BACKSPACE
         
      ' Clear Calculator
      Case vbKeyEscape
         intIndex = KEY_CLEAR
         
      ' Clear Entry
      Case vbKeyDelete
         intIndex = KEY_CLEARENTRY
         
      ' Toggle Receipt On/Off
      Case vbKeyF2
         intIndex = KEY_RECEIPT
         
      ' Clear Receipt
      Case vbKeyF3
         intIndex = KEY_RECEIPTCLEAR
         
      ' Inverse
      Case vbKeyF9
         intIndex = KEY_INVERSE
         
      ' Clear Memory
      Case vbKeyC
         If blnMemory Then
            intIndex = KEY_MEMORY_CLEAR
            
            If ButtonState = ButtonUp Then blnMemory = False
            
         Else
            intIndex = -1
         End If
         
      ' Double Null
      Case vbKeyD
         intIndex = KEY_DOUBLENULL
         
      ' Invert
      Case vbKeyI
         intIndex = KEY_INVERT
         
      ' Activate Memory keys
      Case vbKeyM
         intIndex = -2
         
         If ButtonState = ButtonDown Then blnMemory = True
         
      ' Percent
      Case vbKeyP
         intIndex = KEY_PERCENT
         
      ' PowerTo
      Case vbKeyPowerTo
         intIndex = KEY_POWERTO
         
      ' Recall Memory
      Case vbKeyR
         If blnMemory Then
            intIndex = KEY_MEMORY_RECALL
            
            If ButtonState = ButtonUp Then blnMemory = False
            
         Else
            intIndex = -1
         End If
         
      ' SquareRoot
      Case vbKeyS
         intIndex = KEY_SQUAREROOT
         
      ' Decimal count
      Case vbKeyX
         intIndex = KEY_DECIMALCOUNT
         
      ' Invalid key is pressed
      Case Else
         intIndex = -1
   End Select
   
   If intIndex > -1 Then
      Call DrawButton(intIndex, ButtonState)
      
      If ButtonState = ButtonDown Then Call HandlePressedKey(intIndex)
      
   ElseIf (intIndex = -1) And (ButtonState = ButtonUp) Then
      blnMemory = False
   End If

End Sub

Private Sub SetToClipboard()

Dim strValue As String

   If Not CalculationError Then
      If DecimalCount Then
         strValue = lblDisplay.Caption
         
      Else
         strValue = Val(lblDisplay.Caption)
      End If
      
      Clipboard.Clear
      Clipboard.SetText strValue
   End If

End Sub

Private Sub ShowButtons()

Dim intIndex As Integer

   For intIndex = 0 To UBound(CalcButton)
      Call DrawButton(intIndex, ButtonUp)
   Next 'intIndex

End Sub

Private Sub ShowMemory()

Dim strSymbol As String

   With picMemory
      strSymbol = "M"
      
      If (MemoryValue = 0) Or (CDbl(lblDisplay.Caption) = 0) Then
         .ForeColor = m_DisplaysOffColor
         
      ElseIf MemoryError Then
         .ForeColor = m_ErrorColor
         strSymbol = "E"
         
      Else
         .ForeColor = m_MemoryOnColor
      End If
      
      .Cls
      .CurrentX = (31 - .TextWidth(strSymbol)) / 2 - 7
      .CurrentY = (.ScaleHeight - .TextHeight(strSymbol)) / 2
   End With
   
   picMemory.Print strSymbol
   
   With lblMemory
      If MemoryValue = Int(MemoryValue) Then
         .Caption = Format(MemoryValue, "0.")
         
      Else
         .Caption = MemoryValue
      End If
      
      If Not MemoryError Then .ForeColor = picMemory.ForeColor
      
      .Left = picMemory.ScaleWidth - .Width - 3
   End With

End Sub

Private Sub ShowKeyPressed(Optional ByVal Index As Integer)

Dim lngColor  As Long
Dim strSymbol As String

   If Not m_ShowPressedKey Then Exit Sub
   
   With picDisplay
      If CalculationError Then
         .FontSize = 12
         lngColor = m_ErrorColor
         strSymbol = "E"
         
      Else
         .FontSize = 10
         lngColor = m_PressedKeyColor
         
         With CalcButton(Index)
            If .Tag = "" Then
               strSymbol = .Caption
               
            Else
               strSymbol = .Tag
            End If
         End With
      End If
      
      .Picture = Nothing
      .Cls
      .ForeColor = lngColor
      .CurrentX = (31 - .TextWidth(strSymbol)) / 2
      .CurrentY = (.ScaleHeight - .TextHeight(strSymbol)) / 2 + 2 - (1 And Not CalculationError)
   End With
   
   If Not CalculationError And ((Index = KEY_SUBTRACT) Or (Index = KEY_BACKSPACE) Or (Index = KEY_SQUAREROOT)) Then
      Call DrawButtonImage(Index, ButtonUp, lngColor, 0, 2, picDisplay)
      
   Else
      picDisplay.Print strSymbol
   End If

End Sub

Private Sub ShowReceipt()

   If m_ReceiptVisible Then
      FillStyle = vbFSSolid
      FillColor = m_BackColor
      ExtFloodFill hDC, 400, 2, &H404040, FLOODFILLBORDER
      FillStyle = vbFSTransparent
   End If
   
   Call UserControl_Resize
   Call DrawBorder
   
   RaiseEvent ReceiptState(m_ReceiptVisible)

End Sub

Private Sub ToggleReceipt()

   Call ShowReceipt
   Call DrawButton(KEY_RECEIPT, ButtonUp)

End Sub

Private Sub picCalculator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then
      If ButtonIndex = -1 Then Call CheckAllButtons(CLng(X), CLng(Y))
      
      If ButtonIndex > -1 Then
         Call DrawButton(ButtonIndex, ButtonDown)
         
         MouseOut = False
      End If
   End If

End Sub

Private Sub picCalculator_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If ButtonIndex > -1 Then
      If IsInButton(CalcButton(ButtonIndex).Rect, CLng(X), CLng(Y)) Then
         If Button = vbLeftButton Then Call DrawButton(ButtonIndex, ButtonDown)
         
         MouseOut = False
         
      Else
         Call DrawButton(ButtonIndex, ButtonUp)
         
         MouseOut = True
         
         If Button <> vbLeftButton Then ButtonIndex = -1
      End If
      
   Else
      MouseOut = True
   End If

End Sub

Private Sub picCalculator_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If (Button = vbLeftButton) And Not MouseOut Then
      If ButtonIndex > -1 Then
         Call DrawButton(ButtonIndex, ButtonUp)
         Call HandlePressedKey(ButtonIndex)
      End If
      
   ElseIf MouseOut Then
      ButtonIndex = -1
   End If

End Sub

Private Sub UserControl_ExitFocus()

   RaiseEvent ExitFocus

End Sub

Private Sub UserControl_InitProperties()

   m_BackColor = &HA9B198
   m_BorderColor = &H404040
   m_BorderStyle = Colored
   m_CalculatorColor = &H89A59F
   m_CancelKeysBackColor = &HF2C6FF
   m_CancelKeysForeColor = &H800080
   m_DecimalCount = -1
   m_DisplaysBackColor = &HFF6800
   m_DisplaysOffColor = &HC0C000
   m_CalculatorOnColor = &HFFFF00
   m_EnterKeyBackColor = &HC6C6C6
   m_EnterKeyForeColor = &HC00000
   m_ErrorColor = &H8080FF
   m_FunctionKeysBackColor = &HE6FFFF
   m_FunctionKeysForeColor = &H800080
   m_MemoryKeysBackColor = &HC0FFDF
   m_MemoryKeysForeColor = &HC00000
   m_MemoryOnColor = &HFFFFC0
   m_NumberKeysBackColor = &HFFE0E0
   m_NumberKeysForeColor = &HC00000
   m_OperandKeysBackColor = &HEFEFEF
   m_OperandKeysForeColor = &H800080
   m_PressedKeyColor = &HC0C000
   m_ReceiptBackColor = &HE6FFFF
   m_ReceiptForeColor = &H800080
   m_ReceiptVisible = False
   m_RepeatKey = True
   m_ShowPressedKey = True
   
   Call SetDefaults
   Call Refresh

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   If (Not m_RepeatKey And KeyDown) Or ((PressedKey <> vbDefault) And (PressedKey <> KeyCode)) Or (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or (KeyCode = vbKeyPageUp) Or (KeyCode = vbKeyPageDown) Or (KeyCode = vbKeyHome) Or (KeyCode = vbKeyEnd) Then Exit Sub
   If ((KeyCode = vbKeyC) Or (KeyCode = vbKeyInsert)) And (Shift = vbCtrlMask) Then Call SetToClipboard
   If (KeyCode > vbKeySpace) Or (KeyCode = vbKeyBack) Or (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyEscape) Then PressedKey = KeyCode
   
   KeyDown = True
   
   Call SetPressedKey(KeyCode, Shift, ButtonDown)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or (KeyCode = vbKeyPageUp) Or (KeyCode = vbKeyPageDown) Or (KeyCode = vbKeyHome) Or (KeyCode = vbKeyEnd) Then Exit Sub
   If (PressedKey <> KeyCode) Or ((KeyCode < vbKeySpace) And (KeyCode <> vbKeyBack) And (KeyCode <> vbKeyDelete) And (KeyCode <> vbKeyReturn) And (KeyCode <> vbKeyEscape)) Then Exit Sub
   
   Call SetPressedKey(KeyCode, Shift, ButtonUp)
   
   KeyDown = False
   PressedKey = vbDefault

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_BackColor = .ReadProperty("BackColor", &HA9B198)
      m_BorderColor = .ReadProperty("BorderColor", &H404040)
      m_BorderStyle = .ReadProperty("BorderStyle", Colored)
      m_CalculatorColor = .ReadProperty("CalculatorColor", &H89A59F)
      m_CancelKeysBackColor = .ReadProperty("CancelKeysBackColor", &HF2C6FF)
      m_CancelKeysForeColor = .ReadProperty("CancelKeysForeColor", &H800080)
      m_DecimalCount = .ReadProperty("DecimalCount", -1)
      m_DisplaysBackColor = .ReadProperty("DisplaysBackColor", &HFF6800)
      m_DisplaysOffColor = .ReadProperty("DisplaysOffColor", &HC0C000)
      m_CalculatorOnColor = .ReadProperty("CalculatorOnColor", &HFFFF00)
      m_EnterKeyBackColor = .ReadProperty("EnterKeyBackColor", &HC6C6C6)
      m_EnterKeyForeColor = .ReadProperty("EnterKeyForeColor", &HC00000)
      m_ErrorColor = .ReadProperty("ErrorColor", &H8080FF)
      m_FunctionKeysBackColor = .ReadProperty("FunctionKeysBackColor", &HE6FFFF)
      m_FunctionKeysForeColor = .ReadProperty("FunctionKeysForeColor", &H800080)
      m_MemoryKeysBackColor = .ReadProperty("MemoryKeysBackColor", &HC0FFDF)
      m_MemoryKeysForeColor = .ReadProperty("MemoryKeysForeColor", &HC00000)
      m_MemoryOnColor = .ReadProperty("MemoryOnColor", &HFFFFC0)
      m_NumberKeysBackColor = .ReadProperty("NumberKeysBackColor", &HFFE0E0)
      m_NumberKeysForeColor = .ReadProperty("NumberKeysForeColor", &HC00000)
      m_OperandKeysBackColor = .ReadProperty("OperandKeysBackColor", &HEFEFEF)
      m_OperandKeysForeColor = .ReadProperty("OperandKeysForeColor", &H800080)
      m_PressedKeyColor = .ReadProperty("PressedKeyColor", &HC0C000)
      m_ReceiptBackColor = .ReadProperty("ReceiptBackColor", &HE6FFFF)
      m_ReceiptForeColor = .ReadProperty("ReceiptForeColor", &H800080)
      m_ReceiptVisible = .ReadProperty("ReceiptVisible", False)
      m_RepeatKey = .ReadProperty("RepeatKey", True)
      m_ShowPressedKey = .ReadProperty("ShowPressedKey", True)
   End With
   
   Call SetDefaults
   Call UserControl_Resize
   Call Refresh
   Call SetDisplays

End Sub

Private Sub UserControl_Resize()

Dim blnBusy   As Boolean
Dim lngHeight As Long
Dim lngWidth  As Long

   If blnBusy Then Exit Sub
   
   blnBusy = True
   lngHeight = picCalculator.Height + 4
   lngWidth = picCalculator.Width + 4
   
   With lstReceipt
      .Left = lngWidth
      .Width = (TextWidth("X") * 22) + 4
      .Height = ScaleHeight - 4
      .Top = (ScaleHeight - .Height) \ 2
      
      If m_ReceiptVisible Then lngWidth = .Left + .Width + 6
   End With
   
   Width = lngWidth * Screen.TwipsPerPixelX
   Height = lngHeight * Screen.TwipsPerPixelY
   blnBusy = False

End Sub

Private Sub UserControl_Terminate()

   Erase CalcButton

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BackColor", m_BackColor, &HA9B198
      .WriteProperty "BorderColor", m_BorderColor, &H404040
      .WriteProperty "BorderStyle", m_BorderStyle, Colored
      .WriteProperty "CalculatorColor", m_CalculatorColor, &H89A59F
      .WriteProperty "CancelKeysBackColor", m_CancelKeysBackColor, &HF2C6FF
      .WriteProperty "CancelKeysForeColor", m_CancelKeysForeColor, &H800080
      .WriteProperty "DecimalCount", m_DecimalCount, -1
      .WriteProperty "DisplaysBackColor", m_DisplaysBackColor, &HFF6800
      .WriteProperty "DisplaysOffColor", m_DisplaysOffColor, &HC0C000
      .WriteProperty "CalculatorOnColor", m_CalculatorOnColor, &HFFFF00
      .WriteProperty "EnterKeyBackColor", m_EnterKeyBackColor, &HC6C6C6
      .WriteProperty "EnterKeyForeColor", m_EnterKeyForeColor, &HC00000
      .WriteProperty "ErrorColor", m_ErrorColor, &H8080FF
      .WriteProperty "FunctionKeysBackColor", m_FunctionKeysBackColor, &HE6FFFF
      .WriteProperty "FunctionKeysForeColor", m_FunctionKeysForeColor, &H800080
      .WriteProperty "MemoryKeysBackColor", m_MemoryKeysBackColor, &HC0FFDF
      .WriteProperty "MemoryKeysForeColor", m_MemoryKeysForeColor, &HC00000
      .WriteProperty "MemoryOnColor", m_MemoryOnColor, &HFFFFC0
      .WriteProperty "NumberKeysBackColor", m_NumberKeysBackColor, &HFFE0E0
      .WriteProperty "NumberKeysForeColor", m_NumberKeysForeColor, &HC00000
      .WriteProperty "OperandKeysBackColor", m_OperandKeysBackColor, &HEFEFEF
      .WriteProperty "OperandKeysForeColor", m_OperandKeysForeColor, &H800080
      .WriteProperty "PressedKeyColor", m_PressedKeyColor, &HC0C000
      .WriteProperty "ReceiptBackColor", m_ReceiptBackColor, &HE6FFFF
      .WriteProperty "ReceiptForeColor", m_ReceiptForeColor, &H800080
      .WriteProperty "ReceiptVisible", m_ReceiptVisible, False
      .WriteProperty "RepeatKey", m_RepeatKey, True
      .WriteProperty "ShowPressedKey", m_ShowPressedKey, True
   End With

End Sub

