VERSION 5.00
Begin VB.Form frmCalculator 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   6492
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   7932
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   StartUpPosition =   2  'CenterScreen
   Begin Caclulator.Calculator Calculator1 
      Height          =   4656
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   8213
      ReceiptVisible  =   -1  'True
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calculator1_Closing(Value As String)

   Calculator1.Visible = False

End Sub

Private Sub Form_Click()

   If Not Calculator1.Visible Then Calculator1.Visible = True

End Sub

