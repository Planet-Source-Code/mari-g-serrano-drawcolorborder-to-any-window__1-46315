VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MaRiØ_DrawBorder!"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "STOP && GO"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1380
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'MaRiØ 2003

Public Function GetActualColor(ByVal OleValue As OLE_COLOR) As Long
    'Get True Color  from System Color
    If (OleValue And &H80000000) Then
        GetActualColor = GetSysColor(OleValue And &H7FF)
    Else
        GetActualColor = OleValue
    End If
End Function
Private Sub GetRGB(ByVal color As Long, Red As Long, Green As Long, Blue As Long)
    Dim c As Long
    
    c = (color And &HFF&)
    Red = CByte(c)
    
    c = ((color And &HFF00&) / &H100&)
    Green = CByte(c)
    
    c = ((color And &HFF0000) / &H10000)
    Blue = CByte(c)
    
End Sub
Public Sub Smooth(ByRef color As Long, Optional toWhite As Boolean = True)
    Dim R As Long, G As Long, b As Long
    GetRGB GetActualColor(color), R, G, b
    If toWhite Then
        R = Abs(R + 19)
        G = Abs(G + 19)
        b = Abs(b + 19)
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If b > 255 Then b = 255
    Else
        R = Abs(R - 19)
        G = Abs(G - 19)
        b = Abs(b - 19)
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If b < 0 Then b = 0
    End If
    
    color = RGB(R, G, b)
End Sub

Public Function Smooth2(ByRef color As Long) As Long
    Static toWhite As Boolean
    Const VA As Byte = 15
    Dim R As Long, G As Long, b As Long
    GetRGB GetActualColor(color), R, G, b
    
    If toWhite Then
        R = (R + VA)
        G = (G + VA)
        b = (b + VA)
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If b > 255 Then b = 255
        If R = 255 And G = 255 And b = 255 Then toWhite = True
    Else
        R = (R - VA)
        G = (G - VA)
        b = (b - VA)
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If b < 0 Then b = 0
        If R = 0 And G = 0 And b = 0 Then toWhite = False
    End If
    
    Smooth2 = RGB(R, G, b)
End Function

Private Sub Command1_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Resize()
    DrawBorder Me.hwnd, &H8000000D
End Sub

Private Sub Timer1_Timer()
    Static color As Long
    If color = 0 Then color = &H8000000D
    color = Smooth2(color)
    DrawBorder Me.hwnd, color
End Sub

