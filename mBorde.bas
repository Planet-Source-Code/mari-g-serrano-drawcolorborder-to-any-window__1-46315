Attribute VB_Name = "mBorde"
Option Explicit
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const CLR_INVALID As Integer = 0
Private Const PS_SOLID    As Integer = 0
Private Const SM_CXBORDER As Integer = 5
Private Const SM_CYBORDER As Integer = 6


Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private rc As RECT
Public Sub DrawBorder(lhWnd As Long, color As OLE_COLOR, Optional mWidth As Integer = 0)

  Dim hWindowDC As Long
  Dim hOldPen   As Long
  Dim nLeft     As Long
  Dim nRight    As Long
  Dim nTop      As Long
  Dim nBottom   As Long
  Dim Ret       As Long
  Dim hMyPen    As Long
  Dim WidthX    As Long
  Dim rgbColor  As Long
    rgbColor = TranslateColor(color)
    ' border width
    If mWidth = 0 Then
        WidthX = GetSystemMetrics(SM_CYBORDER) * 5
    Else
        WidthX = mWidth
    End If
    ' window DC
    hWindowDC = GetWindowDC(lhWnd)   'this is outside the form
    ' create a pen
    hMyPen = CreatePen(PS_SOLID, WidthX, rgbColor)
    ' Init variables
    GetWindowRect lhWnd, rc
    
    nRight = rc.Right - rc.Left 'for HwNd
    nBottom = rc.Bottom - rc.Top
    
    ''nRight = frmTarget.Width / Screen.TwipsPerPixelX 'if you want to do the same with
                                                       'Form as a Param
    ''nBottom = frmTarget.Height / Screen.TwipsPerPixelY
    ' selecciona borde del pen
    hOldPen = SelectObject(hWindowDC, hMyPen)
    ' draw the border
    Ret = LineTo(hWindowDC, nLeft, nBottom)
    Ret = LineTo(hWindowDC, nRight, nBottom)
    Ret = LineTo(hWindowDC, nRight, nTop)
    Ret = LineTo(hWindowDC, nLeft, nTop)
    ' select old pen
    Ret = SelectObject(hWindowDC, hOldPen)
    Ret = DeleteObject(hMyPen)
    Ret = ReleaseDC(lhWnd, hWindowDC)

End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hpal As Long = 0) As Long
  ' OLE_COLOR 2 RGB
  If OleTranslateColor(clr, hpal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function

