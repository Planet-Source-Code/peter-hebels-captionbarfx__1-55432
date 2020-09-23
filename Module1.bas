Attribute VB_Name = "CaptionbarFX"
'***********************************************************************************************************
'                                                                                                          *
'CaptionbarFX V5 project by Peter Hebels, Website "http://www.phsoft.nl"                                   *
'The author of this code cannot be held responsible for any damages may caused by this project.            *
'                                                                                                          *
'***********************************************************************************************************

'---------------Begin API Calls--------------
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
    ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
'---------------End API Calls--------------

'---------------Begin Const----------------
Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const LF_FACESIZE = 32
Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZENESW = 32643&
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
Private Const WS_DLGFRAME = &H400000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_POPUP = &H80000000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_END_ELLIPSIS = &H8000&
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const TRANSPARENT = 1
Private Const SM_CXBORDER = 5
Private Const SM_CXDLGFRAME = 7
Private Const SM_CXFRAME = 32
Private Const SM_CXICON = 11
Private Const SM_CXSMSIZE = 30
Private Const SM_CYBORDER = 6
Private Const SM_CYCAPTION = 4
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYFRAME = 33
Private Const SM_CYICON = 12
Private Const SM_CYMENU = 15
Private Const SM_CYSMSIZE = 31
Private Const DFC_CAPTION = 1
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONHELP = &H4
Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_INACTIVE = &H100
Private Const WM_SIZE = &H5
Private Const WM_SETCURSOR = &H20
Private Const WM_GETICON = &H7F
Private Const WM_SETICON = &H80
Private Const WM_NCACTIVATE = &H86
Private Const WM_MDIACTIVATE = &H222
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_MDIGETACTIVE = &H229
Private Const MA_ACTIVATE = 1
Private Const WM_SETTEXT = &HC
Private Const WM_NCPAINT = &H85
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_SYSCOMMAND = &H112
Private Const WM_INITMENUPOPUP = &H117
Private Const SC_MOUSEMENU = &HF090&
Private Const SC_MOVE = &HF010&
Private Const HTCAPTION = 2
Private Const HTSYSMENU = 3
Private Const HTLEFT = 10
Private Const HTRIGHT = 11
Private Const HTTOP = 12
Private Const HTTOPLEFT = 13
Private Const HTTOPRIGHT = 14
Private Const HTBOTTOM = 15
Private Const HTBOTTOMLEFT = 16
Private Const HTBOTTOMRIGHT = 17

Public Const MERGEPAINT = &HBB0226
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
'---------------End Const------------

'---------------Begin Variables--------------
Public GradForceColors As Boolean
Public GradVerticalGradient As Boolean
Public GradForcedText As Long, GradForcedTextA As Long
Public GradForcedFirst As Long, GradForcedSecond As Long
Public GradForcedFirstA As Long, GradForcedSecondA As Long
Public ButtonRdown As Boolean
Public AddBitmap As Boolean
Public BitmapDC As Long
Public BitmapW As Long
Public BitmapH As Long

Dim GradhWnd As Long
Dim GradIcon As Long
Dim DrawDC As Long
Dim tmpDC As Long
Dim hRgn As Long
Dim tmpGradFont As Long
Dim CaptionFont As LOGFONT
Dim I As Long
'---------------End Variables--------------

'---------------Begin Types----------------
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'---------------End Types------------------

'---------------Begin Objects--------------
Public WithForm As Object
Public MenuFrm As Object
'---------------End Objects----------------

Private Sub BarGetColors(IsActive As Boolean, LColor As Long, RColor As Long)
    If IsActive Then
        If GradForceColors Then
            LColor = GradForcedFirst
            RColor = GradForcedSecond
        Else
            LColor = vbBlack
            RColor = GetSysColor(COLOR_ACTIVECAPTION)
        End If
    Else
        If GradForceColors Then
            LColor = GradForcedFirstA
            RColor = GradForcedSecondA
        Else
            LColor = vbBlack
            RColor = GetSysColor(COLOR_INACTIVECAPTION)
        End If
    End If
End Sub

Public Sub GradientCaption(TheForm As Form)
    If TheForm.BorderStyle = 0 Then Exit Sub
    If TheForm.BorderStyle = 4 Then Exit Sub
    If TheForm.BorderStyle = 5 Then Exit Sub
    
    Dim ProcTmp As Long

    ProcTmp = SetWindowLong(TheForm.hwnd, GWL_WNDPROC, AddressOf GradientCallback)
    SetProp TheForm.hwnd, "OldMeProc", ProcTmp
End Sub

Public Sub GradientGetCapsFont()
    Dim NCM As NONCLIENTMETRICS
    Dim lfNew As LOGFONT

    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)

    CaptionFont = NCM.lfCaptionFont
End Sub

Private Sub GetCaptionRect(hwnd As Long, rct As RECT)
    Dim XCapBorder As Long
    Dim fCapStyle As Long
    Dim YCapHeight As Long

    YCapHeight = GetSystemMetrics(SM_CYCAPTION)
    fCapStyle = GetWindowLong(hwnd, GWL_STYLE)
    Select Case fCapStyle And &H80
    Case &H80
        XCapBorder = GetSystemMetrics(SM_CXDLGFRAME)
    Case Else
        XCapBorder = GetSystemMetrics(SM_CXFRAME)
    End Select

    rct.Left = XCapBorder
    rct.Right = XCapBorder
    rct.Top = XCapBorder
    rct.Bottom = rct.Top + YCapHeight - 1
End Sub

Private Sub GradateColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)
    Dim dblR As Double, dblG As Double, dblB As Double
    Dim addR As Double, addG As Double, addB As Double
    Dim bckR As Double, bckG As Double, bckB As Double

    dblR = CDbl(Color1 And &HFF)
    dblG = CDbl(Color1 And &HFF00&) / 255
    dblB = CDbl(Color1 And &HFF0000) / &HFF00&
    bckR = CDbl(Color2 And &HFF&)
    bckG = CDbl(Color2 And &HFF00&) / 255
    bckB = CDbl(Color2 And &HFF0000) / &HFF00&

    addR = (bckR - dblR) / UBound(Colors)
    addG = (bckG - dblG) / UBound(Colors)
    addB = (bckB - dblB) / UBound(Colors)

    For I = 0 To UBound(Colors)
        dblR = dblR + addR
        dblG = dblG + addG
        dblB = dblB + addB
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Colors(I) = RGB(dblR, dblG, dblB)
    Next
End Sub

Private Function DrawGradient(ByVal Color1 As Long, ByVal Color2 As Long) As Long
    Dim DestWidth As Long, DestHeight As Long
    Dim StartPnt As Long, EndPnt As Long
    Dim PixelStep As Long, XBorder As Long
    Dim WndRect As RECT
    Dim OldFont As Long
    Dim fStyle As Long, fText As String
    Dim SMSize As Long, SMSizeY As Long
    
    On Error Resume Next
    
    SMSize = GetSystemMetrics(SM_CXSMSIZE)
    SMSizeY = GetSystemMetrics(SM_CYSMSIZE)
    
    GetWindowRect GradhWnd, WndRect
    
    DestWidth = WndRect.Right - WndRect.Left
    
    DestHeight = GetSystemMetrics(SM_CYCAPTION)
    fText = Space$(255)
    
    Call GetWindowText(GradhWnd, fText, 255)
    
    fText = Trim$(fText)
    fStyle = GetWindowLong(GradhWnd, GWL_STYLE)
    
    Select Case fStyle And &H80
    
    Case &H80
       
        XBorder = GetSystemMetrics(SM_CXDLGFRAME)
        DestWidth = (DestWidth - XBorder)
    Case Else
        XBorder = GetSystemMetrics(SM_CXFRAME)
        DestWidth = DestWidth - XBorder
    End Select
    
    StartPnt = XBorder
    EndPnt = XBorder + DestWidth - 4
    
    Dim rct As RECT
    Dim hBr As Long
    
    If Not GradVerticalGradient Then
        PixelStep = DestWidth \ 8
        ReDim Colors(PixelStep) As Long
        GradateColors Colors(), Color1, Color2
    
        rct.Top = XBorder
        rct.Left = XBorder
        rct.Right = XBorder + (DestWidth \ PixelStep)
        rct.Bottom = (XBorder + DestHeight - 1)
    
        If (fStyle And &H80) = &H80 Then EndPnt = EndPnt + 1
    
        For I = 0 To PixelStep - 1
            hBr = CreateSolidBrush(Colors(I))
            FillRect DrawDC, rct, hBr
            DeleteObject hBr
            OffsetRect rct, (DestWidth \ PixelStep), 0
            If I = PixelStep - 2 Then rct.Right = EndPnt
        Next
    Else
        PixelStep = DestHeight \ 1
        ReDim Colors(PixelStep) As Long
        GradateColors Colors(), Color2, Color1
    
        rct.Top = XBorder
        rct.Left = XBorder
        If (fStyle And &H80) = &H80 Then
            rct.Right = (XBorder * 2) + DestWidth + 2
        Else
            rct.Right = (XBorder * 2) + DestWidth
        End If
        rct.Bottom = XBorder + (DestHeight \ PixelStep)
    
        For I = 0 To PixelStep - 1
            hBr = CreateSolidBrush(Colors(I))
            FillRect DrawDC, rct, hBr
            DeleteObject hBr
            OffsetRect rct, 0, (DestHeight \ PixelStep)
            If I = PixelStep - 2 Then rct.Bottom = XBorder + (DestHeight - 1)
            rct.Bottom = XBorder + (DestHeight - 1)
        Next
    End If
    rct.Top = XBorder
    
    If AddBitmap = True Then
    BitBlt DrawDC, 0, 0, BitmapW, BitmapH, BitmapDC, 0, 0, SRCCOPY
    End If
        
    If GradIcon <> 0 Then
        rct.Left = XBorder + SMSize + 2
        DrawIconEx DrawDC, XBorder + 1, XBorder + 1, GradIcon, SMSize - 2, SMSize - 2, ByVal 0&, ByVal 0&, 2
    Else
        rct.Left = XBorder
    End If
    tmpGradFont = CreateFontIndirect(CaptionFont)
    OldFont = SelectObject(DrawDC, tmpGradFont)
    
    SetBkMode DrawDC, TRANSPARENT
    
    If GradForceColors Then
        If Color1 = GradForcedFirst Then
            SetTextColor DrawDC, GradForcedText
        Else
            SetTextColor DrawDC, GradForcedTextA
        End If
    Else
        If Color2 = GetSysColor(COLOR_ACTIVECAPTION) Then
            SetTextColor DrawDC, GetSysColor(COLOR_CAPTIONTEXT)
        Else
            SetTextColor DrawDC, GetSysColor(COLOR_INACTIVECAPTIONTEXT)
        End If
    End If
    
    rct.Left = rct.Left + 2
    rct.Right = rct.Right - 10
    DrawText DrawDC, fText, Len(fText) - 1, rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
    SelectObject DrawDC, OldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0
    
    Dim frct As RECT
    If (fStyle And WS_SYSMENU) = WS_SYSMENU Then
        Dim CurMaxPic As Long
        If IsZoomed(GradhWnd) Then
            CurMaxPic = DFCS_CAPTIONRESTORE
        Else
            CurMaxPic = DFCS_CAPTIONMAX
        End If
           
         frct.Right = DestWidth - 2
         frct.Left = frct.Right - SMSize + 2
         frct.Top = XBorder + 2
         frct.Bottom = frct.Top + (DestHeight - 5)
         
        DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONCLOSE
       
        OffsetRect frct, -(SMSize), 0
        
        If (fStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX And (fStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
            DrawFrameControl DrawDC, frct, DFC_CAPTION, CurMaxPic
            OffsetRect frct, -(SMSize) + 2, 0
            DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONMIN
        ElseIf (fStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Then
            DrawFrameControl DrawDC, frct, DFC_CAPTION, CurMaxPic
            OffsetRect frct, -(SMSize) + 2, 0
            DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONMIN Or DFCS_INACTIVE
        ElseIf (fStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
            DrawFrameControl DrawDC, frct, DFC_CAPTION, CurMaxPic Or DFCS_INACTIVE
            OffsetRect frct, -(SMSize) + 2, 0
            DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONMIN
        End If
    End If
    
    rct.Left = XBorder
    rct.Right = rct.Right + 12
    
    If tmpDC <> 0 Then
        BitBlt tmpDC, rct.Left, rct.Top, rct.Right - rct.Left - 10, rct.Bottom - rct.Top, DrawDC, rct.Left, rct.Top, SRCCOPY
        ExcludeClipRect tmpDC, XBorder, XBorder, DestWidth, XBorder + (DestHeight - 1)
    End If
End Function

Public Function GradientCallback(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim OldGradProc As Long
    Dim OldBMP As Long, NewBMP As Long
    Dim rcWnd As RECT
    Dim tmpFrm As Form
    Dim tmpCol1 As Long, tmpCol2 As Long
    Static GettingIcon As Boolean
    
    GradhWnd = hwnd
    OldGradProc = GetProp(GradhWnd, "OldMeProc")
    
    If Not GettingIcon Then
        GettingIcon = True
        GradIcon = SendMessage(hwnd, WM_GETICON, 0, ByVal 0&)
        GettingIcon = False
    End If
    
    Select Case wMsg
    Case WM_NCACTIVATE, WM_MDIACTIVATE, WM_KILLFOCUS, WM_MOUSEACTIVATE
        GetWindowRect GradhWnd, rcWnd
        tmpDC = GetWindowDC(GradhWnd)
        DrawDC = CreateCompatibleDC(tmpDC)
        NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.Left, 50)
        OldBMP = SelectObject(DrawDC, NewBMP)
        
        hRgn = CreateRectRgn(rcWnd.Left, rcWnd.Top, rcWnd.Right, rcWnd.Bottom)
        SelectClipRgn tmpDC, hRgn
        OffsetClipRgn tmpDC, -rcWnd.Left, -rcWnd.Top
            
        If wMsg = WM_KILLFOCUS And GetParent(GradhWnd) <> 0 Then
            BarGetColors False, tmpCol1, tmpCol2
        ElseIf wMsg = WM_NCACTIVATE And wParam And _
        (GetParent(GradhWnd) = 0) Then
            BarGetColors True, tmpCol1, tmpCol2
        ElseIf wMsg = WM_NCACTIVATE And wParam = 0 And _
        (GetParent(GradhWnd) = 0) Then
            BarGetColors False, tmpCol1, tmpCol2
        ElseIf wParam = GradhWnd And GetParent(GradhWnd) <> 0 Then
            BarGetColors False, tmpCol1, tmpCol2
        ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, _
        0, 0) = GradhWnd Then
            BarGetColors True, tmpCol1, tmpCol2
        ElseIf GetActiveWindow() = GradhWnd Then
            BarGetColors True, tmpCol1, tmpCol2
        Else
            BarGetColors False, tmpCol1, tmpCol2
        End If
        
        DrawGradient tmpCol1, tmpCol2
        SelectObject DrawDC, OldBMP
        DeleteObject NewBMP
        DeleteDC DrawDC
        OffsetClipRgn tmpDC, rcWnd.Left, rcWnd.Top
        GetClipRgn tmpDC, hRgn
        
        If wMsg = WM_MOUSEACTIVATE Then
            GradientCallback = MA_ACTIVATE
        Else
            GradientCallback = 1
        End If
        
        ReleaseDC GradhWnd, tmpDC
        DeleteObject hRgn
        tmpDC = 0
        Exit Function
    
    Case WM_SETTEXT, WM_NCPAINT, WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_SYSCOMMAND, WM_INITMENUPOPUP
        GetWindowRect GradhWnd, rcWnd
        tmpDC = GetWindowDC(GradhWnd)
        DrawDC = CreateCompatibleDC(tmpDC)
        NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.Left, 50)
        OldBMP = SelectObject(DrawDC, NewBMP)
         
        hRgn = CreateRectRgn(rcWnd.Left, rcWnd.Top, rcWnd.Right, rcWnd.Bottom)
        SelectClipRgn tmpDC, hRgn
        OffsetClipRgn tmpDC, -rcWnd.Left, -rcWnd.Top
        
        If (GetActiveWindow() = GradhWnd) Then
            BarGetColors True, tmpCol1, tmpCol2
        ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, 0, 0) = GradhWnd Then
            BarGetColors True, tmpCol1, tmpCol2
        Else
            BarGetColors False, tmpCol1, tmpCol2
        End If
        
        DrawGradient tmpCol1, tmpCol2
        SelectObject DrawDC, OldBMP
        DeleteObject NewBMP
        DeleteDC DrawDC
        
        OffsetClipRgn tmpDC, rcWnd.Left, rcWnd.Top
        GetClipRgn tmpDC, hRgn
        GradientCallback = CallWindowProc(OldGradProc, hwnd, WM_NCPAINT, hRgn, lParam)
        ReleaseDC GradhWnd, tmpDC
        DeleteObject hRgn
        tmpDC = 0
        
        If wMsg = (WM_NCLBUTTONDOWN And wParam <> HTSYSMENU And wParam <> HTCAPTION) Or wMsg = (WM_SYSCOMMAND And Not (wParam = SC_MOUSEMENU)) Then
            GetCaptionRect GradhWnd, rcWnd
            ExcludeClipRect tmpDC, rcWnd.Left, rcWnd.Top, rcWnd.Right, rcWnd.Bottom
       
        ElseIf wMsg = WM_NCLBUTTONDOWN And wParam = HTCAPTION Then
         
        Else
           
           If wMsg = (WM_NCRBUTTONDOWN) Then
            WithForm.PopupMenu MenuFrm.MnuPop, , 0, 0
           End If
            
           Exit Function
        End If
    
    Case WM_SIZE
        If hwnd = GradhWnd Then
            SendMessage GradhWnd, WM_NCPAINT, 0, 0
        End If
    
    Case WM_SETCURSOR
       
        Select Case LoWord(lParam)
        Case HTTOP, HTBOTTOM
            SetCursor LoadCursor(ByVal 0&, IDC_SIZENS)
        Case HTLEFT, HTRIGHT
            SetCursor LoadCursor(ByVal 0&, IDC_SIZEWE)
        Case HTTOPLEFT, HTBOTTOMRIGHT
            SetCursor LoadCursor(ByVal 0&, IDC_SIZENWSE)
        Case HTTOPRIGHT, HTBOTTOMLEFT
            SetCursor LoadCursor(ByVal 0&, IDC_SIZENESW)
        Case Else
        
            GoTo JustCallBack
        End Select
        GradientCallback = 1
        Exit Function
    
    End Select
    
JustCallBack:
    GradientCallback = CallWindowProc(OldGradProc, hwnd, wMsg, wParam, lParam)

End Function

Public Sub GradientReleaseForm(frm As Form)
    Dim tmpProc As Long
    tmpProc = GetProp(frm.hwnd, "OldMeProc")
    RemoveProp frm.hwnd, "OldMeProc"

    If tmpProc = 0 Then Exit Sub

    SetWindowLong frm.hwnd, GWL_WNDPROC, tmpProc
End Sub

Private Function LoWord(LongIn As Long) As Integer
   If (LongIn And &HFFFF&) > &H7FFF Then
      LoWord = (LongIn And &HFFFF&) - &H10000
   Else
      LoWord = LongIn And &HFFFF&
   End If
End Function

Public Function RedrawBar(TheForm As Form)
    If TheForm.BorderStyle = 0 Then Exit Function
    If TheForm.BorderStyle = 4 Then Exit Function
    If TheForm.BorderStyle = 5 Then Exit Function
        
    SendMessage GradhWnd, WM_NCPAINT, 0, 0
End Function

Public Function TilePicture(PicInputDC As Long, PicOutputDC As Long, PicSrcWidth As Long, PicSrcHeight As Long, PicTarWidth As Long, PicTarHeight As Long)
    For yDraw = 0 To PicTarHeight Step PicSrcHeight
        For Xdraw = 0 To PicTarWidth Step PicSrcWidth
            BitBlt PicOutputDC, Xdraw, yDraw, PicSrcWidth, PicSrcHeight, PicInputDC, 0, 0, SRCCOPY
        Next Xdraw
    Next yDraw
End Function
