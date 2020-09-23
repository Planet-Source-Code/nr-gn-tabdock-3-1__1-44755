Attribute VB_Name = "modCaption"
Option Explicit
'DefInt A-Z
Private Type DRAWTEXTPARAMS
    cbSize                                         As Long
    iTabLength                                     As Long
    iLeftMargin                                    As Long
    iRightMargin                                   As Long
    uiLengthDrawn                                  As Long
End Type
Public Style                                   As Long
Private Type LOGFONT
    lfHeight                                       As Long
    lfWidth                                        As Long
    lfEscapement                                   As Long
    lfOrientation                                  As Long
    lfWeight                                       As Long
    lfItalic                                       As Byte
    lfUnderline                                    As Byte
    lfStrikeOut                                    As Byte
    lfCharSet                                      As Byte
    lfOutPrecision                                 As Byte
    lfClipPrecision                                As Byte
    lfQuality                                      As Byte
    lfPitchAndFamily                               As Byte
    lfFaceName                                     As String * 32
End Type
Private Type NONCLIENTMETRICS
    cbSize                                         As Long
    iBorderWidth                                   As Long
    iScrollWidth                                   As Long
    iScrollHeight                                  As Long
    iCaptionWidth                                  As Long
    iCaptionHeight                                 As Long
    lfCaptionFont                                  As LOGFONT
    iSMCaptionWidth                                As Long
    iSMCaptionHeight                               As Long
    lfSMCaptionFont                                As LOGFONT
    iMenuWidth                                     As Long
    iMenuHeight                                    As Long
    lfMenuFont                                     As LOGFONT
    lfStatusFont                                   As LOGFONT
    lfMessageFont                                  As LOGFONT
End Type
Private Const DT_SINGLELINE                    As Long = &H20
Private Const DT_VCENTER                       As Long = &H4
Private Const DT_END_ELLIPSIS                  As Long = &H8000&
Private Const DT_CENTER                        As Long = &H1
Private Const DT_BOTTOM                        As Long = &H8
Private Const DT_RIGHT                         As Long = &H2

Private Const TRANSPARENT                      As Integer = 1
Private Const OPAQUE                           As Integer = 2
Private Const SPI_GETNONCLIENTMETRICS          As Integer = 41
Private Const SM_CYSMCAPTION                   As Integer = 51
Private captionFont                            As LOGFONT
Public m_bCloseOver                            As Boolean
Public m_bCloseDown                            As Boolean
Public m_bChevronOver                          As Boolean
Public m_bChevronDown                          As Boolean
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, _
                                                         lpRect As RECT, _
                                                         ByVal un1 As Long, _
                                                         ByVal un2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                   ByVal lpStr As String, _
                                                                   ByVal nCount As Long, _
                                                                   lpRect As RECT, _
                                                                   ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, _
                                                                       ByVal lpsz As String, _
                                                                       ByVal n As Long, _
                                                                       lpRect As RECT, _
                                                                       ByVal un As Long, _
                                                                       lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                  ByVal crColor As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                           ByVal uParam As Long, _
                                                                                           lpvParam As Any, _
                                                                                           ByVal fuWinIni As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Sub drawChevronButton(ByVal lHDC As Long, _
                              ByVal hBrTitle As Long, _
                              m_tChevronR As RECT, _
                              Optional IsCollapsed As Boolean = False, _
                              Optional dPanel As TDockForm)

    Dim lX               As Long
    Dim lY               As Long
    Dim hbr              As Long
    Dim hPen             As Long
    Dim hPenOld          As Long
    Dim tJunk            As POINTAPI
    Dim m_bOfficeXpStyle As Boolean
    m_bOfficeXpStyle = True
    
    If (m_bOfficeXpStyle) Then
        If (m_bChevronOver) Then
            If (m_bChevronDown) Then
                hbr = CreateSolidBrush(VSNetControlColor)
            Else '(M_BCHEVRONDOWN) = FALSE/0
                hbr = CreateSolidBrush(VSNetSelectionColor)
            End If
            FillRect lHDC, m_tChevronR, hbr
            DeleteObject hbr
            If (m_bChevronDown) Then
                hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
                hPenOld = SelectObject(lHDC, hPen)
                MoveToEx lHDC, m_tChevronR.Left, m_tChevronR.Bottom - 1, tJunk
                LineTo lHDC, m_tChevronR.Left, m_tChevronR.Top
                LineTo lHDC, m_tChevronR.Right - 1, m_tChevronR.Top
                LineTo lHDC, m_tChevronR.Right - 1, m_tChevronR.Bottom
                SelectObject lHDC, hPenOld
                DeleteObject hPen
            Else '(M_BCHEVRONDOWN) = FALSE/0
                hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbHighlight And &H1F&))
                hPenOld = SelectObject(lHDC, hPen)
                MoveToEx lHDC, m_tChevronR.Left, m_tChevronR.Bottom - 1, tJunk
                LineTo lHDC, m_tChevronR.Left, m_tChevronR.Top
                LineTo lHDC, m_tChevronR.Right - 1, m_tChevronR.Top
                LineTo lHDC, m_tChevronR.Right - 1, m_tChevronR.Bottom - 1
                LineTo lHDC, m_tChevronR.Left, m_tChevronR.Bottom - 1
                SelectObject lHDC, hPenOld
                DeleteObject hPen
            End If
        Else '(M_BCHEVRONOVER) = FALSE/0
            FillRect lHDC, m_tChevronR, hBrTitle
        End If
    Else '(M_BOFFICEXPSTYLE) = FALSE/0
        FillRect lHDC, m_tChevronR, hBrTitle
        If (m_bChevronOver) Then
            If (m_bChevronDown) Then
                DrawEdge lHDC, m_tChevronR, BF_FLAT, BF_RECT
            Else '(M_BCHEVRONDOWN) = FALSE/0
                DrawEdge lHDC, m_tChevronR, BF_FLAT, BF_RECT
            End If
        End If
    End If
    
    If IsCollapsed Then
        If Not dPanel.Panel.Expanded Then
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbApplicationWorkspace And &H1F&))
        Else
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonFace And &H1F&))
            
        End If
    End If
    
    ' Draw chevron glyph:
    hPenOld = SelectObject(lHDC, hPen)
    lX = m_tChevronR.Left + (m_tChevronR.Right - m_tChevronR.Left - 5) / 2
    lY = m_tChevronR.Top + (m_tChevronR.Bottom - m_tChevronR.Top - 3) / 2
    MoveToEx lHDC, lX, lY, tJunk
    LineTo lHDC, lX + 5, lY
    MoveToEx lHDC, lX + 1, lY + 1, tJunk
    LineTo lHDC, lX + 4, lY + 1
    MoveToEx lHDC, lX + 2, lY, tJunk
    LineTo lHDC, lX + 2, lY + 3
    SelectObject lHDC, hPenOld
    DeleteObject hPen
    ' m_bChevronDown = False
    ' m_bChevronOver = False

End Sub

Public Sub drawCloseButton(ByVal lHDC As Long, _
                            ByVal hBrTitle As Long, _
                            m_tCloseR As RECT, _
                            Optional IsCollapsed As Boolean = False, _
                            Optional dPanel As TDockForm)

    Dim lX               As Long
    Dim lY               As Long
    Dim hbr              As Long
    Dim hPen             As Long
    Dim hPenOld          As Long
    Dim tJunk            As POINTAPI
    Dim m_bOfficeXpStyle As Boolean

    m_bOfficeXpStyle = True
    
    If (m_bOfficeXpStyle) Then
        If (m_bCloseOver) Then
            If (m_bCloseDown) Then
                hbr = CreateSolidBrush(VSNetControlColor)
            Else '(m_bCloseDown) = FALSE/0
                hbr = CreateSolidBrush(VSNetSelectionColor)
            End If
            FillRect lHDC, m_tCloseR, hbr
            DeleteObject hbr
            If (m_bCloseDown) Then
                hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
                hPenOld = SelectObject(lHDC, hPen)
                MoveToEx lHDC, m_tCloseR.Left, m_tCloseR.Bottom - 1, tJunk
                LineTo lHDC, m_tCloseR.Left, m_tCloseR.Top
                LineTo lHDC, m_tCloseR.Right - 1, m_tCloseR.Top
                LineTo lHDC, m_tCloseR.Right - 1, m_tCloseR.Bottom
                SelectObject lHDC, hPenOld
                DeleteObject hPen
            Else '(m_bCloseDown) = FALSE/0
                hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbHighlight And &H1F&))
                hPenOld = SelectObject(lHDC, hPen)
                MoveToEx lHDC, m_tCloseR.Left, m_tCloseR.Bottom - 1, tJunk
                LineTo lHDC, m_tCloseR.Left, m_tCloseR.Top
                LineTo lHDC, m_tCloseR.Right - 1, m_tCloseR.Top
                LineTo lHDC, m_tCloseR.Right - 1, m_tCloseR.Bottom - 1
                LineTo lHDC, m_tCloseR.Left, m_tCloseR.Bottom - 1
                SelectObject lHDC, hPenOld
                DeleteObject hPen
            End If
        Else '(m_bcloseover) = FALSE/0
            FillRect lHDC, m_tCloseR, hBrTitle
        End If
    Else '(M_BOFFICEXPSTYLE) = FALSE/0
        FillRect lHDC, m_tCloseR, hBrTitle
        If (m_bCloseOver) Then
            If (m_bCloseDown) Then
                DrawEdge lHDC, m_tCloseR, BF_FLAT, BF_RECT
            Else '(m_bCloseDown) = FALSE/0
                DrawEdge lHDC, m_tCloseR, BF_FLAT, BF_RECT
            End If
        End If
    End If

    If IsCollapsed Then
        If Not dPanel.Panel.Expanded Then
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbApplicationWorkspace And &H1F&))
        Else
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonFace And &H1F&))
            
        End If
    End If
    
    hPenOld = SelectObject(lHDC, hPen)
    lX = m_tCloseR.Left + (m_tCloseR.Right - m_tCloseR.Left - 8) / 1.5
    lY = m_tCloseR.Top + (m_tCloseR.Bottom - m_tCloseR.Top - 7) / 2
    MoveToEx lHDC, lX, lY, tJunk
    LineTo lHDC, lX + 6, lY + 6
    MoveToEx lHDC, lX + 1, lY, tJunk
    LineTo lHDC, lX + 7, lY + 6
    MoveToEx lHDC, lX + 5, lY, tJunk
    LineTo lHDC, lX - 1, lY + 6
    MoveToEx lHDC, lX + 6, lY, tJunk
    LineTo lHDC, lX, lY + 6
    SelectObject lHDC, hPenOld
    DeleteObject hPen
    '  m_bCloseDown = False
    '  m_bCloseOver = False

End Sub

Public Sub drawGradient(captionRect As RECT, _
                         hdc As Long, _
                         captionText As String, _
                         bActive As Boolean, _
                         gradient As Boolean, _
                         Optional captionOrientation As Integer, _
                         Optional captionForm As Form)

    Dim hbr               As Long
    Dim bar               As Long
    Dim width             As Long
    Dim pixelStep         As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont       As Long
    Dim oldFont           As Long
    Dim hDCTemp           As Long

    hDCTemp = hdc
    ''debug.print captionText, captionOrientation, hDC, hDCTemp
    storedCaptionRect = captionRect
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
        Else 'NOT CAPTIONORIENTATION...
        width = captionRect.Bottom - captionRect.Top
    End If
    If gradient Then
        pixelStep = width / 4
        Else 'GRADIENT = FALSE/0
        pixelStep = 2
    End If
    ReDim Colors(pixelStep) As Long
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        If gradient Then
            gradateColors Colors(), GradClr1, GradClr2
            Else 'GRADIENT = FALSE/0
            gradateColors Colors(), TranslateColor(vbActiveTitleBar), TranslateColor(vbActiveTitleBar)
        End If
        Else 'BACTIVE = FALSE/0
        If gradient Then
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbButtonFace)
            Else 'GRADIENT = FALSE/0
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbInactiveTitleBar)
        End If
    End If
    For bar = 1 To pixelStep - 1
        hbr = CreateSolidBrush(Colors(bar))
        FillRect hDCTemp, captionRect, hbr
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            captionRect.Left = captionRect.Left + 4
            Else 'NOT CAPTIONORIENTATION...
            captionRect.Bottom = captionRect.Bottom - 4
        End If
        DeleteObject hbr
    Next bar
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    'get caption font information
    getCapsFont
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    If captionText = "Form6" Then
        '  Beep
    End If
    If tmpGradFont = 0 Then
        'tmpGradFont = CreateFontIndirect(captionFont)
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
            'hDCTemp = captionForm.hDC
            ''debug.print "gradient font hdc set"
        End If
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    SetBkMode hDCTemp, TRANSPARENT
    If (bActive) Then
        SetTextColor hDCTemp, TranslateColor(vbActiveTitleBarText)
        Else '(BACTIVE) = FALSE/0
        SetTextColor hDCTemp, TranslateColor(vbInactiveTitleBarText)
    End If
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        'captionForm.CurrentX = 50
        'captionForm.CurrentY = captionForm.ScaleHeight - 100
        'captionForm.Print captionText
        ''debug.print "caption text drawn", captionForm.CurrentX
        storedCaptionRect.Right = storedCaptionRect.Bottom - (getCaptionHeight * 2)
        storedCaptionRect.Bottom = 8 + (captionForm.Height / Screen.TwipsPerPixelY)
        ''debug.print "pixel height = "; captionForm.height / Screen.TwipsPerPixelY
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else 'NOT CAPTIONORIENTATION...
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - (getCaptionHeight * 2)
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS 'Or DT_HCENTER
    End If
    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub

Public Sub drawGradientx(captionRect As RECT, _
                          hdc As Long, _
                          captionText As String, _
                          bActive As Boolean, _
                          gradient As Boolean, _
                          Optional captionOrientation As Integer, _
                          Optional captionForm As Form, _
                          Optional borderSurround As Boolean)

    Dim hbr               As Long
    Dim bar               As Long
    Dim width             As Long
    Dim pixelStep         As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont       As Long
    Dim oldFont           As Long
    Dim hDCTemp           As Long

    hDCTemp = hdc
    ''debug.print captionText, captionOrientation, hDC, hDCTemp
    storedCaptionRect = captionRect
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
        Else 'NOT CAPTIONORIENTATION...
        width = captionRect.Bottom - captionRect.Top
    End If

    pixelStep = width / 4
    ReDim Colors(pixelStep) As Long
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        If gradient Then
            gradateColors Colors(), GradClr1, GradClr2
            Else 'GRADIENT = FALSE/0
            gradateColors Colors(), TranslateColor(vbActiveTitleBar), TranslateColor(vbActiveTitleBar)
        End If
        Else 'BACTIVE = FALSE/0
        If gradient Then
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbButtonFace)
            Else 'GRADIENT = FALSE/0
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbInactiveTitleBar)
        End If
    End If
    
    
    If borderSurround Then
        hbr = CreateSolidBrush(TranslateColor(vb3DShadow))
        FillRect hDCTemp, captionRect, hbr
        With captionRect
            .Top = .Top + 1
            .Bottom = .Bottom - 1
            .Left = .Left + 1
            .Right = .Right - 1
        End With
        DeleteObject hbr
        hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
        FillRect hDCTemp, captionRect, hbr
        DeleteObject hbr
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            DrawIconEx hDCTemp, captionRect.Left + 1, captionRect.Top + 1, captionForm.Icon, 16, 16, 0&, 0&, DI_NORMAL
        Else
            DrawIconEx hDCTemp, captionRect.Left + 1, captionRect.Top, captionForm.Icon, 16, 16, 0&, 0&, DI_NORMAL
        End If
    Else
        For bar = 1 To pixelStep - 1
            hbr = CreateSolidBrush(Colors(bar))
            FillRect hDCTemp, captionRect, hbr
            If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
                captionRect.Left = captionRect.Left + 4
                Else 'NOT CAPTIONORIENTATION...
                captionRect.Bottom = captionRect.Bottom - 4
            End If
            DeleteObject hbr
        Next bar
    End If
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    'get caption font information
    getCapsFont
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    If tmpGradFont = 0 Then
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
        End If
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    SetBkMode hDCTemp, TRANSPARENT
    
    If borderSurround Then
        SetTextColor hDCTemp, TranslateColor(vbBlack)
    Else
        If (bActive) Then
            SetTextColor hDCTemp, TranslateColor(vbActiveTitleBarText)
            Else '(BACTIVE) = FALSE/0
            SetTextColor hDCTemp, TranslateColor(vbInactiveTitleBarText)
        End If
    End If
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        storedCaptionRect.Right = storedCaptionRect.Bottom + 40
        storedCaptionRect.Bottom = storedCaptionRect.Bottom + 10
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else 'NOT CAPTIONORIENTATION...
        storedCaptionRect.Bottom = storedCaptionRect.Bottom - 3
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - 5
        If borderSurround Then
            storedCaptionRect.Left = storedCaptionRect.Left + getCaptionHeight
        End If

        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or IIf(borderSurround, DT_RIGHT, 0) Or DT_BOTTOM
    End If
    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub

Public Sub drawGripper(captionRect As RECT, _
                        hdc As Long, _
                        gripStyle As Long, _
                        gripSides As Long, _
                        oneBar As Boolean, _
                        captionHeight As Long, _
                        Optional captionOrientation As Integer, _
                        Optional maximiseButton As Boolean, _
                        Optional CloseButton As Boolean)

    Dim numOfButtons As Integer

    If maximiseButton And CloseButton Then
        numOfButtons = 2
        ElseIf maximiseButton And Not CloseButton Then 'NOT MAXIMISEBUTTON...
        numOfButtons = 1
        ElseIf CloseButton And Not maximiseButton Then 'NOT MAXIMISEBUTTON...
        numOfButtons = 1
        ElseIf Not maximiseButton And Not CloseButton Then 'NOT CLOSEBUTTON...
        numOfButtons = 0
        Else 'NOT NOT...
        numOfButtons = 0
    End If
    If oneBar Then
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Top + ((captionHeight - 11) / 2) - 1 ' 4
                .Left = .Left + 1
                .Right = .Right - (getCaptionHeight * numOfButtons) - 3
                .Bottom = .Top + 4
            End With 'CAPTIONRECT
            Else 'NOT CAPTIONORIENTATION...
            With captionRect
                .Top = .Top + (getCaptionHeight * numOfButtons) + 4
                .Left = .Left + ((captionHeight - 14) / 2) + 2
                .Right = .Left + 4
                .Bottom = .Bottom - 2
            End With 'CAPTIONRECT
        End If
        DrawEdge hdc, captionRect, gripStyle, gripSides
        Else 'ONEBAR = FALSE/0
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Top + ((captionHeight - 16) / 2) + 0.5
                .Left = .Left + 1
                .Right = .Right - (getCaptionHeight * numOfButtons) - 2 '+ 5
                .Bottom = .Top + 4
            End With 'CAPTIONRECT
            Else 'NOT CAPTIONORIENTATION...
            With captionRect
                .Top = .Top + (getCaptionHeight * numOfButtons / 1.1) + 2 ' 3
                .Left = .Left + ((captionHeight - 20) / 2) + 2
                .Right = .Left + 4
                .Bottom = .Bottom - 2
            End With 'CAPTIONRECT
        End If
        DrawEdge hdc, captionRect, gripStyle, gripSides
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Bottom + 1
                .Bottom = .Bottom + 5
            End With 'CAPTIONRECT
            Else 'NOT CAPTIONORIENTATION...
            With captionRect
                .Left = .Right + 1
                .Right = .Left + 4
            End With 'CAPTIONRECT
        End If
        DrawEdge hdc, captionRect, gripStyle, gripSides
    End If

End Sub

Public Sub drawOfficeXP(captionRect As RECT, _
                         hdc As Long, _
                         captionText As String, _
                         bActive As Boolean, _
                         gradient As Boolean, _
                         Optional captionOrientation As Integer, _
                         Optional captionForm As Form)

    Dim hbr               As Long
    Dim width             As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont       As Long
    Dim oldFont           As Long
    Dim hDCTemp           As Long
    Dim colorOutline      As Long
    Dim colorInline       As Long

    hDCTemp = hdc
    ''debug.print captionText, captionOrientation, hDC, hDCTemp
    storedCaptionRect = captionRect
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
        Else 'NOT CAPTIONORIENTATION...
        width = captionRect.Bottom - captionRect.Top
    End If
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        colorOutline = TranslateColor(vbActiveTitleBar)
        colorInline = TranslateColor(vbActiveTitleBar)
        Else 'BACTIVE = FALSE/0
        colorOutline = TranslateColor(vbInactiveTitleBar)
        colorInline = TranslateColor(vbButtonFace)
    End If
    hbr = CreateSolidBrush(colorOutline)
    FillRect hDCTemp, captionRect, hbr
    DeleteObject hbr
    
    With captionRect
        .Top = .Top + 1
        .Left = .Left + 1
        .Right = .Right - 1
        .Bottom = .Bottom - 1
    End With 'CAPTIONRECT
    hbr = CreateSolidBrush(colorInline)
    FillRect hDCTemp, captionRect, hbr
    DeleteObject hbr
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    'get caption font information
    getCapsFont
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    If captionText = "Form6" Then
        '  Beep
    End If
    If tmpGradFont = 0 Then
        'tmpGradFont = CreateFontIndirect(captionFont)
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
            'hDCTemp = captionForm.hDC
            ''debug.print "gradient font hdc set"
        End If
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    SetBkMode hDCTemp, TRANSPARENT
    If (bActive) Then
        SetTextColor hDCTemp, TranslateColor(vbActiveTitleBarText)
        Else '(BACTIVE) = FALSE/0
        SetTextColor hDCTemp, TranslateColor(vbInactiveTitleBarText)
    End If
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        storedCaptionRect.Right = storedCaptionRect.Bottom - (getCaptionHeight * 2)
        storedCaptionRect.Bottom = 8 + (captionForm.Height / Screen.TwipsPerPixelY)
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else 'NOT CAPTIONORIENTATION...
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - (getCaptionHeight * 2)
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS 'Or DT_HCENTER
    End If

    
    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub

Public Sub drawVSNet(captionRect As RECT, _
                      hdc As Long, _
                      captionText As String, _
                      bActive As Boolean, _
                      gradient As Boolean, _
                      Optional captionOrientation As Integer, _
                      Optional captionForm As Form)

    Dim hbr               As Long
    Dim width             As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont       As Long
    Dim oldFont           As Long
    Dim hDCTemp           As Long
    Dim colorOutline      As Long
    Dim colorInline       As Long

    
    hDCTemp = hdc
    ''debug.print captionText, captionOrientation, hDC, hDCTemp
    storedCaptionRect = captionRect
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
    Else 'NOT CAPTIONORIENTATION...
        width = captionRect.Bottom - captionRect.Top
    End If
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        colorOutline = TranslateColor(vbInactiveTitleBar)
        colorInline = TranslateColor(vbInactiveTitleBar) 'vbButtonFace)
        '     colorOutline = TranslateColor(vbActiveTitleBar)
        '     colorInline = TranslateColor(vbActiveTitleBar)
    Else 'BACTIVE = FALSE/0
        colorOutline = TranslateColor(vbInactiveTitleBar)
        colorInline = TranslateColor(vbInactiveTitleBar) 'vbButtonFace)
    End If
    hbr = CreateSolidBrush(colorOutline)
    FillRect hDCTemp, captionRect, hbr
    DeleteObject hbr
    
    With captionRect
        .Top = .Top + 1
        .Left = .Left + 1
        .Right = .Right - 1
        .Bottom = .Bottom - 1
    End With 'CAPTIONRECT
    hbr = CreateSolidBrush(colorInline)
    FillRect hDCTemp, captionRect, hbr
    DeleteObject hbr
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    'get caption font information
    getCapsFont
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    If captionText = "Form6" Then
        '  Beep
    End If
    If tmpGradFont = 0 Then
        'tmpGradFont = CreateFontIndirect(captionFont)
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
            'hDCTemp = captionForm.hDC
            ''debug.print "gradient font hdc set"
        End If
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    SetBkMode hDCTemp, TRANSPARENT
    If (bActive) Then
        SetTextColor hDCTemp, TranslateColor(vbWhite)
        Else '(BACTIVE) = FALSE/0
        SetTextColor hDCTemp, TranslateColor(&H8000000F)
    End If
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        storedCaptionRect.Right = storedCaptionRect.Bottom - (getCaptionHeight * 2)
        storedCaptionRect.Bottom = 8 + (captionForm.Height / Screen.TwipsPerPixelY)
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else 'NOT CAPTIONORIENTATION...
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - (getCaptionHeight * 2)
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS 'Or DT_HCENTER
    End If

    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub

Private Sub getCapsFont()

    Dim NCM As NONCLIENTMETRICS

    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
    If NCM.iCaptionHeight = 0 Then
        captionFont.lfHeight = 0
        Else 'NOT NCM.ICAPTIONHEIGHT...
        captionFont = NCM.lfSMCaptionFont
        'If captionFont.lfHeight < 10 Then
        ' captionFont.lfHeight = 14
        'End If
    End If

End Sub

Public Function getCaptionButtonHeight() As Long

    Dim NCM As NONCLIENTMETRICS

    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
    If NCM.iCaptionHeight = 0 Then
        'captionFont.lfHeight = 0
        getCaptionButtonHeight = 14
        Else 'NOT NCM.ICAPTIONHEIGHT...
        'captionFont = NCM.lfSMCaptionFont
        getCaptionButtonHeight = NCM.iSMCaptionHeight
    End If

End Function

Public Function getCaptionHeight() As Long

    getCaptionHeight = GetSystemMetrics(SM_CYSMCAPTION)
    'If getCaptionHeight < 20 Then getCaptionHeight = 15

End Function

Public Sub gradateColors(Colors() As Long, _
                          ByVal color1 As Long, _
                          ByVal Color2 As Long)

    Dim i    As Integer
    Dim dblR As Double
    Dim dblG As Double
    Dim dblB As Double
    Dim addR As Double
    Dim addG As Double
    Dim addB As Double
    Dim bckR As Double
    Dim bckG As Double
    Dim bckB As Double

    'Alright, I admit -- this routine was
    'taken from a VBPJ issue a few months back.
    dblR = CDbl(color1 And &HFF)
    dblG = CDbl(color1 And &HFF00&) / 255
    dblB = CDbl(color1 And &HFF0000) / &HFF00&
    bckR = CDbl(Color2 And &HFF&)
    bckG = CDbl(Color2 And &HFF00&) / 255
    bckB = CDbl(Color2 And &HFF0000) / &HFF00&
    addR = (bckR - dblR) / UBound(Colors)
    addG = (bckG - dblG) / UBound(Colors)
    addB = (bckB - dblB) / UBound(Colors)
    For i = 0 To UBound(Colors)
        dblR = dblR + addR
        dblG = dblG + addG
        dblB = dblB + addB
        If dblR > 255 Then
            dblR = 255
        End If
        If dblG > 255 Then
            dblG = 255
        End If
        If dblB > 255 Then
            dblB = 255
        End If
        If dblR < 0 Then
            dblR = 0
        End If
        If dblG < 0 Then
            dblG = 0
        End If
        If dblG < 0 Then
            dblB = 0
        End If
        Colors(i) = RGB(dblR, dblG, dblB)
    Next i

End Sub
