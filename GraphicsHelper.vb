Public Class GraphicsHelper

    ' This class contains information gathered from :
    '   http://www.mentalis.org/apilist/apilist.php
    '   http://www.codeguru.com/cs_graphics/flicker_free.html

#Region " Enumerations "
    Public Enum Bool
        bFalse
        bTrue
    End Enum
    Public Enum RasterOperations
        SRCCOPY = &HCC0020      ', /* dest = source                   */
        SRCPAINT = &HEE0086     ', /* dest = source OR dest           */
        SRCAND = &H8800C6       ', /* dest = source AND dest          */
        SRCINVERT = &H660046    ', /* dest = source XOR dest          */
        SRCERASE = &H440328     ', /* dest = source AND (NOT dest )   */
        NOTSRCCOPY = &H330008   ', /* dest = (NOT source)             */
        NOTSRCERASE = &H1100A6  ', /* dest = (NOT src) AND (NOT dest) */
        MERGECOPY = &HC000CA    ', /* dest = (source AND pattern)     */
        MERGEPAINT = &HBB0226   ', /* dest = (NOT source) OR dest     */
        PATCOPY = &HF00021      ', /* dest = pattern                  */
        PATPAINT = &HFB0A09     ', /* dest = DPSnoo                   */
        PATINVERT = &H5A0049    ', /* dest = pattern XOR dest         */
        DSTINVERT = &H550009    ', /* dest = (NOT dest)               */
        BLACKNESS = &H42        ', /* dest = BLACK                    */
        WHITENESS = &HFF0062    ', /* dest = WHITE                    */
    End Enum
    Public Enum StretchModes
        StretchAndScans = 1
        StretchOrScans = 2
        StretchDeleteScans = 3
        StretchHalftone = 4
        BlackOnWhite = 1
        WhiteOnBlack = 2
        ColorOnColor = 3
        Halftone = 4
    End Enum
#End Region

#Region " GDI32 Functions "

    Declare Function CreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" (ByVal hdc As IntPtr) As IntPtr
    Declare Function DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As IntPtr) As Bool

    Declare Function SelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    Declare Function DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As IntPtr) As Bool

    Declare Function CreateCompatibleBitmap Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal hdc As IntPtr, ByVal nWidth As Long, ByVal nHeight As Long) As IntPtr

    Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As IntPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As IntPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As System.Int32) As Boolean
    Declare Function MaskBlt Lib "gdi32" (ByVal hdcDest As IntPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As IntPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As IntPtr, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As System.Int32) As Long
    Declare Function StretchBlt Lib "gdi32" Alias "StretchBlt" (ByVal hdc As IntPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As IntPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal iStretchMode As Integer) As Integer

    Declare Function FillRect Lib "user32" Alias "FillRect" (ByVal hdc As IntPtr, ByVal lpRect As Rectangle, ByVal hBrush As IntPtr) As Long


#End Region

#Region " VB Helper Functions "
    'Function BuildColoredMask(ByVal hDC As IntPtr, ByVal bmpMask As Bitmap, ByVal col As Color) As IntPtr
    '    Dim sb As New SolidBrush(col)
    '    dim ipsb as IntPtr=sb.
    '    Me.FillRect(hDC, New Rectangle(0, 0, bmpMask.Width, bmpMask.Height), sb)
    '    Dim maskDC As IntPtr = Me.CreateCompatibleDC(hDC)
    '    Dim hMask As IntPtr = bmpMask.GetHbitmap
    '    Me.MaskBlt(maskDC, 0, 0, bmpMask.Width, bmpMask.Height, hDC, 0, 0, hMask, 0, 0, Me.RasterOperations.SRCPAINT)
    '    Me.DeleteObject(hMask)
    '    Return maskDC
    'End Function
    Function BuildColoredMask(ByVal bmp As Bitmap, ByVal col As Color) As Bitmap


    End Function

    Public Sub SetColorKeyExample(ByVal e As PaintEventArgs)
        ' Open an Image file, and draw it to the screen.
        Dim myImage As Image = Image.FromFile("Circle.bmp")
        e.Graphics.DrawImage(myImage, 20, 20)
        ' Create an ImageAttributes object and set the color key.
        Dim lowerColor As Color = Color.FromArgb(245, 0, 0)
        Dim upperColor As Color = Color.FromArgb(255, 0, 0)
        Dim imageAttr As New System.Drawing.Imaging.ImageAttributes()
        imageAttr.SetColorKey(lowerColor, upperColor, Drawing.Imaging.ColorAdjustType.Default)
        ' Draw the image with the color key set.
        Dim rect As New Rectangle(150, 20, 100, 100)
        e.Graphics.DrawImage(myImage, rect, 0, 0, 100, 100, _
        GraphicsUnit.Pixel, imageAttr) ' Image
    End Sub

#End Region

End Class
