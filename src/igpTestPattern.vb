Option Explicit On 
Option Strict On

Namespace Zlosk.Patterns

    Public Class GoodPattern ' GoodPattern is the original working Pattern

        Const SRCCOPY As Integer = &HCC0020 ' Hex
        Const SRCINVERT As Integer = &H660046
        Const SRCAND As Integer = &H8800C6
        Const SRCPAINT As Integer = &HEE0086

        Private Declare Auto Function BitBlt Lib "GDI32.DLL" (ByVal hdcDest As IntPtr, ByVal nXDest As Integer, _
ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, _
ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Int32) As Boolean

        Public MainTile As Tile, SubTiles() As SubTile
        Private bmpBase, bmpRendered, bmpMem As Bitmap
        Private graBase, graRendered As Graphics
        Private mcolBaseColor, mcolBackColor, mcolForeColor As Color
        Public mctrlClient As Control

        Event Invalidated(ByVal sender As Object, ByVal e As InvalidateEventArgs)

        Property BackColor() As Color
            Get
                Return mcolBackColor
            End Get
            Set(ByVal Value As Color)
                If mcolBackColor.ToArgb <> Value.ToArgb Then
                    mcolBackColor = Value
                    UpdateBackColor()
                End If
            End Set
        End Property
        Property ForeColor() As Color
            Get
                Return mcolForeColor
            End Get
            Set(ByVal Value As Color)
                If mcolForeColor.ToArgb <> Value.ToArgb Then
                    mcolForeColor = Value
                    UpdateForeColor()
                End If
            End Set
        End Property
        Property BaseColor() As Color
            Get
                Return mcolBaseColor
            End Get
            Set(ByVal Value As Color)
                If mcolBaseColor.ToArgb <> Value.ToArgb Then
                    mcolBaseColor = Value
                End If
            End Set
        End Property
        Property BaseImage() As Image
            Get
                Return bmpBase
            End Get
            Set(ByVal Value As Image)
                bmpBase = Nothing
                If Value Is Nothing Then
                    bmpRendered = Nothing
                Else
                    bmpBase = New Bitmap(Value)
                End If
            End Set
        End Property
        ReadOnly Property RenderedImage() As Bitmap
            Get
                Return bmpRendered
            End Get
        End Property

        Structure PatternCoordinate
            Public x, y, Index As Integer
            Shadows ReadOnly Property ToString() As String
                Get
                    Dim str As String
                    str = x.ToString + ", " + y.ToString + ", " + Index.ToString
                    Return str
                End Get
            End Property
        End Structure
        Structure Tile
            Dim Image, ImageOutline, ImageBackground As Bitmap
            Dim Width, Height As Integer
            Dim XRepeat, YRepeat As Integer
            Dim NextRowXOffset, NextRowYOffset As Integer
            Dim NumSubTiles As Integer
            Dim XReadWrite, YReadWrite As Integer
            ReadOnly Property Size() As Size
                Get
                    Return New Size(Width, Height)
                End Get
            End Property
            Shadows ReadOnly Property ToString() As String
                Get
                    Dim str As String = ""
                    str += "Tile Properties:"
                    str += " W" + Me.Width.ToString
                    str += " H" + Me.Height.ToString
                    str += " XR" + Me.XRepeat.ToString
                    str += " YR" + Me.YRepeat.ToString
                    str += " NRXO" + Me.NextRowXOffset.ToString
                    str += " NRYO" + Me.NextRowYOffset.ToString
                    str += " NST" + Me.NumSubTiles.ToString
                    str += " XRW" + Me.XReadWrite.ToString
                    str += " YRW" + Me.YReadWrite.ToString
                    Return str
                End Get
            End Property
        End Structure
        Structure SubTile
            Dim Image As Bitmap
            Dim PatternXOffset, PatternYOffset As Integer
            Dim ReadWriteXOffset, ReadWriteYOffset As Integer
            ReadOnly Property Width() As Integer
                Get
                    Return Image.Width
                End Get
            End Property
            ReadOnly Property Height() As Integer
                Get
                    Return Image.Height
                End Get
            End Property
            ReadOnly Property Size() As Size
                Get
                    Return New Size(Width, Height)
                End Get
            End Property
            ReadOnly Property PatternOffset() As Point
                Get
                    Return New Point(PatternXOffset, PatternYOffset)
                End Get
            End Property
        End Structure

        Public Function GetPixel(ByVal pc As PatternCoordinate) As Color
            Dim x, y As Integer
            Try
                x = pc.x * Me.MainTile.XReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteXOffset
                y = pc.y * Me.MainTile.YReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteYOffset
                Return Me.bmpBase.GetPixel(x, y)
            Catch
                Return mcolBaseColor
            End Try
        End Function
        Public Sub SetPixel(ByVal pc As PatternCoordinate, ByVal col As Color)
            Try
                Dim x, y As Integer
                x = pc.x * Me.MainTile.XReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteXOffset
                y = pc.y * Me.MainTile.YReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteYOffset
                Me.bmpBase.SetPixel(x, y, col)
                DrawTile(pc, col)
                Dim e As New InvalidateEventArgs(New Rectangle(Me.TileLocationToImagePoint(pc.x, pc.y), Me.MainTile.Size))
                RaiseEvent Invalidated(Me, e)
            Catch
                ' Bad SetPixel - Caused when bmpBase.Height / MainTile.YReadWrite
                ' is not an integer
            End Try
            '            mctrlClient.Invalidate(New Rectangle(Me.TileLocationToImagePoint(pc.x, pc.y), Me.MainTile.Size))
        End Sub

        Public Sub LoadIni(ByVal FileName As String)
            Dim ini As New Org.Mentalis.Files.IniReader(FileName)
            Dim iniPath As String = System.IO.Path.GetDirectoryName(FileName) & "\"
            Dim strTmp As String

            Me.MainTile.Width = ini.ReadInteger("General", "Width", 0)
            Me.MainTile.Height = ini.ReadInteger("General", "Height", 0)
            Me.MainTile.XRepeat = ini.ReadInteger("General", "XRepeat", 0)
            Me.MainTile.YRepeat = ini.ReadInteger("General", "YRepeat", 0)
            Me.MainTile.NextRowXOffset = ini.ReadInteger("General", "NextRowXOffset", 0)
            Me.MainTile.NextRowYOffset = ini.ReadInteger("General", "NextRowYOffset", 0)
            Me.MainTile.NumSubTiles = ini.ReadInteger("General", "NumSubTiles", 0)
            Me.MainTile.XReadWrite = ini.ReadInteger("General", "XReadWrite", 0)
            Me.MainTile.YReadWrite = ini.ReadInteger("General", "YReadWrite", 0)

            Me.MainTile.Image = Nothing
            strTmp = ini.ReadString("General", "Filename", "")
            If strTmp.Length > 0 Then Me.MainTile.Image = New Bitmap(iniPath & strTmp)

            Me.MainTile.ImageOutline = Nothing
            strTmp = ini.ReadString("General", "OutlineFilename", "")
            If strTmp.Length > 0 Then Me.MainTile.ImageOutline = New Bitmap(iniPath & strTmp)

            Me.MainTile.ImageBackground = Nothing
            strTmp = ini.ReadString("General", "BackgroundFilename", "")
            If strTmp.Length > 0 Then Me.MainTile.ImageBackground = New Bitmap(iniPath & strTmp)

            ReDim SubTiles(Me.MainTile.NumSubTiles - 1)
            Dim i As Integer
            Dim strSectionPath As String
            For i = 0 To Me.MainTile.NumSubTiles - 1
                strSectionPath = "Subtile" & (i + 1).ToString("000")
                Me.SubTiles(i).Image = New Bitmap(iniPath & ini.ReadString(strSectionPath, "Filename", ""))
                Me.SubTiles(i).PatternXOffset = ini.ReadInteger(strSectionPath, "PatternXOffset", 0)
                Me.SubTiles(i).PatternYOffset = ini.ReadInteger(strSectionPath, "PatternYOffset", 0)
                Me.SubTiles(i).ReadWriteXOffset = ini.ReadInteger(strSectionPath, "ReadWriteXOffset", 0)
                Me.SubTiles(i).ReadWriteYOffset = ini.ReadInteger(strSectionPath, "ReadWriteYOffset", 0)
            Next
        End Sub

        Public Sub UpdateForeColor()
            If Not Me.MainTile.ImageOutline Is Nothing Then
                DrawBackground(graRendered, Me.MainTile.ImageOutline, Me.ForeColor)
            End If
        End Sub
        Public Sub UpdateBackColor()
            If Not Me.MainTile.ImageBackground Is Nothing Then
                DrawBackground(graRendered, Me.MainTile.ImageBackground, Me.BackColor)
            End If
        End Sub

        Public Sub BuildImage()
            If BaseImage Is Nothing Then Exit Sub
            Static sw As New Org.Mentalis.Utilities.StopWatch
            sw.Reset()

            bmpRendered = Nothing
            bmpRendered = New Bitmap(Me.MainTile.XRepeat * Me.bmpBase.Width \ Me.MainTile.XReadWrite + (Me.MainTile.Width - Me.MainTile.XRepeat), Me.MainTile.YRepeat * Me.bmpBase.Height \ Me.MainTile.YReadWrite + (Me.MainTile.Height - Me.MainTile.YRepeat))
            graRendered = Graphics.FromImage(bmpRendered)
            Dim sb As New SolidBrush(Me.bmpBase.GetPixel(0, 0))
            graRendered.FillRectangle(sb, graRendered.ClipBounds)
            'SaveRender()
            UpdateForeColor()
            'SaveRender()
            UpdateBackColor()
            Dim i, j, idx As Integer
            'SaveRender()
            'i = 1
            For j = 0 To Me.BaseImage.Height - 1 Step Me.MainTile.YReadWrite
                For i = 0 To Me.BaseImage.Width - 1 Step Me.MainTile.XReadWrite
                    For idx = 0 To Me.MainTile.NumSubTiles - 1
                        'DrawSubTile(graRendered, Me.SubTiles(idx).Image, Me.bmpBase.GetPixel(i + Me.SubTiles(idx).ReadWriteXOffset, j + Me.SubTiles(idx).ReadWriteYOffset), i \ Me.SubTiles(idx).ReadWriteXOffset, j \ Me.SubTiles(idx).ReadWriteYOffset, 0, 0)
                        Try
                            DrawSubTile(graRendered, Me.SubTiles(idx), Me.bmpBase.GetPixel(i + Me.SubTiles(idx).ReadWriteXOffset, j + Me.SubTiles(idx).ReadWriteYOffset), i \ Me.MainTile.XReadWrite, j \ Me.MainTile.YReadWrite)
                        Catch
                            ' This catches illegal bmpBase sizes;
                            ' sizes must act in proper multiples of numSubtiles
                        End Try
                    Next
                Next
                'SaveRender()
            Next
            'SaveRender()
            MsgBox(sw.Peek.ToString)
        End Sub

        Private Function ColorRemap(ByVal Col As Color) As System.Drawing.Imaging.ImageAttributes
            ' Create an ImageAttributes object and set the color remap
            Dim imageAttr As New System.Drawing.Imaging.ImageAttributes
            Dim cMap(0) As System.Drawing.Imaging.ColorMap
            cMap(0) = New System.Drawing.Imaging.ColorMap
            cMap(0).OldColor = Color.Black
            cMap(0).NewColor = Col
            imageAttr.SetRemapTable(cMap)
            Return imageAttr
        End Function
        Private Sub DrawBackground(ByRef graDest As Graphics, ByVal TextureBitmap As Bitmap, ByVal Col As Color)
            If graDest Is Nothing Then Exit Sub
            Dim imageAttr As System.Drawing.Imaging.ImageAttributes = ColorRemap(Col)

            ' Build TextureBrush bitmap
            Dim rect As New Rectangle(0, 0, TextureBitmap.Width, TextureBitmap.Height)
            Dim bmpTexture As New Bitmap(TextureBitmap)
            Dim graTexture As Graphics = Graphics.FromImage(bmpTexture)
            graTexture.DrawImage(TextureBitmap, rect, 0, 0, rect.Width, rect.Height, GraphicsUnit.Pixel, imageAttr)

            ' Make a Temporary Bitmap filled with the colored Outline tile
            Dim tb As New TextureBrush(bmpTexture)
            Dim bmpTemp As New Bitmap(CInt(graDest.VisibleClipBounds.Width), CInt(graDest.VisibleClipBounds.Height))
            Dim graTemp As Graphics = Graphics.FromImage(bmpTemp)
            graTemp.FillRectangle(tb, graTemp.ClipBounds)

            ' Reset image attributes
            imageAttr = New System.Drawing.Imaging.ImageAttributes
            imageAttr.SetColorKey(Color.White, Color.White, System.Drawing.Imaging.ColorAdjustType.Bitmap)
            rect = New Rectangle(0, 0, bmpTemp.Width, bmpTemp.Height)
            graDest.DrawImage(bmpTemp, rect, 0, 0, rect.Width, rect.Height, GraphicsUnit.Pixel, imageAttr)
        End Sub

        Public Sub DrawTile(ByVal pc As Zlosk.Patterns.Pattern.PatternCoordinate, ByVal Col As Color)
            DrawSubTile(Me.graRendered, Me.SubTiles(pc.Index - 1), Col, pc.x, pc.y)
        End Sub
        Private Sub DrawSubTile2(ByRef graDest As Graphics, ByVal stTemp As SubTile, ByVal Col As Color, ByVal x As Integer, ByVal y As Integer)
            ' Create an ImageAttributes object and set the color remap
            Dim imageAttr As System.Drawing.Imaging.ImageAttributes = ColorRemap(Col)

            ' Make a Temporary Bitmap of the colored Subtile
            Dim w, h As Integer
            w = stTemp.Width
            h = stTemp.Height
            Dim bmpTemp As New Bitmap(w, h)
            Dim graTemp As Graphics = Graphics.FromImage(bmpTemp)
            Dim rect As Rectangle = New Rectangle(0, 0, w, h)
            graTemp.DrawImage(stTemp.Image, rect, 0, 0, w, h, GraphicsUnit.Pixel, imageAttr)

            ' Reset image attributes
            imageAttr = New System.Drawing.Imaging.ImageAttributes
            imageAttr.SetColorKey(Color.White, Color.White, System.Drawing.Imaging.ColorAdjustType.Bitmap)
            ' Build destination rectangle
            Dim pt As Point = TileLocationToImagePoint(x, y)
            pt.Offset(stTemp.PatternXOffset, stTemp.PatternYOffset)
            rect = New Rectangle(pt.X, pt.Y, w, h)
            graDest.DrawImage(bmpTemp, rect, 0, 0, w, h, GraphicsUnit.Pixel, imageAttr)
        End Sub
        Private Sub DrawSubTile2(ByRef dcDest As IntPtr, ByVal idx As Integer, ByVal st As SubTile, ByVal Col As Color, ByVal x As Integer, ByVal y As Integer)
            Static sintW, sintH As Integer
            Static sintStIndex As Integer
            Static sbmpSubTileMatte, sbmpSubTile, sbmpColor As Bitmap
            Static sgraSubTileMatte, sgraSubTile, sgraColor As Graphics
            Static scolOld As Color
            Dim dcSubTileMatte, dcSubTile As IntPtr
            If x = 0 And y = 0 Then
                'Probably rebuilding the image; reset static variables
                sintStIndex = -1
                scolOld = Nothing
            End If

            ' Determine where to place colored subtile bitmap
            Dim pt As Point = TileLocationToImagePoint(x, y)
            pt.Offset(st.PatternXOffset, st.PatternYOffset)

            ' If the subtile index is the same, it is not necessary
            ' to recreate the subtile matte bitmaps
            If sintStIndex <> idx Then
                ' Subtile is different; recreate the colored bitmap image
                ' Get the hDC of the matte
                sbmpSubTileMatte = New Bitmap(st.Image)
                sgraSubTileMatte = Graphics.FromImage(sbmpSubTileMatte)
                dcSubTileMatte = sgraSubTileMatte.GetHdc
                ' Create inverted matte
                sintW = sbmpSubTileMatte.Width
                sintH = sbmpSubTileMatte.Height
                sbmpSubTile = New Bitmap(sintW, sintH)
                sgraSubTile = Graphics.FromImage(sbmpSubTile)
                dcSubTile = sgraSubTile.GetHdc
                BitBlt(dcSubTile, 0, 0, sintW, sintH, dcSubTileMatte, 0, 0, SRCINVERT)
                ' Create colored rectangle
                sbmpColor = New Bitmap(sintW, sintH)
                sgraColor = Graphics.FromImage(sbmpColor)
                Dim sb As New SolidBrush(Col)
                sgraColor.FillRectangle(sb, 0, 0, sintW, sintH)
                ' Create colored subtile
                Dim dcColor As IntPtr = sgraColor.GetHdc
                BitBlt(dcSubTile, 0, 0, sintW, sintH, dcColor, 0, 0, SRCAND)
                ' WooHoo! dcSubtile now holds the finished subtile

                ' Release the color hdc
                sgraColor.ReleaseHdc(dcColor)
            Else
                If scolOld.ToArgb <> Col.ToArgb Then
                    ' Subtile is same; color is different

                    ' Get the hDC of the matte
                    dcSubTileMatte = sgraSubTileMatte.GetHdc
                    ' Get the hDC of the colored subtile
                    dcSubTile = sgraSubTile.GetHdc
                    ' Create inverted matte
                    sintW = sbmpSubTileMatte.Width
                    sintH = sbmpSubTileMatte.Height
                    sbmpSubTile = New Bitmap(sintW, sintH)
                    sgraSubTile = Graphics.FromImage(sbmpSubTile)
                    dcSubTile = sgraSubTile.GetHdc
                    BitBlt(dcSubTile, 0, 0, sintW, sintH, dcSubTileMatte, 0, 0, SRCINVERT)
                    ' Create colored rectangle
                    sbmpColor = New Bitmap(sintW, sintH)
                    sgraColor = Graphics.FromImage(sbmpColor)
                    Dim sb As New SolidBrush(Col)
                    sgraColor.FillRectangle(sb, 0, 0, sintW, sintH)
                    ' Create colored subtile
                    Dim dcColor As IntPtr = sgraColor.GetHdc
                    BitBlt(dcSubTile, 0, 0, sintW, sintH, dcColor, 0, 0, SRCAND)
                    ' WooHoo! dcSubtile now holds the finished subtile

                    ' Draw the subtile...First matte it
                    BitBlt(dcDest, pt.X, pt.Y, sintW, sintH, dcSubTileMatte, 0, 0, SRCAND)
                    ' Then paint it
                    BitBlt(dcDest, pt.X, pt.Y, sintW, sintH, dcSubTile, 0, 0, SRCPAINT)
                End If
            End If

            ' Draw the subtile...First matte it
            BitBlt(dcDest, pt.X, pt.Y, sintW, sintH, dcSubTileMatte, 0, 0, SRCAND)
            ' Then paint it
            BitBlt(dcDest, pt.X, pt.Y, sintW, sintH, dcSubTile, 0, 0, SRCPAINT)
            ' Release the resources
            sgraSubTileMatte.ReleaseHdc(dcSubTileMatte)
            sgraSubTile.ReleaseHdc(dcSubTile)
            ' We aren't disposing of the sgra's because they're static

            ' Set the vars so we know what's going on next time through this sub
            sintStIndex = idx
            scolOld = Col
        End Sub

        Private Sub DrawSubTile(ByRef ClientDC As Graphics, ByVal stTemp As SubTile, ByVal Col As Color, ByVal x As Integer, ByVal y As Integer)
            ' Create an ImageAttributes object and set the color remap
            Dim imageAttr As System.Drawing.Imaging.ImageAttributes = ColorRemap(Col)

            ' Make a Temporary Bitmap of the colored Subtile
            Dim w, h As Integer
            w = stTemp.Width
            h = stTemp.Height
            Dim bmpTemp As New Bitmap(w, h)
            Dim graTemp As Graphics = Graphics.FromImage(bmpTemp)
            Dim rect As Rectangle = New Rectangle(0, 0, w, h)
            graTemp.DrawImage(stTemp.Image, rect, 0, 0, w, h, GraphicsUnit.Pixel, imageAttr)

            ' Reset image attributes
            imageAttr = New System.Drawing.Imaging.ImageAttributes
            imageAttr.SetColorKey(Color.White, Color.White, System.Drawing.Imaging.ColorAdjustType.Bitmap)
            ' Build destination rectangle
            Dim pt As Point = TileLocationToImagePoint(x, y)
            pt.Offset(stTemp.PatternXOffset, stTemp.PatternYOffset)
            rect = New Rectangle(pt.X, pt.Y, w, h)
            ClientDC.DrawImage(bmpTemp, rect, 0, 0, w, h, GraphicsUnit.Pixel, imageAttr)
            ' Clean up
            imageAttr.Dispose()
            graTemp.Dispose()
            bmpTemp.Dispose()
        End Sub
        Private Sub DrawSubTileWorkingVersion(ByRef ClientDC As Graphics, ByVal stTemp As SubTile, ByVal Col As Color, ByVal x As Integer, ByVal y As Integer)
            ' Loading TCAT.bmp takes approx 24000 ticks on work machine
            ' Create an ImageAttributes object and set the color remap
            Dim imageAttr As System.Drawing.Imaging.ImageAttributes = ColorRemap(Col)

            ' Make a Temporary Bitmap of the colored Subtile
            Dim w, h As Integer
            w = stTemp.Width
            h = stTemp.Height
            Dim bmpTemp As New Bitmap(w, h)
            Dim graTemp As Graphics = Graphics.FromImage(bmpTemp)
            Dim rect As Rectangle = New Rectangle(0, 0, w, h)
            graTemp.DrawImage(stTemp.Image, rect, 0, 0, w, h, GraphicsUnit.Pixel, imageAttr)

            ' Reset image attributes
            imageAttr = New System.Drawing.Imaging.ImageAttributes
            imageAttr.SetColorKey(Color.White, Color.White, System.Drawing.Imaging.ColorAdjustType.Bitmap)
            ' Build destination rectangle
            Dim pt As Point = TileLocationToImagePoint(x, y)
            pt.Offset(stTemp.PatternXOffset, stTemp.PatternYOffset)
            rect = New Rectangle(pt.X, pt.Y, w, h)
            ClientDC.DrawImage(bmpTemp, rect, 0, 0, w, h, GraphicsUnit.Pixel, imageAttr)
        End Sub
        Private Sub DrawSubTile2(ByRef graDest As Graphics, ByVal idx As Integer, ByVal st As SubTile, ByVal Col As Color, ByVal x As Integer, ByVal y As Integer)
            Static sintW, sintH As Integer
            Static sintStIndex As Integer
            Static sbmpSubTile, sbmpTemp As Bitmap
            Static sgraTemp As Graphics
            Static scolOld As Color

            If x = 0 And y = 0 Then
                'Probably rebuilding the image; reset static variables
                sintStIndex = -1
                scolOld = Nothing
            End If

            ' Determine where to place colored subtile bitmap
            Dim pt As Point = TileLocationToImagePoint(x, y)
            pt.Offset(st.PatternXOffset, st.PatternYOffset)

            ' If the subtile index is the same, it is not necessary
            ' to recreate the subtile matte bitmaps
            If sintStIndex <> idx Then
                ' Subtile is different; recreate everything
                'Make a Temporary Bitmap of the Colored SubTile
                sbmpSubTile = New Bitmap(st.Image)
                sbmpSubTile.MakeTransparent(Color.Black)
                sintW = sbmpSubTile.Width
                sintH = sbmpSubTile.Height
                sbmpTemp = New Bitmap(sintW, sintH)
                sgraTemp = Graphics.FromImage(sbmpTemp)

                Dim sb As New SolidBrush(Col)
                sgraTemp.FillRectangle(sb, 0, 0, sintW, sintH)

                ' These 6 lines replace the 7th in an attempt to speed things up
                'Dim SrcDC, DestDC As IntPtr
                'SrcDC = sgraTemp.GetHdc()
                'DestDC = graDest.GetHdc()
                'BitBlt(DestDC, 0, 0, sintW, sintH, SrcDC, 0, 0, SRCCOPY)
                'sgraTemp.ReleaseHdc(SrcDC)
                'graDest.ReleaseHdc(DestDC)
                sgraTemp.DrawImage(sbmpSubTile, 0, 0)

                sgraTemp.Dispose()
                sintStIndex = idx
                scolOld = Col
            Else
                If scolOld.ToArgb <> Col.ToArgb Then
                    ' Subtile is same; color is different
                    Dim sb As New SolidBrush(Col)
                    sgraTemp = Graphics.FromImage(sbmpTemp)
                    sgraTemp.FillRectangle(sb, 0, 0, sintW, sintH)

                    ' These 6 lines replace the 7th in an attempt to speed things up
                    'Dim SrcDC, DestDC As IntPtr
                    'SrcDC = sgraTemp.GetHdc()
                    'DestDC = graDest.GetHdc()
                    'BitBlt(DestDC, 0, 0, sintW, sintH, SrcDC, 0, 0, SRCCOPY)
                    'sgraTemp.ReleaseHdc(SrcDC)
                    'graDest.ReleaseHdc(DestDC)
                    sgraTemp.DrawImage(sbmpSubTile, 0, 0)

                    scolOld = Col
                    sgraTemp.Dispose()
                End If
            End If

            ' Build destination rectangle
            Dim rect As Rectangle = New Rectangle(pt.X, pt.Y, sintW, sintH)
            sbmpTemp.MakeTransparent(Color.White)
            graDest.DrawImage(sbmpTemp, rect, 0, 0, sintW, sintH, GraphicsUnit.Pixel)
        End Sub

        Private Function TileLocationToImagePoint(ByVal pt As Point) As Point
            Return TileLocationToImagePoint(pt.X, pt.Y)
        End Function
        Private Function TileLocationToImagePoint(ByVal x As Integer, ByVal y As Integer) As Point
            Dim xOffset, yOffset As Integer
            xOffset = (y * Me.MainTile.NextRowXOffset) Mod Me.MainTile.XRepeat
            yOffset = (x * Me.MainTile.NextRowYOffset) Mod Me.MainTile.YRepeat
            Return New Point(Me.MainTile.XRepeat * x + xOffset, Me.MainTile.YRepeat * y + yOffset)
        End Function

        Public Sub SaveImages()
            If Not bmpRendered Is Nothing Then bmpRendered.Save("C:\temp\bmpRendered.png")
        End Sub

        Public Function PointToPatternCoordinate(ByVal pt As Point) As Zlosk.Patterns.Pattern.PatternCoordinate
            Dim pc As Zlosk.Patterns.Pattern.PatternCoordinate
            If Me.BaseImage Is Nothing Then
                pc.x = -1
                Return pc
            End If
            Dim i, j, n As Integer
            Dim ImagePoint As Point
            Dim rect As Rectangle
            For j = 0 To (BaseImage.Height - 1) \ Me.MainTile.YReadWrite
                For i = 0 To (BaseImage.Width - 1) \ Me.MainTile.XReadWrite
                    ' Check every Tile bounding box
                    ImagePoint = TileLocationToImagePoint(i, j)
                    rect = New Rectangle(ImagePoint, Me.MainTile.Size)
                    If IsInBoundingBox(pt, rect) Then
                        ' Check every SubTile bounding box
                        For n = 0 To Me.MainTile.NumSubTiles - 1
                            rect = New Rectangle(ImagePoint.X + Me.SubTiles(n).PatternXOffset, ImagePoint.Y + Me.SubTiles(n).PatternYOffset, Me.SubTiles(n).Width, Me.SubTiles(n).Height)
                            If IsInBoundingBox(pt, rect) Then
                                ' Check the subtile image
                                If Me.SubTiles(n).Image.GetPixel(pt.X - rect.Left, pt.Y - rect.Top).ToArgb = Color.Black.ToArgb Then
                                    pc.x = i
                                    pc.y = j
                                    pc.Index = n + 1
                                    Return pc
                                End If
                            End If
                        Next
                    End If
                Next
            Next
            pc.x = -1
            Return pc
        End Function
        Private Function IsInTile(ByVal pt As Point, ByVal x As Integer, ByVal y As Integer) As Boolean
            Dim xOffset, yOffset As Integer
            xOffset = (y * Me.MainTile.NextRowXOffset) Mod Me.MainTile.XRepeat
            yOffset = (x * Me.MainTile.NextRowYOffset) Mod Me.MainTile.YRepeat
            Dim rect As Rectangle = New Rectangle(Me.MainTile.XRepeat * x + xOffset, Me.MainTile.YRepeat * y + yOffset, Me.MainTile.Width, Me.MainTile.Height)
            If IsInBoundingBox(pt, rect) Then Return True
        End Function
        Private Function IsInSubTile(ByVal pt As Point, ByVal x As Integer, ByVal y As Integer, ByVal Index As Integer) As Boolean
            Dim xOffset, yOffset As Integer
            xOffset = Me.SubTiles(Index).PatternXOffset + (y * Me.MainTile.NextRowXOffset) Mod Me.MainTile.XRepeat
            yOffset = Me.SubTiles(Index).PatternYOffset + (x * Me.MainTile.NextRowYOffset) Mod Me.MainTile.YRepeat
            Dim rect As Rectangle = New Rectangle(Me.MainTile.XRepeat * x + xOffset, Me.MainTile.YRepeat * y + yOffset, Me.MainTile.Width, Me.MainTile.Height)
            If IsInBoundingBox(pt, rect) Then
                ' Check subtile bitmap
                If Me.SubTiles(Index).Image.GetPixel(pt.X - xOffset, pt.Y - yOffset).ToArgb = Color.Black.ToArgb Then
                    Return True
                End If
            End If
        End Function
        Private Function IsInBoundingBox(ByVal pt As Point, ByVal rect As Rectangle) As Boolean
            If pt.X < rect.Left Then Return False
            If pt.Y < rect.Top Then Return False
            If pt.X >= rect.Right Then Return False
            If pt.Y >= rect.Bottom Then Return False
            Return True
        End Function
        Public Sub NewBaseImage(ByVal w As Integer, ByVal h As Integer)
            Dim sw As Org.Mentalis.Utilities.StopWatch = New Org.Mentalis.Utilities.StopWatch
            sw.Reset()
            'Try
            bmpBase = New Bitmap(Me.MainTile.XReadWrite * w, Me.MainTile.YReadWrite * h)
            graBase = Graphics.FromImage(bmpBase)
            Dim rect As Rectangle = New Rectangle(0, 0, bmpBase.Width, bmpBase.Height)
            Dim sb As New SolidBrush(mcolBaseColor)
            graBase.FillRectangle(sb, rect)
            BuildImage()
            'Catch ex As Exception
            '    Dim str As String = ""
            '    str += MainTile.ToString
            '    MsgBox(str, MsgBoxStyle.Information, "Error")
            'End Try
            MsgBox(sw.Peek.ToString)
        End Sub
    End Class

End Namespace