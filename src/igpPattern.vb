Option Explicit On 
Option Strict On

Namespace Zlosk.Patterns

    Public Class Pattern
        Public MainTile As Tile, SubTiles() As SubTile
        Private bmpBase, bmpRendered, bmpMem As Bitmap
        Private graBase, graRendered As Graphics
        Private mcolBaseColor, mcolBackColor, mcolForeColor As Color
        Public mctrlClient As Control
        Private mblnTileInitialized As Boolean
#If DEBUG Then
        Private ctBuildImage, ctDrawBackground, ctDrawSubTile, ctDrawTile, ctLoadIni, ctNewBaseImage, ctSaveRender, ctSetPixel, ctUpdateBackColor, ctUpdateForeColor, ctColorRemap, ctGetPixel, ctIsInBoundingBox, ctIsInSubTile, ctIsInTile, ctPointToPatternCoordinate, ctTileLocationToImagePoint, ctDrawImage As Long
#End If

        Event Invalidated(ByVal sender As Object, ByVal e As InvalidateEventArgs)

        Public Sub DebugPrintCounts()
#If DEBUG Then
            Dim str As String
            str = "ctBuildImage: " + ctBuildImage.ToString + vbCrLf
            str += "ctDrawBackground: " + ctDrawBackground.ToString + vbCrLf
            str += "ctDrawSubTile: " + ctDrawSubTile.ToString + vbCrLf
            str += "ctDrawTile: " + ctDrawTile.ToString + vbCrLf
            str += "ctLoadIni: " + ctLoadIni.ToString + vbCrLf
            str += "ctNewBaseImage: " + ctNewBaseImage.ToString + vbCrLf
            str += "ctSaveRender: " + ctSaveRender.ToString + vbCrLf
            str += "ctSetPixel: " + ctSetPixel.ToString + vbCrLf
            str += "ctUpdateBackColor: " + ctUpdateBackColor.ToString + vbCrLf
            str += "ctUpdateForeColor: " + ctUpdateForeColor.ToString + vbCrLf
            str += "ctColorRemap: " + ctColorRemap.ToString + vbCrLf
            str += "ctGetPixel: " + ctGetPixel.ToString + vbCrLf
            str += "ctIsInBoundingBox: " + ctIsInBoundingBox.ToString + vbCrLf
            str += "ctIsInSubTile: " + ctIsInSubTile.ToString + vbCrLf
            str += "ctIsInTile: " + ctIsInTile.ToString + vbCrLf
            str += "ctPointToPatternCoordinate: " + ctPointToPatternCoordinate.ToString + vbCrLf
            str += "ctTileLocationToImagePoint: " + ctTileLocationToImagePoint.ToString + vbCrLf
            str += "ctDrawImage: " + ctDrawImage.ToString + vbCrLf
            Debug.WriteLine(str)
#End If
        End Sub


        Public Structure PatternCoordinate
            Dim x, y, Index As Integer
        End Structure

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
                bmpBase = New Bitmap(Value)
            End Set
        End Property
        ReadOnly Property RenderedImage() As Bitmap
            Get
                Return bmpRendered
            End Get
        End Property
        ReadOnly Property IsInitialized() As Boolean
            Get
                If Me.BaseImage Is Nothing Then Return False
                If Not Me.mblnTileInitialized Then Return False
                Return True
            End Get
        End Property

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
#If DEBUG Then
            ctGetPixel += 1
#End If
            Dim x, y As Integer
            x = pc.x * Me.MainTile.XReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteXOffset
            y = pc.y * Me.MainTile.YReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteYOffset
            Return Me.bmpBase.GetPixel(x, y)
        End Function
        Public Sub SetPixel(ByVal pc As PatternCoordinate, ByVal col As Color)
#If DEBUG Then
            ctSetPixel += 1
#End If
            Dim x, y As Integer
            x = pc.x * Me.MainTile.XReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteXOffset
            y = pc.y * Me.MainTile.YReadWrite + Me.SubTiles(pc.Index - 1).ReadWriteYOffset
            Me.bmpBase.SetPixel(x, y, col)
            DrawTile(pc, col)
            Dim e As New InvalidateEventArgs(New Rectangle(Me.TileLocationToImagePoint(pc.x, pc.y), Me.MainTile.Size))
            RaiseEvent Invalidated(Me, e)
            '            mctrlClient.Invalidate(New Rectangle(Me.TileLocationToImagePoint(pc.x, pc.y), Me.MainTile.Size))
        End Sub

        Public Sub LoadIni(ByVal FileName As String)
#If DEBUG Then
            ctLoadIni += 1
#End If
            Dim ini As New INIReader(FileName)
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
            Me.mblnTileInitialized = True
        End Sub

        Public Sub UpdateForeColor()
#If DEBUG Then
            ctUpdateForeColor += 1
#End If
            If Not Me.MainTile.ImageOutline Is Nothing Then
                DrawBackground(graRendered, Me.MainTile.ImageOutline, Me.ForeColor)
            End If
        End Sub
        Public Sub UpdateBackColor()
#If DEBUG Then
            ctUpdateBackColor += 1
#End If
            If Not Me.MainTile.ImageBackground Is Nothing Then
                DrawBackground(graRendered, Me.MainTile.ImageBackground, Me.BackColor)
            End If
        End Sub

        Public Sub BuildImage()
#If DEBUG Then
            ctBuildImage += 1
#End If
            If Not Me.IsInitialized Then Exit Sub

            bmpRendered = Nothing
            bmpRendered = New Bitmap(Me.MainTile.XRepeat * Me.bmpBase.Width \ Me.MainTile.XReadWrite + (Me.MainTile.Width - Me.MainTile.XRepeat), Me.MainTile.YRepeat * Me.bmpBase.Height \ Me.MainTile.YReadWrite + (Me.MainTile.Height - Me.MainTile.YRepeat))
            graRendered = Graphics.FromImage(bmpRendered)
            BaseColor = Me.bmpBase.GetPixel(0, 0)
            Dim sb As New SolidBrush(BaseColor)
            graRendered.FillRectangle(sb, graRendered.ClipBounds)
            'SaveRender()
            UpdateForeColor()
            'SaveRender()
            UpdateBackColor()
            Dim i, j, idx As Integer
            'SaveRender()
            'i = 1
            Dim colTemp As Color
            For idx = 0 To Me.MainTile.NumSubTiles - 1
                ' These 2 Draws are in to fix a rendering optimization problem
                ' The IF-THEN optimization here interferes with an optimization in 
                ' the DrawSubtile subroutine. This is a kludge.
                j = 0 : i = 0
                colTemp = Me.bmpBase.GetPixel(i + Me.SubTiles(idx).ReadWriteXOffset, j + Me.SubTiles(idx).ReadWriteYOffset)
                DrawSubTile(graRendered, Me.SubTiles(idx), colTemp, i \ Me.MainTile.XReadWrite, j \ Me.MainTile.YReadWrite)
                DrawSubTile(graRendered, Me.SubTiles(idx), colTemp, i \ Me.MainTile.XReadWrite, j \ Me.MainTile.YReadWrite)
                For j = 0 To Me.BaseImage.Height - 1 Step Me.MainTile.YReadWrite
                    For i = 0 To Me.BaseImage.Width - 1 Step Me.MainTile.XReadWrite
                        colTemp = Me.bmpBase.GetPixel(i + Me.SubTiles(idx).ReadWriteXOffset, j + Me.SubTiles(idx).ReadWriteYOffset)
                        'DrawSubTile(graRendered, Me.SubTiles(idx).Image, Me.bmpBase.GetPixel(i + Me.SubTiles(idx).ReadWriteXOffset, j + Me.SubTiles(idx).ReadWriteYOffset), i \ Me.SubTiles(idx).ReadWriteXOffset, j \ Me.SubTiles(idx).ReadWriteYOffset, 0, 0)
                        If Not colTemp.Equals(mcolBaseColor) Then
                            DrawSubTile(graRendered, Me.SubTiles(idx), colTemp, i \ Me.MainTile.XReadWrite, j \ Me.MainTile.YReadWrite)
                        End If
                    Next
                Next
                'SaveRender()
            Next
            'SaveRender()
        End Sub

        Private Function ColorRemap(ByVal Col As Color) As System.Drawing.Imaging.ImageAttributes
#If DEBUG Then
            ctColorRemap += 1
#End If
            Static imageAttr As System.Drawing.Imaging.ImageAttributes
            Static cMap(0) As System.Drawing.Imaging.ColorMap

            ' Create an ImageAttributes object and set the color remap
            imageAttr = New System.Drawing.Imaging.ImageAttributes
            cMap(0) = New System.Drawing.Imaging.ColorMap
            cMap(0).OldColor = Color.Black
            cMap(0).NewColor = Col
            imageAttr.SetRemapTable(cMap)

            Return imageAttr
        End Function
        Private Sub DrawBackground(ByRef ClientDC As Graphics, ByVal TextureBitmap As Bitmap, ByVal Col As Color)
#If DEBUG Then
            ctDrawBackground += 1
#End If
            If ClientDC Is Nothing Then Exit Sub
            Dim imageAttr As System.Drawing.Imaging.ImageAttributes = ColorRemap(Col)

            ' Build TextureBrush bitmap
            Dim rect As New Rectangle(0, 0, TextureBitmap.Width, TextureBitmap.Height)
            Dim bmpTexture As New Bitmap(TextureBitmap)
            Dim graTexture As Graphics = Graphics.FromImage(bmpTexture)
            graTexture.DrawImage(TextureBitmap, rect, 0, 0, rect.Width, rect.Height, GraphicsUnit.Pixel, imageAttr)
#If DEBUG Then
            ctDrawImage += 1
#End If
            ' Make a Temporary Bitmap filled with the colored Outline tile
            Dim tb As New TextureBrush(bmpTexture)
            Dim bmpTemp As New Bitmap(CInt(ClientDC.VisibleClipBounds.Width), CInt(ClientDC.VisibleClipBounds.Height))
            Dim graTemp As Graphics = Graphics.FromImage(bmpTemp)
            graTemp.FillRectangle(tb, graTemp.ClipBounds)

            ' Reset image attributes
            imageAttr = New System.Drawing.Imaging.ImageAttributes
            imageAttr.SetColorKey(Color.White, Color.White, System.Drawing.Imaging.ColorAdjustType.Bitmap)
            rect = New Rectangle(0, 0, bmpTemp.Width, bmpTemp.Height)
            ClientDC.DrawImage(bmpTemp, rect, 0, 0, rect.Width, rect.Height, GraphicsUnit.Pixel, imageAttr)
#If DEBUG Then
            ctDrawImage += 1
#End If
        End Sub

        Public Sub DrawTile(ByVal pc As Zlosk.Patterns.Pattern.PatternCoordinate, ByVal Col As Color)
#If DEBUG Then
            ctDrawTile += 1
#End If
            DrawSubTile(Me.graRendered, Me.SubTiles(pc.Index - 1), Col, pc.x, pc.y)
        End Sub

        Private Sub DrawSubTile(ByRef ClientDC As Graphics, ByVal stTemp As SubTile, ByVal Col As Color, ByVal x As Integer, ByVal y As Integer)
#If DEBUG Then
            ctDrawSubTile += 1
#End If
            ' Create an ImageAttributes object and set the color remap
            Static imageAttrTemp, imageAttrMain As System.Drawing.Imaging.ImageAttributes
            Static colOld As Color
            Static bmpTemp As Bitmap
            Static graTemp As Graphics
            Static stOld As SubTile
            Static oldW, oldH As Integer
            Static oldRect As Rectangle

            Dim blnNewCol As Boolean = False
            Dim blnNewSubTile As Boolean = False

            ' If the pixel color doesn't change, 
            ' why bother changing the color map?
            If Not (Col.Equals(colOld)) Then
                imageAttrTemp = ColorRemap(Col)
                colOld = Col
                blnNewCol = True
            End If

            ' If the SubTile hasn't changed, 
            ' why bother building a new bitmap? 
            If Not (stOld.Equals(stTemp)) Then
                Dim w, h As Integer
                w = stTemp.Width
                h = stTemp.Height
                If (w <> oldW) Or (h <> oldH) Then
                    bmpTemp = New Bitmap(w, h)
                    oldW = w
                    oldH = h
                    oldRect = New Rectangle(0, 0, oldW, oldH)
                    graTemp = Graphics.FromImage(bmpTemp)
                End If
                stOld = stTemp
                blnNewSubTile = True
            End If

            ' If color or subtile changed, we need to build a new
            ' colored Subtile
            If blnNewCol Or blnNewCol Then
                graTemp.DrawImage(stTemp.Image, oldRect, 0, 0, oldW, oldH, GraphicsUnit.Pixel, imageAttrTemp)
#If DEBUG Then
                ctDrawImage += 1
#End If
                imageAttrMain = New System.Drawing.Imaging.ImageAttributes
                imageAttrMain.SetColorKey(Color.White, Color.White, System.Drawing.Imaging.ColorAdjustType.Bitmap)
            End If

            ' Build destination rectangle
            Dim pt As Point = TileLocationToImagePoint(x, y)
            Dim rect As Rectangle
            pt.Offset(stTemp.PatternXOffset, stTemp.PatternYOffset)
            rect = New Rectangle(pt.X, pt.Y, oldW, oldH)
            ClientDC.DrawImage(bmpTemp, rect, 0, 0, oldW, oldH, GraphicsUnit.Pixel, imageAttrMain)
#If DEBUG Then
            ctDrawImage += 1
#End If
        End Sub

        Private Function TileLocationToImagePoint(ByVal pt As Point) As Point
            Return TileLocationToImagePoint(pt.X, pt.Y)
        End Function
        Private Function TileLocationToImagePoint(ByVal x As Integer, ByVal y As Integer) As Point
#If DEBUG Then
            ctTileLocationToImagePoint += 1
#End If
            Dim xOffset, yOffset As Integer
            xOffset = (y * Me.MainTile.NextRowXOffset) Mod Me.MainTile.XRepeat
            yOffset = (x * Me.MainTile.NextRowYOffset) Mod Me.MainTile.YRepeat
            Return New Point(Me.MainTile.XRepeat * x + xOffset, Me.MainTile.YRepeat * y + yOffset)
        End Function

        Public Sub SaveRender()
#If DEBUG Then
            ctSaveRender += 1
#End If
            Dim x As Image = bmpRendered
            x.Save("C:\Data\VB.NET\Patterns\bin\test.png")
        End Sub

        Public Function PointToPatternCoordinate(ByVal pt As Point) As Zlosk.Patterns.Pattern.PatternCoordinate
#If DEBUG Then
            ctPointToPatternCoordinate += 1
#End If
            Dim pc As Zlosk.Patterns.Pattern.PatternCoordinate
            If Not Me.IsInitialized Then
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
#If DEBUG Then
            ctIsInTile += 1
#End If
            Dim xOffset, yOffset As Integer
            xOffset = (y * Me.MainTile.NextRowXOffset) Mod Me.MainTile.XRepeat
            yOffset = (x * Me.MainTile.NextRowYOffset) Mod Me.MainTile.YRepeat
            Dim rect As Rectangle = New Rectangle(Me.MainTile.XRepeat * x + xOffset, Me.MainTile.YRepeat * y + yOffset, Me.MainTile.Width, Me.MainTile.Height)
            If IsInBoundingBox(pt, rect) Then Return True
        End Function
        Private Function IsInSubTile(ByVal pt As Point, ByVal x As Integer, ByVal y As Integer, ByVal Index As Integer) As Boolean
#If DEBUG Then
            ctIsInSubTile += 1
#End If
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
#If DEBUG Then
            ctIsInBoundingBox += 1
#End If
            If pt.X < rect.Left Then Return False
            If pt.Y < rect.Top Then Return False
            If pt.X >= rect.Right Then Return False
            If pt.Y >= rect.Bottom Then Return False
            Return True
        End Function
        Public Sub NewBaseImage(ByVal w As Integer, ByVal h As Integer)
#If DEBUG Then
            ctNewBaseImage += 1
#End If
            bmpBase = New Bitmap(w, h)
            graBase = Graphics.FromImage(bmpBase)
            Dim rect As Rectangle = New Rectangle(0, 0, bmpBase.Width, bmpBase.Height)
            Dim sb As New SolidBrush(mcolBaseColor)
            graBase.FillRectangle(sb, rect)
            BuildImage()
        End Sub
    End Class

End Namespace