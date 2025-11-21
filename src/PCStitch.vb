Namespace Zlosk

    Public Class PCStitch

        Private mbmp As Bitmap
        Private mfl() As Floss

        Structure Floss
            Dim Color As Color
            Dim ID As String
            Dim Description As String
            Dim Symbol As Char
            Dim PaletteIndex As Int16
            Public Overrides Function ToString() As String
                Dim str As String
                str = Me.PaletteIndex.ToString & ", Color: " & Me.Color.ToString & ", ID: " & Me.ID.ToString.TrimEnd & ", " & "Desc: " & Me.Description.TrimEnd
                Return str
            End Function
        End Structure
        Structure RleStream
            Dim Count As Int16
            Dim MatlIdx As Byte
            Dim Type As Byte '01 = full square, 03 = half square, etc.
            Public Overrides Function ToString() As String
                Dim str As String
                str = "Ct: " & Me.Count.ToString & ", Idx: " & Me.MatlIdx.ToString & ", Type: " & Me.Type.ToString
                Return str
            End Function
        End Structure

        Public ReadOnly Property Bitmap() As Bitmap
            Get
                Return mbmp
            End Get
        End Property

        Public Sub New()
        End Sub
        Public Sub New(ByVal Filename As String)
            Me.FromFile(Filename)
        End Sub

        Public Sub FromFile(ByVal Filename As String)
            Dim fs As New System.IO.FileStream(Filename, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fs)
            Try
                ' Get pattern info
                Dim strFileType As String = System.Convert.ToString(br.ReadChars(55))
                Dim intVersion As Int16
                Select Case strFileType
                    Case Is = "PCStitch 5 Pattern File                                "
                        intVersion = 5
                        Me.FromFile5(br)
                    Case Is = "PCStitch 6 Pattern File                                "
                        intVersion = 6
                        Me.FromFile6(br)
                    Case Is = "PCStitch 7 Pattern File                                "
                        intVersion = 7
                        Me.FromFile7(br)
                    Case Else
                        MsgBox("Invalid version: Format must be 'PCStitch 5, 6 or 7 Pattern File'.")
                        Exit Sub
                End Select
            Catch
                MsgBox("Error reading PCStitch file " & Filename)
            End Try
            fs.Close()
        End Sub

        Private Sub FromFile5(ByVal br As System.IO.BinaryReader)
            ' This sub loads a PCStitch 5 File Pattern
            br.BaseStream.Position = &H108
            Dim intWidth As Int16 = br.ReadInt16                    ' Width
            Dim intHeight As Int16 = br.ReadInt16                   ' Height
            Dim cc1 As Int16 = br.ReadInt16()                       ' Cloth count
            Dim cc2 As Int16 = br.ReadInt16()                       ' Cloth count
            Dim strAuthor As String = Me.ReadLenString(br)          ' Author
            Dim strCompany As String = Me.ReadLenString(br)         ' Company
            Dim strTitle As String = Me.ReadLenString(br)           ' Title
            Dim strFabric As String = Me.ReadLenString(br)          ' Fabric
            Dim strInstructions As String = Me.ReadLenString(br)    ' Instructions
            Dim strPalette As String = System.Convert.ToString(br.ReadChars(25))
            Dim intNumColors As Int16 = br.ReadInt16()
            Dim strDMC1 As String = System.Convert.ToString(br.ReadChars(10))
            Dim strDMC2 As String = System.Convert.ToString(br.ReadChars(10))
            Dim strDMC3 As String = System.Convert.ToString(br.ReadChars(10))
            Dim strPCStitchSymbols As String = Me.ReadLenString(br) ' PC Stitch Symbols
            Dim lng1 As Long = br.ReadInt32                         ' Unknown: 0

            ' Read floss library
            Dim i As Int16
            ReDim mfl(intNumColors)
            For i = 1 To intNumColors
                mfl(i) = ReadFloss5(br)
            Next i

            ' Read bitmap
            Dim rles As RleStream
            Dim blnInitBmp As Boolean = True, blnBmpFull As Boolean = False
            mbmp = New Bitmap(intWidth, intHeight)
            Do Until blnBmpFull
                rles = Me.ReadRLEStream(br)
                blnBmpFull = AddToBmp(rles, blnInitBmp)
                blnInitBmp = False
            Loop
        End Sub
        Private Sub FromFile6(ByVal br As System.IO.BinaryReader)
            ' This sub loads a PCStitch 6 File Pattern
            br.BaseStream.Position = &H108
            Dim intWidth As Int16 = br.ReadInt16                    ' Width
            Dim intHeight As Int16 = br.ReadInt16                   ' Height
            Dim cc1 As Int16 = br.ReadInt16()                       ' Cloth count
            Dim cc2 As Int16 = br.ReadInt16()                       ' Cloth count
            Dim strAuthor As String = Me.ReadLenString(br)          ' Author
            Dim strCompany As String = Me.ReadLenString(br)         ' Company
            Dim strTitle As String = Me.ReadLenString(br)           ' Title
            Dim strFabric As String = Me.ReadLenString(br)          ' Fabric
            Dim strInstructions As String = Me.ReadLenString(br)    ' Instructions
            Dim strPalette As String = System.Convert.ToString(br.ReadChars(25))
            Dim intNumColors As Int16 = br.ReadInt16()
            Dim strDMC1 As String = System.Convert.ToString(br.ReadChars(10))
            Dim strDMC2 As String = System.Convert.ToString(br.ReadChars(10))
            Dim strDMC3 As String = System.Convert.ToString(br.ReadChars(10))
            Dim strDMC4 As String = System.Convert.ToString(br.ReadChars(10))
            intNumColors = br.ReadInt16()
            Dim strPCStitchSymbols As String = Me.ReadLenString(br) ' PC Stitch Symbols
            Dim lng1 As Long = br.ReadInt32                         ' Unknown: 0

            ' Read floss library
            Dim i As Int16
            ReDim mfl(intNumColors)
            For i = 1 To intNumColors
                mfl(i) = ReadFloss6(br)
            Next i

            ' Read bitmap
            Dim rles As RleStream
            Dim blnInitBmp As Boolean = True, blnBmpFull As Boolean = False
            mbmp = New Bitmap(intWidth, intHeight)
            Do Until blnBmpFull
                rles = Me.ReadRLEStream(br)
                blnBmpFull = AddToBmp(rles, blnInitBmp)
                blnInitBmp = False
            Loop
        End Sub

        Private Sub FromFile7(ByVal br As System.IO.BinaryReader)
            ' This sub loads a PCStitch 6 File Pattern
            br.BaseStream.Position = &H108
            Dim intWidth As Int16 = br.ReadInt16                    ' Width
            Dim intHeight As Int16 = br.ReadInt16                   ' Height
            Dim cc1 As Int16 = br.ReadInt16()                       ' Cloth count
            Dim cc2 As Int16 = br.ReadInt16()                       ' Cloth count
            Dim colCloth As Color = Me.ReadColor(br)                ' Cloth Color
            Dim strAuthor As String = Me.ReadLenString(br)          ' Author
            Dim strCopyright As String = Me.ReadLenString(br)       ' Copyright
            Dim strTitle As String = Me.ReadLenString(br)           ' Title
            Dim strFabric As String = Me.ReadLenString(br)          ' Fabric
            Dim strInstructions As String = Me.ReadLenString(br)    ' Instructions
            Dim strKeywords As String = Me.ReadLenString(br)        ' Keywords
            Dim strWebsite As String = Me.ReadLenString(br)         ' Website
            br.ReadBytes(4)                                         ' 2,0,2,0 - Possible Floss Count?
            Dim strPalette As String = System.Convert.ToString(br.ReadChars(25))
            br.ReadBytes(4)                                         ' &H0001117A = 70010
            Dim intNumColors As Int16 = br.ReadInt16()

            ' Read floss library
            Dim i As Int16
            ReDim mfl(intNumColors)
            For i = 1 To intNumColors
                mfl(i) = ReadFloss7(br)
            Next i

            ' Read bitmap
            Dim rles As RleStream
            Dim blnInitBmp As Boolean = True, blnBmpFull As Boolean = False
            mbmp = New Bitmap(intWidth, intHeight)
            Do Until blnBmpFull
                rles = Me.ReadRLEStream(br)
                blnBmpFull = AddToBmp(rles, blnInitBmp)
                blnInitBmp = False
            Loop
        End Sub

        Private Function ReadLenString(ByVal br As System.IO.BinaryReader) As String
            Dim intNum As Int16
            Dim str As String = ""
            Dim b() As Byte
            Dim c() As Char
            Dim d As System.Text.Decoder = System.Text.Encoding.UTF8.GetDecoder()

            intNum = br.ReadInt16 ' Num of characters
            If intNum > 0 Then
                b = br.ReadBytes(intNum)
                ReDim c(intNum - 1)
                Dim charLen As Integer = d.GetChars(b, 0, b.Length, c, 0)
                str = System.Convert.ToString(c)
            End If
            Return str
        End Function
        Private Function ReadColor(ByVal br As System.IO.BinaryReader) As Color
            Dim r, g, b, a As Byte
            r = br.ReadByte()
            g = br.ReadByte()
            b = br.ReadByte()
            a = br.ReadByte()
            Return Color.FromArgb(a, r, g, b)
        End Function
        Private Function ReadFloss5(ByVal br As System.IO.BinaryReader) As Floss
            Dim fl As Floss
            Dim b() As Byte
            Dim i1, i2 As Int16
            Dim str1, str2 As String
            fl.Color = Me.ReadColor(br)
            str1 = System.Convert.ToString(br.ReadChars(10))
            str2 = System.Convert.ToString(br.ReadChars(10))
            fl.ID = System.Convert.ToString(br.ReadChars(10))
            fl.Description = System.Convert.ToString(br.ReadChars(30))
            b = br.ReadBytes(20)                                ' Unknown: 0
            fl.Symbol = System.Convert.ToChar(br.ReadByte)      ' B/W grid symbol
            i1 = br.ReadInt16()                                 ' Unknown: 2
            i2 = br.ReadInt16()                                 ' Unknown: 2
            Return fl
        End Function
        Private Function ReadFloss6(ByVal br As System.IO.BinaryReader) As Floss
            Dim fl As Floss
            Dim b() As Byte
            'Dim c() As Char
            Dim i1, i2, i3, i4 As Int16
            Dim str1, str2 As String
            fl.Description = System.Convert.ToString(br.ReadChars(30))
            fl.Color = Me.ReadColor(br)
            fl.ID = System.Convert.ToString(br.ReadChars(10))
            i1 = br.ReadInt16()                                 ' Unknown: 1
            i2 = br.ReadInt16()                                 ' Unknown: Stitch Count?
            b = br.ReadBytes(36)                                ' Unknown: 0
            b = br.ReadBytes(10)                                ' Unknown: ' '
            b = br.ReadBytes(8)                                 ' Unknown: 0
            fl.Symbol = System.Convert.ToChar(br.ReadByte)      ' B/W grid symbol
            i3 = br.ReadInt16()                                 ' Unknown: 2
            i4 = br.ReadInt16()                                 ' Unknown: 2
            fl.PaletteIndex = br.ReadInt16()                    ' Starts at 0
            str1 = System.Convert.ToString(br.ReadChars(30))
            b = br.ReadBytes(5)                                 ' Unknown: 0
            str2 = System.Convert.ToString(br.ReadChars(25))
            Return fl
        End Function
        Private Function ReadFloss7(ByVal br As System.IO.BinaryReader) As Floss
            ' Each floss uses 250 bytes (&HFA)
            Dim fl As Floss
            Dim b() As Byte
            'Dim c() As Char
            'Dim i1, i2, i3, i4, i5 As Int16
            'Dim str1, str2 As String
            b = br.ReadBytes(199)
            fl.Description = System.Convert.ToString(br.ReadChars(30))  'Description
            fl.Color = Me.ReadColor(br)
            fl.ID = System.Convert.ToString(br.ReadChars(10))   ' DMC Number
            b = br.ReadBytes(7)                                ' Unknown: 0
            Return fl
        End Function
        Private Function ReadRLEStream(ByVal br As System.IO.BinaryReader) As RleStream
            Dim rles As RleStream
            rles.Count = br.ReadInt16
            rles.MatlIdx = br.ReadByte
            rles.Type = br.ReadByte
            Return rles
        End Function

        Private Function AddToBmp(ByVal rles As RleStream, ByVal blnInit As Boolean) As Boolean
            Static x, y As Int16
            Dim i As Int16
            Dim col As Color
            Try
                If blnInit Then
                    x = 0
                    y = 0
                End If
                If rles.Type = 255 Then
                    col = Color.WhiteSmoke
                Else
                    col = mfl(rles.MatlIdx).Color
                End If
                For i = 1 To rles.Count
                    mbmp.SetPixel(x, y, col)
                    y = y + 1
                    If y = mbmp.Height Then
                        x = x + 1
                        y = 0
                    End If
                Next
                If x >= mbmp.Width Then
                    Return True
                Else
                    Return False
                End If
            Catch
                Dim str As String
                str = "  Class: PCStitch" & vbCrLf
                str = "  Function: AddToBmp" & vbCrLf
                str = "  x:" & x.ToString & vbCrLf
                str = "  y:" & y.ToString & vbCrLf
                str = "  i:" & i.ToString & vbCrLf
                str = "  col:" & col.ToString & vbCrLf
                str = "  rles:" & rles.ToString & vbCrLf
                str = "  blnInit:" & blnInit.ToString & vbCrLf
                MsgBox(str, MsgBoxStyle.Critical, "Error")
            End Try
        End Function
    End Class

End Namespace
