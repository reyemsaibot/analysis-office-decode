Imports System.IO
Imports System.IO.Compression
Imports System.Text

Module functions

    Friend Function DecompressString(ByVal compressedText As String) As String
        Dim gZipBuffer() As Byte = Convert.FromBase64String(compressedText)
        Using memoryStream = New MemoryStream()
            Dim dataLength As Integer = BitConverter.ToInt32(gZipBuffer, 0)
            memoryStream.Write(gZipBuffer, 4, gZipBuffer.Length - 4)

            Dim buffer = New Byte(dataLength - 1) {}

            memoryStream.Position = 0
            Using gZipStream = New GZipStream(memoryStream, CompressionMode.Decompress)
                gZipStream.Read(buffer, 0, buffer.Length)
            End Using

            Return Encoding.UTF8.GetString(buffer)
        End Using
    End Function

    Friend Function Compress(ByVal text As String) As String
        Dim buffer() As Byte = Encoding.UTF8.GetBytes(text)
        Dim memoryStream = New MemoryStream()
        Using gZipStream = New GZipStream(memoryStream, CompressionMode.Compress, True)
            gZipStream.Write(buffer, 0, buffer.Length)
        End Using

        memoryStream.Position = 0

        Dim compressedData = New Byte(CInt(memoryStream.Length - 1)) {}
        memoryStream.Read(compressedData, 0, compressedData.Length)

        Dim gZipBuffer = New Byte(compressedData.Length + 4 - 1) {}
        System.Buffer.BlockCopy(compressedData, 0, gZipBuffer, 4, compressedData.Length)
        System.Buffer.BlockCopy(BitConverter.GetBytes(buffer.Length), 0, gZipBuffer, 0, 4)
        Return Convert.ToBase64String(gZipBuffer)
    End Function

    Friend Function GetCustomPropertyFromFile(ByVal pfad As String) As String
        Dim dir As New DirectoryInfo(pfad)
        Dim filelist = dir.GetFiles("*custom*.bin", SearchOption.AllDirectories)
        Dim maxSize = Aggregate aFile In filelist Into Max(GetFileLength(aFile))
        For Each datei In filelist
            If datei.Length = maxSize Then
                Return datei.Name
            End If
        Next
        Return Nothing
    End Function

    Private Function GetFileLength(ByVal fi As System.IO.FileInfo) As Long
        Dim retval As Long
        Try
            retval = fi.Length
        Catch ex As FileNotFoundException
            retval = 0
        End Try
        Return retval
    End Function

End Module
