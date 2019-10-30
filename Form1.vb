Option Explicit On
Option Strict On

Imports System.IO
Imports System.IO.Compression
Imports System.Threading
Imports System.Xml

Public Class Main

    Private _menge As Integer
    Private _variablesXml As Boolean = False
    Private ReadOnly _temppath As String = Path.GetTempPath
    Private _xlsxToZip As String                         'Excel zu Zip
    Private _folder As String                            'Tempordner für das Editieren der Binary-Datei
    Private _workbook As String                          'Originalfile aus BW
    Private _xmlScriptDoc As XmlDocument = Nothing


    Friend Function GetAnalysisOfficeXml(tempfilename As String) As Boolean

        Dim s As String
        Dim xmlText As String

        _xlsxToZip = _temppath & Path.GetFileNameWithoutExtension(tempfilename) & ".zip"
        _workbook = tempfilename
        _folder = _temppath & Path.GetFileNameWithoutExtension(tempfilename)

        Try

            'Umbenennen von XLSX zu Zip
            File.Move(_workbook, _xlsxToZip)

            'Erstellen des Orderns 
            Directory.CreateDirectory(_folder)

            'Festlegen des Ausgabeordners
            ZipFile.ExtractToDirectory(_xlsxToZip, _folder)

            Thread.Sleep(500)

            'Löschen der alten Zip-Datei
            File.Delete(_xlsxToZip)

            'Get CustomProperty
            Dim customProperty As String = GetCustomPropertyFromFile(_folder & "\xl")

            'Binary File 
            Dim binaryFileInput() As Byte = File.ReadAllBytes(_folder & "\xl\" & customProperty)

            'Konvertiere Binary File zu String
            s = System.Text.Encoding.Unicode.GetString(binaryFileInput)
            _menge = 0
            Dim i = 0, a = 0
            Do
                Dim beforeHeader As Integer = s.IndexOf("<BICS_VIEW_HEADER", i) + 36
                Dim afterHeader As Integer = s.IndexOf("</BICS_VIEW_HEADER>", a)

                If Not beforeHeader = -1 And Not afterHeader = -1 Then

                    'Decompress Header Information
                    xmlText = DecompressString(s.Substring(beforeHeader, afterHeader - beforeHeader))
                    _xmlScriptDoc = New XmlDocument()
                    _xmlScriptDoc.LoadXml(xmlText)
                    _xmlScriptDoc.Save(Path.GetTempPath & Path.DirectorySeparatorChar & Path.GetFileNameWithoutExtension(txt_path.Text) & _menge & "__XML-Repository.xml")
                    _menge += 1
                    i = beforeHeader + 1
                    a = afterHeader + 1
                Else
                    Exit Do
                End If
            Loop

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Friend Function SetAnalysisOfficeXml(tempfilename As String, xmltext As String) As Boolean
        Dim s As String
        _xlsxToZip = _temppath & Path.GetFileNameWithoutExtension(tempfilename) & ".zip"
        _workbook = tempfilename
        _folder = _temppath & Path.GetFileNameWithoutExtension(tempfilename)

        Try

            'Get CustomProperty
            Dim customProperty As String = GetCustomPropertyFromFile(_folder & "\xl")

            'Binary File 
            Dim binaryFileInput() As Byte = File.ReadAllBytes(_folder & "\xl\" & customProperty)

            'Konvertiere Binary File zu String
            s = System.Text.Encoding.Unicode.GetString(binaryFileInput)

            Dim beforeHeader As Integer = s.IndexOf("<BICS_VIEW_HEADER") + 36
            Dim afterHeader As Integer = s.IndexOf("</BICS_VIEW_HEADER>")

            'Decompress Header Information
            s = s.Replace((s.Substring(beforeHeader, afterHeader - beforeHeader)), xmltext)
            s = s.Replace("<StorePromptsInDocument>False</StorePromptsInDocument>", "<StorePromptsInDocument>True</StorePromptsInDocument>")

            'Konvertiere String zu Binary File
            Dim binaryFileOutput() As Byte = System.Text.Encoding.Unicode.GetBytes(s)

            'Speichere Binary File ab
            File.WriteAllBytes(_folder & "\xl\" & customProperty, binaryFileOutput)

            ZipFile.CreateFromDirectory(_folder, _xlsxToZip)

            'Löschen des Tempordners
            Directory.Delete(_folder, True)

            'Umbennen von Zip zu XLSX
            File.Move(_xlsxToZip, _workbook)

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

#Region "Buttons"
    Private Sub btn_open_Click(sender As Object, e As EventArgs) Handles btn_open.Click
        My.Application.Log.WriteEntry("Action complete.")

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            txt_path.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub btn_decompress_Click(sender As Object, e As EventArgs) Handles btn_decompress.Click
        If GetAnalysisOfficeXml(txt_path.Text) Then
            MessageBox.Show("Dekompression erfolgreich")
        End If
    End Sub

    Private Sub btn_compress_Click(sender As Object, e As EventArgs) Handles btn_compress.Click
        Try
            For i = 0 To _menge - 1
                _xmlScriptDoc = New XmlDocument()
                _xmlScriptDoc.Load(Path.GetTempPath & Path.DirectorySeparatorChar & Path.GetFileNameWithoutExtension(txt_path.Text) & i & "__XML-Repository.xml")
                Dim xmlText As String = Compress(_xmlScriptDoc.InnerXml)
                If SetAnalysisOfficeXml(txt_path.Text, xmlText) Then
                    MessageBox.Show("Kompression erfolgreich")
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Process.Start("explorer.exe", Path.GetTempPath())

    End Sub
End Class
