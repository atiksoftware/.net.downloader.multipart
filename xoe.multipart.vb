Imports System.Net
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks

Public Class Downloader

    Public Percent As Long = 0
    Public Speed As Long = 0
    Public response As String = "wait"
    Public description As String = ""

    Private Temp As String = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Atiksoftware\MultipartDownloader"
    Private Part = 7
    Private Downloaded_size As Long = 0
    Private Downloaded_part As Integer = 0

    Private Crashed As Boolean = False

    Private Function url_len(ByVal adres As String)
        Try
            Dim oRequest As HttpWebRequest = WebRequest.Create(adres)
            Dim oResponse As WebResponse = oRequest.GetResponse
            Dim boyut As Integer = oResponse.ContentLength
            oResponse.Close()
            Return boyut
        Catch ex As Exception
            Return -1
        End Try
    End Function

    Public Function Download(URL As String, PATH As String, Optional ByVal overwrite As Boolean = False)

        response = "downloading"
        If Not My.Computer.FileSystem.DirectoryExists(Temp) Then MkDir(Temp)
        Dim target_len As Long = url_len(URL)
        If Not target_len > 0 Then response = "failed" : Return False

        If My.Computer.FileSystem.FileExists(PATH) Then
            If FileLen(PATH) = target_len And overwrite = False Then response = "success" : Return True
        End If

        Dim range_start(Part)
        Dim range_stop(Part)
        Dim range_size As Long = (target_len - (target_len Mod Part)) / Part
        For i As Integer = 1 To Part
            range_start(i) = ((i - 1) * range_size)
            range_stop(i) = i * range_size - 1
        Next
        If (range_stop(Part) < range_size) Then range_stop(Part) += target_len - range_stop(Part)

        Dim new_id = str_md5((Rnd() * 287653) & (Rnd() * 287653) & (Rnd() * 287653) & (Rnd() * 287653) & Now.Millisecond)
        For i As Integer = 1 To Part
            Dim file_temp = Temp & "\" & new_id & "_" & i & ".dat"
            Try
                Dim thread As New System.ComponentModel.BackgroundWorker
                AddHandler thread.DoWork, AddressOf Downloading_Running
                AddHandler thread.RunWorkerCompleted, AddressOf Downloading_Complated
                thread.RunWorkerAsync({URL, range_start(i), range_stop(i), file_temp})
            Catch ex As Exception
                description = "Multipart başlarken bir haa oluştu:" & ex.Message
                response = "failed" : Return False
            End Try
        Next

        'PROGRESS 
        Dim last_bytes As Long = 0
        Dim timeout = 0
        Do While Downloaded_part < Part And Crashed = False
            Try
                Speed = Downloaded_size - last_bytes
                last_bytes = Downloaded_size
                Percent = Math.Round(Downloaded_size / target_len * 100, 2)
                If Speed = 0 Then timeout += 1
                If timeout > 30 Then response = "failed" : description = "Bağlantı zaman aşımına uğradı" : Return False
                Thread.Sleep(1000)
            Catch ex As Exception
                response = "failed"
                description = "Oran hesaplanırken bir hata olutu"
                Return False
            End Try
        Loop
        Speed = 0
        Percent = 100
        Thread.Sleep(300)
        If Crashed = True Then
            response = "failed"
            description = "Multipart crashed: " & description
            Return False
        End If


        ''Birleştir
        Try

            Dim TotalSizes As Long = 0
            Dim totalReaded As Long = 0
            For i As Integer = 1 To Part
                Dim file_temp = Temp & "\" & new_id & "_" & i & ".dat"
                Dim nn As New FileInfo(file_temp)
                TotalSizes += nn.Length
            Next
            Dim bytesRead As Integer

            Dim buffer(4194304) As Byte
            Using outFile As New System.IO.FileStream(PATH, IO.FileMode.Append, IO.FileAccess.Write)
                For i As Integer = 1 To Part
                    Dim file_temp = Temp & "\" & new_id & "_" & i & ".dat"

                    Using inFile As New System.IO.FileStream(file_temp, IO.FileMode.Open, IO.FileAccess.Read)
                        Do
                            bytesRead = inFile.Read(buffer, 0, 4194304)
                            totalReaded += bytesRead
                            Percent = Math.Round((totalReaded / TotalSizes) * 100, 2)
                            outFile.Write(buffer, 0, bytesRead)
                        Loop While bytesRead > 0
                    End Using
                Next
            End Using
        Catch ex As Exception
            description = "Birleştirmede bir hata oluştu: " & ex.Message
            response = "failed" : Return False
        End Try

        For i As Integer = 1 To Part
            Dim file_temp = Temp & "\" & new_id & "_" & i & ".dat"
            If My.Computer.FileSystem.FileExists(file_temp) Then Kill(file_temp)
        Next

        response = "success"
        Return True
    End Function

    Private Sub Downloading_Running(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
        Dim stroutput As String = e.Argument.ToString
        Dim url As String = e.Argument(0).ToString
        Dim range_start As Integer = e.Argument(1)
        Dim range_stop As Integer = e.Argument(2)
        Dim Path As String = e.Argument(3).ToString
        Try
            Dim oRequest As HttpWebRequest = WebRequest.Create(url)
            oRequest.AllowAutoRedirect = True
            oRequest.AddRange(range_start, range_stop)
            Dim oResponse As WebResponse
            oResponse = oRequest.GetResponse
            Dim responseStream As IO.Stream = oResponse.GetResponseStream
            Dim fs As New IO.FileStream(Path, FileMode.Create, FileAccess.Write)
            Dim buffer(2047) As Byte
            Dim read As Integer
            Do
                read = responseStream.Read(buffer, 0, buffer.Length)
                Downloaded_size += read
                fs.Write(buffer, 0, read)
            Loop Until read = 0
            responseStream.Close()
            fs.Flush()
            fs.Close()
            responseStream.Close()
            oResponse.Close()
            e.Result = True
        Catch ex As Exception
            description = ex.Message
            e.Result = False
        End Try
    End Sub
    Private Sub Downloading_Complated(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
        If e.Result = True Then
            Downloaded_part += 1
        Else
            Crashed = True
        End If
    End Sub








#Region "Fonksiyonlar"
    Private Function str_md5(ByVal yazi As String) As String
        Try
            Dim MD5şifreleyici As New System.Security.Cryptography.MD5CryptoServiceProvider
            Dim baytlar As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(yazi)
            Dim hash As Byte() = MD5şifreleyici.ComputeHash(baytlar)
            Dim kapasite As Integer = (hash.Length * 2 + (hash.Length / 8))
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder(kapasite)
            Dim I As Integer
            For I = 0 To hash.Length - 1
                sb.Append(BitConverter.ToString(hash, I, 1))
            Next I
            Return sb.ToString().TrimEnd(New Char() {" "c}).ToLower
        Catch ex As Exception
            Return "0"
        End Try
    End Function

    Private Function GetUnixTimestamp() As Double
        Return (DateTime.Now - New DateTime(1970, 1, 1, 0, 0, 0, 0)).TotalSeconds
    End Function

#End Region




End Class
