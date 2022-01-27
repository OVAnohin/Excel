Imports System.IO
Imports System.Net.Sockets
Imports System.Text
Imports Microsoft.Office.Interop

Module Module1

    Dim HasAttachment As Boolean = False
    Dim aryAttachments() As String
    Dim networkStream As NetworkStream
    Dim readStream As StreamReader

    Sub Main()
        Dim popServer As New TcpClient()
        Dim popHost As String = "your_host"
        Dim user As String = "your_username"
        Dim pass As String = "your_password"

        Try
            popServer.Connect(popHost, 110)
            networkStream = popServer.GetStream()
            readStream = New StreamReader(networkStream)

            Dim returnString As String

            returnString = readStream.ReadLine() + vbCrLf
            Console.WriteLine(returnString)
            Console.WriteLine(PopCommand(networkStream, "USER " & user) & vbCrLf)
            Console.WriteLine(PopCommand(networkStream, "PASS " & pass) & vbCrLf)

            Call QuitServer()

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Private Function PopCommand(networkStream As NetworkStream, serverCommand As String) As String

        Try
            serverCommand = serverCommand & vbCrLf
            Dim serverBytes() As Byte = Encoding.ASCII.GetBytes(serverCommand)
            Dim readStreamBytes As StreamReader
            Dim serverResponse As String

            networkStream.Write(serverBytes, 0, serverBytes.Length)
            readStreamBytes = New StreamReader(networkStream)
            serverResponse = readStreamBytes.ReadLine()

            Return serverResponse
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    Private Sub QuitServer()
        Try
            Dim serverResponse As String = ""
            serverResponse = PopCommand(networkStream, "QUIT")
            Console.WriteLine(serverResponse)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

End Module
