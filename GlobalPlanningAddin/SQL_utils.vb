Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Module SQL_utils

    Public Function GetSQLConnection(SQLConnectionString As String) As SqlConnection

        Try
            Dim Connexion As New SqlConnection(SQLConnectionString)
            Connexion.Open()
            Return Connexion
        Catch ex As Exception
            'Error stop here. In case of issue, the problem will be handled in CheckSQLConnectionAndReconnect that should be call straight after.
        End Try

        Return Nothing

    End Function

    Public Function CheckSQLConnectionAndReconnect(Connexion As SqlConnection, NbRetries As Integer) As Boolean
        Dim i As Integer

        If Connexion Is Nothing Then Return False

        Do
            Select Case Connexion.State
                Case ConnectionState.Broken
                    Connexion.Close()
                Case ConnectionState.Closed
                'We will try to open it below
                Case ConnectionState.Open, ConnectionState.Executing, ConnectionState.Fetching
                    Return True
                Case ConnectionState.Connecting
                    For i = 1 To 5
                        Thread.Sleep(1000)
                        If Connexion.State = ConnectionState.Open Then Return True
                    Next i
                    Return False
            End Select
            Try
                Connexion.Open()
            Catch ex As Exception
                'Error stop here
            End Try
            If Connexion.State = ConnectionState.Open Then Return True
            NbRetries -= 1
            Thread.Sleep(1000) 'Wait 1s before retry
        Loop While NbRetries > 0
        Return False

    End Function

End Module



