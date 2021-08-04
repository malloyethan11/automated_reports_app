' --------------------------------------------------------------------------------
' Options
' --------------------------------------------------------------------------------
Option Explicit On

Imports System.Data.OleDb

' This module adapted from the database utilities module provided in Programming 2
Public Module modDatabaseUtilities




    ' --------------------------------------------------------------------------------
    ' Properties
    ' --------------------------------------------------------------------------------
    ' In a 2-Tier app we connect once during FMain_Load and hold
    ' the connection open while until the application closes
    Public m_conAdministrator As System.Data.OleDb.OleDbConnection

    ' SQL Server Connection string with integrated login v1
    Private m_strDatabaseConnectionStringSQLServerV1 As String = "Provider=SQLOLEDB;" &
                                                                 "Server=itd2.cincinnatistate.edu;" &
                                                                 "Database=CPDM-GroupB;" &
                                                                 "Integrated Security=SSPI;"

    ' SQL Server Connection string with integrated login v2
    Private m_strDatabaseConnectionStringSQLServerV2 As String = "Provider=SQLOLEDB;" &
                                                                 "Server=itd2.cincinnatistate.edu;" &
                                                                 "Database=CPDM-GroupB;" &
                                                                 "Trusted_Connection=True;"

    ' SQL Express Connection string                             
    Private m_strDatabaseConnectionString As String = "Provider=SQLOLEDB;" &
                                                      "Server=itd2.cincinnatistate.edu\SQLEXPRESS;" &
                                                      "Database=CPDM-GroupB;" &
                                                      "User ID=sa;" &
                                                      "Password=;"

    ' SQL Express Connection string                             
    Public m_strDatabaseConnectionStringPassword As String = "Provider=SQLOLEDB;" &
                                                      "Server=itd2.cincinnatistate.edu;" &
                                                      "Database=CPDM-GroupB;" &
                                                      "User ID=" & strConnectionUsername & ";" &
                                                      "Password=" & strConnectionPassword & ";"





#Region "Open/Close Connection"

    ' --------------------------------------------------------------------------------
    ' Name: OpenDatabaseConnectionMSAccess
    ' Abstract: Open a connection to the database.

    ' --------------------------------------------------------------------------------
    Public Function OpenDatabaseConnectionSQLServer(ByVal blnQuiet As Boolean) As Boolean

        Dim blnResult As Boolean = False

        Try

            ' Open a connection to the database
            m_conAdministrator = New OleDb.OleDbConnection
            m_conAdministrator.ConnectionString = m_strDatabaseConnectionStringPassword
            m_conAdministrator.Open()


            ' Success
            blnResult = True

        Catch excError As Exception

            ' Log and display error message
            If (blnQuiet = False) Then
                MessageBox.Show(excError.Message)
            Else
                Console.WriteLine(excError.Message)
            End If

        End Try

        Return blnResult

    End Function



    ' --------------------------------------------------------------------------------
    ' Name: CloseDatabaseConnection
    ' Abstract: If the database connection is open then close it.  Release the
    '           memory.
    ' --------------------------------------------------------------------------------
    Public Function CloseDatabaseConnection() As Boolean

        Dim blnResult As Boolean = False

        Try

            ' Anything there?
            If m_conAdministrator IsNot Nothing Then

                ' Open?
                If m_conAdministrator.State <> ConnectionState.Closed Then

                    ' Yes, close it
                    m_conAdministrator.Close()

                End If

                ' Clean up
                m_conAdministrator = Nothing

            End If

            ' Success
            blnResult = True

        Catch excError As Exception

            ' Log and display error message
            MessageBox.Show(excError.Message)

        End Try

        Return blnResult

    End Function

#End Region





End Module
