Public Class frmGetDatabaseCredentials
    Private Sub frmGetDatabaseCredentials_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '' Prompt for database connection credentials
        'strConnectionUsername = InputBox("Please provide database login username: ")
        'strConnectionPassword = InputBox("Please provide database login password: ")

        If (My.Computer.FileSystem.FileExists("dbcred.txt") = True) Then

            ' Get encrypted text
            Dim strEncryptedFile As String = My.Computer.FileSystem.ReadAllText("dbcred.txt")
            ' Start encryption service
            Dim clsEncrypt As New clsEncryption("09b023thboiu3h2rf02p3098t2b03g98uh92oeryugbs0e59y72nb09348tv02934vthm0")
            ' Decrypt text
            Dim strDecryptedFile As String = clsEncrypt.DecryptData(strEncryptedFile)

            ' DId getting the data succeed?
            If (strDecryptedFile <> Nothing) Then
                ' Load Credentials
                Dim strTempUsername As String = ""
                Dim strTempPassword As String = ""
                Dim intCount As Integer = 0

                ' Get all text until pipe
                While (strDecryptedFile.Chars(intCount) <> "|")
                    strTempUsername += strDecryptedFile.Chars(intCount)
                    intCount += 1
                End While

                ' Move passed pipe
                intCount += 1

                ' Get all text until end
                While (intCount < strDecryptedFile.Length)
                    strTempPassword += strDecryptedFile.Chars(intCount)
                    intCount += 1
                End While

                ' Get credentials
                strConnectionUsername = strTempUsername
                strConnectionPassword = strTempPassword

                ' Test connection
                If OpenDatabaseConnectionSQLServer(True) = False Then

                    CloseDatabaseConnection()
                    Console.WriteLine("Unable to connect to database at this time. Either enter new credentials or close the program and try again later.", "Error")

                    Me.CenterToScreen()

                    For Each Control In Controls
                        If Control.GetType() = GetType(System.Windows.Forms.Button) Then
                            Control.FlatStyle = FlatStyle.Flat
                            Control.ForeColor = BackColor
                            Control.FlatAppearance.BorderColor = BackColor
                            Control.FlatAppearance.MouseOverBackColor = BackColor
                            Control.FlatAppearance.MouseDownBackColor = BackColor
                        End If
                    Next

                    txtPassword.PasswordChar = "*"

                    SetImages()

                Else

                    CloseDatabaseConnection()
                    OpenFormKillParent(Me, frmMain)

                    SetImages()

                End If
            Else

                ' Else, get the username and password
                Me.CenterToScreen()

                For Each Control In Controls
                    If Control.GetType() = GetType(System.Windows.Forms.Button) Then
                        Control.FlatStyle = FlatStyle.Flat
                        Control.ForeColor = BackColor
                        Control.FlatAppearance.BorderColor = BackColor
                        Control.FlatAppearance.MouseOverBackColor = BackColor
                        Control.FlatAppearance.MouseDownBackColor = BackColor
                    End If
                Next

                txtPassword.PasswordChar = "*"

                SetImages()

            End If


        Else

            Me.CenterToScreen()

            For Each Control In Controls
                If Control.GetType() = GetType(System.Windows.Forms.Button) Then
                    Control.FlatStyle = FlatStyle.Flat
                    Control.ForeColor = BackColor
                    Control.FlatAppearance.BorderColor = BackColor
                    Control.FlatAppearance.MouseOverBackColor = BackColor
                    Control.FlatAppearance.MouseDownBackColor = BackColor
                End If
            Next

            txtPassword.PasswordChar = "*"

            SetImages()

        End If

    End Sub

    Private Sub SetImages()

        ' Change color - Cite: https://www.daniweb.com/programming/software-development/threads/486625/convert-color-image-to-grayscale
        Dim x As Integer = btmButtonDefault.Width
        Dim y As Integer = btmButtonDefault.Height
        btmButtonDefaultGray = New Bitmap(btmButtonDefault.Width, btmButtonDefault.Height)

        For x = 0 To (btmButtonDefault.Width) - 1
            For y = 0 To (btmButtonDefault.Height) - 1
                Dim c As Color = btmButtonDefault.GetPixel(x, y)
                Dim r As Integer = c.R
                Dim g As Integer = c.G
                Dim b As Integer = c.B
                Dim a As Integer = c.A
                Dim d As Integer = (r + g + b) \ 3
                btmButtonDefaultGray.SetPixel(x, y, Color.FromArgb(a, d, d, d))
            Next
        Next

    End Sub

    Private Sub StepAction_Tick(sender As Object, e As EventArgs) Handles StepAction.Tick

        ButtonColor(MousePosition, btnLogin, Me, btmButtonDefaultGray, btmButtonDefault)

    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        strConnectionUsername = txtUsername.Text
        strConnectionPassword = txtPassword.Text

        m_strDatabaseConnectionStringPassword = "Provider=SQLOLEDB;" &
                                                "Server=itd2.cincinnatistate.edu;" &
                                                "Database=CPDM-GroupB;" &
                                                "User ID=" & strConnectionUsername & ";" &
                                                "Password=" & strConnectionPassword & ";"

        ' Start encryption service
        Dim clsEncrypt As New clsEncryption("09b023thboiu3h2rf02p3098t2b03g98uh92oeryugbs0e59y72nb09348tv02934vthm0")



        ' Open the DB
        If OpenDatabaseConnectionSQLServer(False) = False Then

            CloseDatabaseConnection()
            Console.WriteLine("Unable to connect to database. Please try again.", "Error")

        Else

            My.Computer.FileSystem.WriteAllText("dbcred.txt", clsEncrypt.EncryptData(strConnectionUsername + "|" + strConnectionPassword), False)
            CloseDatabaseConnection()
            OpenFormKillParent(Me, frmMain)

        End If

    End Sub
End Class