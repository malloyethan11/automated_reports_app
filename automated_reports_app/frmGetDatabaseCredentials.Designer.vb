﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGetDatabaseCredentials
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.btnLogin = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(66, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(231, 28)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Enter Database Credentials:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblUsername
        '
        Me.lblUsername.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Location = New System.Drawing.Point(39, 66)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(58, 13)
        Me.lblUsername.TabIndex = 23
        Me.lblUsername.Text = "Username:"
        '
        'txtUsername
        '
        Me.txtUsername.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.txtUsername.Location = New System.Drawing.Point(97, 63)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(200, 20)
        Me.txtUsername.TabIndex = 18
        '
        'lblPassword
        '
        Me.lblPassword.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Location = New System.Drawing.Point(41, 102)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(56, 13)
        Me.lblPassword.TabIndex = 21
        Me.lblPassword.Text = "Password:"
        '
        'txtPassword
        '
        Me.txtPassword.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.txtPassword.Location = New System.Drawing.Point(97, 99)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(200, 20)
        Me.txtPassword.TabIndex = 19
        '
        'btnLogin
        '
        Me.btnLogin.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnLogin.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnLogin.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLogin.Location = New System.Drawing.Point(27, 135)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(284, 42)
        Me.btnLogin.TabIndex = 20
        Me.btnLogin.Text = "Connect"
        Me.btnLogin.UseVisualStyleBackColor = True
        '
        'frmGetDatabaseCredentials
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(368, 354)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.btnLogin)
        Me.Name = "frmGetDatabaseCredentials"
        Me.Text = "frmGetDatabaseCredentials"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents btnLogin As System.Windows.Forms.Button
End Class
