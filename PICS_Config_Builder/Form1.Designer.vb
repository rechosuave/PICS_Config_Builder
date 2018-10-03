<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnDataAndRun = New System.Windows.Forms.Button()
        Me.BtnClearAllSheets = New System.Windows.Forms.Button()
        Me.BtnExit = New System.Windows.Forms.Button()
        Me.CPU_PREFIX = New System.Windows.Forms.TextBox()
        Me.PleaseWaitForm = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(56, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Topic Name:"
        '
        'BtnDataAndRun
        '
        Me.BtnDataAndRun.Location = New System.Drawing.Point(77, 139)
        Me.BtnDataAndRun.Name = "BtnDataAndRun"
        Me.BtnDataAndRun.Size = New System.Drawing.Size(137, 75)
        Me.BtnDataAndRun.TabIndex = 1
        Me.BtnDataAndRun.Text = " Select IO Sheet " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "   &&" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Generate Output"
        Me.BtnDataAndRun.UseVisualStyleBackColor = True
        '
        'BtnClearAllSheets
        '
        Me.BtnClearAllSheets.Location = New System.Drawing.Point(290, 139)
        Me.BtnClearAllSheets.Name = "BtnClearAllSheets"
        Me.BtnClearAllSheets.Size = New System.Drawing.Size(137, 75)
        Me.BtnClearAllSheets.TabIndex = 2
        Me.BtnClearAllSheets.Text = "Clear All Sheets"
        Me.BtnClearAllSheets.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Location = New System.Drawing.Point(339, 266)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(75, 23)
        Me.BtnExit.TabIndex = 3
        Me.BtnExit.Text = "Exit"
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'CPU_PREFIX
        '
        Me.CPU_PREFIX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CPU_PREFIX.Location = New System.Drawing.Point(159, 65)
        Me.CPU_PREFIX.Name = "CPU_PREFIX"
        Me.CPU_PREFIX.Size = New System.Drawing.Size(221, 22)
        Me.CPU_PREFIX.TabIndex = 4
        '
        'PleaseWaitForm
        '
        Me.PleaseWaitForm.AutoSize = True
        Me.PleaseWaitForm.BackColor = System.Drawing.Color.Coral
        Me.PleaseWaitForm.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PleaseWaitForm.Location = New System.Drawing.Point(194, 235)
        Me.PleaseWaitForm.Name = "PleaseWaitForm"
        Me.PleaseWaitForm.Size = New System.Drawing.Size(119, 20)
        Me.PleaseWaitForm.TabIndex = 5
        Me.PleaseWaitForm.Text = "Please Wait..."
        Me.PleaseWaitForm.UseWaitCursor = True
        Me.PleaseWaitForm.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(473, 307)
        Me.Controls.Add(Me.PleaseWaitForm)
        Me.Controls.Add(Me.CPU_PREFIX)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.BtnClearAllSheets)
        Me.Controls.Add(Me.BtnDataAndRun)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form1"
        Me.Text = "PICS Config Builder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BtnDataAndRun As System.Windows.Forms.Button
    Friend WithEvents BtnClearAllSheets As System.Windows.Forms.Button
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents CPU_PREFIX As System.Windows.Forms.TextBox
    Friend WithEvents PleaseWaitForm As Label
End Class
