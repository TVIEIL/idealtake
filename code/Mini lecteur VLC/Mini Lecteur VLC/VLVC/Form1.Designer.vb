<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ButtonPlay = New System.Windows.Forms.Button()
        Me.ButtonPause = New System.Windows.Forms.Button()
        Me.ButtonStop = New System.Windows.Forms.Button()
        Me.ButtonMute = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ImageListButton = New System.Windows.Forms.ImageList(Me.components)
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonPlay
        '
        Me.ButtonPlay.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonPlay.ImageIndex = 0
        Me.ButtonPlay.ImageList = Me.ImageListButton
        Me.ButtonPlay.Location = New System.Drawing.Point(5, 12)
        Me.ButtonPlay.Name = "ButtonPlay"
        Me.ButtonPlay.Size = New System.Drawing.Size(123, 23)
        Me.ButtonPlay.TabIndex = 0
        Me.ButtonPlay.Text = "Play"
        Me.ButtonPlay.UseVisualStyleBackColor = True
        '
        'ButtonPause
        '
        Me.ButtonPause.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonPause.ImageIndex = 1
        Me.ButtonPause.ImageList = Me.ImageListButton
        Me.ButtonPause.Location = New System.Drawing.Point(5, 41)
        Me.ButtonPause.Name = "ButtonPause"
        Me.ButtonPause.Size = New System.Drawing.Size(123, 23)
        Me.ButtonPause.TabIndex = 1
        Me.ButtonPause.Text = "Pause"
        Me.ButtonPause.UseVisualStyleBackColor = True
        '
        'ButtonStop
        '
        Me.ButtonStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonStop.ImageIndex = 2
        Me.ButtonStop.ImageList = Me.ImageListButton
        Me.ButtonStop.Location = New System.Drawing.Point(5, 70)
        Me.ButtonStop.Name = "ButtonStop"
        Me.ButtonStop.Size = New System.Drawing.Size(123, 23)
        Me.ButtonStop.TabIndex = 2
        Me.ButtonStop.Text = "Stop"
        Me.ButtonStop.UseVisualStyleBackColor = True
        '
        'ButtonMute
        '
        Me.ButtonMute.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonMute.ImageIndex = 3
        Me.ButtonMute.ImageList = Me.ImageListButton
        Me.ButtonMute.Location = New System.Drawing.Point(5, 99)
        Me.ButtonMute.Name = "ButtonMute"
        Me.ButtonMute.Size = New System.Drawing.Size(123, 23)
        Me.ButtonMute.TabIndex = 3
        Me.ButtonMute.Text = "Mute"
        Me.ButtonMute.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 137)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(135, 22)
        Me.StatusStrip1.TabIndex = 4
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(121, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'ImageListButton
        '
        Me.ImageListButton.ImageStream = CType(resources.GetObject("ImageListButton.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListButton.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageListButton.Images.SetKeyName(0, "Lecture.png")
        Me.ImageListButton.Images.SetKeyName(1, "Pause.png")
        Me.ImageListButton.Images.SetKeyName(2, "Stop.png")
        Me.ImageListButton.Images.SetKeyName(3, "Muet.png")
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(135, 159)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.ButtonMute)
        Me.Controls.Add(Me.ButtonStop)
        Me.Controls.Add(Me.ButtonPause)
        Me.Controls.Add(Me.ButtonPlay)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Mini VLC Player"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonPlay As System.Windows.Forms.Button
    Friend WithEvents ButtonPause As System.Windows.Forms.Button
    Friend WithEvents ButtonStop As System.Windows.Forms.Button
    Friend WithEvents ButtonMute As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ImageListButton As System.Windows.Forms.ImageList

End Class
