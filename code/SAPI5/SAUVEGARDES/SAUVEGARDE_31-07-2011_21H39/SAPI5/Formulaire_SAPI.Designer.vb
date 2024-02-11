<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Formulaire_SAPI
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Formulaire_SAPI))
        Me.Label_Valeur_volume = New System.Windows.Forms.Label()
        Me.Label_Valeur_vitesse = New System.Windows.Forms.Label()
        Me.Button_Volume_moin = New System.Windows.Forms.Button()
        Me.Button_Volume_plus = New System.Windows.Forms.Button()
        Me.Button_Rate_moin = New System.Windows.Forms.Button()
        Me.Button_Rate_plus = New System.Windows.Forms.Button()
        Me.Label_Volume = New System.Windows.Forms.Label()
        Me.Label_Rate = New System.Windows.Forms.Label()
        Me.idsAudioOutput = New System.Windows.Forms.ComboBox()
        Me.idsVoices = New System.Windows.Forms.ComboBox()
        Me.Label_Audio_Output = New System.Windows.Forms.Label()
        Me.Label_Voice = New System.Windows.Forms.Label()
        Me.idSpeakText = New System.Windows.Forms.Button()
        Me.idTextBox = New System.Windows.Forms.TextBox()
        Me.Timer_Lecture_instructions_Synthèse_vocale = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Label_Valeur_volume
        '
        Me.Label_Valeur_volume.AutoSize = True
        Me.Label_Valeur_volume.Location = New System.Drawing.Point(260, 185)
        Me.Label_Valeur_volume.Name = "Label_Valeur_volume"
        Me.Label_Valeur_volume.Size = New System.Drawing.Size(22, 13)
        Me.Label_Valeur_volume.TabIndex = 47
        Me.Label_Valeur_volume.Text = "- - -"
        '
        'Label_Valeur_vitesse
        '
        Me.Label_Valeur_vitesse.AutoSize = True
        Me.Label_Valeur_vitesse.Location = New System.Drawing.Point(105, 185)
        Me.Label_Valeur_vitesse.Name = "Label_Valeur_vitesse"
        Me.Label_Valeur_vitesse.Size = New System.Drawing.Size(16, 13)
        Me.Label_Valeur_vitesse.TabIndex = 46
        Me.Label_Valeur_vitesse.Text = "- -"
        '
        'Button_Volume_moin
        '
        Me.Button_Volume_moin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Volume_moin.Location = New System.Drawing.Point(251, 200)
        Me.Button_Volume_moin.Name = "Button_Volume_moin"
        Me.Button_Volume_moin.Size = New System.Drawing.Size(23, 23)
        Me.Button_Volume_moin.TabIndex = 45
        Me.Button_Volume_moin.Text = "-"
        Me.Button_Volume_moin.UseVisualStyleBackColor = True
        '
        'Button_Volume_plus
        '
        Me.Button_Volume_plus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Volume_plus.Location = New System.Drawing.Point(221, 200)
        Me.Button_Volume_plus.Name = "Button_Volume_plus"
        Me.Button_Volume_plus.Size = New System.Drawing.Size(23, 23)
        Me.Button_Volume_plus.TabIndex = 44
        Me.Button_Volume_plus.Text = "+"
        Me.Button_Volume_plus.UseVisualStyleBackColor = True
        '
        'Button_Rate_moin
        '
        Me.Button_Rate_moin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Rate_moin.Location = New System.Drawing.Point(73, 201)
        Me.Button_Rate_moin.Name = "Button_Rate_moin"
        Me.Button_Rate_moin.Size = New System.Drawing.Size(23, 23)
        Me.Button_Rate_moin.TabIndex = 43
        Me.Button_Rate_moin.Text = "-"
        Me.Button_Rate_moin.UseVisualStyleBackColor = True
        '
        'Button_Rate_plus
        '
        Me.Button_Rate_plus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Rate_plus.Location = New System.Drawing.Point(44, 201)
        Me.Button_Rate_plus.Name = "Button_Rate_plus"
        Me.Button_Rate_plus.Size = New System.Drawing.Size(23, 23)
        Me.Button_Rate_plus.TabIndex = 42
        Me.Button_Rate_plus.Text = "+"
        Me.Button_Rate_plus.UseVisualStyleBackColor = True
        '
        'Label_Volume
        '
        Me.Label_Volume.AutoSize = True
        Me.Label_Volume.Location = New System.Drawing.Point(206, 184)
        Me.Label_Volume.Name = "Label_Volume"
        Me.Label_Volume.Size = New System.Drawing.Size(48, 13)
        Me.Label_Volume.TabIndex = 41
        Me.Label_Volume.Text = "Volume :"
        '
        'Label_Rate
        '
        Me.Label_Rate.AutoSize = True
        Me.Label_Rate.Location = New System.Drawing.Point(12, 185)
        Me.Label_Rate.Name = "Label_Rate"
        Me.Label_Rate.Size = New System.Drawing.Size(36, 13)
        Me.Label_Rate.TabIndex = 40
        Me.Label_Rate.Text = "Rate :"
        '
        'idsAudioOutput
        '
        Me.idsAudioOutput.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.idsAudioOutput.FormattingEnabled = True
        Me.idsAudioOutput.Location = New System.Drawing.Point(12, 160)
        Me.idsAudioOutput.Margin = New System.Windows.Forms.Padding(2)
        Me.idsAudioOutput.Name = "idsAudioOutput"
        Me.idsAudioOutput.Size = New System.Drawing.Size(298, 21)
        Me.idsAudioOutput.TabIndex = 39
        '
        'idsVoices
        '
        Me.idsVoices.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.idsVoices.FormattingEnabled = True
        Me.idsVoices.Location = New System.Drawing.Point(12, 119)
        Me.idsVoices.Margin = New System.Windows.Forms.Padding(2)
        Me.idsVoices.Name = "idsVoices"
        Me.idsVoices.Size = New System.Drawing.Size(298, 21)
        Me.idsVoices.TabIndex = 38
        '
        'Label_Audio_Output
        '
        Me.Label_Audio_Output.AutoSize = True
        Me.Label_Audio_Output.Location = New System.Drawing.Point(12, 145)
        Me.Label_Audio_Output.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label_Audio_Output.Name = "Label_Audio_Output"
        Me.Label_Audio_Output.Size = New System.Drawing.Size(69, 13)
        Me.Label_Audio_Output.TabIndex = 37
        Me.Label_Audio_Output.Text = "Audio Output"
        '
        'Label_Voice
        '
        Me.Label_Voice.AutoSize = True
        Me.Label_Voice.Location = New System.Drawing.Point(12, 104)
        Me.Label_Voice.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label_Voice.Name = "Label_Voice"
        Me.Label_Voice.Size = New System.Drawing.Size(34, 13)
        Me.Label_Voice.TabIndex = 36
        Me.Label_Voice.Text = "Voice"
        '
        'idSpeakText
        '
        Me.idSpeakText.Location = New System.Drawing.Point(84, 72)
        Me.idSpeakText.Margin = New System.Windows.Forms.Padding(2)
        Me.idSpeakText.Name = "idSpeakText"
        Me.idSpeakText.Size = New System.Drawing.Size(131, 24)
        Me.idSpeakText.TabIndex = 35
        Me.idSpeakText.Text = "Speak Text"
        Me.idSpeakText.UseVisualStyleBackColor = True
        '
        'idTextBox
        '
        Me.idTextBox.Location = New System.Drawing.Point(12, 6)
        Me.idTextBox.Margin = New System.Windows.Forms.Padding(2)
        Me.idTextBox.Multiline = True
        Me.idTextBox.Name = "idTextBox"
        Me.idTextBox.Size = New System.Drawing.Size(298, 62)
        Me.idTextBox.TabIndex = 34
        Me.idTextBox.Text = "Enter text you wish spoken here"
        '
        'Timer_Lecture_instructions_Synthèse_vocale
        '
        Me.Timer_Lecture_instructions_Synthèse_vocale.Enabled = True
        Me.Timer_Lecture_instructions_Synthèse_vocale.Interval = 800
        '
        'Formulaire_SAPI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(323, 230)
        Me.Controls.Add(Me.Label_Valeur_volume)
        Me.Controls.Add(Me.Label_Valeur_vitesse)
        Me.Controls.Add(Me.Button_Volume_moin)
        Me.Controls.Add(Me.Button_Volume_plus)
        Me.Controls.Add(Me.Button_Rate_moin)
        Me.Controls.Add(Me.Button_Rate_plus)
        Me.Controls.Add(Me.Label_Volume)
        Me.Controls.Add(Me.Label_Rate)
        Me.Controls.Add(Me.idsAudioOutput)
        Me.Controls.Add(Me.idsVoices)
        Me.Controls.Add(Me.Label_Audio_Output)
        Me.Controls.Add(Me.Label_Voice)
        Me.Controls.Add(Me.idSpeakText)
        Me.Controls.Add(Me.idTextBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Formulaire_SAPI"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label_Valeur_volume As System.Windows.Forms.Label
    Friend WithEvents Label_Valeur_vitesse As System.Windows.Forms.Label
    Friend WithEvents Button_Volume_moin As System.Windows.Forms.Button
    Friend WithEvents Button_Volume_plus As System.Windows.Forms.Button
    Friend WithEvents Button_Rate_moin As System.Windows.Forms.Button
    Friend WithEvents Button_Rate_plus As System.Windows.Forms.Button
    Friend WithEvents Label_Volume As System.Windows.Forms.Label
    Friend WithEvents Label_Rate As System.Windows.Forms.Label
    Friend WithEvents idsAudioOutput As System.Windows.Forms.ComboBox
    Friend WithEvents idsVoices As System.Windows.Forms.ComboBox
    Friend WithEvents Label_Audio_Output As System.Windows.Forms.Label
    Friend WithEvents Label_Voice As System.Windows.Forms.Label
    Friend WithEvents idSpeakText As System.Windows.Forms.Button
    Friend WithEvents idTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Timer_Lecture_instructions_Synthèse_vocale As System.Windows.Forms.Timer

End Class
