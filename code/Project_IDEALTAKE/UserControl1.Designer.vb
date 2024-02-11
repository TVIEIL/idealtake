<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserControl1
    Inherits System.Windows.Forms.UserControl

    'UserControl remplace la méthode Dispose pour nettoyer la liste des composants.
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
        Me.components = New System.ComponentModel.Container
        Me.Effacer_les_Références = New System.Windows.Forms.Button
        Me.Liste_des_références_Menu_Strip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Liste_des_références_TextBox = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Effacer_les_Références
        '
        Me.Effacer_les_Références.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Effacer_les_Références.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Effacer_les_Références.ForeColor = System.Drawing.Color.Red
        Me.Effacer_les_Références.Location = New System.Drawing.Point(153, -1)
        Me.Effacer_les_Références.Name = "Effacer_les_Références"
        Me.Effacer_les_Références.Size = New System.Drawing.Size(23, 27)
        Me.Effacer_les_Références.TabIndex = 0
        Me.Effacer_les_Références.Text = "x"
        Me.Effacer_les_Références.UseVisualStyleBackColor = False
        '
        'Liste_des_références_Menu_Strip
        '
        Me.Liste_des_références_Menu_Strip.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Table
        Me.Liste_des_références_Menu_Strip.Name = "Liste_des_références_Menu_Strip"
        Me.Liste_des_références_Menu_Strip.Size = New System.Drawing.Size(61, 4)
        '
        'Liste_des_références_TextBox
        '
        Me.Liste_des_références_TextBox.ContextMenuStrip = Me.Liste_des_références_Menu_Strip
        Me.Liste_des_références_TextBox.Location = New System.Drawing.Point(5, 3)
        Me.Liste_des_références_TextBox.Name = "Liste_des_références_TextBox"
        Me.Liste_des_références_TextBox.Size = New System.Drawing.Size(143, 20)
        Me.Liste_des_références_TextBox.TabIndex = 2
        '
        'UserControl1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Effacer_les_Références)
        Me.Controls.Add(Me.Liste_des_références_TextBox)
        Me.Name = "UserControl1"
        Me.Size = New System.Drawing.Size(180, 25)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Effacer_les_Références As System.Windows.Forms.Button
    Friend WithEvents Liste_des_références_Menu_Strip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents Liste_des_références_TextBox As System.Windows.Forms.TextBox

    Private Sub Liste_des_références_Menu_Strip_Closed(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosedEventArgs) Handles Liste_des_références_Menu_Strip.Closed
        Liste_des_références_TextBox.Text = Liste_des_références_TextBox.Text & My.Settings.RA_Référence_à_ajouter
        My.Settings.RA_Référence_à_ajouter = ""

    End Sub
End Class
