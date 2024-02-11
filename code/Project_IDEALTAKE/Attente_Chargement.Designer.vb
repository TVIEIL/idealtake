<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Attente_Chargement
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
    Friend WithEvents ApplicationTitle As System.Windows.Forms.Label
    Friend WithEvents Table_à_charger As System.Windows.Forms.Label
    Friend WithEvents MainLayoutPanel As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents DetailsLayoutPanel As System.Windows.Forms.TableLayoutPanel

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MainLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.DetailsLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.Etat_Chargement = New System.Windows.Forms.Label()
        Me.Table_à_charger = New System.Windows.Forms.Label()
        Me.ApplicationTitle = New System.Windows.Forms.Label()
        Me.MainLayoutPanel.SuspendLayout()
        Me.DetailsLayoutPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainLayoutPanel
        '
        Me.MainLayoutPanel.BackColor = System.Drawing.SystemColors.Control
        Me.MainLayoutPanel.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Pict0017j
        Me.MainLayoutPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.MainLayoutPanel.ColumnCount = 1
        Me.MainLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.MainLayoutPanel.Controls.Add(Me.DetailsLayoutPanel, 0, 1)
        Me.MainLayoutPanel.Controls.Add(Me.ApplicationTitle, 0, 0)
        Me.MainLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MainLayoutPanel.Location = New System.Drawing.Point(0, 0)
        Me.MainLayoutPanel.Name = "MainLayoutPanel"
        Me.MainLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 34.76395!))
        Me.MainLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 65.23605!))
        Me.MainLayoutPanel.Size = New System.Drawing.Size(262, 106)
        Me.MainLayoutPanel.TabIndex = 0
        '
        'DetailsLayoutPanel
        '
        Me.DetailsLayoutPanel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DetailsLayoutPanel.BackColor = System.Drawing.Color.LightCyan
        Me.DetailsLayoutPanel.ColumnCount = 1
        Me.DetailsLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 372.0!))
        Me.DetailsLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 142.0!))
        Me.DetailsLayoutPanel.Controls.Add(Me.Etat_Chargement, 0, 1)
        Me.DetailsLayoutPanel.Controls.Add(Me.Table_à_charger, 0, 0)
        Me.DetailsLayoutPanel.Location = New System.Drawing.Point(3, 39)
        Me.DetailsLayoutPanel.Name = "DetailsLayoutPanel"
        Me.DetailsLayoutPanel.RowCount = 2
        Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 39.13044!))
        Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 60.86956!))
        Me.DetailsLayoutPanel.Size = New System.Drawing.Size(256, 64)
        Me.DetailsLayoutPanel.TabIndex = 1
        '
        'Etat_Chargement
        '
        Me.Etat_Chargement.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Etat_Chargement.BackColor = System.Drawing.Color.Transparent
        Me.Etat_Chargement.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Etat_Chargement.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Etat_Chargement.Location = New System.Drawing.Point(51, 31)
        Me.Etat_Chargement.Name = "Etat_Chargement"
        Me.Etat_Chargement.Size = New System.Drawing.Size(269, 27)
        Me.Etat_Chargement.TabIndex = 4
        Me.Etat_Chargement.Text = "- - -"
        '
        'Table_à_charger
        '
        Me.Table_à_charger.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Table_à_charger.BackColor = System.Drawing.Color.Transparent
        Me.Table_à_charger.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Table_à_charger.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Table_à_charger.Location = New System.Drawing.Point(54, 0)
        Me.Table_à_charger.Name = "Table_à_charger"
        Me.Table_à_charger.Size = New System.Drawing.Size(263, 25)
        Me.Table_à_charger.TabIndex = 1
        Me.Table_à_charger.Text = "- - -"
        '
        'ApplicationTitle
        '
        Me.ApplicationTitle.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.ApplicationTitle.BackColor = System.Drawing.Color.Transparent
        Me.ApplicationTitle.Font = New System.Drawing.Font("Courier New", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ApplicationTitle.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.ApplicationTitle.Image = Global.IDEALTAKE.My.Resources.Resources.new_icone
        Me.ApplicationTitle.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ApplicationTitle.Location = New System.Drawing.Point(19, 0)
        Me.ApplicationTitle.Name = "ApplicationTitle"
        Me.ApplicationTitle.Size = New System.Drawing.Size(223, 36)
        Me.ApplicationTitle.TabIndex = 0
        Me.ApplicationTitle.Text = "IDEALTAKE"
        Me.ApplicationTitle.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Attente_Chargement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(262, 106)
        Me.ControlBox = False
        Me.Controls.Add(Me.MainLayoutPanel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Attente_Chargement"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.MainLayoutPanel.ResumeLayout(False)
        Me.DetailsLayoutPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Etat_Chargement As System.Windows.Forms.Label

End Class
