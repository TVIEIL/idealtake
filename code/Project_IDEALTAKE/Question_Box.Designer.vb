<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Message_Box_Question
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
        Me.NonButton = New System.Windows.Forms.Button()
        Me.OuiButton = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'NonButton
        '
        Me.NonButton.BackColor = System.Drawing.SystemColors.Control
        Me.NonButton.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.NonButton.DialogResult = System.Windows.Forms.DialogResult.No
        Me.NonButton.Dock = System.Windows.Forms.DockStyle.Right
        Me.NonButton.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.NonButton.Location = New System.Drawing.Point(115, 0)
        Me.NonButton.Name = "NonButton"
        Me.NonButton.Size = New System.Drawing.Size(85, 28)
        Me.NonButton.TabIndex = 3
        Me.NonButton.Text = "Non"
        Me.NonButton.UseVisualStyleBackColor = True
        '
        'OuiButton
        '
        Me.OuiButton.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.OuiButton.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.OuiButton.Dock = System.Windows.Forms.DockStyle.Left
        Me.OuiButton.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.OuiButton.Location = New System.Drawing.Point(0, 0)
        Me.OuiButton.Name = "OuiButton"
        Me.OuiButton.Size = New System.Drawing.Size(85, 28)
        Me.OuiButton.TabIndex = 2
        Me.OuiButton.Text = "Oui"
        Me.OuiButton.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ListBox1.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.ListBox1.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Items.AddRange(New Object() {"", "", "    Question?"})
        Me.ListBox1.Location = New System.Drawing.Point(0, 0)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.ListBox1.Size = New System.Drawing.Size(253, 65)
        Me.ListBox1.TabIndex = 4
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.NonButton)
        Me.Panel1.Controls.Add(Me.OuiButton)
        Me.Panel1.Location = New System.Drawing.Point(31, 9)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 28)
        Me.Panel1.TabIndex = 5
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(9, 104)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(234, 40)
        Me.TableLayoutPanel1.TabIndex = 6
        '
        'Message_Box_Question
        '
        Me.AcceptButton = Me.OuiButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.NonButton
        Me.ClientSize = New System.Drawing.Size(252, 153)
        Me.ControlBox = False
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.ListBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Message_Box_Question"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.Text = "Titre"
        Me.Panel1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents NonButton As System.Windows.Forms.Button
    Friend WithEvents OuiButton As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel

End Class
