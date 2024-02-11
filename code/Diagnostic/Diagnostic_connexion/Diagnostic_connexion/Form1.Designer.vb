<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Diagnostic
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_Diagnostic))
        Me.TableLayoutPanel_Diagnostic = New System.Windows.Forms.TableLayoutPanel()
        Me.ProgressBar_Diagnostic = New System.Windows.Forms.ProgressBar()
        Me.RichTextBox_Diagnostic = New System.Windows.Forms.RichTextBox()
        Me.Tâche_de_fond = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker_Diagnostic = New System.ComponentModel.BackgroundWorker()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TableLayoutPanel_Diagnostic.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel_Diagnostic
        '
        Me.TableLayoutPanel_Diagnostic.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel_Diagnostic.ColumnCount = 1
        Me.TableLayoutPanel_Diagnostic.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel_Diagnostic.Controls.Add(Me.ProgressBar_Diagnostic, 0, 1)
        Me.TableLayoutPanel_Diagnostic.Controls.Add(Me.RichTextBox_Diagnostic, 0, 0)
        Me.TableLayoutPanel_Diagnostic.Location = New System.Drawing.Point(1, 3)
        Me.TableLayoutPanel_Diagnostic.Name = "TableLayoutPanel_Diagnostic"
        Me.TableLayoutPanel_Diagnostic.RowCount = 2
        Me.TableLayoutPanel_Diagnostic.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel_Diagnostic.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel_Diagnostic.Size = New System.Drawing.Size(640, 405)
        Me.TableLayoutPanel_Diagnostic.TabIndex = 0
        '
        'ProgressBar_Diagnostic
        '
        Me.ProgressBar_Diagnostic.Dock = System.Windows.Forms.DockStyle.Top
        Me.ProgressBar_Diagnostic.Location = New System.Drawing.Point(3, 373)
        Me.ProgressBar_Diagnostic.Maximum = 75
        Me.ProgressBar_Diagnostic.Name = "ProgressBar_Diagnostic"
        Me.ProgressBar_Diagnostic.Size = New System.Drawing.Size(634, 25)
        Me.ProgressBar_Diagnostic.TabIndex = 0
        '
        'RichTextBox_Diagnostic
        '
        Me.RichTextBox_Diagnostic.BackColor = System.Drawing.Color.White
        Me.RichTextBox_Diagnostic.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox_Diagnostic.Location = New System.Drawing.Point(3, 3)
        Me.RichTextBox_Diagnostic.Name = "RichTextBox_Diagnostic"
        Me.RichTextBox_Diagnostic.ReadOnly = True
        Me.RichTextBox_Diagnostic.Size = New System.Drawing.Size(634, 364)
        Me.RichTextBox_Diagnostic.TabIndex = 1
        Me.RichTextBox_Diagnostic.Text = ""
        '
        'Tâche_de_fond
        '
        Me.Tâche_de_fond.WorkerReportsProgress = True
        Me.Tâche_de_fond.WorkerSupportsCancellation = True
        '
        'BackgroundWorker_Diagnostic
        '
        Me.BackgroundWorker_Diagnostic.WorkerReportsProgress = True
        Me.BackgroundWorker_Diagnostic.WorkerSupportsCancellation = True
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1500
        '
        'Form_Diagnostic
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(641, 405)
        Me.Controls.Add(Me.TableLayoutPanel_Diagnostic)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form_Diagnostic"
        Me.Text = "IDEALTAKE"
        Me.TableLayoutPanel_Diagnostic.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel_Diagnostic As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents ProgressBar_Diagnostic As System.Windows.Forms.ProgressBar
    Friend WithEvents RichTextBox_Diagnostic As System.Windows.Forms.RichTextBox
    Friend WithEvents Tâche_de_fond As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker_Diagnostic As System.ComponentModel.BackgroundWorker
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
