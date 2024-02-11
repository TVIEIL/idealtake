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
        Me.LabelChrono = New System.Windows.Forms.Label()
        Me.LabelNuméroChrono = New System.Windows.Forms.Label()
        Me.LabelNuméroEnregistrement = New System.Windows.Forms.Label()
        Me.LabelEnregistrement = New System.Windows.Forms.Label()
        Me.PictureBoxGermanFlag = New System.Windows.Forms.PictureBox()
        Me.PictureBoxEnglishgFlag = New System.Windows.Forms.PictureBox()
        Me.PictureBoxSpanishFlag = New System.Windows.Forms.PictureBox()
        Me.PictureBoxFrenchFlag = New System.Windows.Forms.PictureBox()
        Me.PictureBoxItalianFlag = New System.Windows.Forms.PictureBox()
        Me.ButtonAddDocument = New System.Windows.Forms.Button()
        Me.ComboBoxTypeDocumentMaintenance = New System.Windows.Forms.ComboBox()
        Me.RadioButtonFTPMaintenance = New System.Windows.Forms.RadioButton()
        Me.RadioButtonTransfertLocalMaintenance = New System.Windows.Forms.RadioButton()
        Me.LabelUserMaintenance = New System.Windows.Forms.Label()
        Me.TextBoxUserNameMaintenance = New System.Windows.Forms.TextBox()
        Me.LabelPasswordMaintenance = New System.Windows.Forms.Label()
        Me.TextBoxPassword = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBoxMaintenance = New System.Windows.Forms.GroupBox()
        Me.ButtonDestinationLocalMaintenance = New System.Windows.Forms.Button()
        Me.LabelDestinationMaintenance = New System.Windows.Forms.Label()
        Me.TextBoxDestinationMaintenance = New System.Windows.Forms.TextBox()
        Me.ButtonMiseàJourMaintenance = New System.Windows.Forms.Button()
        Me.LabelTypeDocumentMaintenance = New System.Windows.Forms.Label()
        Me.LabelDataBaseName = New System.Windows.Forms.Label()
        Me.LabelNomDeLaBase = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBoxInfosDocument = New System.Windows.Forms.GroupBox()
        Me.RadioButtonImagePreviewITA = New System.Windows.Forms.RadioButton()
        Me.RadioButtonImagePreviewFRA = New System.Windows.Forms.RadioButton()
        Me.RadioButtonImagePreviewESP = New System.Windows.Forms.RadioButton()
        Me.RadioButtonImagePreviewANG = New System.Windows.Forms.RadioButton()
        Me.RadioButtonImagePreviewALL = New System.Windows.Forms.RadioButton()
        Me.ComboBoxTypeDocument = New System.Windows.Forms.ComboBox()
        Me.LabelTypeDocDocument = New System.Windows.Forms.Label()
        Me.TextBoxNomFichier = New System.Windows.Forms.TextBox()
        Me.LabelNomFichier = New System.Windows.Forms.Label()
        Me.TextBoxNomDuLienDuFichier = New System.Windows.Forms.TextBox()
        Me.LabelNomLienDuFichier = New System.Windows.Forms.Label()
        Me.LabelNomFichierOrigine = New System.Windows.Forms.Label()
        Me.ButtonSelectFile = New System.Windows.Forms.Button()
        Me.ListBoxDocuments = New System.Windows.Forms.ListBox()
        Me.LabelListeDesDocuments = New System.Windows.Forms.Label()
        Me.ButtonEffaceUnDocument = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BackgroundWorkerChargementTypeDocument = New System.ComponentModel.BackgroundWorker()
        Me.TimerChargementTypeDocument = New System.Windows.Forms.Timer(Me.components)
        Me.ButtonNouveauDocument = New System.Windows.Forms.Button()
        Me.BackgroundWorkerChargementListeDocuments = New System.ComponentModel.BackgroundWorker()
        Me.TimerChargementListeDocuments = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker_MaintenanceTransfert = New System.ComponentModel.BackgroundWorker()
        Me.TimerEtatMaintenanceTransfert = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker_SupprimeDocument = New System.ComponentModel.BackgroundWorker()
        Me.TimerEtatSupprimeDocument = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker_AjoutDocument = New System.ComponentModel.BackgroundWorker()
        Me.TimerEtatAjoutDocument = New System.Windows.Forms.Timer(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CheckBoxActivationDetail = New System.Windows.Forms.CheckBox()
        CType(Me.PictureBoxGermanFlag, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxEnglishgFlag, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxSpanishFlag, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxFrenchFlag, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxItalianFlag, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxMaintenance.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBoxInfosDocument.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'LabelChrono
        '
        Me.LabelChrono.AutoSize = True
        Me.LabelChrono.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelChrono.Location = New System.Drawing.Point(12, 56)
        Me.LabelChrono.Name = "LabelChrono"
        Me.LabelChrono.Size = New System.Drawing.Size(112, 13)
        Me.LabelChrono.TabIndex = 0
        Me.LabelChrono.Text = "Numéro de Chrono"
        '
        'LabelNuméroChrono
        '
        Me.LabelNuméroChrono.AutoSize = True
        Me.LabelNuméroChrono.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNuméroChrono.Location = New System.Drawing.Point(216, 56)
        Me.LabelNuméroChrono.Name = "LabelNuméroChrono"
        Me.LabelNuméroChrono.Size = New System.Drawing.Size(22, 15)
        Me.LabelNuméroChrono.TabIndex = 1
        Me.LabelNuméroChrono.Text = "---"
        '
        'LabelNuméroEnregistrement
        '
        Me.LabelNuméroEnregistrement.AutoSize = True
        Me.LabelNuméroEnregistrement.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNuméroEnregistrement.Location = New System.Drawing.Point(216, 34)
        Me.LabelNuméroEnregistrement.Name = "LabelNuméroEnregistrement"
        Me.LabelNuméroEnregistrement.Size = New System.Drawing.Size(22, 15)
        Me.LabelNuméroEnregistrement.TabIndex = 3
        Me.LabelNuméroEnregistrement.Text = "---"
        '
        'LabelEnregistrement
        '
        Me.LabelEnregistrement.AutoSize = True
        Me.LabelEnregistrement.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelEnregistrement.Location = New System.Drawing.Point(12, 34)
        Me.LabelEnregistrement.Name = "LabelEnregistrement"
        Me.LabelEnregistrement.Size = New System.Drawing.Size(147, 13)
        Me.LabelEnregistrement.TabIndex = 2
        Me.LabelEnregistrement.Text = "Numéro d'enregistrement"
        '
        'PictureBoxGermanFlag
        '
        Me.PictureBoxGermanFlag.Image = Global.TransfertIdealtake.My.Resources.Resources.Flag_of_Germany
        Me.PictureBoxGermanFlag.Location = New System.Drawing.Point(23, 80)
        Me.PictureBoxGermanFlag.Name = "PictureBoxGermanFlag"
        Me.PictureBoxGermanFlag.Size = New System.Drawing.Size(50, 33)
        Me.PictureBoxGermanFlag.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBoxGermanFlag.TabIndex = 4
        Me.PictureBoxGermanFlag.TabStop = False
        '
        'PictureBoxEnglishgFlag
        '
        Me.PictureBoxEnglishgFlag.Image = Global.TransfertIdealtake.My.Resources.Resources.Flag_of_the_United_States
        Me.PictureBoxEnglishgFlag.Location = New System.Drawing.Point(79, 80)
        Me.PictureBoxEnglishgFlag.Name = "PictureBoxEnglishgFlag"
        Me.PictureBoxEnglishgFlag.Size = New System.Drawing.Size(50, 33)
        Me.PictureBoxEnglishgFlag.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBoxEnglishgFlag.TabIndex = 5
        Me.PictureBoxEnglishgFlag.TabStop = False
        '
        'PictureBoxSpanishFlag
        '
        Me.PictureBoxSpanishFlag.Image = Global.TransfertIdealtake.My.Resources.Resources.Flag_of_Spain
        Me.PictureBoxSpanishFlag.Location = New System.Drawing.Point(135, 80)
        Me.PictureBoxSpanishFlag.Name = "PictureBoxSpanishFlag"
        Me.PictureBoxSpanishFlag.Size = New System.Drawing.Size(50, 33)
        Me.PictureBoxSpanishFlag.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBoxSpanishFlag.TabIndex = 6
        Me.PictureBoxSpanishFlag.TabStop = False
        '
        'PictureBoxFrenchFlag
        '
        Me.PictureBoxFrenchFlag.Image = Global.TransfertIdealtake.My.Resources.Resources.Flag_of_France
        Me.PictureBoxFrenchFlag.Location = New System.Drawing.Point(191, 80)
        Me.PictureBoxFrenchFlag.Name = "PictureBoxFrenchFlag"
        Me.PictureBoxFrenchFlag.Size = New System.Drawing.Size(50, 33)
        Me.PictureBoxFrenchFlag.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBoxFrenchFlag.TabIndex = 7
        Me.PictureBoxFrenchFlag.TabStop = False
        '
        'PictureBoxItalianFlag
        '
        Me.PictureBoxItalianFlag.Image = Global.TransfertIdealtake.My.Resources.Resources.Flag_of_Italy
        Me.PictureBoxItalianFlag.Location = New System.Drawing.Point(247, 80)
        Me.PictureBoxItalianFlag.Name = "PictureBoxItalianFlag"
        Me.PictureBoxItalianFlag.Size = New System.Drawing.Size(50, 33)
        Me.PictureBoxItalianFlag.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBoxItalianFlag.TabIndex = 8
        Me.PictureBoxItalianFlag.TabStop = False
        '
        'ButtonAddDocument
        '
        Me.ButtonAddDocument.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAddDocument.Location = New System.Drawing.Point(11, 279)
        Me.ButtonAddDocument.Name = "ButtonAddDocument"
        Me.ButtonAddDocument.Size = New System.Drawing.Size(304, 23)
        Me.ButtonAddDocument.TabIndex = 15
        Me.ButtonAddDocument.Text = "Transfert"
        Me.ButtonAddDocument.UseVisualStyleBackColor = True
        '
        'ComboBoxTypeDocumentMaintenance
        '
        Me.ComboBoxTypeDocumentMaintenance.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxTypeDocumentMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxTypeDocumentMaintenance.FormattingEnabled = True
        Me.ComboBoxTypeDocumentMaintenance.Location = New System.Drawing.Point(162, 11)
        Me.ComboBoxTypeDocumentMaintenance.Name = "ComboBoxTypeDocumentMaintenance"
        Me.ComboBoxTypeDocumentMaintenance.Size = New System.Drawing.Size(284, 23)
        Me.ComboBoxTypeDocumentMaintenance.TabIndex = 1
        '
        'RadioButtonFTPMaintenance
        '
        Me.RadioButtonFTPMaintenance.AutoSize = True
        Me.RadioButtonFTPMaintenance.Checked = True
        Me.RadioButtonFTPMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonFTPMaintenance.Location = New System.Drawing.Point(274, 66)
        Me.RadioButtonFTPMaintenance.Name = "RadioButtonFTPMaintenance"
        Me.RadioButtonFTPMaintenance.Size = New System.Drawing.Size(48, 17)
        Me.RadioButtonFTPMaintenance.TabIndex = 5
        Me.RadioButtonFTPMaintenance.TabStop = True
        Me.RadioButtonFTPMaintenance.Text = "FTP"
        Me.RadioButtonFTPMaintenance.UseVisualStyleBackColor = True
        '
        'RadioButtonTransfertLocalMaintenance
        '
        Me.RadioButtonTransfertLocalMaintenance.AutoSize = True
        Me.RadioButtonTransfertLocalMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonTransfertLocalMaintenance.Location = New System.Drawing.Point(348, 66)
        Me.RadioButtonTransfertLocalMaintenance.Name = "RadioButtonTransfertLocalMaintenance"
        Me.RadioButtonTransfertLocalMaintenance.Size = New System.Drawing.Size(56, 17)
        Me.RadioButtonTransfertLocalMaintenance.TabIndex = 6
        Me.RadioButtonTransfertLocalMaintenance.Text = "Local"
        Me.RadioButtonTransfertLocalMaintenance.UseVisualStyleBackColor = True
        '
        'LabelUserMaintenance
        '
        Me.LabelUserMaintenance.AutoSize = True
        Me.LabelUserMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelUserMaintenance.Location = New System.Drawing.Point(33, 69)
        Me.LabelUserMaintenance.Name = "LabelUserMaintenance"
        Me.LabelUserMaintenance.Size = New System.Drawing.Size(64, 13)
        Me.LabelUserMaintenance.TabIndex = 13
        Me.LabelUserMaintenance.Text = "Utilisateur"
        '
        'TextBoxUserNameMaintenance
        '
        Me.TextBoxUserNameMaintenance.Location = New System.Drawing.Point(162, 66)
        Me.TextBoxUserNameMaintenance.Name = "TextBoxUserNameMaintenance"
        Me.TextBoxUserNameMaintenance.Size = New System.Drawing.Size(100, 20)
        Me.TextBoxUserNameMaintenance.TabIndex = 4
        '
        'LabelPasswordMaintenance
        '
        Me.LabelPasswordMaintenance.AutoSize = True
        Me.LabelPasswordMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPasswordMaintenance.Location = New System.Drawing.Point(33, 95)
        Me.LabelPasswordMaintenance.Name = "LabelPasswordMaintenance"
        Me.LabelPasswordMaintenance.Size = New System.Drawing.Size(83, 13)
        Me.LabelPasswordMaintenance.TabIndex = 15
        Me.LabelPasswordMaintenance.Text = "Mot de passe"
        '
        'TextBoxPassword
        '
        Me.TextBoxPassword.Location = New System.Drawing.Point(162, 92)
        Me.TextBoxPassword.Name = "TextBoxPassword"
        Me.TextBoxPassword.Size = New System.Drawing.Size(100, 20)
        Me.TextBoxPassword.TabIndex = 7
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog1"
        '
        'GroupBoxMaintenance
        '
        Me.GroupBoxMaintenance.Controls.Add(Me.ButtonDestinationLocalMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.LabelDestinationMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.TextBoxDestinationMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.ButtonMiseàJourMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.RadioButtonTransfertLocalMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.LabelTypeDocumentMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.TextBoxPassword)
        Me.GroupBoxMaintenance.Controls.Add(Me.RadioButtonFTPMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.LabelPasswordMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.ComboBoxTypeDocumentMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.LabelUserMaintenance)
        Me.GroupBoxMaintenance.Controls.Add(Me.TextBoxUserNameMaintenance)
        Me.GroupBoxMaintenance.Location = New System.Drawing.Point(18, 78)
        Me.GroupBoxMaintenance.Name = "GroupBoxMaintenance"
        Me.GroupBoxMaintenance.Size = New System.Drawing.Size(701, 128)
        Me.GroupBoxMaintenance.TabIndex = 17
        Me.GroupBoxMaintenance.TabStop = False
        Me.GroupBoxMaintenance.Visible = False
        '
        'ButtonDestinationLocalMaintenance
        '
        Me.ButtonDestinationLocalMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDestinationLocalMaintenance.Location = New System.Drawing.Point(474, 35)
        Me.ButtonDestinationLocalMaintenance.Name = "ButtonDestinationLocalMaintenance"
        Me.ButtonDestinationLocalMaintenance.Size = New System.Drawing.Size(174, 23)
        Me.ButtonDestinationLocalMaintenance.TabIndex = 3
        Me.ButtonDestinationLocalMaintenance.Text = "Destination Local"
        Me.ButtonDestinationLocalMaintenance.UseVisualStyleBackColor = True
        '
        'LabelDestinationMaintenance
        '
        Me.LabelDestinationMaintenance.AutoSize = True
        Me.LabelDestinationMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDestinationMaintenance.Location = New System.Drawing.Point(33, 41)
        Me.LabelDestinationMaintenance.Name = "LabelDestinationMaintenance"
        Me.LabelDestinationMaintenance.Size = New System.Drawing.Size(71, 13)
        Me.LabelDestinationMaintenance.TabIndex = 19
        Me.LabelDestinationMaintenance.Text = "Destination"
        '
        'TextBoxDestinationMaintenance
        '
        Me.TextBoxDestinationMaintenance.Location = New System.Drawing.Point(162, 38)
        Me.TextBoxDestinationMaintenance.Name = "TextBoxDestinationMaintenance"
        Me.TextBoxDestinationMaintenance.Size = New System.Drawing.Size(284, 20)
        Me.TextBoxDestinationMaintenance.TabIndex = 2
        '
        'ButtonMiseàJourMaintenance
        '
        Me.ButtonMiseàJourMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonMiseàJourMaintenance.Location = New System.Drawing.Point(268, 90)
        Me.ButtonMiseàJourMaintenance.Name = "ButtonMiseàJourMaintenance"
        Me.ButtonMiseàJourMaintenance.Size = New System.Drawing.Size(173, 23)
        Me.ButtonMiseàJourMaintenance.TabIndex = 8
        Me.ButtonMiseàJourMaintenance.Text = "Mise à jour"
        Me.ButtonMiseàJourMaintenance.UseVisualStyleBackColor = True
        '
        'LabelTypeDocumentMaintenance
        '
        Me.LabelTypeDocumentMaintenance.AutoSize = True
        Me.LabelTypeDocumentMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTypeDocumentMaintenance.Location = New System.Drawing.Point(33, 14)
        Me.LabelTypeDocumentMaintenance.Name = "LabelTypeDocumentMaintenance"
        Me.LabelTypeDocumentMaintenance.Size = New System.Drawing.Size(96, 13)
        Me.LabelTypeDocumentMaintenance.TabIndex = 20
        Me.LabelTypeDocumentMaintenance.Text = "Type Document"
        '
        'LabelDataBaseName
        '
        Me.LabelDataBaseName.AutoSize = True
        Me.LabelDataBaseName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDataBaseName.Location = New System.Drawing.Point(12, 9)
        Me.LabelDataBaseName.Name = "LabelDataBaseName"
        Me.LabelDataBaseName.Size = New System.Drawing.Size(165, 13)
        Me.LabelDataBaseName.TabIndex = 21
        Me.LabelDataBaseName.Text = "Nom de la base de données"
        '
        'LabelNomDeLaBase
        '
        Me.LabelNomDeLaBase.AutoSize = True
        Me.LabelNomDeLaBase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNomDeLaBase.Location = New System.Drawing.Point(216, 9)
        Me.LabelNomDeLaBase.Name = "LabelNomDeLaBase"
        Me.LabelNomDeLaBase.Size = New System.Drawing.Size(22, 15)
        Me.LabelNomDeLaBase.TabIndex = 22
        Me.LabelNomDeLaBase.Text = "---"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.LabelEnregistrement)
        Me.Panel1.Controls.Add(Me.LabelNomDeLaBase)
        Me.Panel1.Controls.Add(Me.LabelChrono)
        Me.Panel1.Controls.Add(Me.LabelDataBaseName)
        Me.Panel1.Controls.Add(Me.LabelNuméroChrono)
        Me.Panel1.Controls.Add(Me.LabelNuméroEnregistrement)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(753, 75)
        Me.Panel1.TabIndex = 23
        '
        'GroupBoxInfosDocument
        '
        Me.GroupBoxInfosDocument.Controls.Add(Me.CheckBoxActivationDetail)
        Me.GroupBoxInfosDocument.Controls.Add(Me.RadioButtonImagePreviewITA)
        Me.GroupBoxInfosDocument.Controls.Add(Me.RadioButtonImagePreviewFRA)
        Me.GroupBoxInfosDocument.Controls.Add(Me.RadioButtonImagePreviewESP)
        Me.GroupBoxInfosDocument.Controls.Add(Me.RadioButtonImagePreviewANG)
        Me.GroupBoxInfosDocument.Controls.Add(Me.RadioButtonImagePreviewALL)
        Me.GroupBoxInfosDocument.Controls.Add(Me.ComboBoxTypeDocument)
        Me.GroupBoxInfosDocument.Controls.Add(Me.LabelTypeDocDocument)
        Me.GroupBoxInfosDocument.Controls.Add(Me.TextBoxNomFichier)
        Me.GroupBoxInfosDocument.Controls.Add(Me.LabelNomFichier)
        Me.GroupBoxInfosDocument.Controls.Add(Me.TextBoxNomDuLienDuFichier)
        Me.GroupBoxInfosDocument.Controls.Add(Me.LabelNomLienDuFichier)
        Me.GroupBoxInfosDocument.Controls.Add(Me.LabelNomFichierOrigine)
        Me.GroupBoxInfosDocument.Controls.Add(Me.ButtonSelectFile)
        Me.GroupBoxInfosDocument.Controls.Add(Me.PictureBoxSpanishFlag)
        Me.GroupBoxInfosDocument.Controls.Add(Me.ButtonAddDocument)
        Me.GroupBoxInfosDocument.Controls.Add(Me.PictureBoxGermanFlag)
        Me.GroupBoxInfosDocument.Controls.Add(Me.PictureBoxEnglishgFlag)
        Me.GroupBoxInfosDocument.Controls.Add(Me.PictureBoxFrenchFlag)
        Me.GroupBoxInfosDocument.Controls.Add(Me.PictureBoxItalianFlag)
        Me.GroupBoxInfosDocument.Enabled = False
        Me.GroupBoxInfosDocument.Location = New System.Drawing.Point(385, 245)
        Me.GroupBoxInfosDocument.Name = "GroupBoxInfosDocument"
        Me.GroupBoxInfosDocument.Size = New System.Drawing.Size(334, 310)
        Me.GroupBoxInfosDocument.TabIndex = 24
        Me.GroupBoxInfosDocument.TabStop = False
        '
        'RadioButtonImagePreviewITA
        '
        Me.RadioButtonImagePreviewITA.AutoSize = True
        Me.RadioButtonImagePreviewITA.Enabled = False
        Me.RadioButtonImagePreviewITA.Location = New System.Drawing.Point(266, 118)
        Me.RadioButtonImagePreviewITA.Name = "RadioButtonImagePreviewITA"
        Me.RadioButtonImagePreviewITA.Size = New System.Drawing.Size(14, 13)
        Me.RadioButtonImagePreviewITA.TabIndex = 34
        Me.ToolTip1.SetToolTip(Me.RadioButtonImagePreviewITA, "Image_Preview")
        Me.RadioButtonImagePreviewITA.UseVisualStyleBackColor = True
        '
        'RadioButtonImagePreviewFRA
        '
        Me.RadioButtonImagePreviewFRA.AutoSize = True
        Me.RadioButtonImagePreviewFRA.Enabled = False
        Me.RadioButtonImagePreviewFRA.Location = New System.Drawing.Point(210, 118)
        Me.RadioButtonImagePreviewFRA.Name = "RadioButtonImagePreviewFRA"
        Me.RadioButtonImagePreviewFRA.Size = New System.Drawing.Size(14, 13)
        Me.RadioButtonImagePreviewFRA.TabIndex = 33
        Me.ToolTip1.SetToolTip(Me.RadioButtonImagePreviewFRA, "Image_Preview")
        Me.RadioButtonImagePreviewFRA.UseVisualStyleBackColor = True
        '
        'RadioButtonImagePreviewESP
        '
        Me.RadioButtonImagePreviewESP.AutoSize = True
        Me.RadioButtonImagePreviewESP.Enabled = False
        Me.RadioButtonImagePreviewESP.Location = New System.Drawing.Point(154, 118)
        Me.RadioButtonImagePreviewESP.Name = "RadioButtonImagePreviewESP"
        Me.RadioButtonImagePreviewESP.Size = New System.Drawing.Size(14, 13)
        Me.RadioButtonImagePreviewESP.TabIndex = 32
        Me.ToolTip1.SetToolTip(Me.RadioButtonImagePreviewESP, "Image_Preview")
        Me.RadioButtonImagePreviewESP.UseVisualStyleBackColor = True
        '
        'RadioButtonImagePreviewANG
        '
        Me.RadioButtonImagePreviewANG.AutoSize = True
        Me.RadioButtonImagePreviewANG.Enabled = False
        Me.RadioButtonImagePreviewANG.Location = New System.Drawing.Point(98, 118)
        Me.RadioButtonImagePreviewANG.Name = "RadioButtonImagePreviewANG"
        Me.RadioButtonImagePreviewANG.Size = New System.Drawing.Size(14, 13)
        Me.RadioButtonImagePreviewANG.TabIndex = 31
        Me.ToolTip1.SetToolTip(Me.RadioButtonImagePreviewANG, "Image_Preview")
        Me.RadioButtonImagePreviewANG.UseVisualStyleBackColor = True
        '
        'RadioButtonImagePreviewALL
        '
        Me.RadioButtonImagePreviewALL.AutoSize = True
        Me.RadioButtonImagePreviewALL.Enabled = False
        Me.RadioButtonImagePreviewALL.Location = New System.Drawing.Point(42, 118)
        Me.RadioButtonImagePreviewALL.Name = "RadioButtonImagePreviewALL"
        Me.RadioButtonImagePreviewALL.Size = New System.Drawing.Size(14, 13)
        Me.RadioButtonImagePreviewALL.TabIndex = 30
        Me.ToolTip1.SetToolTip(Me.RadioButtonImagePreviewALL, "Image_Preview")
        Me.RadioButtonImagePreviewALL.UseVisualStyleBackColor = True
        '
        'ComboBoxTypeDocument
        '
        Me.ComboBoxTypeDocument.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxTypeDocument.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxTypeDocument.FormattingEnabled = True
        Me.ComboBoxTypeDocument.Location = New System.Drawing.Point(11, 29)
        Me.ComboBoxTypeDocument.Name = "ComboBoxTypeDocument"
        Me.ComboBoxTypeDocument.Size = New System.Drawing.Size(304, 23)
        Me.ComboBoxTypeDocument.TabIndex = 11
        '
        'LabelTypeDocDocument
        '
        Me.LabelTypeDocDocument.AutoSize = True
        Me.LabelTypeDocDocument.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTypeDocDocument.Location = New System.Drawing.Point(8, 12)
        Me.LabelTypeDocDocument.Name = "LabelTypeDocDocument"
        Me.LabelTypeDocDocument.Size = New System.Drawing.Size(96, 13)
        Me.LabelTypeDocDocument.TabIndex = 29
        Me.LabelTypeDocDocument.Text = "Type Document"
        '
        'TextBoxNomFichier
        '
        Me.TextBoxNomFichier.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxNomFichier.Location = New System.Drawing.Point(11, 245)
        Me.TextBoxNomFichier.Name = "TextBoxNomFichier"
        Me.TextBoxNomFichier.Size = New System.Drawing.Size(304, 20)
        Me.TextBoxNomFichier.TabIndex = 14
        '
        'LabelNomFichier
        '
        Me.LabelNomFichier.AutoSize = True
        Me.LabelNomFichier.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNomFichier.Location = New System.Drawing.Point(8, 228)
        Me.LabelNomFichier.Name = "LabelNomFichier"
        Me.LabelNomFichier.Size = New System.Drawing.Size(92, 13)
        Me.LabelNomFichier.TabIndex = 27
        Me.LabelNomFichier.Text = "Nom du Fichier"
        '
        'TextBoxNomDuLienDuFichier
        '
        Me.TextBoxNomDuLienDuFichier.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxNomDuLienDuFichier.Location = New System.Drawing.Point(11, 201)
        Me.TextBoxNomDuLienDuFichier.Name = "TextBoxNomDuLienDuFichier"
        Me.TextBoxNomDuLienDuFichier.Size = New System.Drawing.Size(304, 20)
        Me.TextBoxNomDuLienDuFichier.TabIndex = 13
        '
        'LabelNomLienDuFichier
        '
        Me.LabelNomLienDuFichier.AutoSize = True
        Me.LabelNomLienDuFichier.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNomLienDuFichier.Location = New System.Drawing.Point(8, 184)
        Me.LabelNomLienDuFichier.Name = "LabelNomLienDuFichier"
        Me.LabelNomLienDuFichier.Size = New System.Drawing.Size(134, 13)
        Me.LabelNomLienDuFichier.TabIndex = 11
        Me.LabelNomLienDuFichier.Text = "Nom du lien du Fichier"
        '
        'LabelNomFichierOrigine
        '
        Me.LabelNomFichierOrigine.AutoSize = True
        Me.LabelNomFichierOrigine.Location = New System.Drawing.Point(17, 167)
        Me.LabelNomFichierOrigine.Name = "LabelNomFichierOrigine"
        Me.LabelNomFichierOrigine.Size = New System.Drawing.Size(16, 13)
        Me.LabelNomFichierOrigine.TabIndex = 10
        Me.LabelNomFichierOrigine.Text = "---"
        '
        'ButtonSelectFile
        '
        Me.ButtonSelectFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSelectFile.Location = New System.Drawing.Point(11, 141)
        Me.ButtonSelectFile.Name = "ButtonSelectFile"
        Me.ButtonSelectFile.Size = New System.Drawing.Size(304, 23)
        Me.ButtonSelectFile.TabIndex = 12
        Me.ButtonSelectFile.Text = "Sélectionner le Fichier "
        Me.ButtonSelectFile.UseVisualStyleBackColor = True
        '
        'ListBoxDocuments
        '
        Me.ListBoxDocuments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBoxDocuments.FormattingEnabled = True
        Me.ListBoxDocuments.Location = New System.Drawing.Point(18, 273)
        Me.ListBoxDocuments.Name = "ListBoxDocuments"
        Me.ListBoxDocuments.ScrollAlwaysVisible = True
        Me.ListBoxDocuments.Size = New System.Drawing.Size(345, 212)
        Me.ListBoxDocuments.TabIndex = 9
        '
        'LabelListeDesDocuments
        '
        Me.LabelListeDesDocuments.AutoSize = True
        Me.LabelListeDesDocuments.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelListeDesDocuments.Location = New System.Drawing.Point(15, 245)
        Me.LabelListeDesDocuments.Name = "LabelListeDesDocuments"
        Me.LabelListeDesDocuments.Size = New System.Drawing.Size(287, 16)
        Me.LabelListeDesDocuments.TabIndex = 26
        Me.LabelListeDesDocuments.Text = "Liste des documents de l'enregistrement"
        '
        'ButtonEffaceUnDocument
        '
        Me.ButtonEffaceUnDocument.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonEffaceUnDocument.Location = New System.Drawing.Point(18, 491)
        Me.ButtonEffaceUnDocument.Name = "ButtonEffaceUnDocument"
        Me.ButtonEffaceUnDocument.Size = New System.Drawing.Size(345, 23)
        Me.ButtonEffaceUnDocument.TabIndex = 10
        Me.ButtonEffaceUnDocument.Text = "Efface le document sélectionné"
        Me.ButtonEffaceUnDocument.UseVisualStyleBackColor = True
        '
        'StatusStrip
        '
        Me.StatusStrip.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 573)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(753, 22)
        Me.StatusStrip.TabIndex = 27
        Me.StatusStrip.Text = "---"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(22, 17)
        Me.ToolStripStatusLabel1.Text = "---"
        '
        'BackgroundWorkerChargementTypeDocument
        '
        '
        'TimerChargementTypeDocument
        '
        Me.TimerChargementTypeDocument.Interval = 1000
        '
        'ButtonNouveauDocument
        '
        Me.ButtonNouveauDocument.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNouveauDocument.Location = New System.Drawing.Point(18, 520)
        Me.ButtonNouveauDocument.Name = "ButtonNouveauDocument"
        Me.ButtonNouveauDocument.Size = New System.Drawing.Size(345, 23)
        Me.ButtonNouveauDocument.TabIndex = 30
        Me.ButtonNouveauDocument.Text = "Nouveau Document"
        Me.ButtonNouveauDocument.UseVisualStyleBackColor = False
        '
        'BackgroundWorkerChargementListeDocuments
        '
        '
        'TimerChargementListeDocuments
        '
        Me.TimerChargementListeDocuments.Interval = 1000
        '
        'BackgroundWorker_MaintenanceTransfert
        '
        '
        'TimerEtatMaintenanceTransfert
        '
        Me.TimerEtatMaintenanceTransfert.Interval = 1000
        '
        'BackgroundWorker_SupprimeDocument
        '
        '
        'TimerEtatSupprimeDocument
        '
        Me.TimerEtatSupprimeDocument.Interval = 1000
        '
        'BackgroundWorker_AjoutDocument
        '
        '
        'TimerEtatAjoutDocument
        '
        Me.TimerEtatAjoutDocument.Interval = 1000
        '
        'CheckBoxActivationDetail
        '
        Me.CheckBoxActivationDetail.AutoSize = True
        Me.CheckBoxActivationDetail.Checked = True
        Me.CheckBoxActivationDetail.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxActivationDetail.Location = New System.Drawing.Point(23, 57)
        Me.CheckBoxActivationDetail.Name = "CheckBoxActivationDetail"
        Me.CheckBoxActivationDetail.Size = New System.Drawing.Size(111, 17)
        Me.CheckBoxActivationDetail.TabIndex = 35
        Me.CheckBoxActivationDetail.Text = "Détails par langue"
        Me.CheckBoxActivationDetail.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(753, 595)
        Me.Controls.Add(Me.ButtonNouveauDocument)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.ButtonEffaceUnDocument)
        Me.Controls.Add(Me.LabelListeDesDocuments)
        Me.Controls.Add(Me.ListBoxDocuments)
        Me.Controls.Add(Me.GroupBoxInfosDocument)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.GroupBoxMaintenance)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Transfert Idealtake"
        CType(Me.PictureBoxGermanFlag, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxEnglishgFlag, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxSpanishFlag, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxFrenchFlag, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxItalianFlag, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxMaintenance.ResumeLayout(False)
        Me.GroupBoxMaintenance.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBoxInfosDocument.ResumeLayout(False)
        Me.GroupBoxInfosDocument.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelChrono As System.Windows.Forms.Label
    Friend WithEvents LabelNuméroChrono As System.Windows.Forms.Label
    Friend WithEvents LabelNuméroEnregistrement As System.Windows.Forms.Label
    Friend WithEvents LabelEnregistrement As System.Windows.Forms.Label
    Friend WithEvents PictureBoxGermanFlag As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBoxEnglishgFlag As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBoxSpanishFlag As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBoxFrenchFlag As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBoxItalianFlag As System.Windows.Forms.PictureBox
    Friend WithEvents ButtonAddDocument As System.Windows.Forms.Button
    Friend WithEvents ComboBoxTypeDocumentMaintenance As System.Windows.Forms.ComboBox
    Friend WithEvents RadioButtonFTPMaintenance As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonTransfertLocalMaintenance As System.Windows.Forms.RadioButton
    Friend WithEvents LabelUserMaintenance As System.Windows.Forms.Label
    Friend WithEvents TextBoxUserNameMaintenance As System.Windows.Forms.TextBox
    Friend WithEvents LabelPasswordMaintenance As System.Windows.Forms.Label
    Friend WithEvents TextBoxPassword As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents GroupBoxMaintenance As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonMiseàJourMaintenance As System.Windows.Forms.Button
    Friend WithEvents LabelDestinationMaintenance As System.Windows.Forms.Label
    Friend WithEvents TextBoxDestinationMaintenance As System.Windows.Forms.TextBox
    Friend WithEvents LabelTypeDocumentMaintenance As System.Windows.Forms.Label
    Friend WithEvents LabelDataBaseName As System.Windows.Forms.Label
    Friend WithEvents LabelNomDeLaBase As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBoxInfosDocument As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonSelectFile As System.Windows.Forms.Button
    Friend WithEvents TextBoxNomDuLienDuFichier As System.Windows.Forms.TextBox
    Friend WithEvents LabelNomLienDuFichier As System.Windows.Forms.Label
    Friend WithEvents LabelNomFichierOrigine As System.Windows.Forms.Label
    Friend WithEvents TextBoxNomFichier As System.Windows.Forms.TextBox
    Friend WithEvents LabelNomFichier As System.Windows.Forms.Label
    Friend WithEvents ListBoxDocuments As System.Windows.Forms.ListBox
    Friend WithEvents LabelListeDesDocuments As System.Windows.Forms.Label
    Friend WithEvents ButtonEffaceUnDocument As System.Windows.Forms.Button
    Friend WithEvents ButtonDestinationLocalMaintenance As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents LabelTypeDocDocument As System.Windows.Forms.Label
    Friend WithEvents ComboBoxTypeDocument As System.Windows.Forms.ComboBox
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents BackgroundWorkerChargementTypeDocument As System.ComponentModel.BackgroundWorker
    Friend WithEvents TimerChargementTypeDocument As System.Windows.Forms.Timer
    Friend WithEvents ButtonNouveauDocument As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorkerChargementListeDocuments As System.ComponentModel.BackgroundWorker
    Friend WithEvents TimerChargementListeDocuments As System.Windows.Forms.Timer
    Friend WithEvents BackgroundWorker_MaintenanceTransfert As System.ComponentModel.BackgroundWorker
    Friend WithEvents TimerEtatMaintenanceTransfert As System.Windows.Forms.Timer
    Friend WithEvents BackgroundWorker_SupprimeDocument As System.ComponentModel.BackgroundWorker
    Friend WithEvents TimerEtatSupprimeDocument As System.Windows.Forms.Timer
    Friend WithEvents BackgroundWorker_AjoutDocument As System.ComponentModel.BackgroundWorker
    Friend WithEvents TimerEtatAjoutDocument As System.Windows.Forms.Timer
    Friend WithEvents RadioButtonImagePreviewITA As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonImagePreviewFRA As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonImagePreviewESP As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonImagePreviewANG As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonImagePreviewALL As System.Windows.Forms.RadioButton
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents CheckBoxActivationDetail As System.Windows.Forms.CheckBox

End Class
