<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Ajout_Suppression_connexion_base
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ajout_Suppression_connexion_base))
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Description_connexion_Label = New System.Windows.Forms.Label()
        Me.Flag_Italy = New System.Windows.Forms.PictureBox()
        Me.Flag_France = New System.Windows.Forms.PictureBox()
        Me.Flag_Spain = New System.Windows.Forms.PictureBox()
        Me.Flag_USA = New System.Windows.Forms.PictureBox()
        Me.Flag_Germany = New System.Windows.Forms.PictureBox()
        Me.Selection_connexion_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Connect_base = New System.Windows.Forms.Button()
        Me.Titre_de_la_connexion_Label = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Etat_Ajout_Suppression_TextBox = New System.Windows.Forms.TextBox()
        Me.Supprimer_une_connexion_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Supprimer_une_connexion = New System.Windows.Forms.Button()
        Me.Titre_de_la_connexion_Label1 = New System.Windows.Forms.Label()
        Me.Selection_connexion_ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Ajouter_une_connexion_GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Description_connexion_TextBox = New System.Windows.Forms.TextBox()
        Me.Description_Connexion_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Localisation_Photo_produit_TextBox = New System.Windows.Forms.TextBox()
        Me.Label_Photos_Produit = New System.Windows.Forms.Label()
        Me.Localisation_Fiche_Produit_TextBox = New System.Windows.Forms.TextBox()
        Me.Label_Fiche_Produit = New System.Windows.Forms.Label()
        Me.Localisation_Notice_TextBox = New System.Windows.Forms.TextBox()
        Me.Label_Notice = New System.Windows.Forms.Label()
        Me.Label_Chemins_fichiers = New System.Windows.Forms.Label()
        Me.Localisation_Synoptique_TextBox = New System.Windows.Forms.TextBox()
        Me.Label_Chemin_Synoptiques = New System.Windows.Forms.Label()
        Me.Localisation_Schéma_TextBox = New System.Windows.Forms.TextBox()
        Me.Label_Chemin_Schéma = New System.Windows.Forms.Label()
        Me.Ajouter_une_connexion = New System.Windows.Forms.Button()
        Me.Nom_de_la_base_TextBox = New System.Windows.Forms.TextBox()
        Me.Confimer_mot_de_passe_MaskedTextBox = New System.Windows.Forms.MaskedTextBox()
        Me.mot_de_passe_MaskedTextBox = New System.Windows.Forms.MaskedTextBox()
        Me.Confirmer_mot_de_passe = New System.Windows.Forms.Label()
        Me.Utilisateur_TextBox = New System.Windows.Forms.TextBox()
        Me.Adresse_Serveur_TextBox = New System.Windows.Forms.TextBox()
        Me.Nom_Connexion_TextBox = New System.Windows.Forms.TextBox()
        Me.Creer_base = New System.Windows.Forms.CheckBox()
        Me.Nom_de_la_base = New System.Windows.Forms.Label()
        Me.Mot_de_passe_Utilisateur = New System.Windows.Forms.Label()
        Me.Nom_Utilisateur = New System.Windows.Forms.Label()
        Me.IP_SERVEUR = New System.Windows.Forms.Label()
        Me.Nom_Connexion = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.Flag_Italy, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Flag_France, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Flag_Spain, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Flag_USA, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Flag_Germany, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.Supprimer_une_connexion_GroupBox.SuspendLayout()
        Me.Ajouter_une_connexion_GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.Color.Transparent
        Me.StatusStrip1.BackgroundImage = CType(resources.GetObject("StatusStrip1.BackgroundImage"), System.Drawing.Image)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 429)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(630, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.ForeColor = System.Drawing.Color.Black
        Me.TabControl1.Location = New System.Drawing.Point(6, 1)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(618, 425)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.LightCyan
        Me.TabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.TabPage1.Controls.Add(Me.Button2)
        Me.TabPage1.Controls.Add(Me.Description_connexion_Label)
        Me.TabPage1.Controls.Add(Me.Flag_Italy)
        Me.TabPage1.Controls.Add(Me.Flag_France)
        Me.TabPage1.Controls.Add(Me.Flag_Spain)
        Me.TabPage1.Controls.Add(Me.Flag_USA)
        Me.TabPage1.Controls.Add(Me.Flag_Germany)
        Me.TabPage1.Controls.Add(Me.Selection_connexion_ComboBox)
        Me.TabPage1.Controls.Add(Me.Connect_base)
        Me.TabPage1.Controls.Add(Me.Titre_de_la_connexion_Label)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(610, 399)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Connexion"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(424, 43)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(156, 23)
        Me.Button2.TabIndex = 11
        Me.Button2.Text = "Logiciel de dessin"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Description_connexion_Label
        '
        Me.Description_connexion_Label.AutoSize = True
        Me.Description_connexion_Label.Location = New System.Drawing.Point(124, 169)
        Me.Description_connexion_Label.Name = "Description_connexion_Label"
        Me.Description_connexion_Label.Size = New System.Drawing.Size(0, 13)
        Me.Description_connexion_Label.TabIndex = 10
        '
        'Flag_Italy
        '
        Me.Flag_Italy.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Flag_of_Italy
        Me.Flag_Italy.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Flag_Italy.Location = New System.Drawing.Point(280, 374)
        Me.Flag_Italy.Name = "Flag_Italy"
        Me.Flag_Italy.Size = New System.Drawing.Size(37, 22)
        Me.Flag_Italy.TabIndex = 9
        Me.Flag_Italy.TabStop = False
        '
        'Flag_France
        '
        Me.Flag_France.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Flag_of_France
        Me.Flag_France.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Flag_France.Location = New System.Drawing.Point(215, 374)
        Me.Flag_France.Name = "Flag_France"
        Me.Flag_France.Size = New System.Drawing.Size(37, 22)
        Me.Flag_France.TabIndex = 8
        Me.Flag_France.TabStop = False
        '
        'Flag_Spain
        '
        Me.Flag_Spain.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Flag_of_Spain
        Me.Flag_Spain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Flag_Spain.Location = New System.Drawing.Point(150, 374)
        Me.Flag_Spain.Name = "Flag_Spain"
        Me.Flag_Spain.Size = New System.Drawing.Size(37, 22)
        Me.Flag_Spain.TabIndex = 7
        Me.Flag_Spain.TabStop = False
        '
        'Flag_USA
        '
        Me.Flag_USA.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Flag_of_the_United_States
        Me.Flag_USA.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Flag_USA.Location = New System.Drawing.Point(85, 374)
        Me.Flag_USA.Name = "Flag_USA"
        Me.Flag_USA.Size = New System.Drawing.Size(37, 22)
        Me.Flag_USA.TabIndex = 6
        Me.Flag_USA.TabStop = False
        '
        'Flag_Germany
        '
        Me.Flag_Germany.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Flag_of_Germany
        Me.Flag_Germany.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Flag_Germany.Location = New System.Drawing.Point(20, 374)
        Me.Flag_Germany.Name = "Flag_Germany"
        Me.Flag_Germany.Size = New System.Drawing.Size(37, 22)
        Me.Flag_Germany.TabIndex = 5
        Me.Flag_Germany.TabStop = False
        '
        'Selection_connexion_ComboBox
        '
        Me.Selection_connexion_ComboBox.BackColor = System.Drawing.SystemColors.Window
        Me.Selection_connexion_ComboBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Selection_connexion_ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Selection_connexion_ComboBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Selection_connexion_ComboBox.FormattingEnabled = True
        Me.Selection_connexion_ComboBox.Location = New System.Drawing.Point(111, 127)
        Me.Selection_connexion_ComboBox.Name = "Selection_connexion_ComboBox"
        Me.Selection_connexion_ComboBox.Size = New System.Drawing.Size(150, 21)
        Me.Selection_connexion_ComboBox.TabIndex = 1
        '
        'Connect_base
        '
        Me.Connect_base.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Connect_base.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Connect_base.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Connect_base.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Connect_base.Location = New System.Drawing.Point(295, 125)
        Me.Connect_base.Name = "Connect_base"
        Me.Connect_base.Size = New System.Drawing.Size(75, 23)
        Me.Connect_base.TabIndex = 3
        Me.Connect_base.Text = "Connexion"
        Me.Connect_base.UseVisualStyleBackColor = False
        '
        'Titre_de_la_connexion_Label
        '
        Me.Titre_de_la_connexion_Label.AutoSize = True
        Me.Titre_de_la_connexion_Label.BackColor = System.Drawing.Color.Transparent
        Me.Titre_de_la_connexion_Label.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Titre_de_la_connexion_Label.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Titre_de_la_connexion_Label.Location = New System.Drawing.Point(124, 105)
        Me.Titre_de_la_connexion_Label.Name = "Titre_de_la_connexion_Label"
        Me.Titre_de_la_connexion_Label.Size = New System.Drawing.Size(106, 13)
        Me.Titre_de_la_connexion_Label.TabIndex = 0
        Me.Titre_de_la_connexion_Label.Text = "Titre de la connexion"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.LightCyan
        Me.TabPage2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.TabPage2.Controls.Add(Me.Etat_Ajout_Suppression_TextBox)
        Me.TabPage2.Controls.Add(Me.Supprimer_une_connexion_GroupBox)
        Me.TabPage2.Controls.Add(Me.Ajouter_une_connexion_GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(610, 399)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Ajout - Suppression"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Etat_Ajout_Suppression_TextBox
        '
        Me.Etat_Ajout_Suppression_TextBox.BackColor = System.Drawing.SystemColors.Window
        Me.Etat_Ajout_Suppression_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Etat_Ajout_Suppression_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Etat_Ajout_Suppression_TextBox.Location = New System.Drawing.Point(386, 12)
        Me.Etat_Ajout_Suppression_TextBox.Multiline = True
        Me.Etat_Ajout_Suppression_TextBox.Name = "Etat_Ajout_Suppression_TextBox"
        Me.Etat_Ajout_Suppression_TextBox.ReadOnly = True
        Me.Etat_Ajout_Suppression_TextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.Etat_Ajout_Suppression_TextBox.Size = New System.Drawing.Size(218, 274)
        Me.Etat_Ajout_Suppression_TextBox.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.Etat_Ajout_Suppression_TextBox, "Informations de création de base et des tables")
        '
        'Supprimer_une_connexion_GroupBox
        '
        Me.Supprimer_une_connexion_GroupBox.BackColor = System.Drawing.Color.DarkTurquoise
        Me.Supprimer_une_connexion_GroupBox.Controls.Add(Me.Supprimer_une_connexion)
        Me.Supprimer_une_connexion_GroupBox.Controls.Add(Me.Titre_de_la_connexion_Label1)
        Me.Supprimer_une_connexion_GroupBox.Controls.Add(Me.Selection_connexion_ComboBox2)
        Me.Supprimer_une_connexion_GroupBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Supprimer_une_connexion_GroupBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Supprimer_une_connexion_GroupBox.Location = New System.Drawing.Point(386, 292)
        Me.Supprimer_une_connexion_GroupBox.Name = "Supprimer_une_connexion_GroupBox"
        Me.Supprimer_une_connexion_GroupBox.Size = New System.Drawing.Size(218, 101)
        Me.Supprimer_une_connexion_GroupBox.TabIndex = 17
        Me.Supprimer_une_connexion_GroupBox.TabStop = False
        Me.Supprimer_une_connexion_GroupBox.Text = "Supprimer une connexion"
        '
        'Supprimer_une_connexion
        '
        Me.Supprimer_une_connexion.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Supprimer_une_connexion.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Supprimer_une_connexion.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Supprimer_une_connexion.Location = New System.Drawing.Point(31, 66)
        Me.Supprimer_une_connexion.Name = "Supprimer_une_connexion"
        Me.Supprimer_une_connexion.Size = New System.Drawing.Size(150, 23)
        Me.Supprimer_une_connexion.TabIndex = 19
        Me.Supprimer_une_connexion.Text = "Suppression connexion"
        Me.Supprimer_une_connexion.UseVisualStyleBackColor = False
        '
        'Titre_de_la_connexion_Label1
        '
        Me.Titre_de_la_connexion_Label1.AutoSize = True
        Me.Titre_de_la_connexion_Label1.BackColor = System.Drawing.Color.Transparent
        Me.Titre_de_la_connexion_Label1.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Titre_de_la_connexion_Label1.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Titre_de_la_connexion_Label1.Location = New System.Drawing.Point(28, 23)
        Me.Titre_de_la_connexion_Label1.Name = "Titre_de_la_connexion_Label1"
        Me.Titre_de_la_connexion_Label1.Size = New System.Drawing.Size(106, 13)
        Me.Titre_de_la_connexion_Label1.TabIndex = 1
        Me.Titre_de_la_connexion_Label1.Text = "Titre de la connexion"
        '
        'Selection_connexion_ComboBox2
        '
        Me.Selection_connexion_ComboBox2.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Selection_connexion_ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Selection_connexion_ComboBox2.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Selection_connexion_ComboBox2.FormattingEnabled = True
        Me.Selection_connexion_ComboBox2.Location = New System.Drawing.Point(31, 39)
        Me.Selection_connexion_ComboBox2.Name = "Selection_connexion_ComboBox2"
        Me.Selection_connexion_ComboBox2.Size = New System.Drawing.Size(150, 21)
        Me.Selection_connexion_ComboBox2.TabIndex = 18
        '
        'Ajouter_une_connexion_GroupBox1
        '
        Me.Ajouter_une_connexion_GroupBox1.BackColor = System.Drawing.Color.DarkTurquoise
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Description_connexion_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Description_Connexion_ComboBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Localisation_Photo_produit_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Label_Photos_Produit)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Localisation_Fiche_Produit_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Label_Fiche_Produit)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Localisation_Notice_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Label_Notice)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Label_Chemins_fichiers)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Localisation_Synoptique_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Label_Chemin_Synoptiques)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Localisation_Schéma_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Label_Chemin_Schéma)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Ajouter_une_connexion)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Nom_de_la_base_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Confimer_mot_de_passe_MaskedTextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.mot_de_passe_MaskedTextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Confirmer_mot_de_passe)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Utilisateur_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Adresse_Serveur_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Nom_Connexion_TextBox)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Creer_base)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Nom_de_la_base)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Mot_de_passe_Utilisateur)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Nom_Utilisateur)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.IP_SERVEUR)
        Me.Ajouter_une_connexion_GroupBox1.Controls.Add(Me.Nom_Connexion)
        Me.Ajouter_une_connexion_GroupBox1.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Ajouter_une_connexion_GroupBox1.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Ajouter_une_connexion_GroupBox1.Location = New System.Drawing.Point(5, 6)
        Me.Ajouter_une_connexion_GroupBox1.Name = "Ajouter_une_connexion_GroupBox1"
        Me.Ajouter_une_connexion_GroupBox1.Size = New System.Drawing.Size(375, 387)
        Me.Ajouter_une_connexion_GroupBox1.TabIndex = 0
        Me.Ajouter_une_connexion_GroupBox1.TabStop = False
        Me.Ajouter_une_connexion_GroupBox1.Text = "Ajouter une connexion"
        '
        'Description_connexion_TextBox
        '
        Me.Description_connexion_TextBox.Location = New System.Drawing.Point(162, 46)
        Me.Description_connexion_TextBox.Name = "Description_connexion_TextBox"
        Me.Description_connexion_TextBox.Size = New System.Drawing.Size(204, 20)
        Me.Description_connexion_TextBox.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.Description_connexion_TextBox, "Description de la connexion")
        '
        'Description_Connexion_ComboBox
        '
        Me.Description_Connexion_ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Description_Connexion_ComboBox.FormattingEnabled = True
        Me.Description_Connexion_ComboBox.Location = New System.Drawing.Point(5, 46)
        Me.Description_Connexion_ComboBox.Name = "Description_Connexion_ComboBox"
        Me.Description_Connexion_ComboBox.Size = New System.Drawing.Size(155, 21)
        Me.Description_Connexion_ComboBox.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Description_Connexion_ComboBox, "Sélectionnez la langue de la description")
        '
        'Localisation_Photo_produit_TextBox
        '
        Me.Localisation_Photo_produit_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Localisation_Photo_produit_TextBox.Enabled = False
        Me.Localisation_Photo_produit_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Localisation_Photo_produit_TextBox.Location = New System.Drawing.Point(162, 357)
        Me.Localisation_Photo_produit_TextBox.Name = "Localisation_Photo_produit_TextBox"
        Me.Localisation_Photo_produit_TextBox.Size = New System.Drawing.Size(204, 20)
        Me.Localisation_Photo_produit_TextBox.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.Localisation_Photo_produit_TextBox, "c:\PHOTOS_PRODUITS ou http://www.site.fr/data/Photos_produits")
        '
        'Label_Photos_Produit
        '
        Me.Label_Photos_Produit.AutoSize = True
        Me.Label_Photos_Produit.BackColor = System.Drawing.Color.Transparent
        Me.Label_Photos_Produit.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label_Photos_Produit.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Label_Photos_Produit.Location = New System.Drawing.Point(2, 357)
        Me.Label_Photos_Produit.Name = "Label_Photos_Produit"
        Me.Label_Photos_Produit.Size = New System.Drawing.Size(71, 13)
        Me.Label_Photos_Produit.TabIndex = 25
        Me.Label_Photos_Produit.Text = "Photo Produit"
        '
        'Localisation_Fiche_Produit_TextBox
        '
        Me.Localisation_Fiche_Produit_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Localisation_Fiche_Produit_TextBox.Enabled = False
        Me.Localisation_Fiche_Produit_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Localisation_Fiche_Produit_TextBox.Location = New System.Drawing.Point(162, 333)
        Me.Localisation_Fiche_Produit_TextBox.Name = "Localisation_Fiche_Produit_TextBox"
        Me.Localisation_Fiche_Produit_TextBox.Size = New System.Drawing.Size(204, 20)
        Me.Localisation_Fiche_Produit_TextBox.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.Localisation_Fiche_Produit_TextBox, "e.g : c:\FICHES_PRODUITS ou http://www.site.fr/data/fiches_produits")
        '
        'Label_Fiche_Produit
        '
        Me.Label_Fiche_Produit.AutoSize = True
        Me.Label_Fiche_Produit.BackColor = System.Drawing.Color.Transparent
        Me.Label_Fiche_Produit.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label_Fiche_Produit.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Label_Fiche_Produit.Location = New System.Drawing.Point(2, 333)
        Me.Label_Fiche_Produit.Name = "Label_Fiche_Produit"
        Me.Label_Fiche_Produit.Size = New System.Drawing.Size(69, 13)
        Me.Label_Fiche_Produit.TabIndex = 23
        Me.Label_Fiche_Produit.Text = "Fiche Produit"
        '
        'Localisation_Notice_TextBox
        '
        Me.Localisation_Notice_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Localisation_Notice_TextBox.Enabled = False
        Me.Localisation_Notice_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Localisation_Notice_TextBox.Location = New System.Drawing.Point(162, 309)
        Me.Localisation_Notice_TextBox.Name = "Localisation_Notice_TextBox"
        Me.Localisation_Notice_TextBox.Size = New System.Drawing.Size(204, 20)
        Me.Localisation_Notice_TextBox.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.Localisation_Notice_TextBox, "e.g : c:\NOTICES ou http://www.site.fr/data/notices")
        '
        'Label_Notice
        '
        Me.Label_Notice.AutoSize = True
        Me.Label_Notice.BackColor = System.Drawing.Color.Transparent
        Me.Label_Notice.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label_Notice.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Label_Notice.Location = New System.Drawing.Point(2, 307)
        Me.Label_Notice.Name = "Label_Notice"
        Me.Label_Notice.Size = New System.Drawing.Size(38, 13)
        Me.Label_Notice.TabIndex = 21
        Me.Label_Notice.Text = "Notice"
        '
        'Label_Chemins_fichiers
        '
        Me.Label_Chemins_fichiers.AutoSize = True
        Me.Label_Chemins_fichiers.BackColor = System.Drawing.Color.Transparent
        Me.Label_Chemins_fichiers.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label_Chemins_fichiers.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Chemins_fichiers.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Label_Chemins_fichiers.Location = New System.Drawing.Point(3, 240)
        Me.Label_Chemins_fichiers.Name = "Label_Chemins_fichiers"
        Me.Label_Chemins_fichiers.Size = New System.Drawing.Size(146, 13)
        Me.Label_Chemins_fichiers.TabIndex = 20
        Me.Label_Chemins_fichiers.Text = "Les chemins des fichiers"
        '
        'Localisation_Synoptique_TextBox
        '
        Me.Localisation_Synoptique_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Localisation_Synoptique_TextBox.Enabled = False
        Me.Localisation_Synoptique_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Localisation_Synoptique_TextBox.Location = New System.Drawing.Point(162, 285)
        Me.Localisation_Synoptique_TextBox.Name = "Localisation_Synoptique_TextBox"
        Me.Localisation_Synoptique_TextBox.Size = New System.Drawing.Size(204, 20)
        Me.Localisation_Synoptique_TextBox.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.Localisation_Synoptique_TextBox, "e.g : c:\SYNOPIQUES ou http://www.site.fr/data/synoptiques")
        '
        'Label_Chemin_Synoptiques
        '
        Me.Label_Chemin_Synoptiques.AutoSize = True
        Me.Label_Chemin_Synoptiques.BackColor = System.Drawing.Color.Transparent
        Me.Label_Chemin_Synoptiques.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label_Chemin_Synoptiques.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Label_Chemin_Synoptiques.Location = New System.Drawing.Point(2, 284)
        Me.Label_Chemin_Synoptiques.Name = "Label_Chemin_Synoptiques"
        Me.Label_Chemin_Synoptiques.Size = New System.Drawing.Size(60, 13)
        Me.Label_Chemin_Synoptiques.TabIndex = 18
        Me.Label_Chemin_Synoptiques.Text = "Synoptique"
        '
        'Localisation_Schéma_TextBox
        '
        Me.Localisation_Schéma_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Localisation_Schéma_TextBox.Enabled = False
        Me.Localisation_Schéma_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Localisation_Schéma_TextBox.Location = New System.Drawing.Point(162, 261)
        Me.Localisation_Schéma_TextBox.Name = "Localisation_Schéma_TextBox"
        Me.Localisation_Schéma_TextBox.Size = New System.Drawing.Size(205, 20)
        Me.Localisation_Schéma_TextBox.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.Localisation_Schéma_TextBox, "e.g : c:\SCHEMAS ou http://www.site.fr/data/schemas")
        '
        'Label_Chemin_Schéma
        '
        Me.Label_Chemin_Schéma.AutoSize = True
        Me.Label_Chemin_Schéma.BackColor = System.Drawing.Color.Transparent
        Me.Label_Chemin_Schéma.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Label_Chemin_Schéma.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Label_Chemin_Schéma.Location = New System.Drawing.Point(2, 264)
        Me.Label_Chemin_Schéma.Name = "Label_Chemin_Schéma"
        Me.Label_Chemin_Schéma.Size = New System.Drawing.Size(46, 13)
        Me.Label_Chemin_Schéma.TabIndex = 16
        Me.Label_Chemin_Schéma.Text = "Schéma"
        '
        'Ajouter_une_connexion
        '
        Me.Ajouter_une_connexion.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Ajouter_une_connexion.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Ajouter_une_connexion.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Ajouter_une_connexion.Location = New System.Drawing.Point(213, 202)
        Me.Ajouter_une_connexion.Name = "Ajouter_une_connexion"
        Me.Ajouter_une_connexion.Size = New System.Drawing.Size(75, 23)
        Me.Ajouter_une_connexion.TabIndex = 9
        Me.Ajouter_une_connexion.Text = "Ajouter"
        Me.Ajouter_une_connexion.UseVisualStyleBackColor = False
        '
        'Nom_de_la_base_TextBox
        '
        Me.Nom_de_la_base_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Nom_de_la_base_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Nom_de_la_base_TextBox.Location = New System.Drawing.Point(162, 176)
        Me.Nom_de_la_base_TextBox.Name = "Nom_de_la_base_TextBox"
        Me.Nom_de_la_base_TextBox.Size = New System.Drawing.Size(204, 20)
        Me.Nom_de_la_base_TextBox.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.Nom_de_la_base_TextBox, "Le nom de la base")
        '
        'Confimer_mot_de_passe_MaskedTextBox
        '
        Me.Confimer_mot_de_passe_MaskedTextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Confimer_mot_de_passe_MaskedTextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Confimer_mot_de_passe_MaskedTextBox.Location = New System.Drawing.Point(162, 150)
        Me.Confimer_mot_de_passe_MaskedTextBox.Name = "Confimer_mot_de_passe_MaskedTextBox"
        Me.Confimer_mot_de_passe_MaskedTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Confimer_mot_de_passe_MaskedTextBox.Size = New System.Drawing.Size(204, 20)
        Me.Confimer_mot_de_passe_MaskedTextBox.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.Confimer_mot_de_passe_MaskedTextBox, "confirmation du mot de passe")
        '
        'mot_de_passe_MaskedTextBox
        '
        Me.mot_de_passe_MaskedTextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.mot_de_passe_MaskedTextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.mot_de_passe_MaskedTextBox.Location = New System.Drawing.Point(162, 124)
        Me.mot_de_passe_MaskedTextBox.Name = "mot_de_passe_MaskedTextBox"
        Me.mot_de_passe_MaskedTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.mot_de_passe_MaskedTextBox.Size = New System.Drawing.Size(204, 20)
        Me.mot_de_passe_MaskedTextBox.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.mot_de_passe_MaskedTextBox, "Le mot de passe")
        '
        'Confirmer_mot_de_passe
        '
        Me.Confirmer_mot_de_passe.AutoSize = True
        Me.Confirmer_mot_de_passe.BackColor = System.Drawing.Color.Transparent
        Me.Confirmer_mot_de_passe.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Confirmer_mot_de_passe.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Confirmer_mot_de_passe.Location = New System.Drawing.Point(2, 150)
        Me.Confirmer_mot_de_passe.Name = "Confirmer_mot_de_passe"
        Me.Confirmer_mot_de_passe.Size = New System.Drawing.Size(128, 13)
        Me.Confirmer_mot_de_passe.TabIndex = 10
        Me.Confirmer_mot_de_passe.Text = "Confirmer le mot de passe"
        '
        'Utilisateur_TextBox
        '
        Me.Utilisateur_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Utilisateur_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Utilisateur_TextBox.Location = New System.Drawing.Point(162, 98)
        Me.Utilisateur_TextBox.Name = "Utilisateur_TextBox"
        Me.Utilisateur_TextBox.Size = New System.Drawing.Size(205, 20)
        Me.Utilisateur_TextBox.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.Utilisateur_TextBox, "Le nom de l'utilisateur")
        '
        'Adresse_Serveur_TextBox
        '
        Me.Adresse_Serveur_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Adresse_Serveur_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Adresse_Serveur_TextBox.Location = New System.Drawing.Point(162, 72)
        Me.Adresse_Serveur_TextBox.Name = "Adresse_Serveur_TextBox"
        Me.Adresse_Serveur_TextBox.Size = New System.Drawing.Size(205, 20)
        Me.Adresse_Serveur_TextBox.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.Adresse_Serveur_TextBox, "e.g : 2.5.1.9 ou www.mon_serveur_mysql.fr")
        '
        'Nom_Connexion_TextBox
        '
        Me.Nom_Connexion_TextBox.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Nom_Connexion_TextBox.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Nom_Connexion_TextBox.Location = New System.Drawing.Point(162, 20)
        Me.Nom_Connexion_TextBox.Name = "Nom_Connexion_TextBox"
        Me.Nom_Connexion_TextBox.Size = New System.Drawing.Size(204, 20)
        Me.Nom_Connexion_TextBox.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Nom_Connexion_TextBox, "Le nom de la connexion")
        '
        'Creer_base
        '
        Me.Creer_base.AutoSize = True
        Me.Creer_base.BackColor = System.Drawing.Color.Transparent
        Me.Creer_base.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Creer_base.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Creer_base.Location = New System.Drawing.Point(5, 220)
        Me.Creer_base.Name = "Creer_base"
        Me.Creer_base.Size = New System.Drawing.Size(164, 17)
        Me.Creer_base.TabIndex = 10
        Me.Creer_base.Text = "Créer une nouvelle base vide"
        Me.ToolTip1.SetToolTip(Me.Creer_base, "Création de la base, et des tables")
        Me.Creer_base.UseVisualStyleBackColor = False
        '
        'Nom_de_la_base
        '
        Me.Nom_de_la_base.AutoSize = True
        Me.Nom_de_la_base.BackColor = System.Drawing.Color.Transparent
        Me.Nom_de_la_base.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Nom_de_la_base.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Nom_de_la_base.Location = New System.Drawing.Point(2, 179)
        Me.Nom_de_la_base.Name = "Nom_de_la_base"
        Me.Nom_de_la_base.Size = New System.Drawing.Size(81, 13)
        Me.Nom_de_la_base.TabIndex = 4
        Me.Nom_de_la_base.Text = "Nom de la base"
        '
        'Mot_de_passe_Utilisateur
        '
        Me.Mot_de_passe_Utilisateur.AutoSize = True
        Me.Mot_de_passe_Utilisateur.BackColor = System.Drawing.Color.Transparent
        Me.Mot_de_passe_Utilisateur.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Mot_de_passe_Utilisateur.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Mot_de_passe_Utilisateur.Location = New System.Drawing.Point(2, 125)
        Me.Mot_de_passe_Utilisateur.Name = "Mot_de_passe_Utilisateur"
        Me.Mot_de_passe_Utilisateur.Size = New System.Drawing.Size(71, 13)
        Me.Mot_de_passe_Utilisateur.TabIndex = 3
        Me.Mot_de_passe_Utilisateur.Text = "Mot de passe"
        '
        'Nom_Utilisateur
        '
        Me.Nom_Utilisateur.AutoSize = True
        Me.Nom_Utilisateur.BackColor = System.Drawing.Color.Transparent
        Me.Nom_Utilisateur.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Nom_Utilisateur.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Nom_Utilisateur.Location = New System.Drawing.Point(2, 99)
        Me.Nom_Utilisateur.Name = "Nom_Utilisateur"
        Me.Nom_Utilisateur.Size = New System.Drawing.Size(53, 13)
        Me.Nom_Utilisateur.TabIndex = 2
        Me.Nom_Utilisateur.Text = "Utilisateur"
        '
        'IP_SERVEUR
        '
        Me.IP_SERVEUR.AutoSize = True
        Me.IP_SERVEUR.BackColor = System.Drawing.Color.Transparent
        Me.IP_SERVEUR.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.IP_SERVEUR.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.IP_SERVEUR.Location = New System.Drawing.Point(2, 73)
        Me.IP_SERVEUR.Name = "IP_SERVEUR"
        Me.IP_SERVEUR.Size = New System.Drawing.Size(85, 13)
        Me.IP_SERVEUR.TabIndex = 1
        Me.IP_SERVEUR.Text = "Adresse Serveur"
        '
        'Nom_Connexion
        '
        Me.Nom_Connexion.AutoSize = True
        Me.Nom_Connexion.BackColor = System.Drawing.Color.Transparent
        Me.Nom_Connexion.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Nom_Connexion.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Nom_Connexion.Location = New System.Drawing.Point(3, 23)
        Me.Nom_Connexion.Name = "Nom_Connexion"
        Me.Nom_Connexion.Size = New System.Drawing.Size(106, 13)
        Me.Nom_Connexion.TabIndex = 0
        Me.Nom_Connexion.Text = "Titre de la connexion"
        '
        'Ajout_Suppression_connexion_base
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Pict0017j
        Me.ClientSize = New System.Drawing.Size(630, 451)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Ajout_Suppression_connexion_base"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "IDEALTAKE - Sélection - Ajout - Suppression de connexion"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.Flag_Italy, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Flag_France, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Flag_Spain, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Flag_USA, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Flag_Germany, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.Supprimer_une_connexion_GroupBox.ResumeLayout(False)
        Me.Supprimer_une_connexion_GroupBox.PerformLayout()
        Me.Ajouter_une_connexion_GroupBox1.ResumeLayout(False)
        Me.Ajouter_une_connexion_GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Ajouter_une_connexion_GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Nom_Connexion As System.Windows.Forms.Label
    Friend WithEvents IP_SERVEUR As System.Windows.Forms.Label
    Friend WithEvents Nom_Utilisateur As System.Windows.Forms.Label
    Friend WithEvents Mot_de_passe_Utilisateur As System.Windows.Forms.Label
    Friend WithEvents Nom_de_la_base As System.Windows.Forms.Label
    Friend WithEvents Creer_base As System.Windows.Forms.CheckBox
    Friend WithEvents Nom_Connexion_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Adresse_Serveur_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Utilisateur_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Confirmer_mot_de_passe As System.Windows.Forms.Label
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents mot_de_passe_MaskedTextBox As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Confimer_mot_de_passe_MaskedTextBox As System.Windows.Forms.MaskedTextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Nom_de_la_base_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Ajouter_une_connexion As System.Windows.Forms.Button
    Friend WithEvents Supprimer_une_connexion_GroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents Titre_de_la_connexion_Label1 As System.Windows.Forms.Label
    Friend WithEvents Selection_connexion_ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Supprimer_une_connexion As System.Windows.Forms.Button
    Friend WithEvents Etat_Ajout_Suppression_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Selection_connexion_ComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents Connect_base As System.Windows.Forms.Button
    Friend WithEvents Titre_de_la_connexion_Label As System.Windows.Forms.Label
    Friend WithEvents Label_Chemin_Schéma As System.Windows.Forms.Label
    Friend WithEvents Localisation_Schéma_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label_Chemin_Synoptiques As System.Windows.Forms.Label
    Friend WithEvents Localisation_Synoptique_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label_Chemins_fichiers As System.Windows.Forms.Label
    Friend WithEvents Label_Notice As System.Windows.Forms.Label
    Friend WithEvents Localisation_Notice_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label_Fiche_Produit As System.Windows.Forms.Label
    Friend WithEvents Localisation_Fiche_Produit_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label_Photos_Produit As System.Windows.Forms.Label
    Friend WithEvents Localisation_Photo_produit_TextBox As System.Windows.Forms.TextBox

    Private Sub Ajout_Suppresion_connexion_base_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Form_Demarrage.Close() 'ferme complétement le logiciel
    End Sub
    Friend WithEvents Flag_Italy As System.Windows.Forms.PictureBox
    Friend WithEvents Flag_France As System.Windows.Forms.PictureBox
    Friend WithEvents Flag_Spain As System.Windows.Forms.PictureBox
    Friend WithEvents Description_connexion_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Description_Connexion_ComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents Description_connexion_Label As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Flag_Germany As System.Windows.Forms.PictureBox
    Friend WithEvents Flag_USA As System.Windows.Forms.PictureBox
End Class
