<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Parametres
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Parametres))
        Me.Parametre_Gestion_Base = New System.Windows.Forms.TabControl()
        Me.Parametres_TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Activation_debugage_Script = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Sélection_Connexion_Vérouillé = New System.Windows.Forms.CheckBox()
        Me.Titre_Connexion = New System.Windows.Forms.Label()
        Me.Chaine_connexion = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.NIVEAU_ACCES = New System.Windows.Forms.ComboBox()
        Me.Parametres_TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox_Menus_Systeme = New System.Windows.Forms.GroupBox()
        Me.Collections_Menus_LISTE_DES_TABLES_Label2 = New System.Windows.Forms.Label()
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1 = New System.Windows.Forms.Label()
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1 = New System.Windows.Forms.ListBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ActivationMaintenanceFTP = New System.Windows.Forms.CheckBox()
        Me.Activation_Modification_Objet = New System.Windows.Forms.CheckBox()
        Me.Collections_Menus_Lecture_Table = New System.Windows.Forms.Button()
        Me.mise_a_jour_menu_systeme = New System.Windows.Forms.Button()
        Me.Parametre_Gestion_Base.SuspendLayout()
        Me.Parametres_TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Parametres_TabPage2.SuspendLayout()
        Me.GroupBox_Menus_Systeme.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Parametre_Gestion_Base
        '
        Me.Parametre_Gestion_Base.Controls.Add(Me.Parametres_TabPage1)
        Me.Parametre_Gestion_Base.Controls.Add(Me.Parametres_TabPage2)
        Me.Parametre_Gestion_Base.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Parametre_Gestion_Base.Location = New System.Drawing.Point(0, 0)
        Me.Parametre_Gestion_Base.Name = "Parametre_Gestion_Base"
        Me.Parametre_Gestion_Base.SelectedIndex = 0
        Me.Parametre_Gestion_Base.Size = New System.Drawing.Size(742, 531)
        Me.Parametre_Gestion_Base.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.Parametre_Gestion_Base.TabIndex = 0
        '
        'Parametres_TabPage1
        '
        Me.Parametres_TabPage1.BackColor = System.Drawing.Color.LightCyan
        Me.Parametres_TabPage1.Controls.Add(Me.GroupBox2)
        Me.Parametres_TabPage1.Controls.Add(Me.GroupBox1)
        Me.Parametres_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.Parametres_TabPage1.Name = "Parametres_TabPage1"
        Me.Parametres_TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.Parametres_TabPage1.Size = New System.Drawing.Size(734, 505)
        Me.Parametres_TabPage1.TabIndex = 0
        Me.Parametres_TabPage1.Text = "Paramètres Serveur"
        Me.Parametres_TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.DarkTurquoise
        Me.GroupBox2.Controls.Add(Me.ActivationMaintenanceFTP)
        Me.GroupBox2.Controls.Add(Me.Activation_debugage_Script)
        Me.GroupBox2.Controls.Add(Me.Activation_Modification_Objet)
        Me.GroupBox2.Location = New System.Drawing.Point(374, 6)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(333, 466)
        Me.GroupBox2.TabIndex = 29
        Me.GroupBox2.TabStop = False
        '
        'Activation_debugage_Script
        '
        Me.Activation_debugage_Script.AutoSize = True
        Me.Activation_debugage_Script.Location = New System.Drawing.Point(14, 38)
        Me.Activation_debugage_Script.Name = "Activation_debugage_Script"
        Me.Activation_debugage_Script.Size = New System.Drawing.Size(154, 17)
        Me.Activation_debugage_Script.TabIndex = 1
        Me.Activation_debugage_Script.Text = "Activation debugage Script"
        Me.Activation_debugage_Script.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.DarkTurquoise
        Me.GroupBox1.Controls.Add(Me.Sélection_Connexion_Vérouillé)
        Me.GroupBox1.Controls.Add(Me.Titre_Connexion)
        Me.GroupBox1.Controls.Add(Me.Chaine_connexion)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Label57)
        Me.GroupBox1.Controls.Add(Me.Label47)
        Me.GroupBox1.Controls.Add(Me.NIVEAU_ACCES)
        Me.GroupBox1.Location = New System.Drawing.Point(11, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(356, 466)
        Me.GroupBox1.TabIndex = 28
        Me.GroupBox1.TabStop = False
        '
        'Sélection_Connexion_Vérouillé
        '
        Me.Sélection_Connexion_Vérouillé.AutoSize = True
        Me.Sélection_Connexion_Vérouillé.BackColor = System.Drawing.Color.Transparent
        Me.Sélection_Connexion_Vérouillé.Location = New System.Drawing.Point(109, 282)
        Me.Sélection_Connexion_Vérouillé.Name = "Sélection_Connexion_Vérouillé"
        Me.Sélection_Connexion_Vérouillé.Size = New System.Drawing.Size(178, 17)
        Me.Sélection_Connexion_Vérouillé.TabIndex = 8
        Me.Sélection_Connexion_Vérouillé.Text = "Toujours utiliser cette connexion"
        Me.Sélection_Connexion_Vérouillé.UseVisualStyleBackColor = False
        '
        'Titre_Connexion
        '
        Me.Titre_Connexion.AutoSize = True
        Me.Titre_Connexion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Titre_Connexion.Location = New System.Drawing.Point(6, 16)
        Me.Titre_Connexion.Name = "Titre_Connexion"
        Me.Titre_Connexion.Size = New System.Drawing.Size(139, 13)
        Me.Titre_Connexion.TabIndex = 28
        Me.Titre_Connexion.Text = "Titre de la connexion : "
        '
        'Chaine_connexion
        '
        Me.Chaine_connexion.Location = New System.Drawing.Point(6, 59)
        Me.Chaine_connexion.Multiline = True
        Me.Chaine_connexion.Name = "Chaine_connexion"
        Me.Chaine_connexion.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Chaine_connexion.Size = New System.Drawing.Size(344, 155)
        Me.Chaine_connexion.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = Global.IDEALTAKE.My.Resources.Resources.bouton_nue1
        Me.Button1.Location = New System.Drawing.Point(135, 320)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(118, 36)
        Me.Button1.TabIndex = 9
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.BackColor = System.Drawing.Color.Transparent
        Me.Label57.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label57.Location = New System.Drawing.Point(6, 43)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(162, 13)
        Me.Label57.TabIndex = 15
        Me.Label57.Text = "Chaîne de connexion au serveur"
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.BackColor = System.Drawing.Color.Transparent
        Me.Label47.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label47.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label47.Location = New System.Drawing.Point(14, 241)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(95, 13)
        Me.Label47.TabIndex = 20
        Me.Label47.Text = "Niveau d'acces"
        '
        'NIVEAU_ACCES
        '
        Me.NIVEAU_ACCES.Cursor = System.Windows.Forms.Cursors.Hand
        Me.NIVEAU_ACCES.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.NIVEAU_ACCES.FormattingEnabled = True
        Me.NIVEAU_ACCES.Location = New System.Drawing.Point(109, 239)
        Me.NIVEAU_ACCES.Name = "NIVEAU_ACCES"
        Me.NIVEAU_ACCES.Size = New System.Drawing.Size(198, 21)
        Me.NIVEAU_ACCES.TabIndex = 7
        '
        'Parametres_TabPage2
        '
        Me.Parametres_TabPage2.BackColor = System.Drawing.Color.LightCyan
        Me.Parametres_TabPage2.Controls.Add(Me.GroupBox_Menus_Systeme)
        Me.Parametres_TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.Parametres_TabPage2.Name = "Parametres_TabPage2"
        Me.Parametres_TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.Parametres_TabPage2.Size = New System.Drawing.Size(734, 505)
        Me.Parametres_TabPage2.TabIndex = 1
        Me.Parametres_TabPage2.Text = "Collections Menus "
        Me.Parametres_TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox_Menus_Systeme
        '
        Me.GroupBox_Menus_Systeme.BackColor = System.Drawing.Color.DarkTurquoise
        Me.GroupBox_Menus_Systeme.Controls.Add(Me.Collections_Menus_LISTE_DES_TABLES_Label2)
        Me.GroupBox_Menus_Systeme.Controls.Add(Me.Collections_menus_TABLE_EN_LECTURE_Label_1)
        Me.GroupBox_Menus_Systeme.Controls.Add(Me.Collections_Menus_Lecture_Table)
        Me.GroupBox_Menus_Systeme.Controls.Add(Me.Collections_Menus_LISTE_DES_TABLES_ListBox1)
        Me.GroupBox_Menus_Systeme.Controls.Add(Me.mise_a_jour_menu_systeme)
        Me.GroupBox_Menus_Systeme.Controls.Add(Me.DataGridView1)
        Me.GroupBox_Menus_Systeme.Controls.Add(Me.FlowLayoutPanel1)
        Me.GroupBox_Menus_Systeme.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox_Menus_Systeme.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox_Menus_Systeme.Name = "GroupBox_Menus_Systeme"
        Me.GroupBox_Menus_Systeme.Size = New System.Drawing.Size(728, 499)
        Me.GroupBox_Menus_Systeme.TabIndex = 0
        Me.GroupBox_Menus_Systeme.TabStop = False
        '
        'Collections_Menus_LISTE_DES_TABLES_Label2
        '
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.AutoSize = True
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.Location = New System.Drawing.Point(7, 338)
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.Name = "Collections_Menus_LISTE_DES_TABLES_Label2"
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.Size = New System.Drawing.Size(146, 13)
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.TabIndex = 4
        Me.Collections_Menus_LISTE_DES_TABLES_Label2.Text = "Liste des Tables de la Base : "
        '
        'Collections_menus_TABLE_EN_LECTURE_Label_1
        '
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.AutoEllipsis = True
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.AutoSize = True
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.Location = New System.Drawing.Point(513, 357)
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.Name = "Collections_menus_TABLE_EN_LECTURE_Label_1"
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.Size = New System.Drawing.Size(22, 13)
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.TabIndex = 6
        Me.Collections_menus_TABLE_EN_LECTURE_Label_1.Text = "- - -"
        '
        'Collections_Menus_LISTE_DES_TABLES_ListBox1
        '
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.ColumnWidth = 200
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.FormattingEnabled = True
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.HorizontalScrollbar = True
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.Location = New System.Drawing.Point(7, 357)
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.MultiColumn = True
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.Name = "Collections_Menus_LISTE_DES_TABLES_ListBox1"
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.ScrollAlwaysVisible = True
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.Size = New System.Drawing.Size(500, 132)
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.Sorted = True
        Me.Collections_Menus_LISTE_DES_TABLES_ListBox1.TabIndex = 5
        '
        'DataGridView1
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView1.GridColor = System.Drawing.Color.DarkRed
        Me.DataGridView1.Location = New System.Drawing.Point(10, 19)
        Me.DataGridView1.Name = "DataGridView1"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(712, 316)
        Me.DataGridView1.TabIndex = 3
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 16)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(722, 480)
        Me.FlowLayoutPanel1.TabIndex = 9
        '
        'ActivationMaintenanceFTP
        '
        Me.ActivationMaintenanceFTP.AutoSize = True
        Me.ActivationMaintenanceFTP.Checked = Global.IDEALTAKE.My.MySettings.Default.MaintenanceTransfertIdealtake
        Me.ActivationMaintenanceFTP.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Global.IDEALTAKE.My.MySettings.Default, "MaintenanceTransfertIdealtake", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.ActivationMaintenanceFTP.Location = New System.Drawing.Point(14, 59)
        Me.ActivationMaintenanceFTP.Name = "ActivationMaintenanceFTP"
        Me.ActivationMaintenanceFTP.Size = New System.Drawing.Size(177, 17)
        Me.ActivationMaintenanceFTP.TabIndex = 2
        Me.ActivationMaintenanceFTP.Text = "Maintenance TransfertIdealtake"
        Me.ActivationMaintenanceFTP.UseVisualStyleBackColor = True
        '
        'Activation_Modification_Objet
        '
        Me.Activation_Modification_Objet.AutoSize = True
        Me.Activation_Modification_Objet.Checked = Global.IDEALTAKE.My.MySettings.Default.Modification_Objet
        Me.Activation_Modification_Objet.DataBindings.Add(New System.Windows.Forms.Binding("Checked", Global.IDEALTAKE.My.MySettings.Default, "Modification_Objet", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Activation_Modification_Objet.Location = New System.Drawing.Point(14, 18)
        Me.Activation_Modification_Objet.Margin = New System.Windows.Forms.Padding(2)
        Me.Activation_Modification_Objet.Name = "Activation_Modification_Objet"
        Me.Activation_Modification_Objet.Size = New System.Drawing.Size(191, 17)
        Me.Activation_Modification_Objet.TabIndex = 0
        Me.Activation_Modification_Objet.Text = "Modification objet 'Windows.Forms'"
        Me.Activation_Modification_Objet.UseVisualStyleBackColor = True
        '
        'Collections_Menus_Lecture_Table
        '
        Me.Collections_Menus_Lecture_Table.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Collections_Menus_Lecture_Table.BackColor = System.Drawing.Color.Silver
        Me.Collections_Menus_Lecture_Table.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Collections_Menus_Lecture_Table.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Collections_Menus_Lecture_Table.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.Collections_Menus_Lecture_Table.Location = New System.Drawing.Point(513, 466)
        Me.Collections_Menus_Lecture_Table.Name = "Collections_Menus_Lecture_Table"
        Me.Collections_Menus_Lecture_Table.Size = New System.Drawing.Size(80, 22)
        Me.Collections_Menus_Lecture_Table.TabIndex = 7
        Me.Collections_Menus_Lecture_Table.Text = "Lecture"
        Me.Collections_Menus_Lecture_Table.UseVisualStyleBackColor = False
        '
        'mise_a_jour_menu_systeme
        '
        Me.mise_a_jour_menu_systeme.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.mise_a_jour_menu_systeme.BackColor = System.Drawing.Color.Silver
        Me.mise_a_jour_menu_systeme.DataBindings.Add(New System.Windows.Forms.Binding("ForeColor", Global.IDEALTAKE.My.MySettings.Default, "Couleur_Font", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.mise_a_jour_menu_systeme.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mise_a_jour_menu_systeme.ForeColor = Global.IDEALTAKE.My.MySettings.Default.Couleur_Font
        Me.mise_a_jour_menu_systeme.Location = New System.Drawing.Point(599, 466)
        Me.mise_a_jour_menu_systeme.Name = "mise_a_jour_menu_systeme"
        Me.mise_a_jour_menu_systeme.Size = New System.Drawing.Size(92, 22)
        Me.mise_a_jour_menu_systeme.TabIndex = 8
        Me.mise_a_jour_menu_systeme.Text = "Mise à Jour"
        Me.mise_a_jour_menu_systeme.UseVisualStyleBackColor = False
        '
        'Parametres
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.IDEALTAKE.My.Resources.Resources.Pict0017j
        Me.ClientSize = New System.Drawing.Size(742, 531)
        Me.Controls.Add(Me.Parametre_Gestion_Base)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "Parametres"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parametres"
        Me.Parametre_Gestion_Base.ResumeLayout(False)
        Me.Parametres_TabPage1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Parametres_TabPage2.ResumeLayout(False)
        Me.GroupBox_Menus_Systeme.ResumeLayout(False)
        Me.GroupBox_Menus_Systeme.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Parametre_Gestion_Base As System.Windows.Forms.TabControl
    Friend WithEvents Parametres_TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Parametres_TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents NIVEAU_ACCES As System.Windows.Forms.ComboBox
    Friend WithEvents Chaine_connexion As System.Windows.Forms.TextBox
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox_Menus_Systeme As System.Windows.Forms.GroupBox
    Friend WithEvents mise_a_jour_menu_systeme As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Titre_Connexion As System.Windows.Forms.Label
    Friend WithEvents Sélection_Connexion_Vérouillé As System.Windows.Forms.CheckBox
    Friend WithEvents Collections_Menus_Lecture_Table As System.Windows.Forms.Button
    Friend WithEvents Collections_Menus_LISTE_DES_TABLES_ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Collections_menus_TABLE_EN_LECTURE_Label_1 As System.Windows.Forms.Label
    Friend WithEvents Collections_Menus_LISTE_DES_TABLES_Label2 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Activation_Modification_Objet As System.Windows.Forms.CheckBox
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents Activation_debugage_Script As System.Windows.Forms.CheckBox
    Friend WithEvents ActivationMaintenanceFTP As System.Windows.Forms.CheckBox
End Class
