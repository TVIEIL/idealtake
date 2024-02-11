Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe enregistrement regroupe les infos de chaque enregistrement.
    ''' </summary>
    Public Class enregistrement

        Private m_Numero_Enregistrement As String = ""
        Private m_Chrono As Integer = -1
        Private m_Auteur As String = ""
        Private m_Date_de_creation As Date

        Private m_Titre_texte_francais As String = ""
        Private m_Sous_Titre_texte_francais As String = ""
        Private m_Commentaire_francais As String = ""

        Private m_Titre_texte_Allemand As String = ""
        Private m_Sous_Titre_texte_Allemand As String = ""
        Private m_Commentaire_allemand As String = ""

        Private m_Titre_texte_Anglais As String = ""
        Private m_Sous_Titre_texte_Anglais As String = ""
        Private m_Commentaire_anglais As String = ""


        Private m_Titre_texte_Espagnole As String = ""
        Private m_Sous_Titre_texte_Espagnole As String = ""
        Private m_Commentaire_espagnole As String = ""

        Private m_Titre_texte_Italien As String = ""
        Private m_Sous_Titre_texte_Italien As String = ""
        Private m_Commentaire_italien As String = ""

        Private m_Activation_XAML As Boolean = False
        Private m_Type_XAML_Titre_predefini As Integer = -1
        Private m_Type_XAML_Sous_Titre_predefini As Integer = -1
        Private m_Type_XAML_Commentaire_predefini As Integer = -1
        Private m_Type_XAML_Titre_specifique As String = ""
        Private m_Type_XAML_Sous_Titre_specifique As String = ""
        Private m_Type_XAML_Commentaire_specifique As String = ""

        Private m_PARAM_A As Integer = -1    'Anciennement : m_Verrouillage
        Private m_PARAM_B As Integer = -1    'Anciennement : m_Audio_Video
        Private m_PARAM_C As Integer = -1    'Anciennement : m_systeme_Interphonie
        Private m_PARAM_D As Integer = -1    'Anciennement : m_Principale
        Private m_PARAM_E As Integer = -1    'Anciennement : m_Secondaire
        Private m_PARAM_F As Integer = -1
        Private m_PARAM_G As Integer = -1
        Private m_PARAM_H As Integer = -1
        Private m_PARAM_I As Integer = -1

        Private m_Recherche_par_Arbre As String = ""
        Private m_Label_reserve1_francais As String = ""
        Private m_Texte_reserve1_francais As String = ""
        Private m_Label_reserve1_Allemand As String = ""
        Private m_Texte_reserve1_Allemand As String = ""
        Private m_Label_reserve1_Anglais As String = ""
        Private m_Texte_reserve1_Anglais As String = ""
        Private m_Label_reserve1_Espagnole As String = ""
        Private m_Texte_reserve1_Espagnole As String = ""
        Private m_Label_reserve1_Italien As String = ""
        Private m_Texte_reserve1_Italien As String = ""

        Private m_Boolean_reserve1 As Boolean = False
        Private m_Integer_reserve1 As Integer = -1
        Private m_Date_reserve1 As Date
        Private m_IDConstructeur As Integer = -1
        Private m_Niveau_acces As Integer = -1
        Private m_STANDARD As Boolean = False
        Private m_Derniere_modification As Date
        Private m_Auteur_derniere_modification As String = ""
        Private m_NUMERO_MODIFICATION As Integer = -1

        Private m_Type_Doc As Integer = -1




        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Numero_Enregistrement = "ENR001"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Numero_Enregistrement</c> indique par une chaîne le nom unique de l'enregistrement.
        ''' </summary>
        Public Property Numero_Enregistrement() As String
            Get
                Return m_Numero_Enregistrement

            End Get
            Set(ByVal value As String)
                m_Numero_Enregistrement = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Chrono = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Chrono</c> indique par un entier le numéro unique de l'enregistrement.
        ''' </summary>
        Public Property Chrono() As Integer
            Get
                Return m_Chrono

            End Get
            Set(ByVal value As Integer)
                m_Chrono = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Auteur = "TVIEIL"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Auteur</c> indique par une chaîne le nom de l'auteur de l'enregistrement.
        ''' </summary>
        Public Property Auteur() As String
            Get
                Return m_Auteur

            End Get
            Set(ByVal value As String)
                m_Auteur = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''num_Enr1.Date_de_creation = "17/12/2007"
        ''' MsgBox(num_Enr1.Date_ENR.Day)
        ''' MsgBox(num_Enr1.Date_ENR.Month)
        ''' MsgBox(num_Enr1.Date_ENR.Year)
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Date_de_creation</c> indique la date de création d'un enregistrement.
        ''' </summary>
        Public Property Date_de_creation() As Date
            Get
                Return m_Date_de_creation

            End Get
            Set(ByVal value As Date)
                m_Date_de_creation = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Titre_francais_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Titre_francais_Enr1.Titre_texte_francais = "Titre du document en français"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Titre_texte_francais</c> indique par une chaîne en Français le titre d'un enregistrement.
        ''' </summary>
        Public Property Titre_texte_francais() As String
            Get
                Return m_Titre_texte_francais

            End Get
            Set(ByVal value As String)
                m_Titre_texte_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Sous_Titre_francais_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Sous_Titre_francais_Enr1.Sous_Titre_texte_francais = "Sous Titre du document en français"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Sous_Titre_texte_francais</c> indique par une chaîne le sous titre en Français d'un enregistrement.
        ''' </summary>
        Public Property Sous_Titre_texte_francais() As String
            Get
                Return m_Sous_Titre_texte_francais

            End Get
            Set(ByVal value As String)
                m_Sous_Titre_texte_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Commentaire_francais_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Commentaire_francais_Enr1.Commentaire_francais = "Ce document traite de..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Commentaire_francais</c> indique par une chaîne en Français décrivant un enregistrement.
        ''' </summary>
        Public Property Commentaire_francais() As String
            Get
                Return m_Commentaire_francais

            End Get
            Set(ByVal value As String)
                m_Commentaire_francais = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Titre_Allemand_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Titre_francais_Enr1.Titre_texte_Allemand = "Titre du document en Allemand"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Titre_texte_Allemand</c> indique par une chaîne du titre en Allemand d'un enregistrement.
        ''' </summary>
        Public Property Titre_texte_Allemand() As String
            Get
                Return m_Titre_texte_Allemand

            End Get
            Set(ByVal value As String)
                m_Titre_texte_Allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Sous_Titre_Allemand_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Sous_Titre_Allemand_Enr1.Sous_Titre_texte_Allemand = "Sous Titre du document en Allemand"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Sous_Titre_texte_Allemand</c> indique par une chaîne du sous titre en Allemand d'un enregistrement.
        ''' </summary>
        Public Property Sous_Titre_texte_Allemand() As String
            Get
                Return m_Sous_Titre_texte_Allemand

            End Get
            Set(ByVal value As String)
                m_Sous_Titre_texte_Allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Commentaire_allemand_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Commentaire_allemand_Enr1.Commentaire_anglais = "Ce document traite de..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Commentaire_allemand</c> indique par une chaîne en Allemand décrivant un enregistrement.
        ''' </summary>
        Public Property Commentaire_allemand() As String
            Get
                Return m_Commentaire_allemand

            End Get
            Set(ByVal value As String)
                m_Commentaire_allemand = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Titre_Anglais_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Titre_Anglais_Enr1.Titre_texte_Anglais = "Titre du document en Anglais"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Titre_texte_Anglais</c> indique par une chaîne du titre en Anglais d'un enregistrement.
        ''' </summary>
        Public Property Titre_texte_Anglais() As String
            Get
                Return m_Titre_texte_Anglais

            End Get
            Set(ByVal value As String)
                m_Titre_texte_Anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Sous_Titre_Anglais_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Sous_Titre_Anglais_Enr1.Sous_Titre_texte_Anglais = "Sous Titre du document en Anglais"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Sous_Titre_texte_Anglais</c> indique par une chaîne du sous titre en Anglais d'un enregistrement.
        ''' </summary>
        Public Property Sous_Titre_texte_Anglais() As String
            Get
                Return m_Sous_Titre_texte_Anglais

            End Get
            Set(ByVal value As String)
                m_Sous_Titre_texte_Anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Commentaire_anglais_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Commentaire_anglais_Enr1.Commentaire_anglais = "Ce document traite de..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Commentaire_anglais</c> indique par une chaîne en Anglais décrivant un enregistrement.
        ''' </summary>
        Public Property Commentaire_anglais() As String
            Get
                Return m_Commentaire_anglais

            End Get
            Set(ByVal value As String)
                m_Commentaire_anglais = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Titre_texte_Espagnole_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Titre_texte_Espagnole_Enr1.Titre_texte_Espagnole = "Titre du document en Espagnole"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Titre_texte_Espagnole</c> indique par une chaîne du titre en Espagnole d'un enregistrement.
        ''' </summary>
        Public Property Titre_texte_Espagnole() As String
            Get
                Return m_Titre_texte_Espagnole

            End Get
            Set(ByVal value As String)
                m_Titre_texte_Espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Sous_Titre_texte_Espagnole_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Sous_Titre_texte_Espagnole_Enr1.Sous_Titre_texte_Espagnole = "Sous Titre du document en Espagnole"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Sous_Titre_texte_Espagnole</c> indique par une chaîne du sous titre en Espagnole d'un enregistrement.
        ''' </summary>
        Public Property Sous_Titre_texte_Espagnole() As String
            Get
                Return m_Sous_Titre_texte_Espagnole

            End Get
            Set(ByVal value As String)
                m_Sous_Titre_texte_Espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Commentaire_espagnole_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Commentaire_espagnole_Enr1.Commentaire_anglais = "Ce document traite de..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Commentaire_espagnole</c> indique par une chaîne en Espagnole décrivant un enregistrement.
        ''' </summary>
        Public Property Commentaire_espagnole() As String
            Get
                Return m_Commentaire_espagnole

            End Get
            Set(ByVal value As String)
                m_Commentaire_espagnole = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Titre_texte_Italien_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Titre_texte_Italien_Enr1.Titre_texte_Italien = "Titre du document en Italien"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Titre_texte_Italien</c> indique par une chaîne du titre en Italien d'un enregistrement.
        ''' </summary>
        Public Property Titre_texte_Italien() As String
            Get
                Return m_Titre_texte_Italien

            End Get
            Set(ByVal value As String)
                m_Titre_texte_Italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Sous_Titre_texte_Italien_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Sous_Titre_texte_Italien_Enr1.Sous_Titre_texte_Italien = "Sous Titre du document en Italien"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Sous_Titre_texte_Italien</c> indique par une chaîne du sous titre en Italien d'un enregistrement.
        ''' </summary>
        Public Property Sous_Titre_texte_Italien() As String
            Get
                Return m_Sous_Titre_texte_Italien

            End Get
            Set(ByVal value As String)
                m_Sous_Titre_texte_Italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Commentaire_italien_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Commentaire_italien_Enr1.Commentaire_anglais = "Ce document traite de..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Commentaire_italien</c> indique par une chaîne en Italien décrivant un enregistrement.
        ''' </summary>
        Public Property Commentaire_italien() As String
            Get
                Return m_Commentaire_italien

            End Get
            Set(ByVal value As String)
                m_Commentaire_italien = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Activation_XAML = TRUE
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Activation_XAML</c> indique par un booléen l'utilisation de code XAML pour la présentation.
        ''' </summary>
        Public Property Activation_XAML() As Boolean
            Get
                Return m_Activation_XAML

            End Get
            Set(ByVal value As Boolean)
                m_Activation_XAML = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim xaml_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''xaml_Enr1.Type_XAML_Titre_predefini = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_XAML_Titre_predefini</c> entier se rapportant à la table 'table_xaml'codant la présentation du titre en XAML.
        ''' </summary>
        Public Property Type_XAML_Titre_predefini() As Integer
            Get
                Return m_Type_XAML_Titre_predefini

            End Get
            Set(ByVal value As Integer)
                m_Type_XAML_Titre_predefini = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim xaml_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''xaml_Enr1.Type_XAML_Sous_Titre_predefini = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_XAML_Sous_Titre_predefini</c> entier se rapportant à la table 'table_xaml'codant la présentation du sous titre en XAML.
        ''' </summary>
        Public Property Type_XAML_Sous_Titre_predefini() As Integer
            Get
                Return m_Type_XAML_Sous_Titre_predefini

            End Get
            Set(ByVal value As Integer)
                m_Type_XAML_Sous_Titre_predefini = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim xaml_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''xaml_Enr1.Type_XAML_Commentaire_predefini = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_XAML_Commentaire_predefini</c> entier se rapportant à la table 'table_xaml'codant la présentation du commentaire en XAML.
        ''' </summary>
        Public Property Type_XAML_Commentaire_predefini() As Integer
            Get
                Return m_Type_XAML_Commentaire_predefini

            End Get
            Set(ByVal value As Integer)
                m_Type_XAML_Commentaire_predefini = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim xaml_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''xaml_Enr1.Type_XAML_Titre_specifique = "...Canvas xmlns=..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_XAML_Titre_specifique</c> est un chaîne qui code la présentation du titre en XAML.
        ''' </summary>
        Public Property Type_XAML_Titre_specifique() As String
            Get
                Return m_Type_XAML_Titre_specifique

            End Get
            Set(ByVal value As String)
                m_Type_XAML_Titre_specifique = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim xaml_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''xaml_Enr1.Type_XAML_Sous_Titre_specifique = "...Canvas xmlns=..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_XAML_Sous_Titre_specifique</c> est un chaîne qui code la présentation du sous titre en XAML.
        ''' </summary>
        Public Property Type_XAML_Sous_Titre_specifique() As String
            Get
                Return m_Type_XAML_Sous_Titre_specifique

            End Get
            Set(ByVal value As String)
                m_Type_XAML_Sous_Titre_specifique = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim xaml_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''xaml_Enr1.Type_XAML_Commentaire_specifique = "...Canvas xmlns=..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_XAML_Commentaire_specifique</c> est un chaîne qui code la présentation du commentaire en XAML.
        ''' </summary>
        Public Property Type_XAML_Commentaire_specifique() As String
            Get
                Return m_Type_XAML_Commentaire_specifique

            End Get
            Set(ByVal value As String)
                m_Type_XAML_Commentaire_specifique = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim ID_PARAM_A As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''ID_PARAM_A.PARAM_A = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_A</c> indique par un entier le type de verrouillage d'un enregistrement(Lie Table_A).
        ''' </summary>
        Public Property PARAM_A() As Integer
            Get
                Return m_PARAM_A

            End Get
            Set(ByVal value As Integer)
                m_PARAM_A = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_systeme_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Type_systeme_Enr1.PARAM_B = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_B</c> indique par un entier le type d'installation audio, vidéo, mixte d'un enregistrement(Lie Table_B).
        ''' </summary>
        Public Property PARAM_B() As Integer
            Get
                Return m_PARAM_B

            End Get
            Set(ByVal value As Integer)
                m_PARAM_B = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim systeme_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''systeme_Enr1.PARAM_C = 4
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_C</c> indique par un entier le type de système interphonie d'un enregistrement(Lie Table_C).
        ''' </summary>
        Public Property PARAM_C() As Integer
            Get
                Return m_PARAM_C

            End Get
            Set(ByVal value As Integer)
                m_PARAM_C = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Platine_Pri_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Platine_Pri_Enr1.PARAM_D = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_D</c> indique par un entier Lie à Table_D (serie - PLATINE).
        ''' </summary>
        Public Property PARAM_D() As Integer
            Get
                Return m_PARAM_D

            End Get
            Set(ByVal value As Integer)
                m_PARAM_D = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Platine_sec_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Platine_sec_Enr1.PARAM_E = 0
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_E</c> indique par un entier Lie Table_E (CONTROLE_ACCESS).
        ''' </summary>
        Public Property PARAM_E() As Integer
            Get
                Return m_PARAM_E

            End Get
            Set(ByVal value As Integer)
                m_PARAM_E = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Platine_sec_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Platine_sec_Enr1.PARAM_F = 0
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_F</c> indique par un entier Non lie -NBRE_PRI_AUDIO-.
        ''' </summary>
        Public Property PARAM_F() As Integer
            Get
                Return m_PARAM_F

            End Get
            Set(ByVal value As Integer)
                m_PARAM_F = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Platine_sec_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Platine_sec_Enr1.PARAM_G = 0
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_G</c> indique par un entier Non lie -NBRE_PRI_VIDEO-.
        ''' </summary>
        Public Property PARAM_G() As Integer
            Get
                Return m_PARAM_G

            End Get
            Set(ByVal value As Integer)
                m_PARAM_G = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Platine_sec_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Platine_sec_Enr1.PARAM_H = 0
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_H</c> indique par un entier Non lie -NBRE_SEC_AUDIO-.
        ''' </summary>
        Public Property PARAM_H() As Integer
            Get
                Return m_PARAM_H

            End Get
            Set(ByVal value As Integer)
                m_PARAM_H = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Platine_sec_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Platine_sec_Enr1.PARAM_I = 0
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>PARAM_I</c> indique par un entier Non lie -NBRE_SEC_VIDEO-.
        ''' </summary>
        Public Property PARAM_I() As Integer
            Get
                Return m_PARAM_I

            End Get
            Set(ByVal value As Integer)
                m_PARAM_I = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Recherche_par_Arbre = "R0=2;R-1=1;R-2=4"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Recherche_par_Arbre</c> indique par une chaîne la position dans un arbre de recherche de l'enregistrement(cré à partir des
        ''' valeurs des tables Racine_zero à racine-9).
        ''' </summary>
        Public Property Recherche_par_Arbre() As String
            Get
                Return m_Recherche_par_Arbre

            End Get
            Set(ByVal value As String)
                m_Recherche_par_Arbre = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Label_reserve1_francais = "Label de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Label_reserve1_francais</c> indique par une chaîne un label complémentaire en Français dans l'enregistrement.
        ''' </summary>
        Public Property Label_reserve1_francais() As String
            Get
                Return m_Label_reserve1_francais

            End Get
            Set(ByVal value As String)
                m_Label_reserve1_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Texte_reserve1_francais = "texte de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Texte_reserve1_francais</c> indique par une chaîne un texte complémentaire en Français dans l'enregistrement.
        ''' </summary>
        Public Property Texte_reserve1_francais() As String
            Get
                Return m_Texte_reserve1_francais

            End Get
            Set(ByVal value As String)
                m_Texte_reserve1_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Label_reserve1_Allemand = "Label de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Label_reserve1_Allemand</c> indique par une chaîne un label complémentaire en Allemand dans l'enregistrement.
        ''' </summary>
        Public Property Label_reserve1_Allemand() As String
            Get
                Return m_Label_reserve1_Allemand

            End Get
            Set(ByVal value As String)
                m_Label_reserve1_Allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Texte_reserve1_Allemand = "Texte de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Texte_reserve1_Allemand</c> indique par une chaîne un texte complémentaire en Allemand dans l'enregistrement.
        ''' </summary>
        Public Property Texte_reserve1_Allemand() As String
            Get
                Return m_Texte_reserve1_Allemand

            End Get
            Set(ByVal value As String)
                m_Texte_reserve1_Allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Label_reserve1_Anglais = "Label de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Label_reserve1_Anglais</c> indique par une chaîne un label complémentaire en Anglais dans l'enregistrement.
        ''' </summary>
        Public Property Label_reserve1_Anglais() As String
            Get
                Return m_Label_reserve1_Anglais

            End Get
            Set(ByVal value As String)
                m_Label_reserve1_Anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Texte_reserve1_Anglais = "Texte de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Texte_reserve1_Anglais</c> indique par une chaîne un texte complémentaire en Anglais dans l'enregistrement.
        ''' </summary>
        Public Property Texte_reserve1_Anglais() As String
            Get
                Return m_Texte_reserve1_Anglais

            End Get
            Set(ByVal value As String)
                m_Texte_reserve1_Anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Label_reserve1_Espagnole = "Label de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Label_reserve1_Espagnole</c> indique par une chaîne un label complémentaire en Espagnole dans l'enregistrement.
        ''' </summary>
        Public Property Label_reserve1_Espagnole() As String
            Get
                Return m_Label_reserve1_Espagnole

            End Get
            Set(ByVal value As String)
                m_Label_reserve1_Espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Texte_reserve1_Espagnole = "Texte de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Texte_reserve1_Espagnole</c> indique par une chaîne un texte complémentaire en Espagnole dans l'enregistrement.
        ''' </summary>
        Public Property Texte_reserve1_Espagnole() As String
            Get
                Return m_Texte_reserve1_Espagnole

            End Get
            Set(ByVal value As String)
                m_Texte_reserve1_Espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Label_reserve1_Italien = "Label de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Label_reserve1_Italien</c> indique par une chaîne un label complémentaire en Italien dans l'enregistrement.
        ''' </summary>
        Public Property Label_reserve1_Italien() As String
            Get
                Return m_Label_reserve1_Italien

            End Get
            Set(ByVal value As String)
                m_Label_reserve1_Italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Texte_reserve1_Italien = "Texte de réserve"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Texte_reserve1_Italien</c> indique par une chaîne un texte complémentaire en Italien dans l'enregistrement.
        ''' </summary>
        Public Property Texte_reserve1_Italien() As String
            Get
                Return m_Texte_reserve1_Italien

            End Get
            Set(ByVal value As String)
                m_Texte_reserve1_Italien = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Boolean_reserve1 = TRUE
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Boolean_reserve1</c> booléen de réserve.
        ''' </summary>
        Public Property Boolean_reserve1() As Boolean
            Get
                Return m_Boolean_reserve1

            End Get
            Set(ByVal value As Boolean)
                m_Boolean_reserve1 = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim xaml_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''xaml_Enr1.Integer_reserve1 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Integer_reserve1</c> entier de réserve.
        ''' </summary>
        Public Property Integer_reserve1() As Integer
            Get
                Return m_Integer_reserve1

            End Get
            Set(ByVal value As Integer)
                m_Integer_reserve1 = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''num_Enr1.Date_reserve1 = "17/12/2007"
        ''' MsgBox(num_Enr1.Date_ENR.Day)
        ''' MsgBox(num_Enr1.Date_ENR.Month)
        ''' MsgBox(num_Enr1.Date_ENR.Year)
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Date_reserve1</c> date de réserve.
        ''' </summary>
        Public Property Date_reserve1() As Date
            Get
                Return m_Date_reserve1

            End Get
            Set(ByVal value As Date)
                m_Date_reserve1 = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.IDConstructeur = 4
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> indique par un entier le numéro unique d'un constructeur.
        ''' </summary>
        Public Property IDConstructeur() As Integer
            Get
                Return m_IDConstructeur

            End Get
            Set(ByVal value As Integer)
                m_IDConstructeur = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_Niveau_acces_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Type_Niveau_acces_Enr1.Niveau_acces = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Niveau_acces</c> indique par un entier le Niveau d'acces d'un enregistrement.
        ''' </summary>
        Public Property Niveau_acces() As Integer
            Get
                Return m_Niveau_acces

            End Get
            Set(ByVal value As Integer)
                m_Niveau_acces = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.STANDARD = TRUE
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>STANDARD</c> booléen de indique si l'enregistrement est standard.
        ''' </summary>
        Public Property STANDARD() As Boolean
            Get
                Return m_STANDARD

            End Get
            Set(ByVal value As Boolean)
                m_STANDARD = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''num_Enr1.Derniere_modification = "17/12/2007"
        ''' MsgBox(num_Enr1.Date_ENR.Day)
        ''' MsgBox(num_Enr1.Date_ENR.Month)
        ''' MsgBox(num_Enr1.Date_ENR.Year)
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Derniere_modification</c> date de dernière modification.
        ''' </summary>
        Public Property Derniere_modification() As Date
            Get
                Return m_Derniere_modification

            End Get
            Set(ByVal value As Date)
                m_Derniere_modification = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Enr1.Auteur_derniere_modification = "T.VIEIL"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Auteur_derniere_modification</c> indique par une chaîne l'auteur de la dernière modification de l'enregistrement.
        ''' </summary>
        Public Property Auteur_derniere_modification() As String
            Get
                Return m_Auteur_derniere_modification

            End Get
            Set(ByVal value As String)
                m_Auteur_derniere_modification = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_Niveau_acces_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Type_Niveau_acces_Enr1.NUMERO_MODIFICATION = 3
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>NUMERO_MODIFICATION</c> indique par un entier l'indice de modification d'un enregistrement.
        ''' </summary>
        Public Property NUMERO_MODIFICATION() As Integer
            Get
                Return m_NUMERO_MODIFICATION

            End Get
            Set(ByVal value As Integer)
                m_NUMERO_MODIFICATION = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_Doc_Enr1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.enregistrement
        '''Type_Doc_Enr1.Type_Doc = 23
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_Doc</c> indique par un entier le type de verrouillage d'un enregistrement.
        ''' </summary>
        Public Property Type_Doc() As Integer
            Get
                Return m_Type_Doc

            End Get
            Set(ByVal value As Integer)
                m_Type_Doc = value
            End Set
        End Property





    End Class

End Namespace