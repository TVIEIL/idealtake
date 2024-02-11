Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe constructeur fournit des informations sur le construteur de produit.
    ''' </summary>
    Public Class constructeur


        Private m_IDConstructeur As Integer = -1
        Private m_NOM As String = ""
        Private m_SITE_CONSTRUCTEUR As String = ""
        Private m_INFORMATIONS_francais As String = ""
        Private m_INFORMATIONS_anglais As String = ""
        Private m_INFORMATIONS_allemand As String = ""
        Private m_INFORMATIONS_espagnole As String = ""
        Private m_INFORMATIONS_italien As String = ""
        Private m_CHEMIN_LOGOS As String = ""
        Private m_NOM_FICHIER_IMAGE_LOGOS As String = ""
        Private m_LARGEUR_IMAGE_LOGOS As String = ""
        Private m_HAUTEUR_IMAGE_LOGOS As String = ""

        'Private m_Liens_Constructeur As Boolean = False    '' par défaut inactif
        ' Private m_URL_schemas As String = ""
        ' Private m_URL_synoptiques As String = ""
        ' Private m_URL_doc As String = ""
        'Private m_URL_fiches_produits As String = ""
        'Private m_URL_photos_produits As String = ""


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_IDConstructeur As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''num_IDConstructeur.IDConstructeur = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> est un entier identifiant de façon unique le constructeur.
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
        '''Dim NOM1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''NOM1.NOM = "Entreprise Debolitchvi"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>NOM</c> est une chaîne représentatnt le nom de la société du constructeur.
        ''' </summary>
        Public Property NOM() As String
            Get
                Return m_NOM

            End Get
            Set(ByVal value As String)
                m_NOM = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Site_Web1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Site_Web1.SITE_CONSTRUCTEUR = "http://www.SiteDebolitchvi.fr"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>SITE_CONSTRUCTEUR</c> est une chaîne représentatnt l'URL du site web du constructeur.
        ''' </summary>
        Public Property SITE_CONSTRUCTEUR() As String
            Get
                Return m_SITE_CONSTRUCTEUR

            End Get
            Set(ByVal value As String)
                m_SITE_CONSTRUCTEUR = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Infos_Fr As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Infos_Fr.INFORMATIONS_francais = "Fabricant de machine spécialisé dans le..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>INFORMATIONS_francais</c> est une chaîne représentant la description de l'activité de la société en Français.
        ''' </summary>
        Public Property INFORMATIONS_francais() As String
            Get
                Return m_INFORMATIONS_francais

            End Get
            Set(ByVal value As String)
                m_INFORMATIONS_francais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Infos_An As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Infos_An.INFORMATIONS_anglais = "Our company build industrial tools..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>INFORMATIONS_anglais</c> est une chaîne représentant la description de l'activité de la société en Anglais.
        ''' </summary>
        Public Property INFORMATIONS_anglais() As String
            Get
                Return m_INFORMATIONS_anglais

            End Get
            Set(ByVal value As String)
                m_INFORMATIONS_anglais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Infos_Al As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Infos_Al.INFORMATIONS_allemand = "Das Ziel unseres Betriebes..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>INFORMATIONS_allemand</c> est une chaîne représentant la description de l'activité de la société en Allemand.
        ''' </summary>
        Public Property INFORMATIONS_allemand() As String
            Get
                Return m_INFORMATIONS_allemand

            End Get
            Set(ByVal value As String)
                m_INFORMATIONS_allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Infos_Es As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Infos_Es.INFORMATIONS_espagnole = "Buenos dias..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>INFORMATIONS_espagnole</c> est une chaîne représentant la description de l'activité de la société en Espagnole.
        ''' </summary>
        Public Property INFORMATIONS_espagnole() As String
            Get
                Return m_INFORMATIONS_espagnole

            End Get
            Set(ByVal value As String)
                m_INFORMATIONS_espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Infos_It As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Infos_It.INFORMATIONS_italien = "Il scopo della nostra impresa..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>INFORMATIONS_italien</c> est une chaîne représentant la description de l'activité de la société en Italien.
        ''' </summary>
        Public Property INFORMATIONS_italien() As String
            Get
                Return m_INFORMATIONS_italien

            End Get
            Set(ByVal value As String)
                m_INFORMATIONS_italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim URL_Logos As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''URL_Logos.CHEMIN_LOGOS = "http://www.sites/images/"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>CHEMIN_LOGOS</c> est une chaîne représentant l'URL sans le nom du fichier du logos de la société.
        ''' </summary>
        Public Property CHEMIN_LOGOS() As String
            Get
                Return m_CHEMIN_LOGOS

            End Get
            Set(ByVal value As String)
                m_CHEMIN_LOGOS = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Logos As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Logos.NOM_FICHIER_IMAGE_LOGOS = "societe.jpg"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>NOM_FICHIER_IMAGE_LOGOS</c> est une chaîne représentant le nom du fichier logos de la société.
        ''' </summary>
        Public Property NOM_FICHIER_IMAGE_LOGOS() As String
            Get
                Return m_NOM_FICHIER_IMAGE_LOGOS

            End Get
            Set(ByVal value As String)
                m_NOM_FICHIER_IMAGE_LOGOS = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim L_Logos As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''L_Logos.LARGEUR_IMAGE_LOGOS = "220"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>LARGEUR_IMAGE_LOGOS</c> est une chaîne représentant la largeur en pixel du logos de la société.
        ''' </summary>
        Public Property LARGEUR_IMAGE_LOGOS() As String
            Get
                Return m_LARGEUR_IMAGE_LOGOS

            End Get
            Set(ByVal value As String)
                m_LARGEUR_IMAGE_LOGOS = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim H_Logos As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''H_Logos.HAUTEUR_IMAGE_LOGOS = "110"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>HAUTEUR_IMAGE_LOGOS</c> est une chaîne représentant la hauteur en pixel du logos de la société.
        ''' </summary>
        Public Property HAUTEUR_IMAGE_LOGOS() As String
            Get
                Return m_HAUTEUR_IMAGE_LOGOS

            End Get
            Set(ByVal value As String)
                m_HAUTEUR_IMAGE_LOGOS = value
            End Set
        End Property

        'Public Property Liens_Constructeur() As Boolean
        ' Get
        '    Return m_Liens_Constructeur

        'End Get
        ' Set(ByVal value As Boolean)
        '    m_Liens_Constructeur = value
        ' End Set
        ' End Property

        ' Public Property URL_schemas() As String
        ' Get
        '      Return m_URL_schemas
        '
        '  End Get
        ' Set(ByVal value As String)
        '     m_URL_schemas = value
        '   End Set
        ' End Property

        '  Public Property URL_synoptiques() As String
        '   Get
        '      Return m_URL_synoptiques
        '
        '  End Get
        '  Set(ByVal value As String)
        '      m_URL_synoptiques = value
        '  End Set
        'End Property

        ' Public Property URL_doc() As String
        ' Get
        '    Return m_URL_doc

        ' End Get
        '  Set(ByVal value As String)
        '      m_URL_doc = value
        '  End Set
        'End Property

        ' Public Property URL_fiches_produits() As String
        '  Get
        '    Return m_URL_fiches_produits

        ' End Get
        'Set(ByVal value As String)
        '   m_URL_fiches_produits = value
        'End Set
        'End Property

        'Public Property URL_photos_produits() As String
        'Get
        '   Return m_URL_photos_produits

        ' End Get
        'Set(ByVal value As String)
        '     m_URL_photos_produits = value
        '   End Set
        'End Property

    End Class
End Namespace