Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe document stocke le nom de fichier et le type de document pour un enregistrement.
    ''' </summary>
    Public Class document

        Private m_Chrono As Integer = -1
        Private m_Nom_Lien_francais As String = ""
        Private m_Nom_Lien_anglais As String = ""
        Private m_Nom_Lien_allemand As String = ""
        Private m_Nom_Lien_espagnole As String = ""
        Private m_Nom_Lien_italien As String = ""
        Private m_Type_Doc As Integer = -1
        Private m_Nom_Fichier_francais As String = ""
        Private m_Nom_Fichier_anglais As String = ""
        Private m_Nom_Fichier_allemand As String = ""
        Private m_Nom_Fichier_espagnole As String = ""
        Private m_Nom_Fichier_italien As String = ""
        Private m_Taille_fichier_Fr As String = ""
        Private m_Taille_fichier_anglais As String = ""
        Private m_Taille_fichier_allemand As String = ""
        Private m_Taille_fichier_espagnole As String = ""
        Private m_Taille_fichier_italien As String = ""
        Private m_Largeur As String = ""
        Private m_Hauteur As String = ""



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Chrono As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''num_Chrono.Chrono = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Chrono</c> est un entier identifiant de façon unique un document lié à un enregistrement.
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
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Lien_francais = "Notice_du_clavier"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Lien_francais</c> est une chaîne identifiant un document lié à un enregistrement en langue Française.
        ''' </summary>
        Public Property Nom_Lien_francais() As String
            Get
                Return m_Nom_Lien_francais

            End Get
            Set(ByVal value As String)
                m_Nom_Lien_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Lien_anglais = "Notice_du_clavier"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Lien_anglais</c> est une chaîne identifiant un document lié à un enregistrement en langue Anglaise.
        ''' </summary>
        Public Property Nom_Lien_anglais() As String
            Get
                Return m_Nom_Lien_anglais

            End Get
            Set(ByVal value As String)
                m_Nom_Lien_anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Lien_allemand = "Notice_du_clavier"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Lien_allemand</c> est une chaîne identifiant un document lié à un enregistrement en langue Allemande.
        ''' </summary>
        Public Property Nom_Lien_allemand() As String
            Get
                Return m_Nom_Lien_allemand

            End Get
            Set(ByVal value As String)
                m_Nom_Lien_allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Lien_espagnole = "Notice_du_clavier"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Lien_espagnole</c> est une chaîne identifiant un document lié à un enregistrement en langue Espagnole.
        ''' </summary>
        Public Property Nom_Lien_espagnole() As String
            Get
                Return m_Nom_Lien_espagnole

            End Get
            Set(ByVal value As String)
                m_Nom_Lien_espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Lien_italien = "Notice_du_clavier"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Lien_italien</c> est une chaîne identifiant un document lié à un enregistrement en langue Italienne.
        ''' </summary>
        Public Property Nom_Lien_italien() As String
            Get
                Return m_Nom_Lien_italien

            End Get
            Set(ByVal value As String)
                m_Nom_Lien_italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''Type_Doc.Nom_Document_italien = "Notice_du_clavier_001.pdf"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_Doc</c> est un entier identifiant le type de document.
        ''' </summary>
        ''' 
        Public Property Type_Doc() As Integer
            Get
                Return m_Type_Doc

            End Get
            Set(ByVal value As Integer)
                m_Type_Doc = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Fichier_francais = "Notice_du_clavier_001.pdf"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Fichier_francais</c> est une chaîne identifiant le nom d'un document lié à un enregistrement en langue Française.
        ''' </summary>
        Public Property Nom_Fichier_francais() As String
            Get
                Return m_Nom_Fichier_francais

            End Get
            Set(ByVal value As String)
                m_Nom_Fichier_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Fichier_anglais = "Notice_du_clavier_001.pdf"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Fichier_anglais</c> est une chaîne identifiant le nom d'un document lié à un enregistrement en langue Anglaise.
        ''' </summary>
        Public Property Nom_Fichier_anglais() As String
            Get
                Return m_Nom_Fichier_anglais

            End Get
            Set(ByVal value As String)
                m_Nom_Fichier_anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Fichier_allemand = "Notice_du_clavier_001.pdf"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Fichier_allemand</c> est une chaîne identifiant le nom d'un document lié à un enregistrement en langue Allemande.
        ''' </summary>
        Public Property Nom_Fichier_allemand() As String
            Get
                Return m_Nom_Fichier_allemand

            End Get
            Set(ByVal value As String)
                m_Nom_Fichier_allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Fichier_espagnole = "Notice_du_clavier_001.pdf"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Fichier_espagnole</c> est une chaîne identifiant le nom d'un document lié à un enregistrement en langue Espagnole.
        ''' </summary>
        Public Property Nom_Fichier_espagnole() As String
            Get
                Return m_Nom_Fichier_espagnole

            End Get
            Set(ByVal value As String)
                m_Nom_Fichier_espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Nom_Fichier_italien = "Notice_du_clavier_001.pdf"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Fichier_italien</c> est une chaîne identifiant le nom d'un document lié à un enregistrement en langue Italienne.
        ''' </summary>
        Public Property Nom_Fichier_italien() As String
            Get
                Return m_Nom_Fichier_italien

            End Get
            Set(ByVal value As String)
                m_Nom_Fichier_italien = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Taille_fichier_Fr = "200 Ko"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Taille_fichier_Fr</c> est une chaîne représentant la taille du document.
        ''' </summary>
        Public Property Taille_fichier_Fr() As String
            Get
                Return m_Taille_fichier_Fr

            End Get
            Set(ByVal value As String)
                m_Taille_fichier_Fr = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Taille_fichier_anglais = "200 Ko"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Taille_fichier_anglais</c> est une chaîne représentant la taille du document version anglaise.
        ''' </summary>
        Public Property Taille_fichier_anglais() As String
            Get
                Return m_Taille_fichier_anglais

            End Get
            Set(ByVal value As String)
                m_Taille_fichier_anglais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Taille_fichier_allemand = "200 Ko"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Taille_fichier_allemand</c> est une chaîne représentant la taille du document version Allemande.
        ''' </summary>
        Public Property Taille_fichier_allemand() As String
            Get
                Return m_Taille_fichier_allemand

            End Get
            Set(ByVal value As String)
                m_Taille_fichier_allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Taille_fichier_espagnole = "200 Ko"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Taille_fichier_espagnole</c> est une chaîne représentant la taille du document version Espagnole.
        ''' </summary>
        Public Property Taille_fichier_espagnole() As String
            Get
                Return m_Taille_fichier_espagnole

            End Get
            Set(ByVal value As String)
                m_Taille_fichier_espagnole = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Taille_fichier_italien = "200 Ko"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Taille_fichier_italien</c> est une chaîne représentant la taille du document version Italienne.
        ''' </summary>
        Public Property Taille_fichier_italien() As String
            Get
                Return m_Taille_fichier_italien

            End Get
            Set(ByVal value As String)
                m_Taille_fichier_italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Largeur = "200"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Largeur</c> est une chaîne représentant la largeur d'une image prévisualisable en miniature grace au ratio de l'image.
        ''' </summary>
        Public Property Largeur() As String
            Get
                Return m_Largeur

            End Get
            Set(ByVal value As String)
                m_Largeur = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom_Doc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.document
        '''nom_Doc.Hauteur = "200"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Hauteur</c> est une chaîne représentant la hauteur d'une image prévisualisable en miniature grace au ratio de l'image.
        ''' </summary>
        Public Property Hauteur() As String
            Get
                Return m_Hauteur

            End Get
            Set(ByVal value As String)
                m_Hauteur = value
            End Set
        End Property


    End Class
End Namespace