Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe categorie regroupe des références de même fonction.
    ''' </summary>
    Public Class categorie

        Private m_Indice_AUTO As Integer = -1
        Private m_IDConstructeur As Integer = -1
        Private m_TYPE_REFERENCE_francais As String = ""
        Private m_TYPE_REFERENCE_anglais As String = ""
        Private m_TYPE_REFERENCE_allemand As String = ""
        Private m_TYPE_REFERENCE_espagnole As String = ""
        Private m_TYPE_REFERENCE_italien As String = ""

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_categorie As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.categorie
        '''num_categorie.Indice_AUTO = 4
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Indice_AUTO</c> indique le numéro d'une catégorie d'article.
        ''' </summary>
        Public Property Indice_AUTO() As Integer


            Get
                Return m_Indice_AUTO

            End Get
            Set(ByVal value As Integer)
                m_Indice_AUTO = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_constructeur As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.categorie
        '''num_constructeur.IDConstructeur = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> indique le numéro du constructeur.
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
        '''Dim TYPE1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.categorie
        '''TYPE1.TYPE_REFERENCE_francais = "Panneau Solaire"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TYPE_REFERENCE_francais</c> indique le nom de la catégorie en francais.
        ''' </summary>
        Public Property TYPE_REFERENCE_francais() As String
            Get
                Return m_TYPE_REFERENCE_francais

            End Get
            Set(ByVal value As String)
                m_TYPE_REFERENCE_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim TYPE1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.categorie
        '''TYPE1.TYPE_REFERENCE_anglais = "circuit"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        '''La propriété <c>TYPE_REFERENCE_anglais</c> indique le nom de la catégorie en anglais.
        ''' </summary>
        Public Property TYPE_REFERENCE_anglais() As String
            Get
                Return m_TYPE_REFERENCE_anglais

            End Get
            Set(ByVal value As String)
                m_TYPE_REFERENCE_anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim TYPE1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.categorie
        '''TYPE1.TYPE_REFERENCE_allemand = "Panneau Solaire"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TYPE_REFERENCE_allemand</c> indique le nom de la catégorie en allemand.
        ''' </summary>
        Public Property TYPE_REFERENCE_allemand() As String
            Get
                Return m_TYPE_REFERENCE_allemand
            End Get
            Set(ByVal value As String)
                m_TYPE_REFERENCE_allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim TYPE1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.categorie
        '''TYPE1.TYPE_REFERENCE_espagnole = "Panneau Solaire"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TYPE_REFERENCE_espagnole</c> indique le nom de la catégorie en espagnole.
        ''' </summary>
        Public Property TYPE_REFERENCE_espagnole() As String
            Get
                Return m_TYPE_REFERENCE_espagnole

            End Get
            Set(ByVal value As String)
                m_TYPE_REFERENCE_espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim TYPE1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.categorie
        '''TYPE1.TYPE_REFERENCE_italien = "Panneau Solaire"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TYPE_REFERENCE_italien</c> indique le nom de la catégorie en italien.
        ''' </summary>
        Public Property TYPE_REFERENCE_italien() As String
            Get
                Return m_TYPE_REFERENCE_italien
            End Get
            Set(ByVal value As String)
                m_TYPE_REFERENCE_italien = value
            End Set
        End Property

    End Class

End Namespace