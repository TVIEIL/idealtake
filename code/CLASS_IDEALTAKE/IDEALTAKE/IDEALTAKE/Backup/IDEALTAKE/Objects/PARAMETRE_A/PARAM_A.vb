Namespace CLASS_IDEALTAKE.Objects

    ''' <summary>
    ''' La classe PARAM_A (Anciennement : menu_verrouillages) détermine le système fermant les portes.
    ''' le nom PARAM_A rend la propriété utilisable dans d'autres domaines.
    ''' </summary>
    Public Class PARAM_A

        Private m_Indice_AUTO As Integer = -1
        Private m_VALEUR_LOGICIEL_francais As String = ""
        Private m_VALEUR_LOGICIEL_anglais As String = ""
        Private m_VALEUR_LOGICIEL_allemand As String = ""
        Private m_VALEUR_LOGICIEL_espagnole As String = ""
        Private m_VALEUR_LOGICIEL_italien As String = ""
        Private m_VALEUR_REELLE As String = ""
        Private m_IDConstructeur As Integer = -1


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Indice_AUTO As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''num_Indice_AUTO.Indice_AUTO = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Indice_AUTO</c> est un entier identifiant de façon unique le système de verrouillage.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_francais = "Gache 2T"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_francais</c> est une chaîne décrivant le système de verrouillage en Français.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_francais() As String
            Get
                Return m_VALEUR_LOGICIEL_francais

            End Get
            Set(ByVal value As String)
                m_VALEUR_LOGICIEL_francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_anglais = "Gache 2T"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_anglais</c> est une chaîne décrivant le système de verrouillage en Anglais.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_anglais() As String
            Get
                Return m_VALEUR_LOGICIEL_anglais

            End Get
            Set(ByVal value As String)
                m_VALEUR_LOGICIEL_anglais = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_allemand = "Gache 2T"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_allemand</c> est une chaîne décrivant le système de verrouillage en Allemand.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_allemand() As String
            Get
                Return m_VALEUR_LOGICIEL_allemand

            End Get
            Set(ByVal value As String)
                m_VALEUR_LOGICIEL_allemand = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_espagnole = "Gache 2T"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_espagnole</c> est une chaîne décrivant le système de verrouillage en Espagnole.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_espagnole() As String
            Get
                Return m_VALEUR_LOGICIEL_espagnole

            End Get
            Set(ByVal value As String)
                m_VALEUR_LOGICIEL_espagnole = value
            End Set
        End Property




        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_italien = "Gache 2T"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_italien</c> est une chaîne décrivant le système de verrouillage en Italien.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_italien() As String
            Get
                Return m_VALEUR_LOGICIEL_italien

            End Get
            Set(ByVal value As String)
                m_VALEUR_LOGICIEL_italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.VALEUR_REELLE = "Gache 2T"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_REELLE</c> est une chaîne historiquement utilisée comme valeur stockée dans la table SCHEMA maintenant ENREGISTREMENTS.
        ''' </summary>
        Public Property VALEUR_REELLE() As String
            Get
                Return m_VALEUR_REELLE

            End Get
            Set(ByVal value As String)
                m_VALEUR_REELLE = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.IDConstructeur = 7
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> est un entier identifiant un constructeur se servant de ce type de verrouillage.
        ''' </summary>
        Public Property IDConstructeur() As Integer
            Get
                Return m_IDConstructeur

            End Get
            Set(ByVal value As Integer)
                m_IDConstructeur = value
            End Set
        End Property

    End Class
End Namespace