Namespace CLASS_IDEALTAKE.Objects

    ''' <summary>
    ''' La classe PARAM_E (Anciennement : menu_controle_acces) détermine le type de centrale au clavier qui autorise l'entrée.
    ''' le nom PARAM_E rend la propriété utilisable dans d'autres domaines.
    ''' </summary>
    Public Class PARAM_E

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
        '''Dim Centrale As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Centrale.Indice_AUTO = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Indice_AUTO</c> est un entier identifiant de façon unique le type de contrôle d'acces.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_francais = "Centrale Vigik/Résidents"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_francais</c> est une chaîne décrivant le type de contrôle en Français.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_anglais = "Centrale Vigik/Résidents"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_anglais</c> est une chaîne décrivant le type de contrôle en Anglais.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_allemand = "Centrale Vigik/Résidents"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_allemand</c> est une chaîne décrivant le type de contrôle  en Allemand.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_espagnole = "Centrale Vigik/Résidents"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_espagnole</c> est une chaîne décrivant le type de contrôle en Espagnole.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_italien = "Centrale Vigik/Résidents"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_italien</c> est une chaîne décrivant le type de contrôle en Italien.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Str_DESCRIPTION.VALEUR_REELLE = "Centrale Vigik/Résidents"
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_E
        '''Str_DESCRIPTION.IDConstructeur = 7
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> est un entier identifiant un constructeur se servant de ce type de contrôle.
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