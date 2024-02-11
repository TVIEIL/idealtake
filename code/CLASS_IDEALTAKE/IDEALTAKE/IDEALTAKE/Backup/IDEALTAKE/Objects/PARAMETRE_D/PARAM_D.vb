Namespace CLASS_IDEALTAKE.Objects

    ''' <summary>
    ''' La classe PARAM_D (Anciennement : menu_serie) détermine l'aspect des platines de rue 52000 ou KSTEEL.
    ''' le nom PARAM_D rend la propriété utilisable dans d'autres domaines.
    ''' </summary>
    Public Class PARAM_D

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
        '''Dim Platine As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Platine.Indice_AUTO = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Indice_AUTO</c> est un entier identifiant de façon unique le type de platine.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_francais = "KSTEEL/SINTHESI"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_francais</c> est une chaîne décrivant le type de platine.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_anglais = "KSTEEL/SINTHESI"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_anglais</c> est une chaîne décrivant le type de platine en Anglais.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_allemand = "KSTEEL/SINTHESI"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_allemand</c> est une chaîne décrivant le type de platine en Allemand.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_espagnole = "KSTEEL/SINTHESI"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_espagnole</c> est une chaîne décrivant le type de platine en Espagnole.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_italien = "KSTEEL/SINTHESI"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_italien</c> est une chaîne décrivant le type de platine en Italien.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Str_DESCRIPTION.VALEUR_REELLE = "KSTEEL/SINTHESI"
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_D
        '''Str_DESCRIPTION.IDConstructeur = 7
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> est un entier identifiant un constructeur se servant de ce type de platine.
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