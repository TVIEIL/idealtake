
Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe PARAM_C (Anciennement : menu_systemes) détermine la technologie 2fils 5fils etc....
    ''' le nom PARAM_C rend la propriété utilisable dans d'autres domaines.
    ''' </summary>
    Public Class PARAM_C  'Anciennement : menu_systemes

        Private m_Indice_AUTO As Integer = -1
        Private m_IDConstructeur As Integer = -1
        Private m_VISIBLE_LOGICIEL_francais As String = ""
        Private m_VISIBLE_LOGICIEL_anglais As String = ""
        Private m_VISIBLE_LOGICIEL_allemand As String = ""
        Private m_VISIBLE_LOGICIEL_espagnole As String = ""
        Private m_VISIBLE_LOGICIEL_italien As String = ""
        Private m_VALEUR_REELLE As String = ""
        Private m_VALEUR_PARAM_B As Integer = -1           'Anciennement : m_VALEUR_AUDIO_VIDEO
        Private m_FIGER_PARAM_B As Boolean = False         'Anciennement : m_FIGER_AUDIO_VIDEO
        Private m_SUPPRIMER_DE_PARAM_B As Integer = -1     'Anciennement : m_SUPPRIMER_DE_AUDIO_VIDEO


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Indice_AUTO As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''num_Indice_AUTO.Indice_AUTO = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Indice_AUTO</c> est un entier identifiant de façon unique le type de système e.g : Audio 5fils.
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
        '''Dim Int_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''Int_DESCRIPTION.IDConstructeur = 7
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> est un entier identifiant un constructeur se servant de cet ype de système e.g : Audio 5fils.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_francais = "Audio 5 Fils"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_francais</c> est une chaîne décrivant le type de système en Français  e.g : Audio 5fils.
        ''' </summary>
        Public Property VISIBLE_LOGICIEL_francais() As String
            Get
                Return m_VISIBLE_LOGICIEL_francais

            End Get
            Set(ByVal value As String)
                m_VISIBLE_LOGICIEL_francais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''Str_DESCRIPTION.VISIBLE_LOGICIEL_anglais = "Audio 5 wires"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VISIBLE_LOGICIEL_anglais</c> est une chaîne décrivant le type de système en Anglais  e.g : Audio 5 wires.
        ''' </summary>
        Public Property VISIBLE_LOGICIEL_anglais() As String
            Get
                Return m_VISIBLE_LOGICIEL_anglais

            End Get
            Set(ByVal value As String)
                m_VISIBLE_LOGICIEL_anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''Str_DESCRIPTION.VISIBLE_LOGICIEL_allemand = "Audio 5 fils"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VISIBLE_LOGICIEL_allemand</c> est une chaîne décrivant le type de système en Allemand.
        ''' </summary>
        Public Property VISIBLE_LOGICIEL_allemand() As String
            Get
                Return m_VISIBLE_LOGICIEL_allemand

            End Get
            Set(ByVal value As String)
                m_VISIBLE_LOGICIEL_allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''Str_DESCRIPTION.VISIBLE_LOGICIEL_espagnole = "Audio 5 fils"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VISIBLE_LOGICIEL_espagnole</c> est une chaîne décrivant le type de système en Espagnole.
        ''' </summary>
        Public Property VISIBLE_LOGICIEL_espagnole() As String
            Get
                Return m_VISIBLE_LOGICIEL_espagnole

            End Get
            Set(ByVal value As String)
                m_VISIBLE_LOGICIEL_espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''Str_DESCRIPTION.VISIBLE_LOGICIEL_italien = "Audio 5 fils"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VISIBLE_LOGICIEL_italien</c> est une chaîne décrivant le type de système en Espagnole.
        ''' </summary>
        Public Property VISIBLE_LOGICIEL_italien() As String
            Get
                Return m_VISIBLE_LOGICIEL_italien

            End Get
            Set(ByVal value As String)
                m_VISIBLE_LOGICIEL_italien = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_A
        '''Str_DESCRIPTION.VALEUR_REELLE = "Audio 5fils"
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
        '''Dim val_defaut As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''val_defaut.VALEUR_PARAM_B = 3
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_PARAM_B</c> est un entier désignant la valeur par défaut.
        ''' </summary>
        Public Property VALEUR_PARAM_B() As Integer     'VALEUR_AUDIO_VIDEO
            Get
                Return m_VALEUR_PARAM_B

            End Get
            Set(ByVal value As Integer)
                m_VALEUR_PARAM_B = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim val_defaut As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''val_defaut.FIGER_PARAM_B = false
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>FIGER_PARAM_B</c> est un booleen désignant si il faut empécher une autre sélection.
        ''' </summary>
        Public Property FIGER_PARAM_B() As Boolean  'FIGER_AUDIO_VIDEO
            Get
                Return m_FIGER_PARAM_B

            End Get
            Set(ByVal value As Boolean)
                m_FIGER_PARAM_B = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim val_defaut As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.PARAM_C
        '''sup_val.SUPPRIMER_DE_PARAM_B = false
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>SUPPRIMER_DE_PARAM_B</c> est un entier désignant un des éléments ne devant plus faire partit de la liste des valeurs possible de ce paramètre.
        ''' </summary>
        Public Property SUPPRIMER_DE_PARAM_B() As Integer  'SUPPRIMER_DE_AUDIO_VIDEO
            Get
                Return m_SUPPRIMER_DE_PARAM_B

            End Get
            Set(ByVal value As Integer)
                m_SUPPRIMER_DE_PARAM_B = value
            End Set
        End Property
    End Class
End Namespace