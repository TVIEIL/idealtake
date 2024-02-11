Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe niveau_acces détermine le niveau d'acces d'un document
    ''' </summary>
    Public Class niveau_acces


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
        '''Dim num_niveau_acces As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''num_niveau_acces.Indice_AUTO = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Indice_AUTO</c> est un entier identifiant le niveau d'acces du document.
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
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_francais = "-- Facil --"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_francais</c> est une chaîne décrivant le  niveau d'acces du document en Français.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_francais() As Integer
            Get
                Return m_VALEUR_LOGICIEL_francais

            End Get
            Set(ByVal value As Integer)
                m_VALEUR_LOGICIEL_francais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_anglais = "-- Easy --"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_anglais</c> est une chaîne décrivant le  niveau d'acces du document en Anglais.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_anglais() As Integer
            Get
                Return m_VALEUR_LOGICIEL_anglais

            End Get
            Set(ByVal value As Integer)
                m_VALEUR_LOGICIEL_anglais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_allemand = "-- Easy --"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_allemand</c> est une chaîne décrivant le  niveau d'acces du document en Allemand.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_allemand() As Integer
            Get
                Return m_VALEUR_LOGICIEL_allemand

            End Get
            Set(ByVal value As Integer)
                m_VALEUR_LOGICIEL_allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_espagnole = "-- Easy --"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_espagnole</c> est une chaîne décrivant le  niveau d'acces du document en Espagnole.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_espagnole() As Integer
            Get
                Return m_VALEUR_LOGICIEL_espagnole

            End Get
            Set(ByVal value As Integer)
                m_VALEUR_LOGICIEL_espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''Str_DESCRIPTION.VALEUR_LOGICIEL_italien = "-- Easy --"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_LOGICIEL_italien</c> est une chaîne décrivant le  niveau d'acces du document en Italien.
        ''' </summary>
        Public Property VALEUR_LOGICIEL_italien() As Integer
            Get
                Return m_VALEUR_LOGICIEL_italien

            End Get
            Set(ByVal value As Integer)
                m_VALEUR_LOGICIEL_italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''Str_DESCRIPTION.VALEUR_REELLE = "Facil"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>VALEUR_REELLE</c> est une chaîne décrivant le niveau d'acces du document anciennement dans M.R.U.
        ''' </summary>
        Public Property VALEUR_REELLE() As Integer
            Get
                Return m_VALEUR_REELLE

            End Get
            Set(ByVal value As Integer)
                m_VALEUR_REELLE = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.niveau_acces
        '''Str_DESCRIPTION.IDConstructeur = 7
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> est un entier identifiant un constructeur.
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