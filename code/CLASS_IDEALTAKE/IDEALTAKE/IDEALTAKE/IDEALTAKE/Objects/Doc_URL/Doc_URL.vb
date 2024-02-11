
Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe Doc_URL donne l'URL d'un document en fonction du type de document un entier et un identifiant constructeur un entier également.
    ''' </summary>
    Public Class Doc_URL


        Private m_IDConstructeur As Integer = -1
        Private m_IDDoc As Integer = -1
        Private m_URL As String = ""


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
        '''Dim Num_IDDoc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''Num_IDDoc.IDDoc = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDDoc</c> est un entier identifiant de façon unique le type de document.
        ''' </summary>
        Public Property IDDoc() As Integer
            Get
                Return m_IDDoc

            End Get
            Set(ByVal value As Integer)
                m_IDDoc = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim URL1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.constructeur
        '''URL1.URL = "http://www.monsite.fr/doc1/"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>URL</c> est une chaîne fournissant l'URL d'un type de document.
        ''' </summary>
        Public Property URL() As String
            Get
                Return m_URL

            End Get
            Set(ByVal value As String)
                m_URL = value
            End Set
        End Property

    End Class

End Namespace