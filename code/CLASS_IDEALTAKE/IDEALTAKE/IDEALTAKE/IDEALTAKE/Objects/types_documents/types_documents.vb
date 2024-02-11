Namespace CLASS_IDEALTAKE.Objects

    ''' <summary>
    ''' La classe types_documents stocke des informations sur les différents types de documents
    ''' </summary>
    Public Class types_documents


        Private m_Indice_AUTO As Integer = -1
        Private m_Description_francais As String = ""
        Private m_Description_anglais As String = ""
        Private m_Description_allemand As String = ""
        Private m_Description_espagnole As String = ""
        Private m_Description_italien As String = ""
        Private m_Type_Previsualisable As Boolean = False
        Private m_Picto_Doc As String = ""

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim int_typedoc As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''int_typedoc.Indice_AUTO = 7
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Indice_AUTO</c> est un entier identifiant un type de document.
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
        '''Dim Str_DESCRIPTION_ref_francais As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''Str_DESCRIPTION_ref_francais.Description_francais = "Vidéo"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_francais</c> est une chaîne décrivant un type de document en Français.
        ''' </summary>
        Public Property Description_francais() As String
            Get
                Return m_Description_francais

            End Get
            Set(ByVal value As String)
                m_Description_francais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''Str_DESCRIPTION.Description_anglais = "Video"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_anglais</c> est une chaîne décrivant un type de document en Anglais.
        ''' </summary>
        Public Property Description_anglais() As String
            Get
                Return m_Description_anglais

            End Get
            Set(ByVal value As String)
                m_Description_anglais = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''Str_DESCRIPTION.Description_allemand = "Video"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_allemand</c> est une chaîne décrivant un type de document en Allemand.
        ''' </summary>
        Public Property Description_allemand() As String
            Get
                Return m_Description_allemand

            End Get
            Set(ByVal value As String)
                m_Description_allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''Str_DESCRIPTION.Description_espagnole = "Video"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_espagnole</c> est une chaîne décrivant un type de document en Espagnole.
        ''' </summary>
        Public Property Description_espagnole() As String
            Get
                Return m_Description_espagnole

            End Get
            Set(ByVal value As String)
                m_Description_espagnole = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Str_DESCRIPTION As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''Str_DESCRIPTION.Description_italien = "Video"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_italien</c> est une chaîne décrivant un type de document en Italien.
        ''' </summary>
        Public Property Description_italien() As String
            Get
                Return m_Description_italien

            End Get
            Set(ByVal value As String)
                m_Description_italien = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_DOC As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''Type_DOC.Type_Previsualisable = true
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_Previsualisable</c> est un booléen décrivant si ce type de document est prévisualisable comme une image.
        ''' </summary>
        Public Property Type_Previsualisable() As Boolean
            Get
                Return m_Type_Previsualisable

            End Get
            Set(ByVal value As Boolean)
                m_Type_Previsualisable = value
            End Set
        End Property




        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_DOC As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.types_documents
        '''Type_DOC.Picto_Doc = "http://www.monsite.fr/images/type_doc/Son.jpg"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Picto_Doc</c> est une chaîne pointant sur une image illustrant de type de document par exemple d'un son.
        ''' </summary>
        Public Property Picto_Doc() As String
            Get
                Return m_Picto_Doc

            End Get
            Set(ByVal value As String)
                m_Picto_Doc = value
            End Set
        End Property

    End Class
End Namespace