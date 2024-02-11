

Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe Elements_HTML stocke du code HTML permettant de crééer les interafes utilisateurs.
    ''' </summary>
    Public Class Elements_HTML

        Private m_ID_elements_HTML As Integer = -1
        Private m_Name_of_page As String = ""
        Private m_Additional_Name As String = ""
        Private m_Position_in_page As Integer = -1
        Private m_Name_of_TAG As String = ""
        Private m_German_CODE As String = ""
        Private m_English_CODE As String = ""
        Private m_Spanish_CODE As String = ""
        Private m_French_CODE As String = ""
        Private m_Italian_CODE As String = ""

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_ID_elements_HTML As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''num_ID_elements_HTML.ID_elements_HTML = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ID_elements_HTML</c> est un entier identifiant de façon unique un élément HTML.
        ''' </summary>
        Public Property ID_elements_HTML() As Integer
            Get
                Return m_ID_elements_HTML

            End Get
            Set(ByVal value As Integer)
                m_ID_elements_HTML = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_Name_of_page As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj_Name_of_page.Name_of_page = "Advanced_Research"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Name_of_page</c> indique le nom de la page dans lequel se trouve l'élément HTML.
        ''' </summary>
        Public Property Name_of_page() As String
            Get
                Return m_Name_of_page

            End Get
            Set(ByVal value As String)
                m_Name_of_page = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_Name_of_page As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj_Name_of_page.Additional_Name = "Récupération mot de passe"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Additional_Name</c> indique le nom de la page dans lequel se trouve l'élément HTML.
        ''' </summary>
        Public Property Additional_Name() As String
            Get
                Return m_Additional_Name

            End Get
            Set(ByVal value As String)
                m_Additional_Name = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Position_in_page As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''num_Position_in_page.Position_in_page = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Position_in_page</c> est un entier identifiant de façon unique la position de l'élément dans la page HTML.
        ''' </summary>
        Public Property Position_in_page() As Integer
            Get
                Return m_Position_in_page

            End Get
            Set(ByVal value As Integer)
                m_Position_in_page = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj1.Name_of_TAG = "Advanced_Research"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Name_of_TAG</c> indique le nom de la page dans lequel se trouve l'élément HTML.
        ''' </summary>
        Public Property Name_of_TAG() As String
            Get
                Return m_Name_of_TAG

            End Get
            Set(ByVal value As String)
                m_Name_of_TAG = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj1.German_CODE = "<a>Test</a>"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>German_CODE</c> code de l'élément HTML en Allemand.
        ''' </summary>
        Public Property German_CODE() As String
            Get
                Return m_German_CODE

            End Get
            Set(ByVal value As String)
                m_German_CODE = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj1.English_CODE = "<a>Test</a>"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>English_CODE</c> code de l'élément HTML en Anglais.
        ''' </summary>
        Public Property English_CODE() As String
            Get
                Return m_English_CODE

            End Get
            Set(ByVal value As String)
                m_English_CODE = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj1.Spanish_CODE = "<a>Test</a>"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Spanish_CODE</c> code de l'élément HTML en Espagnole.
        ''' </summary>
        Public Property Spanish_CODE() As String
            Get
                Return m_Spanish_CODE

            End Get
            Set(ByVal value As String)
                m_Spanish_CODE = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj1.French_CODE = "<a>Test</a>"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>French_CODE</c> code de l'élément HTML en Français.
        ''' </summary>
        Public Property French_CODE() As String
            Get
                Return m_French_CODE

            End Get
            Set(ByVal value As String)
                m_French_CODE = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_HTML
        '''Obj1.Italian_CODE = "<a>Test</a>"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Italian_CODE</c> code de l'élément HTML en Français.
        ''' </summary>
        Public Property Italian_CODE() As String
            Get
                Return m_Italian_CODE

            End Get
            Set(ByVal value As String)
                m_Italian_CODE = value
            End Set
        End Property


    End Class

End Namespace