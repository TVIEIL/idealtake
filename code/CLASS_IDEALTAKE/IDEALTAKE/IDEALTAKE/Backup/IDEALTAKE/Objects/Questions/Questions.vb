

Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe Questions stocke les questions types pour la récupération du compte.
    ''' </summary>
    Public Class Questions

        Private m_ID As Integer = -1
        Private m_German As String = ""
        Private m_English As String = ""
        Private m_Spanish As String = ""
        Private m_French As String = ""
        Private m_Italian As String = ""



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Questions
        '''obj1.ID = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ID</c> est un entier identifiant de façon unique une question.
        ''' </summary>
        Public Property ID() As Integer
            Get
                Return m_ID

            End Get
            Set(ByVal value As Integer)
                m_ID = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Questions
        '''Obj1.German = "question..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>User_Name</c> question en Allemand.
        ''' </summary>
        Public Property German() As String
            Get
                Return m_German

            End Get
            Set(ByVal value As String)
                m_German = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Questions
        '''Obj1.English = "question..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>English</c> question en Anglais.
        ''' </summary>
        Public Property English() As String
            Get
                Return m_English

            End Get
            Set(ByVal value As String)
                m_English = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Questions
        '''Obj1.Spanish = "question..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Spanish</c> question en Espagnole.
        ''' </summary>
        Public Property Spanish() As String
            Get
                Return m_Spanish

            End Get
            Set(ByVal value As String)
                m_Spanish = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Questions
        '''Obj1.French = "question..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>French</c> question en Français.
        ''' </summary>
        Public Property French() As String
            Get
                Return m_French

            End Get
            Set(ByVal value As String)
                m_French = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Questions
        '''Obj1.Italian = "question..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Italian</c> question en Italien.
        ''' </summary>
        Public Property Italian() As String
            Get
                Return m_Italian

            End Get
            Set(ByVal value As String)
                m_Italian = value
            End Set
        End Property


    End Class

End Namespace