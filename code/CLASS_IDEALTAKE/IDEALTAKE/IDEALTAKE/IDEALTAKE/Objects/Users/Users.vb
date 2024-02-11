

Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe Users stocke les paramètres des utilisateurs sur la base.
    ''' </summary>
    Public Class Users

        Private m_User_ID As Integer = -1
        Private m_User_Name As String = ""
        Private m_Password As String = ""
        Private m_e_mail As String = ""
        Private m_Question_ID As String = ""
        Private m_Answer As String = ""
        Private m_Enable As String = ""


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Users
        '''obj1.User_ID = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>User_ID</c> est un entier identifiant de façon unique l'utilisateur.
        ''' </summary>
        Public Property User_ID() As Integer
            Get
                Return m_User_ID

            End Get
            Set(ByVal value As Integer)
                m_User_ID = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Users
        '''Obj1.User_Name = "my_name"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>User_Name</c> indique le nom de l'utilisateur.
        ''' </summary>
        Public Property User_Name() As String
            Get
                Return m_User_Name

            End Get
            Set(ByVal value As String)
                m_User_Name = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Users
        '''Obj1.Password = "6798521"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Password</c> indique le mot de passe de l'utilisateur.
        ''' </summary>
        Public Property Password() As String
            Get
                Return m_Password

            End Get
            Set(ByVal value As String)
                m_Password = value
            End Set
        End Property




        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Users
        '''Obj1.e_mail = "Mon_email@domaine.com"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>e_mail</c> indique l'e-mail de l'utilisateur.
        ''' </summary>
        Public Property e_mail() As String
            Get
                Return m_e_mail

            End Get
            Set(ByVal value As String)
                m_e_mail = value
            End Set
        End Property




        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Users
        '''Obj1.Question_ID = "How old are you?"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Question_ID</c> question pour la récupération du mot de passe.
        ''' </summary>
        Public Property Question_ID() As String
            Get
                Return m_Question_ID

            End Get
            Set(ByVal value As String)
                m_Question_ID = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Users
        '''Obj1.Answer = "I'am 35 years old"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Answer</c> réponse à la question.
        ''' </summary>
        Public Property Answer() As String
            Get
                Return m_Answer

            End Get
            Set(ByVal value As String)
                m_Answer = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Users
        '''Obj1.Enable = "No"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Enable</c> compte actif.
        ''' </summary>
        Public Property Enable() As String
            Get
                Return m_Enable

            End Get
            Set(ByVal value As String)
                m_Enable = value
            End Set
        End Property








    End Class

End Namespace