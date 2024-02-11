Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe messages_utilisateur manipule des messages utilisateur.
    ''' </summary>
    ''' 
    Public Class messages_utilisateur

        Private m_ID_Message As Integer = -1
        Private m_Type_Message As String = ""
        Private m_Anglais As String = ""
        Private m_Allemand As String = ""
        Private m_Espagnole As String = ""
        Private m_Italien As String = ""
        Private m_Français As String = ""

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim ID_Message As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.messages_utilisateur
        '''ID_Message.ID_Message = 23
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ID_Message</c> est un entier identifiant de façon unique un message utilisateur.
        ''' </summary>
        Public Property ID_Message() As Integer
            Get
                Return m_ID_Message

            End Get
            Set(ByVal value As Integer)
                m_ID_Message = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type_Message As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.messages_utilisateur
        '''Type_Message.Type_Message = "Information"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_Message</c> est une chaîne identifiant la nature du message utilisateur
        ''' </summary>
        Public Property Type_Message() As String
            Get
                Return m_Type_Message

            End Get
            Set(ByVal value As String)
                m_Type_Message = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim description_Anglais As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.messages_utilisateur
        '''description_Anglais.Anglais = "Close this form"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Anglais</c> est une chaîne contenant un message utilisateur en Anglais.
        ''' </summary>
        Public Property Anglais() As String
            Get
                Return m_Anglais

            End Get
            Set(ByVal value As String)
                m_Anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim description_msg As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.messages_utilisateur
        '''description_msg.Allemand = "Das ist gut"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Allemand</c> est une chaîne contenant un message utilisateur en Allemand.
        ''' </summary>
        Public Property Allemand() As String
            Get
                Return m_Allemand

            End Get
            Set(ByVal value As String)
                m_Allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim description_msg As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.messages_utilisateur
        '''description_msg.Espagnole = "hola"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Espagnole</c> est une chaîne contenant un message utilisateur en Espagnole.
        ''' </summary>
        Public Property Espagnole() As String
            Get
                Return m_Espagnole

            End Get
            Set(ByVal value As String)
                m_Espagnole = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim description_msg As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.messages_utilisateur
        '''description_msg.Italien = "Mangare"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Italien</c> est une chaîne contenant un message utilisateur en Italien.
        ''' </summary>
        Public Property Italien() As String
            Get
                Return m_Italien

            End Get
            Set(ByVal value As String)
                m_Italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim description_msg As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.messages_utilisateur
        '''description_msg.Francais = "Bonjour"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Francais</c> est une chaîne contenant un message utilisateur en Français.
        ''' </summary>
        Public Property Francais() As String
            Get
                Return m_Français

            End Get
            Set(ByVal value As String)
                m_Français = value
            End Set
        End Property

    End Class
End Namespace
