

Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe HISTORY_EVENTS stocke l'historique de l'activité sur la base.
    ''' </summary>
    Public Class HISTORY_EVENTS

        Private m_ID_HISTORY_EVENTS As Integer = -1
        Private m_Date As String = ""
        Private m_User_ID As String = ""
        Private m_IP As String = ""
        Private m_Name_of_Computer As String = ""
        Private m_Operating_System As String = ""
        Private m_Action As String = ""
        Private m_Table_Name As String = ""
        Private m_ID_Record As String = ""
        Private m_SQL_String As String = ""
        Private m_Result_Action As String = ""


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''obj1.ID_HISTORY_EVENTS = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ID_HISTORY_EVENTS</c> est un entier identifiant de façon unique un événement de l'historique.
        ''' </summary>
        Public Property ID_HISTORY_EVENTS() As Integer
            Get
                Return m_ID_HISTORY_EVENTS

            End Get
            Set(ByVal value As Integer)
                m_ID_HISTORY_EVENTS = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.Date_evenement = "14/07/1789..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Date_evenement</c> indique la date d'un événement de l'historique.
        ''' </summary>
        Public Property Date_evenement() As String
            Get
                Return m_Date

            End Get
            Set(ByVal value As String)
                m_Date = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.User_ID = "Victor007"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>User_ID</c> indique le nom de l'utilisateur à l'origine de l'événement.
        ''' </summary>
        Public Property User_ID() As String
            Get
                Return m_User_ID

            End Get
            Set(ByVal value As String)
                m_User_ID = value
            End Set
        End Property




        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.IP = "217.235.21.15"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>IP</c> indique l'I.P de l'utilisateur à l'origine de l'événement.
        ''' </summary>
        Public Property IP() As String
            Get
                Return m_IP

            End Get
            Set(ByVal value As String)
                m_IP = value
            End Set
        End Property




        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.Name_of_Computer = "Victor_computer"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Name_of_Computer</c> indique le nom de l'ordinateur à l'origine de l'événement.
        ''' </summary>
        Public Property Name_of_Computer() As String
            Get
                Return m_Name_of_Computer

            End Get
            Set(ByVal value As String)
                m_Name_of_Computer = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.Operating_System = "Windows 7"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Operating_System</c> indique le nom du système d'exploitation à l'origine de l'événement.
        ''' </summary>
        Public Property Operating_System() As String
            Get
                Return m_Operating_System

            End Get
            Set(ByVal value As String)
                m_Operating_System = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.Action = "Delete"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Action</c> indique le type d'action à l'origine de l'événement.
        ''' </summary>
        Public Property Action() As String
            Get
                Return m_Action

            End Get
            Set(ByVal value As String)
                m_Action = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.Table_Name = "Enregistrements"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Table_Name</c> indique le nom de la table lié à l'événement.
        ''' </summary>
        Public Property Table_Name() As String
            Get
                Return m_Table_Name

            End Get
            Set(ByVal value As String)
                m_Table_Name = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.ID_Record = "2153"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>ID_Record</c> indique le numéro d'enregistrement lié à l'événement.
        ''' </summary>
        Public Property ID_Record() As String
            Get
                Return m_ID_Record

            End Get
            Set(ByVal value As String)
                m_ID_Record = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.SQL_String = "SELECT * FROM..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>SQL_String</c> requête SQL de l'événement.
        ''' </summary>
        Public Property SQL_String() As String
            Get
                Return m_SQL_String

            End Get
            Set(ByVal value As String)
                m_SQL_String = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.HISTORY_EVENTS
        '''Obj1.Result_Action = "Failed"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>Result_Action</c> indique le statut de la requête SQL.
        ''' </summary>
        Public Property Result_Action() As String
            Get
                Return m_Result_Action

            End Get
            Set(ByVal value As String)
                m_Result_Action = value
            End Set
        End Property







    End Class

End Namespace