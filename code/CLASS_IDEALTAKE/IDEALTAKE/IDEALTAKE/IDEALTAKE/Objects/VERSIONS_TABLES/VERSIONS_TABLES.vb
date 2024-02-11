Namespace CLASS_IDEALTAKE.Objects

    ''' <summary>
    ''' La classe VERSIONS_TABLES stocke des informations sur les versions des tables dans la base de données.
    ''' </summary>
    Public Class VERSIONS_TABLES



        Private m_ID As Integer = -1
        Private m_Date As Date = Date.Today
        Private m_categories As Integer = -1
        Private m_chrono_idref As Integer = -1
        Private m_CONSTRUCTEURS As Integer = -1
        Private m_doc_url As Integer = -1
        Private m_documents As Integer = -1
        Private m_elements_HTML As Integer = -1
        Private m_elements_winform As Integer = -1
        Private m_enregistrements As Integer = -1
        Private m_HISTORY_EVENTS As Integer = -1
        Private m_messages_utilisateur As Integer = -1
        Private m_niveau_acces As Integer = -1
        Private m_Questions As Integer = -1
        Private m_racine_zero As Integer = -1
        Private m_racine1 As Integer = -1
        Private m_racine2 As Integer = -1
        Private m_racine3 As Integer = -1
        Private m_racine4 As Integer = -1
        Private m_racine5 As Integer = -1
        Private m_racine6 As Integer = -1
        Private m_racine7 As Integer = -1
        Private m_racine8 As Integer = -1
        Private m_racine9 As Integer = -1
        Private m_ref_constructeurs As Integer = -1
        Private m_schema As Integer = -1
        Private m_TABLE_A As Integer = -1
        Private m_TABLE_B As Integer = -1
        Private m_TABLE_C As Integer = -1
        Private m_TABLE_D As Integer = -1
        Private m_TABLE_E As Integer = -1
        Private m_TABLE_XAML As Integer = -1
        Private m_types_documents As Integer = -1
        Private m_Users As Integer = -1
        Private m_Commentaire As String = ""


        'ID
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.ID = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ID</c> est un entier identifiant le numéro de ligne dans la table.
        ''' </summary>
        Public Property ID() As Integer
            Get
                Return m_ID


            End Get
            Set(ByVal value As Integer)
                m_ID = value
            End Set
        End Property



        'Date
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.RECORD_DATE = Date.Today
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>RECORD_DATE</c> est la date d'enregistrement de la ligne dans la base de données.
        ''' </summary>
        Public Property RECORD_DATE() As Date
            Get
                Return m_Date


            End Get
            Set(ByVal value As Date)
                m_Date = value
            End Set
        End Property



        'categories
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.categories = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>categories</c> est un entier identifiant la version de la table categories.
        ''' </summary>
        Public Property categories() As Integer
            Get
                Return m_categories


            End Get
            Set(ByVal value As Integer)
                m_categories = value
            End Set
        End Property


        'chrono_idref
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.chrono_idref = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>chrono_idref</c> est un entier identifiant la version de la table chrono_idref.
        ''' </summary>
        Public Property chrono_idref() As Integer
            Get
                Return m_chrono_idref


            End Get
            Set(ByVal value As Integer)
                m_chrono_idref = value
            End Set
        End Property


        'CONSTRUCTEURS
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.CONSTRUCTEURS = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>CONSTRUCTEURS</c> est un entier identifiant la version de la table CONSTRUCTEURS.
        ''' </summary>
        Public Property CONSTRUCTEURS() As Integer
            Get
                Return m_CONSTRUCTEURS


            End Get
            Set(ByVal value As Integer)
                m_CONSTRUCTEURS = value
            End Set
        End Property


        'doc_url
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.doc_url = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>doc_url</c> est un entier identifiant la version de la table doc_url.
        ''' </summary>
        Public Property doc_url() As Integer
            Get
                Return m_doc_url


            End Get
            Set(ByVal value As Integer)
                m_doc_url = value
            End Set
        End Property


        'documents
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.documents = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>documents</c> est un entier identifiant la version de la table documents.
        ''' </summary>
        Public Property documents() As Integer
            Get
                Return m_documents


            End Get
            Set(ByVal value As Integer)
                m_documents = value
            End Set
        End Property


        'elements_HTML
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.elements_HTML = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>elements_HTML</c> est un entier identifiant la version de la table elements_HTML.
        ''' </summary>
        Public Property elements_HTML() As Integer
            Get
                Return m_elements_HTML


            End Get
            Set(ByVal value As Integer)
                m_elements_HTML = value
            End Set
        End Property

        'elements_winform
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.elements_winform = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>elements_winform</c> est un entier identifiant la version de la table elements_winform.
        ''' </summary>
        Public Property elements_winform() As Integer
            Get
                Return m_elements_winform


            End Get
            Set(ByVal value As Integer)
                m_elements_winform = value
            End Set
        End Property


        'enregistrements
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.enregistrements = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>enregistrements</c> est un entier identifiant la version de la table enregistrements.
        ''' </summary>
        Public Property enregistrements() As Integer
            Get
                Return m_enregistrements


            End Get
            Set(ByVal value As Integer)
                m_enregistrements = value
            End Set
        End Property

        'HISTORY_EVENTS
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.HISTORY_EVENTS = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>HISTORY_EVENTS</c> est un entier identifiant la version de la table HISTORY_EVENTS.
        ''' </summary>
        Public Property HISTORY_EVENTS() As Integer
            Get
                Return m_HISTORY_EVENTS


            End Get
            Set(ByVal value As Integer)
                m_HISTORY_EVENTS = value
            End Set
        End Property

        'messages_utilisateur
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.messages_utilisateur = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>messages_utilisateur</c> est un entier identifiant la version de la table messages_utilisateur.
        ''' </summary>
        Public Property messages_utilisateur() As Integer
            Get
                Return m_messages_utilisateur


            End Get
            Set(ByVal value As Integer)
                m_messages_utilisateur = value
            End Set
        End Property

        'niveau_acces
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.niveau_acces = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>niveau_acces</c> est un entier identifiant la version de la table niveau_acces.
        ''' </summary>
        Public Property niveau_acces() As Integer
            Get
                Return m_niveau_acces


            End Get
            Set(ByVal value As Integer)
                m_niveau_acces = value
            End Set
        End Property

        'Questions
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.Questions = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Questions</c> est un entier identifiant la version de la table Questions.
        ''' </summary>
        Public Property Questions() As Integer
            Get
                Return m_Questions


            End Get
            Set(ByVal value As Integer)
                m_Questions = value
            End Set
        End Property

        'racine_zero
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine_zero = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine_zero</c> est un entier identifiant la version de la table racine_zero.
        ''' </summary>
        Public Property racine_zero() As Integer
            Get
                Return m_racine_zero


            End Get
            Set(ByVal value As Integer)
                m_racine_zero = value
            End Set
        End Property

        'racine1
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine1 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine1</c> est un entier identifiant la version de la table racine1.
        ''' </summary>
        Public Property racine1() As Integer
            Get
                Return m_racine1


            End Get
            Set(ByVal value As Integer)
                m_racine1 = value
            End Set
        End Property

        'racine2
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine2 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine2</c> est un entier identifiant la version de la table racine2.
        ''' </summary>
        Public Property racine2() As Integer
            Get
                Return m_racine2


            End Get
            Set(ByVal value As Integer)
                m_racine2 = value
            End Set
        End Property

        'racine3
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine3 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine3</c> est un entier identifiant la version de la table racine3.
        ''' </summary>
        Public Property racine3() As Integer
            Get
                Return m_racine3


            End Get
            Set(ByVal value As Integer)
                m_racine3 = value
            End Set
        End Property

        'racine4
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine4 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine4</c> est un entier identifiant la version de la table racine4.
        ''' </summary>
        Public Property racine4() As Integer
            Get
                Return m_racine4


            End Get
            Set(ByVal value As Integer)
                m_racine4 = value
            End Set
        End Property

        'racine5
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine5 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine5</c> est un entier identifiant la version de la table racine5.
        ''' </summary>
        Public Property racine5() As Integer
            Get
                Return m_racine5


            End Get
            Set(ByVal value As Integer)
                m_racine5 = value
            End Set
        End Property

        'racine6
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine6 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine6</c> est un entier identifiant la version de la table racine6.
        ''' </summary>
        Public Property racine6() As Integer
            Get
                Return m_racine6


            End Get
            Set(ByVal value As Integer)
                m_racine6 = value
            End Set
        End Property

        'racine7
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine7 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine7</c> est un entier identifiant la version de la table racine7.
        ''' </summary>
        Public Property racine7() As Integer
            Get
                Return m_racine7


            End Get
            Set(ByVal value As Integer)
                m_racine7 = value
            End Set
        End Property

        'racine8
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine8 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine8</c> est un entier identifiant la version de la table racine8.
        ''' </summary>
        Public Property racine8() As Integer
            Get
                Return m_racine8


            End Get
            Set(ByVal value As Integer)
                m_racine8 = value
            End Set
        End Property

        'racine9
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.racine9 = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>racine9</c> est un entier identifiant la version de la table racine9.
        ''' </summary>
        Public Property racine9() As Integer
            Get
                Return m_racine9


            End Get
            Set(ByVal value As Integer)
                m_racine9 = value
            End Set
        End Property

        'ref_constructeurs
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.ref_constructeurs = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ref_constructeurs</c> est un entier identifiant la version de la table ref_constructeurs.
        ''' </summary>
        Public Property ref_constructeurs() As Integer
            Get
                Return m_ref_constructeurs


            End Get
            Set(ByVal value As Integer)
                m_ref_constructeurs = value
            End Set
        End Property

        'schema
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.schema = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>schema</c> est un entier identifiant la version de la table schema.
        ''' </summary>
        Public Property schema() As Integer
            Get
                Return m_schema


            End Get
            Set(ByVal value As Integer)
                m_schema = value
            End Set
        End Property

        'TABLE_A
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.TABLE_A = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TABLE_A</c> est un entier identifiant la version de la table TABLE_A.
        ''' </summary>
        Public Property TABLE_A() As Integer
            Get
                Return m_TABLE_A


            End Get
            Set(ByVal value As Integer)
                m_TABLE_A = value
            End Set
        End Property

        'TABLE_B
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.TABLE_B = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TABLE_B</c> est un entier identifiant la version de la table TABLE_B.
        ''' </summary>
        Public Property TABLE_B() As Integer
            Get
                Return m_TABLE_B


            End Get
            Set(ByVal value As Integer)
                m_TABLE_B = value
            End Set
        End Property

        'TABLE_C
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.TABLE_C = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TABLE_C</c> est un entier identifiant la version de la table TABLE_C.
        ''' </summary>
        Public Property TABLE_C() As Integer
            Get
                Return m_TABLE_C


            End Get
            Set(ByVal value As Integer)
                m_TABLE_C = value
            End Set
        End Property

        'TABLE_D
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.TABLE_D = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TABLE_D</c> est un entier identifiant la version de la table TABLE_D.
        ''' </summary>
        Public Property TABLE_D() As Integer
            Get
                Return m_TABLE_D


            End Get
            Set(ByVal value As Integer)
                m_TABLE_D = value
            End Set
        End Property


        'TABLE_E
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.TABLE_E = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TABLE_E</c> est un entier identifiant la version de la table TABLE_E.
        ''' </summary>
        Public Property TABLE_E() As Integer
            Get
                Return m_TABLE_E


            End Get
            Set(ByVal value As Integer)
                m_TABLE_E = value
            End Set
        End Property

        'TABLE_XAML
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.TABLE_XAML = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TABLE_XAML</c> est un entier identifiant la version de la table TABLE_XAML.
        ''' </summary>
        Public Property TABLE_XAML() As Integer
            Get
                Return m_TABLE_XAML


            End Get
            Set(ByVal value As Integer)
                m_TABLE_XAML = value
            End Set
        End Property

        'types_documents
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.types_documents = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>types_documents</c> est un entier identifiant la version de la table types_documents.
        ''' </summary>
        Public Property types_documents() As Integer
            Get
                Return m_types_documents


            End Get
            Set(ByVal value As Integer)
                m_types_documents = value
            End Set
        End Property

        'Users
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.Users = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Users</c> est un entier identifiant la version de la table Users.
        ''' </summary>
        Public Property Users() As Integer
            Get
                Return m_Users


            End Get
            Set(ByVal value As Integer)
                m_Users = value
            End Set
        End Property

        'Commentaire
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Obj_VERSIONS_TABLES As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.VERSIONS_TABLES
        '''Obj_VERSIONS_TABLES.Commentaire = "Mise à jour des données de TABLE_A et TABLE_B"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Commentaire</c> est une chaîne donnant des informations sur la mise à jour.
        ''' </summary>
        Public Property Commentaire() As String
            Get
                Return m_Commentaire


            End Get
            Set(ByVal value As String)
                m_Commentaire = value
            End Set
        End Property
    End Class

End Namespace