﻿Namespace CLASS_IDEALTAKE.Objects

    ''' <summary>
    ''' La classe racine_moin_2 est utlisée dans la recherche guidée.
    ''' </summary>
    Public Class racine_moin_2

        Private m_ID_SR2 As Integer = -1
        Private m_ID_SR1 As Integer = -1
        Private m_Rubrique_Allemand As String = ""
        Private m_Rubrique_Anglais As String = ""
        Private m_Rubrique_Espagnole As String = ""
        Private m_Rubrique_Francais As String = ""
        Private m_Rubrique_Italien As String = ""
        Private m_Description_Allemand As String = ""
        Private m_Description_Anglais As String = ""
        Private m_Description_Espagnole As String = ""
        Private m_Description_Francais As String = ""
        Private m_Description_Italien As String = ""
        Private m_Symbole_Arbre As String = ""


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.ID_SR2 = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ID_SR2</c> est un entier identifiant de façon unique la rubrique en Racine -2.
        ''' </summary>
        Public Property ID_SR2() As Integer
            Get
                Return m_ID_SR2

            End Get
            Set(ByVal value As Integer)
                m_ID_SR2 = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.ID_SR1 = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ID_SR1</c> est un entier identifiant de façon unique la rubrique en Racine -1.
        ''' </summary>
        Public Property ID_SR1() As Integer
            Get
                Return m_ID_SR1

            End Get
            Set(ByVal value As Integer)
                m_ID_SR1 = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Rubrique_Allemand = "Nom rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Rubrique_Allemand</c> est la chaîne de la rubrique en Allemand.
        ''' </summary>
        Public Property Rubrique_Allemand() As String
            Get
                Return m_Rubrique_Allemand

            End Get
            Set(ByVal value As String)
                m_Rubrique_Allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Rubrique_Anglais = "Nom rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Rubrique_Anglais</c> est la chaîne de la rubrique en Anglais.
        ''' </summary>
        Public Property Rubrique_Anglais() As String
            Get
                Return m_Rubrique_Anglais

            End Get
            Set(ByVal value As String)
                m_Rubrique_Anglais = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Rubrique_Espagnole = "Nom rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Rubrique_Espagnole</c> est la chaîne de la rubrique en Espagnole.
        ''' </summary>
        Public Property Rubrique_Espagnole() As String
            Get
                Return m_Rubrique_Espagnole

            End Get
            Set(ByVal value As String)
                m_Rubrique_Espagnole = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Rubrique_Francais = "Nom rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Rubrique_Francais</c> est la chaîne de la rubrique en Francais.
        ''' </summary>
        Public Property Rubrique_Francais() As String
            Get
                Return m_Rubrique_Francais

            End Get
            Set(ByVal value As String)
                m_Rubrique_Francais = value
            End Set
        End Property





        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Rubrique_Italien = "Nom rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Rubrique_Italien</c> est la chaîne de la rubrique en Italien.
        ''' </summary>
        Public Property Rubrique_Italien() As String
            Get
                Return m_Rubrique_Italien

            End Get
            Set(ByVal value As String)
                m_Rubrique_Italien = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Description_Allemand = "Description rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_Allemand</c> est la chaîne décrivant la rubrique en Allemand.
        ''' </summary>
        Public Property Description_Allemand() As String
            Get
                Return m_Description_Allemand

            End Get
            Set(ByVal value As String)
                m_Description_Allemand = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Description_Anglais = "Description rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_Anglais</c> est la chaîne décrivant la rubrique en Anglais.
        ''' </summary>
        Public Property Description_Anglais() As String
            Get
                Return m_Description_Anglais

            End Get
            Set(ByVal value As String)
                m_Description_Anglais = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Description_Espagnole = "Description rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_Espagnole</c> est la chaîne décrivant la rubrique en Espagnole.
        ''' </summary>
        Public Property Description_Espagnole() As String
            Get
                Return m_Description_Espagnole

            End Get
            Set(ByVal value As String)
                m_Description_Espagnole = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Description_Francais = "Description rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_Francais</c> est la chaîne décrivant la rubrique en Francais.
        ''' </summary>
        Public Property Description_Francais() As String
            Get
                Return m_Description_Francais

            End Get
            Set(ByVal value As String)
                m_Description_Francais = value
            End Set
        End Property





        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_R_moin_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_R_moin_2.Description_Italien = "Description rubrique"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Description_Italien</c> est la chaîne décrivant la rubrique en Italien.
        ''' </summary>
        Public Property Description_Italien() As String
            Get
                Return m_Description_Italien

            End Get
            Set(ByVal value As String)
                m_Description_Italien = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj_2 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.racine_moin_2
        '''obj_2.Symbole_Arbre = "0-1"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Symbole_Arbre</c> est une chaîne permettant la recherche guidée.
        ''' </summary>
        Public Property Symbole_Arbre() As String
            Get
                Return m_Symbole_Arbre

            End Get
            Set(ByVal value As String)
                m_Symbole_Arbre = value
            End Set
        End Property

    End Class
End Namespace