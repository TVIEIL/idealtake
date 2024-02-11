Namespace CLASS_IDEALTAKE.Objects

    Public Class RésultatRechercheDetailsEnregistrement
        Implements IEquatable(Of RésultatRechercheDetailsEnregistrement)

        Public Numéro_enregistrement As String
        Public Chrono As Integer
        Public Auteur As String
        Public DateDeCreation As String
        Public Titre As String
        Public Sous_Titre As String
        Public Commentaire As String
        Public DateDeModification As String
        Public ListeDeRéférence As List(Of RésultatRechercheDetailsEnregistrementRéférence)
        Public ListeDocuments As List(Of RésultatRechercheDetailsEnregistrementDocuments)

        Public Sub New( _
                      ByVal m_NuméroEnregistrement As String, _
                      ByVal m_Chrono As String, _
                      ByVal m_Auteur As String, _
                      ByVal m_DateDeCreation As String, _
                      ByVal m_Titre As String, _
                      ByVal m_Sous_Titre As String, _
                      ByVal m_Commentaire As String, _
                      ByVal m_DateDeModification As String,
                      ByVal m_ListeDeRéférence As List(Of RésultatRechercheDetailsEnregistrementRéférence), _
                      ByVal m_ListeDocuments As List(Of RésultatRechercheDetailsEnregistrementDocuments))

            Numéro_enregistrement = m_NuméroEnregistrement
            Chrono = m_Chrono
            Auteur = m_Auteur
            DateDeCreation = m_DateDeCreation
            Titre = m_Titre
            Sous_Titre = m_Sous_Titre
            Commentaire = m_Commentaire
            DateDeModification = m_DateDeModification
            ListeDeRéférence = m_ListeDeRéférence
            ListeDocuments = m_ListeDocuments

        End Sub

        Public Overloads Function Equals(ByVal other As RésultatRechercheDetailsEnregistrement) _
      As Boolean Implements IEquatable(Of RésultatRechercheDetailsEnregistrement).Equals

            If Me.Numéro_enregistrement = other.Numéro_enregistrement And Me.Chrono = other.Chrono And Me.Auteur = other.Auteur And Me.DateDeCreation = other.DateDeCreation _
              And Me.Titre = other.Titre And Me.Sous_Titre = other.Sous_Titre And Me.Commentaire = other.Commentaire And Me.DateDeModification = other.DateDeModification Then
                Return True
            Else
                Return False
            End If

        End Function

        Private m_Liste_de_resultats_Details As List(Of RésultatRechercheDetailsEnregistrement)

        Public Property Liste_de_resultats_Details() As List(Of RésultatRechercheDetailsEnregistrement)

            Get
                Return m_Liste_de_resultats_Details
            End Get

            Set(ByVal value As List(Of RésultatRechercheDetailsEnregistrement))
                m_Liste_de_resultats_Details = value
            End Set

        End Property



        'Public Structure Référence

        'Public Constructeur As String
        'Public Catégorie As String
        'Public Référence As String
        'Public Description As String
        'Public photo As String

        'End Structure



        ' Public Structure Documents

        'Public Nom As String
        'Public Type As String
        'Public Taille As String
        'Public Lien As String

        ' End Structure

    End Class



    Public Class RésultatRechercheDetailsEnregistrementRéférence
        'Implements IEquatable(Of RésultatRechercheDetailsEnregistrement)

        Public Constructeur As String
        Public Catégorie As String
        Public Référence As String
        Public Description As String
        Public photo As String


        Public Sub New( _
                      ByVal m_Constructeur As String, _
                      ByVal m_Catégorie As String, _
                      ByVal m_Référence As String, _
                      ByVal m_Description As String, _
                      ByVal m_photo As String)


            Constructeur = m_Constructeur
            Catégorie = m_Catégorie
            Référence = m_Référence
            Description = m_Description
            photo = m_photo


        End Sub


    End Class


    Public Class RésultatRechercheDetailsEnregistrementDocuments
        'Implements IEquatable(Of RésultatRechercheDetailsEnregistrement)

        Public Nom As String
        Public Type As String
        Public Taille As String
        Public Lien As String


        Public Sub New( _
                      ByVal m_Nom As String, _
                      ByVal m_Type As String, _
                      ByVal m_Taille As String, _
                      ByVal m_Lien As String)


            Nom = m_Nom
            Type = m_Type
            Taille = m_Taille
            Lien = m_Lien


        End Sub


    End Class

End Namespace

