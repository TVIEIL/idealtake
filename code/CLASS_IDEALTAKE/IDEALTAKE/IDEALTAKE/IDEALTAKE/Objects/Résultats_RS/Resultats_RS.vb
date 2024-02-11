Namespace CLASS_IDEALTAKE.Objects
    Public Class Résultats_recherche_simplifiée
        Implements IEquatable(Of Résultats_recherche_simplifiée)

        Public Numéro_enregistrement As String
        Public Chrono As Integer
        Public Auteur As String
        Public DateEnregistrement As String
        Public Titre As String
        Public Sous_Titre As String
        Public Commentaire As String


        Public Sub New( _
                      ByVal m_Numéro_enregistrement As String, _
                      ByVal m_Chrono As String, _
                      ByVal m_Auteur As String, _
                      ByVal m_DateEnregistrement As String, _
                      ByVal m_Titre As String, _
                      ByVal m_Sous_Titre As String, _
                      ByVal m_Commentaire As String)

            Numéro_enregistrement = m_Numéro_enregistrement
            Chrono = m_Chrono
            Auteur = m_Auteur
            DateEnregistrement = m_DateEnregistrement
            Titre = m_Titre
            Sous_Titre = m_Sous_Titre
            Commentaire = m_Commentaire

        End Sub

        Public Overloads Function Equals(ByVal other As Résultats_recherche_simplifiée) _
      As Boolean Implements IEquatable(Of Résultats_recherche_simplifiée).Equals

            If Me.Numéro_enregistrement = other.Numéro_enregistrement And Me.Chrono = other.Chrono _
              And Me.Titre = other.Titre And Me.Sous_Titre = other.Sous_Titre And Me.Commentaire = other.Commentaire Then
                Return True
            Else
                Return False
            End If

        End Function

        Private m_Liste_de_resultats_RS As List(Of Résultats_recherche_simplifiée)

        Public Property Liste_de_resultats_RS() As List(Of Résultats_recherche_simplifiée)

            Get
                Return m_Liste_de_resultats_RS
            End Get

            Set(ByVal value As List(Of Résultats_recherche_simplifiée))
                m_Liste_de_resultats_RS = value
            End Set

        End Property

    End Class

End Namespace