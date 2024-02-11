Namespace CLASS_IDEALTAKE.Objects

    Public Class Résultats_recherche_avancée
        Implements IEquatable(Of Résultats_recherche_avancée)

        Public Numéro_enregistrement As String
        Public Chrono As Integer
        Public Titre As String
        Public Sous_Titre As String
        Public Commentaire As String

        Public Sub New( _
                      ByVal m_Numéro_enregistrement As String, _
                      ByVal m_Chrono As String, _
                      ByVal m_Titre As String, _
                      ByVal m_Sous_Titre As String, _
                      ByVal m_Commentaire As String)

            Numéro_enregistrement = m_Numéro_enregistrement
            Chrono = m_Chrono
            Titre = m_Titre
            Sous_Titre = m_Sous_Titre
            Commentaire = m_Commentaire

        End Sub

        Public Overloads Function Equals(ByVal other As Résultats_recherche_avancée) _
      As Boolean Implements IEquatable(Of Résultats_recherche_avancée).Equals

            If Me.Numéro_enregistrement = other.Numéro_enregistrement And Me.Chrono = other.Chrono _
              And Me.Titre = other.Titre And Me.Sous_Titre = other.Sous_Titre And Me.Commentaire = other.Commentaire Then
                Return True
            Else
                Return False
            End If

        End Function

        Private m_Liste_de_resultats_RA As List(Of Résultats_recherche_avancée)

        Public Property Liste_de_resultats_RA() As List(Of Résultats_recherche_avancée)

            Get
                Return m_Liste_de_resultats_RA
            End Get

            Set(ByVal value As List(Of Résultats_recherche_avancée))
                m_Liste_de_resultats_RA = value
            End Set

        End Property

    End Class

End Namespace