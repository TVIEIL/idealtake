Public NotInheritable Class Attente_Chargement

    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ''Liste des tables à charger par chargement_collection_menus(...)
        '' table_a
        '' table_b
        '' table_c
        '' table_d
        '' table_e
        '' Table_niveau_acces
        '' Liste_des_Tables
        '' Table_Doc_URL
        '' racine_zero
        '' racine1
        '' racine2
        '' racine3
        '' racine4
        '' racine5
        '' racine6
        '' racine7
        '' racine8
        '' racine9
        '' types_documents
        '' CONSTRUCTEURS
        '' categories
        '' ref_constructeurs
        '' messages_utilisateur
        '' elements_winform

        'Table_à_charger.Text = "Nom de la Table"

        'Etat de chargement de la table
        'Etat_Chargement.Text = "Etat de chargement"
    End Sub

    Function Affichage_Etat_chargement(ByVal état As Integer)

        'état = 0 - - -
        'état = 1 Chargement en cours
        'état = 2 Chargement Terminé
        'état = 3 Mise à jour
        'état = 4 Pas de dialogue

        Me.Etat_Chargement.ForeColor = Color.Black

        Select Case Form_Demarrage.LangueLogiciel
            Case 0 ' Allemand

                If état = 0 Then
                    Me.Etat_Chargement.Text = "- - -"

                End If

                If état = 1 Then
                    Me.Etat_Chargement.Text = "Aktuelles Laden..."
                End If

                If état = 2 Then
                    Me.Etat_Chargement.Text = "Beendetes Laden."
                End If

                If état = 3 Then
                    Me.Etat_Chargement.Text = "Aktualisierung..."
                End If

                If état = 4 Then
                    Me.Etat_Chargement.ForeColor = Color.Red
                    Me.Etat_Chargement.Text = "  Kein Dialog!"
                End If

            Case 1 ' Anglais

                If état = 0 Then
                    Me.Etat_Chargement.Text = "- - -"
                End If

                If état = 1 Then
                    Me.Etat_Chargement.Text = "Loading..."
                End If

                If état = 2 Then
                    Me.Etat_Chargement.Text = "Ended load."
                End If

                If état = 3 Then
                    Me.Etat_Chargement.Text = " Update..."
                End If

                If état = 4 Then
                    Me.Etat_Chargement.ForeColor = Color.Red
                    Me.Etat_Chargement.Text = " No dialogue!"
                End If

            Case 2 ' Espagnole

                If état = 0 Then
                    Me.Etat_Chargement.Text = "- - -"
                End If

                If état = 1 Then
                    Me.Etat_Chargement.Text = "Carga corriente..."
                End If

                If état = 2 Then
                    Me.Etat_Chargement.Text = "Carga acabada."
                End If


                If état = 3 Then
                    Me.Etat_Chargement.Text = "Actualización..."
                End If

                If état = 4 Then
                    Me.Etat_Chargement.ForeColor = Color.Red
                    Me.Etat_Chargement.Text = "Ningún diálogo!"
                End If

            Case 3 ' Français

                If état = 0 Then
                    Me.Etat_Chargement.Text = "- - -"
                End If

                If état = 1 Then
                    Me.Etat_Chargement.Text = "Chargement en cours..."
                End If

                If état = 2 Then
                    Me.Etat_Chargement.Text = "Chargement terminé."
                End If


                If état = 3 Then
                    Me.Etat_Chargement.Text = " Mise à jour..."
                End If

                If état = 4 Then
                    Me.Etat_Chargement.ForeColor = Color.Red
                    Me.Etat_Chargement.Text = "Pas de dialogue!"
                End If

            Case 4 ' Italien

                If état = 0 Then
                    Me.Etat_Chargement.Text = "- - -"
                End If

                If état = 1 Then
                    Me.Etat_Chargement.Text = " Caricando in corso..."
                End If

                If état = 2 Then
                    Me.Etat_Chargement.Text = "Carico finito."
                End If

                If état = 3 Then
                    Me.Etat_Chargement.Text = "Collocamento aggiornato..."
                End If

                If état = 4 Then
                    Me.Etat_Chargement.ForeColor = Color.Red
                    Me.Etat_Chargement.Text = " Non di dialogo!"
                End If

        End Select

        Me.Refresh()


        Return True
    End Function

    Private Sub MainLayoutPanel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainLayoutPanel.Click

        Me.Refresh()


    End Sub


End Class
