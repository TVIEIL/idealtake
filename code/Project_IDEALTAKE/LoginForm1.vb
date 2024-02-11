'Moteur de Recherche IDEALTAKE
'-----------------------------
'
'
'DESCRIPTION :
'
'Moteur de recherche de documents pdf utilisant le SGBD MysQL.
'Ce logiciel a été developpé en  VISUAL BASIC 2005, VISUAL BASIC 2008, VISUAL BASIC 2010.
'
'
'Copyright (C) 2006-2013 VIEIL Thierry Ce programme est un logiciel libre ;
'vous pouvez le redistribuer et/ou le modifier au titre des clauses 
'de la Licence Publique Générale GNU, telle que publiée par la 
'Free Software Foundation ; soit la version 2 de la Licence,
'ou (à votre discrétion) une version ultérieure quelconque.
'Ce programme est distribué dans l'espoir qu'il sera utile,
'mais SANS AUCUNE GARANTIE ; sans même une garantie implicite
'de COMMERCIABILITE ou DE CONFORMITE A UNE UTILISATION PARTICULIERE.
'
'Voir la Licence Publique Générale GNU pour plus de détails.
'Vous devriez avoir reçu un exemplaire de la Licence Publique Générale
' GNU avec ce programme ; si ce n'est pas le cas, écrivez à la  :
'
'    Free Software Foundation Inc.
'   51 Franklin Street, Fifth Floor
'    Boston, MA 02110-1301, USA.





Public Class LoginForm1



    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click



        Dim counter_1 As Integer
        Dim flag1 As Boolean
        Dim nom As String = ""
        Dim Nom_Compte_Utilisateurs_a_tester As String = ""
        Dim mot_de_passe As String = ""

        nom = UsernameTextBox.Text
        nom = LTrim(nom)
        nom = RTrim(nom)

        mot_de_passe = PasswordTextBox.Text
        mot_de_passe = LTrim(mot_de_passe)
        mot_de_passe = RTrim(mot_de_passe)

        For counter_1 = 0 To 100

            If (Microsoft.VisualBasic.Strings.Len(Form_Demarrage.Compte_Utilisateurs(counter_1, 1)) > 2) Then

                Form_Demarrage.Compte_Utilisateurs(counter_1, 1).Replace(vbNewLine, "") 'supprime d'éventuel retour de ligne
                'MsgBox(Form1.Compte_Utilisateurs(counter_1, 1) & "," & Form1.Compte_Utilisateurs(counter_1, 2) & "," & Form1.Compte_Utilisateurs(counter_1, 3))
                ' MsgBox(nom & "+_+" & Form1.Compte_Utilisateurs(counter_1, 1) & "_")

                Nom_Compte_Utilisateurs_a_tester = Form_Demarrage.Compte_Utilisateurs(counter_1, 1)

                If (nom = Nom_Compte_Utilisateurs_a_tester) Then
                    'Utilisateur Valide
                    If (mot_de_passe = Form_Demarrage.Compte_Utilisateurs(counter_1, 2)) Then
                        'mot de passe valide

                        If (Form_Demarrage.menu_Authentification = "Gestion Base") Then

                            Select Case Form_Demarrage.Compte_Utilisateurs(counter_1, 3) 'quel type de compte?

                                Case "Gestionnaire"
                                    Form_Demarrage.TabControl1.Enabled = True


                                Case "Administrateur"
                                    Form_Demarrage.TabControl1.Enabled = True
                                    Form_Demarrage.Gestion_Base_GroupBox.Enabled = True

                                Case Else
                                    'le droit enregistré dans le fichier crypté n'est pas reconnu comme valide
                                    'il est ni Administrateur ni Gestionnaire

                                    'MsgBox("Droit de l'utilisateur non reconnu contacté l'administrateur du logiciel", MsgBoxStyle.Information, "Authentification")
                                    MsgBox(Ajout_Suppression_connexion_base.MessageUtilisateur(145, Form_Demarrage.LangueLogiciel).ToString(), MsgBoxStyle.Information, _
                                           Ajout_Suppression_connexion_base.MessageUtilisateur(169, Form_Demarrage.LangueLogiciel).ToString())

                                    Form_Demarrage.TabControl1.SelectedIndex = 1
                                    Form_Demarrage.TabControl1.Enabled = True
                                    Form_Demarrage.Gestion_Base_GroupBox.Enabled = False
                                    Me.UsernameTextBox.Text = "Gestionnaire"
                                    Me.PasswordTextBox.Text = ""
                                    Me.Hide()
                                    Form_Demarrage.Show()

                            End Select
                        End If 'LoginForm a été appelé par l'onglet Gestion de base

                        'enregistre dans une variable global le nom de l'utilisateur
                        Form_Demarrage.Nom_Utilisateur = Form_Demarrage.Compte_Utilisateurs(counter_1, 1)

                        If (Form_Demarrage.menu_Authentification = "Personnaliser") Then

                            ' MsgBox("Ok ça Marche") 'en effet ça marche

                            Select Case Form_Demarrage.Compte_Utilisateurs(counter_1, 3) 'quel type de compte?

                                Case "Gestionnaire"

                                    'MsgBox("c'est un compte gestionnaire")
                                    Form_Demarrage.Bonjour_Utilisateur.Text = "Bonjour Gestionnaire : " & Form_Demarrage.Nom_Utilisateur
                                    Form_Demarrage.RadioButton2.Checked = True 'sélectionne Gestionnaire dans utilisateur Base

                                    'indique le nom de l'utilisateur
                                    Form_Demarrage.TextBox2.Text = Form_Demarrage.Nom_Utilisateur

                                    ' met le mot de passe actuel pour ce compte dans la zone mot de passe
                                    Form_Demarrage.TextBox3.Text = Form_Demarrage.Compte_Utilisateurs(counter_1, 2)

                                    'permet la modification du mot de passe
                                    Form_Demarrage.TextBox3.Enabled = True



                                Case "Administrateur"
                                    'MsgBox("c'est un compte Administrateur")
                                    Form_Demarrage.Bonjour_Utilisateur.Text = "Bonjour Administrateur : " & Form_Demarrage.Nom_Utilisateur

                                    'indique le nom de l'utilisateur
                                    'Form1.TextBox2.Text = Form1.Nom_Utilisateur

                                    'permet la modification du nom d'utilisateur
                                    Form_Demarrage.TextBox2.Enabled = True

                                    ' met le mot de passe actuel pour ce compte dans la zone mot de passe
                                    'Form1.TextBox3.Text = Form1.Compte_Utilisateurs(counter_1, 2)

                                    'permet la modification du mot de passe
                                    Form_Demarrage.TextBox3.Enabled = True

                                    Form_Demarrage.RadioButton28.Enabled = True 'autorise le bouton radio Ajouter
                                    Form_Demarrage.RadioButton29.Enabled = True 'autorise le bouton radio Modifier
                                    Form_Demarrage.RadioButton30.Enabled = True 'autorise le bouton radio Supprimer

                                    Form_Demarrage.RadioButton1.Checked = True 'sélectionne Administrateur dans utilisateur Base
                                    Form_Demarrage.RadioButton1.Enabled = True 'autorise le bouton Radio Administrateur
                                    Form_Demarrage.RadioButton2.Enabled = True 'autorise le bouton Radio Gestionnaire

                                    'on récupère la liste des utilisateurs à partir du Tableau 
                                    Dim T1 As Integer = 0

                                    ' autorise le ListBox
                                    Form_Demarrage.ListBox1.Enabled = True

                                    'Nettoyage du ListBox Effacement des valeurs déjà présente
                                    Form_Demarrage.ListBox1.Items.Clear()

                                    For T1 = 0 To 100
                                        If (Len(Form_Demarrage.Compte_Utilisateurs(T1, 1)) > 2) Then
                                            'on réinitialise les données du ListBox à partir
                                            'des noms utilisateurs stockés dans le tableau
                                            Form_Demarrage.ListBox1.Items.Add(Form_Demarrage.Compte_Utilisateurs(T1, 1)) '
                                        End If

                                    Next

                                Case Else
                                    'le droit enregistré dans le fichier crypté n'est pas reconnu comme valide
                                    'il est ni Administrateur ni Gestionnaire

                                    'MsgBox("Droit de l'utilisateur non reconnu contacté l'administrateur du logiciel", MsgBoxStyle.Information, "Authentification")
                                    MsgBox(Ajout_Suppression_connexion_base.MessageUtilisateur(145, Form_Demarrage.LangueLogiciel).ToString(), MsgBoxStyle.Information, _
                                           Ajout_Suppression_connexion_base.MessageUtilisateur(169, Form_Demarrage.LangueLogiciel).ToString())

                                    Form_Demarrage.TabControl1.SelectedIndex = 1
                                    Form_Demarrage.TabControl1.Enabled = True
                                    Form_Demarrage.Gestion_Base_GroupBox.Enabled = False
                                    Me.UsernameTextBox.Text = "Gestionnaire"
                                    Me.PasswordTextBox.Text = ""
                                    Me.Hide()
                                    Form_Demarrage.Show()
                            End Select


                            'empéche de pouvoir être authentifié de
                            'nouveau alors qu'on l'est déjà
                            Form_Demarrage.Button27.Enabled = False


                            'active le bouton valider
                            Form_Demarrage.Button32.Enabled = True

                            'active le bouton exit
                            Form_Demarrage.Button31.Enabled = True

                        End If 'LoginForm a été appelé par l'onglet personnaliser



                        'réinitialise la variable permettant de savoir d'ou vient l'appel
                        'de LoginForm par défaut c'est pour aller dans l'onglet Gestion Base
                        Form_Demarrage.menu_Authentification = "Gestion Base"

                        'Réinitialise les paramètres de LoginForm
                        Me.UsernameTextBox.Text = "Gestionnaire"
                        Me.PasswordTextBox.Text = ""
                        Me.Hide()
                        Form_Demarrage.Show()

                        ' Message de bienvenu à l'authentification correct
                        ' si la langue est en français et que la synthèse vocale est activée
                        '
                        'Valeur de Langue
                        ' 0 = Allemand
                        ' 1 = Anglais
                        ' 2 = Espagnole
                        ' 3 = Français
                        ' 4 = Italien

                        ' on va faire parler la synthèse vocale e.g : 'Bonjour Gestionnaire Thierry'

                        'modification le 30/10/2009
                        'Avant : If ((Form_Demarrage.Langue_logiciel = 3) And (My.Settings.Activation_Synthese_Vocale = True)) Then
                        'Maintenant :
                        If (My.Settings.Activation_Synthese_Vocale = True) Then
                            'MsgBox(lecture_du_resultat)
                            Dim a As System.EventArgs
                            a = System.EventArgs.Empty

                            Dim lecture_vocale_du_resultat As String = ""
                            Dim phrase_à_dire0 As String = ""
                            Dim phrase_à_dire1 As String = ""
                            'phrase_à_dire1 = Form_Demarrage.Bonjour_Utilisateur.Text.Replace(":", "")

                            'on rajoute le nom de l'utilisateur
                            phrase_à_dire1 = Form_Demarrage.Nom_Utilisateur.ToString
                            'Form_Demarrage.xvoice.Select(My.Settings.Index_Moteur_Synthèse_Vocale)


                            'If (My.Settings.Version_SAPI = "5.4") Then  'Début SAPI 5.4


                            'SAPI 5.4

                            'MsgBox(Formulaire_SAPI_5_4.idsVoices.Items.Count)
                            'MsgBox(My.Settings.Nom_Moteur_Synthèse_Vocale_SAPI5)
                            'Sélectionne le nom du moteur




                            'If (InStr(sender.ToString, Form_Demarrage.Parle.Text) > 0) Then

                            'Valeur de Langue
                            ' 0 = Allemand
                            ' 1 = Anglais
                            ' 2 = Espagnole
                            ' 3 = Français
                            ' 4 = Italien

                            Select Case Form_Demarrage.LangueLogiciel

                                Case 0
                                    'Allemand
                                    phrase_à_dire0 = "Guten Tag"

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Gestionnaire") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Geschaftsfuhrer " & phrase_à_dire1
                                    End If

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Administrateur") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Verwalter " & phrase_à_dire1
                                    End If
                                    'Form_Demarrage.VoiceObj_SAPI5.Speak("Guten Tag und willkommen IDEALE Tec", 1)

                                Case 1
                                    'Americain Anglais
                                    phrase_à_dire0 = "Hello"

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Gestionnaire") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Controller " & phrase_à_dire1
                                    End If

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Administrateur") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Administrator " & phrase_à_dire1
                                    End If

                                    'Form_Demarrage.VoiceObj_SAPI5.Speak("Hello and Welcome to E D altech", 1)

                                Case 2
                                    'Espagnole
                                    phrase_à_dire0 = "Buenos dias"

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Gestionnaire") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Gerente " & phrase_à_dire1
                                    End If

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Administrateur") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Administrador " & phrase_à_dire1
                                    End If

                                    'Form_Demarrage.VoiceObj_SAPI5.Speak("Buenos dias y bienvenido IDEALE Tec", 1)

                                Case 3
                                    'Français
                                    phrase_à_dire0 = "Bonjour"

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Gestionnaire") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Gestionnaire " & phrase_à_dire1
                                    End If

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Administrateur") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Administrateur " & phrase_à_dire1
                                    End If
                                    'Form_Demarrage.VoiceObj_SAPI5.Speak("Bonjour et bienvenu sur idé al teq", 1)

                                Case 4
                                    'Italien
                                    phrase_à_dire0 = "Buon giorno"

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Gestionnaire") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Gestore " & phrase_à_dire1
                                    End If

                                    If (Form_Demarrage.Compte_Utilisateurs(counter_1, 3).ToString = "Administrateur") Then
                                        phrase_à_dire0 = phrase_à_dire0 & " Amministratore " & phrase_à_dire1
                                    End If

                                    'Form_Demarrage.VoiceObj_SAPI5.Speak("Buon giorno e benvenuto all' IDEALE Tec", 1)

                                Case Else

                            End Select

                            'My.Settings.Texte_à_dire = phrase_à_dire0
                            'Form_Demarrage.Parle_Click(lecture_vocale_du_resultat, a)

                            Form_Demarrage.modification_commandes_SAPI5_TEXTE_A_DIRE(phrase_à_dire0, My.Application.Info.DirectoryPath)


                        End If 'FIN SAPI 5.4

                        Exit Sub 'Utilisateur reconnu

                    Else 'mot de passe non valide
                        flag1 = True
                        'If (counter_1 = 100) Then GoTo label1
                        GoTo label1

                    End If 'FIN  If (mot_de_passe = Form_Demarrage.Compte_Utilisateurs(counter_1, 2))






                Else 'utilisateur non valide
                    flag1 = False
                    If (counter_1 = 100) Then GoTo label1


                End If ' fin If (nom = Nom_Compte_Utilisateurs_a_tester) 



            End If

            ' MsgBox(Form1.Compte_Utilisateurs(counter_1, 1) & "," & Form1.Compte_Utilisateurs(counter_1, 2) & "," & Form1.Compte_Utilisateurs(counter_1, 3))

            'End If
            'End If
        Next counter_1

label1:
        If flag1 = False Then  'utilisateur non valide

            'MsgBox("Utilisateur non valide", MsgBoxStyle.Information, "Authentification")
            MsgBox(Ajout_Suppression_connexion_base.MessageUtilisateur(146, Form_Demarrage.LangueLogiciel).ToString(), MsgBoxStyle.Information, _
                                           Ajout_Suppression_connexion_base.MessageUtilisateur(169, Form_Demarrage.LangueLogiciel).ToString())

            Form_Demarrage.TabControl1.SelectedIndex = 1
            Form_Demarrage.TabControl1.Enabled = True
            Form_Demarrage.Gestion_Base_GroupBox.Enabled = False
            Me.UsernameTextBox.Text = "Gestionnaire"
            Me.PasswordTextBox.Text = ""
            Me.Hide()
            Form_Demarrage.Show()
        End If


        If flag1 = True Then  'mot de passe non valide

            'MsgBox("Mot de passe non valide", MsgBoxStyle.Information, "Authentification")
            MsgBox(Ajout_Suppression_connexion_base.MessageUtilisateur(147, Form_Demarrage.LangueLogiciel).ToString(), MsgBoxStyle.Information, _
                                           Ajout_Suppression_connexion_base.MessageUtilisateur(169, Form_Demarrage.LangueLogiciel).ToString())

            Form_Demarrage.TabControl1.SelectedIndex = 1
            Form_Demarrage.TabControl1.Enabled = True
            Me.UsernameTextBox.Text = "Gestionnaire"
            Me.PasswordTextBox.Text = ""
            Me.Hide()
            Form_Demarrage.Show()
        End If



    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click

        If (Form_Demarrage.menu_Authentification = "Gestion Base") Then

            Form_Demarrage.TabControl1.SelectedIndex = 1
            Form_Demarrage.TabControl1.Enabled = True
            'Form1.GroupBox13.Enabled = False
        End If 'LoginForm a été appelé par l'onglet Gestion de base

        If (Form_Demarrage.menu_Authentification = "Personnaliser") Then

            'Réinitialise la variable d'ou vient l'appel de
            'LoginForm par défaut il est vers l'onglet de Gestion de Base
            Form_Demarrage.menu_Authentification = "Gestion Base"
        End If  'LoginForm a été appelé par l'onglet Personnaliser

        Me.UsernameTextBox.Text = "Gestionnaire"
        Me.PasswordTextBox.Text = ""
        Me.Hide()
        Form_Demarrage.Show()

    End Sub

    Private Sub UsernameLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsernameLabel.Click
        If (Form_Demarrage.FonctionModificationObjet(UsernameLabel, sender.ToString) = True) Then
            Exit Sub
        End If
    End Sub

    Private Sub PasswordLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasswordLabel.Click
        If (Form_Demarrage.FonctionModificationObjet(PasswordLabel, sender.ToString) = True) Then
            Exit Sub
        End If
    End Sub
End Class
