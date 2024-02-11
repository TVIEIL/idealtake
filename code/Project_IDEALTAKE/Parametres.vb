'IDEALTAKE
'---------
'
'
'DESCRIPTION :
'
'Moteur de recherche de documents pdf utilisant le SGBD MysQL.
'Ce logiciel a été developpé en  VISUAL BASIC 2005, VISUAL BASIC 2008, VISUAL BASIC 2010.
'
'
'Copyright (C) 2007-2013 VIEIL Thierry Ce programme est un logiciel libre ;
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
'GNU avec ce programme ; si ce n'est pas le cas, écrivez à la  :
'
'    Free Software Foundation Inc.
'    51 Franklin Street, Fifth Floor
'    Boston, MA 02110-1301, USA.
'

Imports Microsoft.Win32
Imports System
Imports System.String
Imports System.Data
Imports System.Data.Common
Imports System.Collections
Imports MySql.Data.MySqlClient
Imports System.IO
Imports system.xml


''Structure de la Table MENU_SYSTEME
''mysql> USE IDEALTAKE;
''Database changed
''mysql> show tables;
''+-----------------+
''| Tables_in_IDEALTAKE |
''+-----------------+
''| panne           |
''| schema          |
''+-----------------+
''2 rows in set (0.00 sec)

''CREATE TABLE `IDEALTAKE`.`MENU_SYSTEME` (
''  `INDICE_AUTO` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
''  `INFOS_VISIBLE_LOGICIEL` TEXT,
''  `REQUETE_REELLE` TEXT,
''  `VALEUR_AUDIO_VIDEO` TEXT,
''  `FIGER_AUDIO_VIDEO` TEXT,
''  PRIMARY KEY(`INDICE_AUTO`)
'')
''ENGINE = InnoDB;
''
''ALTER TABLE `IDEALTAKE`.`menu_systeme` ADD COLUMN `SUPPRIMER_DE_AUDIO_VIDEO` TEXT AFTER `FIGER_AUDIO_VIDEO`;
''mysql>
''
''mysql> show tables;
''+-----------------+
''| Tables_in_IDEALTAKE |
''+-----------------+
''| menu_systeme    |
''| panne           |
''| schema          |
''+-----------------+
''3 rows in set (0.00 sec)
''
''mysql> describe menu_systeme;
''mysql> describe menu_systeme;
''+--------------------------+------------------+------+-----+---------+----------------+
''| Field                    | Type             | Null | Key | Default | Extra          |
''+--------------------------+------------------+------+-----+---------+----------------+
''| INDICE_AUTO              | int(10) unsigned | NO   | PRI | NULL    | auto_increment |
''| INFOS_VISIBLE_LOGICIEL   | text             | YES  |     | NULL    |                |
''| REQUETE_REELLE           | text             | YES  |     | NULL    |                |
''| VALEUR_AUDIO_VIDEO       | text             | YES  |     | NULL    |                |
''| FIGER_AUDIO_VIDEO        | text             | YES  |     | NULL    |                |
''| SUPPRIMER_DE_AUDIO_VIDEO | text             | YES  |     | NULL    |                |
''+--------------------------+------------------+------+-----+---------+----------------+
''6 rows in set (0.00 sec)
''
''mysql>
''
''Création d'une table pour Audio/Video
''
''CREATE TABLE `IDEALTAKE`.`menu_audio_video` (
''  `INDICE_AUTO` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
''  `VISIBLE LOGICIEL` TEXT,
''  `VALEUR REELLE` TEXT,
''  PRIMARY KEY(`INDICE_AUTO`)
'')
''ENGINE = InnoDB;
''
'mysql> describe menu_audio_video;
''+------------------+------------------+------+-----+---------+----------------+
''| Field            | Type             | Null | Key | Default | Extra          |
''+------------------+------------------+------+-----+---------+----------------+
''| INDICE_AUTO      | int(10) unsigned | NO   | PRI | NULL    | auto_increment |
''| VISIBLE LOGICIEL | text             | YES  |     | NULL    |                |
''| VALEUR REELLE    | text             | YES  |     | NULL    |                |
''+------------------+------------------+------+-----+---------+----------------+
''3 rows in set (0.02 sec)
''
''mysql>
''
''
''
''Création d'une table de contrôle d'acces
''
''CREATE TABLE `IDEALTAKE`.`MENU_CONTROLE_ACCES` (
''  `INDICE AUTO` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
''  `VISIBLE LOGICIEL` TEXT,
''  `VALEUR REELLE` TEXT,
''  PRIMARY KEY(`INDICE AUTO`)
'')
''ENGINE = InnoDB;
''
''mysql> describe menu_controle_acces;
''+------------------+------------------+------+-----+---------+----------------+
''| Field            | Type             | Null | Key | Default | Extra          |
''+------------------+------------------+------+-----+---------+----------------+
''| INDICE AUTO      | int(10) unsigned | NO   | PRI | NULL    | auto_increment |
''| VISIBLE LOGICIEL | text             | YES  |     | NULL    |                |
''| VALEUR REELLE    | text             | YES  |     | NULL    |                |
''+------------------+------------------+------+-----+---------+----------------+
''3 rows in set (0.00 sec)
''
''
''mysql>
''
''
''
''Création d'une table SERIE de platine
''
''CREATE TABLE `IDEALTAKE`.`MENU_SERIE` (
''  `INDICE_AUTO` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
''  `PLATINE_SERIE` TEXT,
''  PRIMARY KEY(`INDICE_AUTO`)
'')
''ENGINE = InnoDB;
''
''
''
''ALTER TABLE `IDEALTAKE`.`menu_serie` CHANGE COLUMN `PLATINE_SERIE` `VISIBLE_CLIENT` TEXT CHARACTER SET latin1 COLLATE latin1_swedish_ci DEFAULT NULL,
'' ADD COLUMN `VALEUR_REELLE` TEXT AFTER `VISIBLE_CLIENT`;
''
''
''
''
''mysql> describe menu_serie;
''+------------------+------------------+------+-----+---------+----------------+
''| Field            | Type             | Null | Key | Default | Extra          |
''+------------------+------------------+------+-----+---------+----------------+
''| INDICE_AUTO      | int(10) unsigned | NO   | PRI | NULL    | auto_increment |
''| VISIBLE_LOGICIEL | text             | YES  |     | NULL    |                |
''| VALEUR_REELLE    | text             | YES  |     | NULL    |                |
''+------------------+------------------+------+-----+---------+----------------+
''3 rows in set (0.00 sec)
''
''
''mysql>
''
''
'' Création Table Verrouillage
''
''CREATE TABLE `IDEALTAKE`.`MENU_VERROUILLAGE` (
''  `INDICE_AUTO` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
''  `VALEUR_LOGICIEL` TEXT,
''  `VALEUR_LISIBLE` TEXT,
''  PRIMARY KEY(`INDICE_AUTO`)
'')
''ENGINE = InnoDB;
''
''mysql> describe menu_VERROUILLAGE;
''+-----------------+------------------+------+-----+---------+----------------+
''| Field           | Type             | Null | Key | Default | Extra          |
''+-----------------+------------------+------+-----+---------+----------------+
''| INDICE_AUTO     | int(10) unsigned | NO   | PRI | NULL    | auto_increment |
''| VALEUR_LOGICIEL | text             | YES  |     | NULL    |                |
''| VALEUR_LISIBLE  | text             | YES  |     | NULL    |                |
''+-----------------+------------------+------+-----+---------+----------------+
''3 rows in set (0.00 sec)
''
''mysql>
''
''



Public Class Parametres

    Dim da As MySqlDataAdapter
    Dim ds As DataSet
    Dim dv As DataView
    Dim cb As MySqlCommandBuilder
    Public Tableau_à_lire As String = ""
    Public index_du_tableau_à_lire As Integer = 0

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (Form_Demarrage.FonctionModificationObjet(Button1, sender.ToString) = True) Then

        End If

        Dim flux As FileStream
        Dim Str1 As String
        Dim ecriture As StreamWriter

        Dim chemin5 As String



        chemin5 = My.Application.Info.DirectoryPath & "\niveau_access_" & Form_Demarrage.Nom_Connexion.Text.ToString & ".txt"


        Try


            '           MsgBox(Form1.Parametrage_REQUETE_Str & "; " & Me.NIVEAU_ACCES.SelectedItem & " : " & Len(Me.NIVEAU_ACCES.SelectedItem))

            'Dim String_Adresse_Serveur_temp As String = Form_Demarrage.connexion
            Dim String_Adresse_Serveur_temp As String = Me.Chaine_connexion.Text

            String_Adresse_Serveur_temp = Mid(String_Adresse_Serveur_temp, (InStr(String_Adresse_Serveur_temp, "server=") + 7), (String_Adresse_Serveur_temp.Length - (InStr(String_Adresse_Serveur_temp, "server=") + 7)))
            String_Adresse_Serveur_temp = Mid(String_Adresse_Serveur_temp, 1, (InStr(String_Adresse_Serveur_temp, ";user id=") - 1))
            Form_Demarrage.Adresse_Serveur.Text = String_Adresse_Serveur_temp 'dans l'onglet aide on réactualise se paramètre
            'MsgBox("Serveur : " & String_Adresse_Serveur)


            Dim String_Utilisateur_temp As String = Me.Chaine_connexion.Text
            String_Utilisateur_temp = Mid(String_Utilisateur_temp, (InStr(String_Utilisateur_temp, "user id=") + 8), (String_Utilisateur_temp.Length - (InStr(String_Utilisateur_temp, "user id=") + 8)))
            String_Utilisateur_temp = Mid(String_Utilisateur_temp, 1, (InStr(String_Utilisateur_temp, ";pwd=") - 1))
            'MsgBox("Utilisateur : " & String_Utilisateur)
            Form_Demarrage.Utilisateur.Text = String_Utilisateur_temp
            Form_Demarrage.Utilisateur_MySql = String_Utilisateur_temp

            Dim String_mot_de_passe_temp As String = Me.Chaine_connexion.Text
            String_mot_de_passe_temp = Mid(String_mot_de_passe_temp, (InStr(String_mot_de_passe_temp, "pwd=") + 4), (String_mot_de_passe_temp.Length + 1 - (InStr(String_mot_de_passe_temp, "pwd=") + 4)))
            Form_Demarrage.mot_de_passe_Utilisateur_MySql = String_mot_de_passe_temp
            'String_mot_de_passe = Mid(String_mot_de_passe, 1, (InStr(String_mot_de_passe, ";pwd=") - 1))
            'MsgBox("Mot de passe : " & String_mot_de_passe)

            Dim String_Nom_de_la_base_temp As String = Me.Chaine_connexion.Text
            String_Nom_de_la_base_temp = Mid(String_Nom_de_la_base_temp, (InStr(String_Nom_de_la_base_temp, "database=") + 9), (String_Nom_de_la_base_temp.Length - (InStr(String_Nom_de_la_base_temp, "database=") + 9)))
            'InStr(String_Nom_de_la_base, "database=") + 10))
            ' - (InStr(String_Nom_de_la_base, ";server="))
            String_Nom_de_la_base_temp = Mid(String_Nom_de_la_base_temp, 1, (InStr(String_Nom_de_la_base_temp, ";server=") - 1))
            Form_Demarrage.Nom_de_la_base.Text = String_Nom_de_la_base_temp  'dans l'onglet aide on réactualise se paramètre


            'Form1.Chargement_Collections_Menus(Form1.connexion)  'ReChargement des collections des menus 
            'MsgBox(Form1.connexion)



            If (Len(Me.NIVEAU_ACCES.SelectedItem) > 0) Then
                Str1 = Me.NIVEAU_ACCES.SelectedItem

                If (Str1 = "Nicht aktiviert" _
                    Or Str1 = "Not activated" _
                    Or Str1 = "No acelerado" _
                    Or Str1 = "Pas Activé" _
                    Or Str1 = "Non Attivato") Then
                    Str1 = "0"
                    GoTo Fin
                End If

                'en cas de connexion sans succes il ne faut pas tenter de positionner
                'le niveau d'acces dans parametres

                '    Public Table_niveau_acces(128, 20)           'Niveau d'acces au document : Bas Niveau, Normal, confidentiel
                'Table_niveau_acces(Position1, 1) = myReader.GetString(0).ToString     'INDICE_AUTO
                'Table_niveau_acces(Position1, 2) = myReader.GetString(1).ToString     'VALEUR_LOGICIEL_francais
                'Table_niveau_acces(Position1, 3) = myReader.GetString(2).ToString     'VALEUR_REELLE Utilisé avant dans M.R.U Administrateur
                'Table_niveau_acces(Position1, 4) = myReader.GetString(3).ToString     'IDConstructeur
                'Table_niveau_acces(Position1, 5) = myReader.GetString(4).ToString     'VALEUR_LOGICIEL_anglais
                'Table_niveau_acces(Position1, 6) = myReader.GetString(5).ToString     'VALEUR_LOGICIEL_allemand
                'Table_niveau_acces(Position1, 7) = myReader.GetString(6).ToString     'VALEUR_LOGICIEL_espagnole
                'Table_niveau_acces(Position1, 8) = myReader.GetString(7).ToString     'VALEUR_LOGICIEL_italien

                Dim T1 As Integer = 0

                For T1 = 0 To 128


                    'Valeur de Langue
                    ' 0 = Allemand
                    ' 1 = Anglais
                    ' 2 = Espagnole
                    ' 3 = Français
                    ' 4 = Italien


                    ' On recherche l'INDICE_AUTO pour le sauvegarder dans le fichier
                    ' niveau_access_" & Form_Demarrage.Nom_Connexion.Text.ToString & ".txt"
                    Select Case Form_Demarrage.LangueLogiciel

                        Case 0
                            If (Form_Demarrage.Table_niveau_acces(T1, 6) = Str1) Then
                                Str1 = Form_Demarrage.Table_niveau_acces(T1, 1).ToString
                                Exit For
                            End If

                        Case 1
                            If (Form_Demarrage.Table_niveau_acces(T1, 5) = Str1) Then
                                Str1 = Form_Demarrage.Table_niveau_acces(T1, 1).ToString
                                Exit For
                            End If

                        Case 2
                            If (Form_Demarrage.Table_niveau_acces(T1, 7) = Str1) Then
                                Str1 = Form_Demarrage.Table_niveau_acces(T1, 1).ToString
                                Exit For
                            End If

                        Case 3
                            If (Form_Demarrage.Table_niveau_acces(T1, 2) = Str1) Then
                                Str1 = Form_Demarrage.Table_niveau_acces(T1, 1).ToString
                                Exit For
                            End If

                        Case 4
                            If (Form_Demarrage.Table_niveau_acces(T1, 8) = Str1) Then
                                Str1 = Form_Demarrage.Table_niveau_acces(T1, 1).ToString
                                Exit For
                            End If

                        Case Else

                    End Select



                Next T1


Fin:
                Form_Demarrage.Parametrage_REQUETE_Str = Str1
                flux = File.Create(chemin5)

                ecriture = New StreamWriter(flux)
                ecriture.WriteLine(Str1)
                ecriture.Close()
            End If



            'ici on est bien connecté au serveur on peut
            'sauver la mise à jour des données dans le fichier 
            'liste_des_connexions.xml

            'Public Liste_Connexions(100, 15) As String
            'Titre de la connexion, Adresse Serveur, Utilisateur, Mot de passe,
            'Confirmer le mot de passe, Nom de la base, chemin Schéma,c hemin synoptique
            'chemin notice, chemin fiche produit,chemin photos produit

            'Analyse de la chaîne de connexion
            'Persist Security Info=true;database=IDEALTAKE;server=localhost;user id=root;pwd=FML6A42V1B
            Dim String_Nom_Connexion As String = Form_Demarrage.Nom_Connexion.Text
            'MsgBox("nom de la connexion : " & String_Nom_Connexion)

            Dim String_Adresse_Serveur As String = Me.Chaine_connexion.Text
            String_Adresse_Serveur = Mid(String_Adresse_Serveur, (InStr(String_Adresse_Serveur, "server=") + 7), (String_Adresse_Serveur.Length - (InStr(String_Adresse_Serveur, "server=") + 7)))
            String_Adresse_Serveur = Mid(String_Adresse_Serveur, 1, (InStr(String_Adresse_Serveur, ";user id=") - 1))
            'MsgBox("Serveur : " & String_Adresse_Serveur)

            Dim String_Utilisateur As String = Me.Chaine_connexion.Text
            String_Utilisateur = Mid(String_Utilisateur, (InStr(String_Utilisateur, "user id=") + 8), (String_Utilisateur.Length - (InStr(String_Utilisateur, "user id=") + 8)))
            String_Utilisateur = Mid(String_Utilisateur, 1, (InStr(String_Utilisateur, ";pwd=") - 1))
            'MsgBox("Utilisateur : " & String_Utilisateur)

            Dim String_mot_de_passe As String = Me.Chaine_connexion.Text
            String_mot_de_passe = Mid(String_mot_de_passe, (InStr(String_mot_de_passe, "pwd=") + 4), (String_mot_de_passe.Length + 1 - (InStr(String_mot_de_passe, "pwd=") + 4)))
            'String_mot_de_passe = Mid(String_mot_de_passe, 1, (InStr(String_mot_de_passe, ";pwd=") - 1))
            'MsgBox("Mot de passe : " & String_mot_de_passe)

            Dim String_Nom_de_la_base As String = Me.Chaine_connexion.Text
            String_Nom_de_la_base = Mid(String_Nom_de_la_base, (InStr(String_Nom_de_la_base, "database=") + 9), (String_Nom_de_la_base.Length - (InStr(String_Nom_de_la_base, "database=") + 9)))
            'InStr(String_Nom_de_la_base, "database=") + 10))
            ' - (InStr(String_Nom_de_la_base, ";server="))
            String_Nom_de_la_base = Mid(String_Nom_de_la_base, 1, (InStr(String_Nom_de_la_base, ";server=") - 1))

            'MsgBox("base : " & String_Nom_de_la_base)

            'Dim String_Localisation_Schéma As String = Form_Demarrage.Chemin_SCHEMA_Str
            'Dim String_Localisation_Synoptique As String = Form_Demarrage.Chemin_SYNOPTIQUE_Str
            'Dim String_Localisation_Notice As String = Form_Demarrage.Chemin_DOCUMENTATION_Str
            'Dim String_Localisation_Fiche_Produit As String = Form_Demarrage.Chemin_Fiche_Produit_Str
            'Dim String_Localisation_Photo_produit As String = Form_Demarrage.Chemin_Photo_Produit_Str

            Dim t As Integer
            Dim string_selection As String = ""
            Dim nmbre_de_connexion As Integer = 0
            Dim position_connexion As Integer = 0

            string_selection = String_Nom_Connexion

            For t = 1 To 100
                If (Len(Form_Demarrage.Liste_Connexions(t, 1)) > 0) Then
                    nmbre_de_connexion += 1
                End If

                If (Form_Demarrage.Liste_Connexions(t, 1) = string_selection) Then
                    position_connexion = t
                    'MsgBox(Selection_connexion_ComboBox2.SelectedItem.ToString & " : " & t)
                End If
            Next


            'ici on réactualise dans l'onglet aide
            'les paramètres du serveur
            Form_Demarrage.Nom_Connexion.Text = String_Nom_Connexion
            Form_Demarrage.Adresse_Serveur.Text = String_Adresse_Serveur
            Form_Demarrage.Nom_de_la_base.Text = String_Nom_de_la_base
            Form_Demarrage.Utilisateur.Text = String_Utilisateur

            Dim Str1_temp As String = ""

            Str1_temp = "Persist Security Info=true;database=" & String_Nom_de_la_base & ";server=" & String_Adresse_Serveur & ";user id=" & String_Utilisateur & ";pwd=" & String_mot_de_passe
            Form_Demarrage.connexion = Str1_temp

            Form_Demarrage.Liste_Connexions(position_connexion, 2) = String_Adresse_Serveur
            Form_Demarrage.Liste_Connexions(position_connexion, 3) = String_Utilisateur
            Form_Demarrage.Liste_Connexions(position_connexion, 4) = String_mot_de_passe
            Form_Demarrage.Liste_Connexions(position_connexion, 5) = String_Nom_de_la_base
            'Form_Demarrage.Liste_Connexions(position_connexion, 6) = String_Localisation_Schéma
            'Form_Demarrage.Liste_Connexions(position_connexion, 7) = String_Localisation_Synoptique
            'Form_Demarrage.Liste_Connexions(position_connexion, 8) = String_Localisation_Notice
            'Form_Demarrage.Liste_Connexions(position_connexion, 9) = String_Localisation_Fiche_Produit
            'Form_Demarrage.Liste_Connexions(position_connexion, 10) = String_Localisation_Photo_produit



            Dim settings As New XmlWriterSettings()
            Dim dossier As String = My.Application.Info.DirectoryPath & "\"
            'ici le fichier liste_des_connexions.xml va être réécrit

            settings.Encoding = System.Text.Encoding.UTF8
            settings.Indent = True
            settings.IndentChars = "    "
            Using writer As XmlWriter = XmlWriter.Create(dossier & "liste_des_connexions.xml", settings)
                ' Write XML data.
                writer.WriteStartElement("Liste_des_connexions")
                writer.WriteStartElement("Connexion")
                writer.WriteEndElement()

                For t = 1 To nmbre_de_connexion

                    writer.WriteStartElement("Connexion")
                    writer.WriteElementString("Nom_Connexion", Form_Demarrage.Liste_Connexions(t, 1))
                    writer.WriteElementString("Adresse_Serveur", Form_Demarrage.Liste_Connexions(t, 2))
                    writer.WriteElementString("Utilisateur", Form_Demarrage.Liste_Connexions(t, 3))
                    writer.WriteElementString("mot_de_passe", Form_Demarrage.Liste_Connexions(t, 4))
                    writer.WriteElementString("Nom_de_la_base", Form_Demarrage.Liste_Connexions(t, 5))
                    writer.WriteElementString("Description_Allemande", Form_Demarrage.Liste_Connexions(t, 6))
                    writer.WriteElementString("Description_Anglaise", Form_Demarrage.Liste_Connexions(t, 7))
                    writer.WriteElementString("Description_Espagnole", Form_Demarrage.Liste_Connexions(t, 8))
                    writer.WriteElementString("Description_Française", Form_Demarrage.Liste_Connexions(t, 9))
                    writer.WriteElementString("Description_Italienne", Form_Demarrage.Liste_Connexions(t, 10))
                    writer.WriteEndElement()
                Next t

                writer.WriteEndElement()
                writer.Flush()

            End Using







            If (Me.Sélection_Connexion_Vérouillé.Checked = True) Then
                'la connexion sélectionné sera toujours utilisée
                'paramètres : 
                'Utiliser_toujours_cette_connexion       string "OUI" ou ""
                'Nom_de_la_connexion                     string  "Nom de la connexion" ou ""

                My.Settings.Nom_de_la_connexion = Form_Demarrage.Nom_Connexion.Text.ToString
                My.Settings.Utiliser_toujours_cette_connexion = "OUI" 'Utiliser toujours cette connexion
            Else
                My.Settings.Nom_de_la_connexion = ""
                My.Settings.Utiliser_toujours_cette_connexion = "" 'On n'utilise pas de connexion particulière
            End If

            'MsgBox("Enregistrement Terminé", MsgBoxStyle.Information, "Paramètrage Serveur")
            MsgBox(Ajout_Suppression_connexion_base.MessageUtilisateur(190, Form_Demarrage.LangueLogiciel), MsgBoxStyle.Information, Ajout_Suppression_connexion_base.MessageUtilisateur(189, Form_Demarrage.LangueLogiciel))

            Dim enabled_combobox6 As Boolean = True
            If (Form_Demarrage.ComboBox_TABLE_B.Enabled = False) Then
                enabled_combobox6 = False
            End If

            Attente_Chargement.Visible = True
            Form_Demarrage.Chargement_Collections_Menus(sender, Form_Demarrage.connexion, "Sauter_les_combobox")  'ReChargement des collections des menus 
            'Form_Demarrage.chargements_constructeurs(Form_Demarrage.connexion)  'Rechargement de la table constructeurs
            Attente_Chargement.Visible = False

            If (enabled_combobox6 = False) Then
                Form_Demarrage.ComboBox_TABLE_B.Enabled = False
            Else
                'rien le chargement de la collection aura
                'réactivé le menu combobox6

            End If


        Catch ex As Exception

            ' erreur dans le formulaire Parametre.vb - Onglet Serveur - Enregistrer
            ' NIVEAU_ACCES System.windows.forms.combobox
            ' InvalidArgument= La valeur '0' n'est pas valide pour 'SelectedIndex'
            ' Nom du paramètre : SelectedIndex

            If (InStr(ex.Message, "InvalidArgument= La valeur '0'") > 0) Then
                'MessageBox.Show("Erreur sur le SelectedIndex valeur nulle")
            Else
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End Try
    End Sub

    Function AfficheTable(ByVal nom_de_la_table As String, ByVal nom_du_datagridview As DataGridView)

        Try

            Dim conn_SQL As MySqlConnection = New MySqlConnection()
            conn_SQL.ConnectionString = Form_Demarrage.connexion
            conn_SQL.Open()
            If (nom_de_la_table = "enregistrements") Then

                da = New MySqlDataAdapter("Select Numero_Enregistrement," & _
                                          "Chrono," & _
                                          "Auteur," & _
                                          "DATE_FORMAT(" & Form_Demarrage.Nom_de_la_base.Text & ".enregistrements.Date_de_creation ,' %d/%c/%Y')," & _
                                          "Titre_texte_francais," & _
                                          "Sous_Titre_texte_francais," & _
                                          "Commentaire_Texte_francais," & _
                                          "Titre_texte_Allemand," & _
                                          "Sous_Titre_texte_Allemand," & _
                                          "Commentaire_Texte_Allemand," & _
                                          "Titre_texte_Anglais," & _
                                          "Sous_Titre_texte_Anglais," & _
                                          "Commentaire_Texte_Anglais," & _
                                          "Titre_texte_Espagnole," & _
                                          "Sous_Titre_texte_Espagnole," & _
                                          "Commentaire_Texte_Espagnole," & _
                                          "Titre_texte_Italien," & _
                                          "Sous_Titre_texte_Italien," & _
                                          "Commentaire_Texte_Italien," & _
                                          "Activation_XAML," & _
                                          "Type_XAML_Titre_predefini," & _
                                          "Type_XAML_Sous_Titre_predefini," & _
                                          "Type_XAML_Commentaire_predefini," & _
                                          "Type_XAML_Titre_specifique," & _
                                          "Type_XAML_Sous_Titre_specifique," & _
                                          "Type_XAML_Commentaire_specifique," & _
                                          "PARAM_A," & _
                                          "PARAM_B," & _
                                          "PARAM_C," & _
                                          "PARAM_D," & _
                                          "PARAM_E," & _
                                          "PARAM_F," & _
                                          "PARAM_G," & _
                                          "PARAM_H," & _
                                          "PARAM_I," & _
                                          "Recherche_par_Arbre," & _
                                          "Label_reserve1_francais," & _
                                          "Texte_reserve1_francais," & _
                                          "Label_reserve1_Allemand," & _
                                          "Texte_reserve1_Allemand," & _
                                          "Label_reserve1_Anglais," & _
                                          "Texte_reserve1_Anglais," & _
                                          "Label_reserve1_Espagnole," & _
                                          "Texte_reserve1_Espagnole," & _
                                          "Label_reserve1_Italien," & _
                                          "Texte_reserve1_Italien," & _
                                          "Boolean_reserve1," & _
                                          "Integer_reserve1," & _
                                          "DATE_FORMAT(" & Form_Demarrage.Nom_de_la_base.Text & ".enregistrements.Date_reserve1 ,' %d/%c/%Y')," & _
                                          "IDConstructeur," & _
                                          "NIVEAU_ACCES," & _
                                          "STANDARD," & _
                                          "DATE_FORMAT(" & Form_Demarrage.Nom_de_la_base.Text & ".enregistrements.Derniere_modification ,' %d/%c/%Y')," & _
                                          "Auteur_derniere_modification," & _
                                          "NUMERO_MODIFICATION" & _
                                          " FROM `" & Form_Demarrage.Nom_de_la_base.Text & "`.`" & nom_de_la_table & "`", conn_SQL)
                GoTo Fin_de_da
            Else

            End If


            If (nom_de_la_table = "VERSIONS_TABLES") Then

                da = New MySqlDataAdapter("SELECT" & _
                                          "`ID`," & _
                                          "DATE_FORMAT(" & Form_Demarrage.Nom_de_la_base.Text & ".VERSIONS_TABLES.Date ,' %d/%c/%Y')," & _
                                          "`categories`," & _
                                          "`chrono_idref`," & _
                                          "`CONSTRUCTEURS`," & _
                                          "`doc_url`," & _
                                          "`documents`," & _
                                          "`elements_HTML`," & _
                                          "`elements_winform`," & _
                                          "`enregistrements`," & _
                                          "`HISTORY_EVENTS`," & _
                                          "`messages_utilisateur`," & _
                                          "`niveau_acces`," & _
                                          "`Questions`," & _
                                          "`racine_zero`," & _
                                          "`racine1`," & _
                                          "`racine2`," & _
                                          "`racine3`," & _
                                          "`racine4`," & _
                                          "`racine5`," & _
                                          "`racine6`," & _
                                          "`racine7`," & _
                                          "`racine8`," & _
                                          "`racine9`," & _
                                          "`ref_constructeurs`," & _
                                          "`schema`," & _
                                          "`TABLE_A`," & _
                                          "`TABLE_B`," & _
                                          "`TABLE_C`," & _
                                          "`TABLE_D`," & _
                                          "`TABLE_E`," & _
                                          "`TABLE_XAML`," & _
                                          "`types_documents`," & _
                                          "`Users`," & _
                                          "`Commentaire`" & _
                                          " FROM  `" & Form_Demarrage.Nom_de_la_base.Text & "`.`" & nom_de_la_table & "`", conn_SQL)
                GoTo Fin_de_da
            End If


            'pour les autres tables on va créer ce DataAdapter
            da = New MySqlDataAdapter("Select * FROM `" & Form_Demarrage.Nom_de_la_base.Text & "`.`" & nom_de_la_table & "`", conn_SQL)
Fin_de_da:

            'da = New MySqlDataAdapter("Select * FROM `" & Form_Demarrage.Nom_de_la_base.Text & "`.`" & nom_de_la_table & "`", conn_SQL)
            ds = New DataSet()
            da.Fill(ds, "mytable")
            dv = ds.Tables("mytable").DefaultView
            conn_SQL.Close()

            'on réinitialise le datagridview
            nom_du_datagridview.Columns.Clear()
            nom_du_datagridview.DataBindings.Clear()
            nom_du_datagridview.DataSource = Nothing
            nom_du_datagridview.DataMember = Nothing

            nom_du_datagridview.DataSource = dv
            nom_du_datagridview.Show()


            dv.AllowEdit = True

            If (nom_de_la_table = "enregistrements") Then

                'On adapte le Datagridview pour la table 'enregistrements'
                'on renomme les colonnes Ou il a des dates
                Me.DataGridView1.Columns.Item(3).HeaderText = "Création"
                'Me.DataGridView1.Columns.Item(47).HeaderText = "Date de Reserve1" '!! 47 pointe sur Integer_reserve1 (26/10/2010)
                Me.DataGridView1.Columns.Item(48).HeaderText = "Date de Reserve1"
                'Me.DataGridView1.Columns.Item(51).HeaderText = "Dernière Modification"
                Me.DataGridView1.Columns.Item(52).HeaderText = "Dernière Modification"

                'désactivation du tri sur les colonnes de date
                Me.DataGridView1.Columns.Item(3).SortMode = DataGridViewColumnSortMode.NotSortable  'Date de création
                'Me.DataGridView1.Columns.Item(47).SortMode = DataGridViewColumnSortMode.NotSortable  'Date de Reserve1
                Me.DataGridView1.Columns.Item(48).SortMode = DataGridViewColumnSortMode.NotSortable  'Date de Reserve1
                'Me.DataGridView1.Columns.Item(51).SortMode = DataGridViewColumnSortMode.NotSortable  'Date de Dernière Modification
                Me.DataGridView1.Columns.Item(52).SortMode = DataGridViewColumnSortMode.NotSortable  'Date de Dernière Modification

            End If


            If (nom_de_la_table = "VERSIONS_TABLES") Then
                'On adapte le Datagridview pour la table 'VERSIONS_TABLES'
                'on renomme la colonne Ou il a une date
                Me.DataGridView1.Columns.Item(1).HeaderText = "Date"
            End If

            'affiche le nom de la table en cours de lecture
            Collections_menus_TABLE_EN_LECTURE_Label_1.Text = _
            """" & nom_de_la_table & """"

            'Liste_des_Tables(1, 1)  'Name
            'Liste_des_Tables(1, 2)  'Engine
            'Liste_des_Tables(1, 3)  'Version
            'Liste_des_Tables(1, 4)  'Row format
            'Liste_des_Tables(1, 5)  'Rows
            'Liste_des_Tables(1, 6)  'Create Time
            'Liste_des_Tables(1, 7)  'Collation

            For x1 As Integer = 1 To 1024
                If (Form_Demarrage.Liste_des_Tables(x1, 1) = nom_de_la_table) Then
                    Collections_menus_TABLE_EN_LECTURE_Label_1.Text = Collections_menus_TABLE_EN_LECTURE_Label_1.Text & _
                    vbCrLf & "Engine : " & Form_Demarrage.Liste_des_Tables(x1, 2) & _
                    vbCrLf & "Version : " & Form_Demarrage.Liste_des_Tables(x1, 3) & _
                    vbCrLf & "Row format : " & Form_Demarrage.Liste_des_Tables(x1, 4) & _
                    vbCrLf & "Rows : " & Form_Demarrage.Liste_des_Tables(x1, 5) & _
                    vbCrLf & "Create Time : " & Form_Demarrage.Liste_des_Tables(x1, 6) & _
                    vbCrLf & "Collation : " & Form_Demarrage.Liste_des_Tables(x1, 7)

                    Exit For
                End If
            Next




        Catch ex As Exception
            'affiche --- pour la table en cours
            Collections_menus_TABLE_EN_LECTURE_Label_1.Text = "- - -"

            MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return True

    End Function


    Private Sub mise_a_jour_menu_systeme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mise_a_jour_menu_systeme.Click

        If (Form_Demarrage.FonctionModificationObjet(mise_a_jour_menu_systeme, sender.ToString) = True) Then

        End If

        'Si le datagridView n'est pas visible aucune raison de
        'faire une mise à jour on sort du sub
        If (Me.DataGridView1.RowCount < 1) Then
            Exit Sub
        End If


        Dim conn_SQL As MySqlConnection = New MySqlConnection()

        Try

            conn_SQL.ConnectionString = Form_Demarrage.connexion
            cb = New MySqlCommandBuilder(da)
            da.Update(ds, "mytable")
            conn_SQL.Close()

            'MsgBox("Table """ & Tableau_à_lire & """" & vbCrLf & " Mise à jour effectuée", MsgBoxStyle.Information, "Mise à jour")
            MsgBox(Ajout_Suppression_connexion_base.MessageUtilisateur(17, Form_Demarrage.LangueLogiciel).ToString() & _
                    " """ & Tableau_à_lire & """" & vbCrLf & " " & _
                   Ajout_Suppression_connexion_base.MessageUtilisateur(179, Form_Demarrage.LangueLogiciel).ToString() & " " _
                   , MsgBoxStyle.Information, _
                   Ajout_Suppression_connexion_base.MessageUtilisateur(178, Form_Demarrage.LangueLogiciel).ToString())

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

            'MsgBox("Cause possible de l'erreur vous n'avez pas éffectué de lecture de la base dans le DataGridView de votre logiciel la mise à jour de la table de la base de donnée est alors impossible, ou la connexion au serveur est indisponible", MsgBoxStyle.Information, "Mise à jour")
            MsgBox(Ajout_Suppression_connexion_base.MessageUtilisateur(5, Form_Demarrage.LangueLogiciel).ToString(), MsgBoxStyle.Information, _
                   Ajout_Suppression_connexion_base.MessageUtilisateur(178, Form_Demarrage.LangueLogiciel).ToString())

            DataGridView1.Hide()

        End Try

        'réactualise le DataSet
        AfficheTable(Tableau_à_lire, DataGridView1)

        'on recharge tout les menus et les tableaux
        Form_Demarrage.Chargement_Collections_Menus(sender, Form_Demarrage.connexion)  'ReChargement des collections des menus 
        Form_Demarrage.chargements_constructeurs(Form_Demarrage.connexion)  'Rechargement de la table constructeurs


        'Liste_des_Tables(1, 1)  'Name
        'Liste_des_Tables(1, 2)  'Engine
        'Liste_des_Tables(1, 3)  'Version
        'Liste_des_Tables(1, 4)  'Row format
        'Liste_des_Tables(1, 5)  'Rows
        'Liste_des_Tables(1, 6)  'Create Time
        'Liste_des_Tables(1, 7)  'Collation

        For x1 As Integer = 1 To 1024
            If (Form_Demarrage.Liste_des_Tables(x1, 1) = Tableau_à_lire) Then
                Collections_menus_TABLE_EN_LECTURE_Label_1.Text = """" & Form_Demarrage.Liste_des_Tables(x1, 1) & """" & _
                vbCrLf & "Engine : " & Form_Demarrage.Liste_des_Tables(x1, 2) & _
                vbCrLf & "Version : " & Form_Demarrage.Liste_des_Tables(x1, 3) & _
                vbCrLf & "Row format : " & Form_Demarrage.Liste_des_Tables(x1, 4) & _
                vbCrLf & "Rows : " & Form_Demarrage.Liste_des_Tables(x1, 5) & _
                vbCrLf & "Create Time : " & Form_Demarrage.Liste_des_Tables(x1, 6) & _
                vbCrLf & "Collation : " & Form_Demarrage.Liste_des_Tables(x1, 7)

                Exit For
            End If
        Next



        Dim enabled_combobox6 As Boolean = True

        If (Form_Demarrage.ComboBox_TABLE_B.Enabled = False) Then
            enabled_combobox6 = False
        End If

        If (enabled_combobox6 = False) Then
            If (InStr(Form_Demarrage.ComboBox_TABLE_B.SelectedItem, "Tou") < 1) Then

                Form_Demarrage.ComboBox_TABLE_B.Enabled = False
                'MsgBox("Tou ?? : " & InStr(Form1.ComboBox6.SelectedItem, "Tou"))

            Else

                Form_Demarrage.ComboBox_TABLE_B.Enabled = True
                ' MsgBox("Tou ?? : " & InStr(Form1.ComboBox6.SelectedItem, "Tou"))

            End If

        Else
            'rien le chargement de la collection aura
            'réactivé le menu combobox6

        End If

    End Sub




    Private Sub Collections_Menus_Lecture_Table_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Collections_Menus_Lecture_Table.Click

        If (Form_Demarrage.FonctionModificationObjet(Collections_Menus_Lecture_Table, sender.ToString) = True) Then

        End If


        If (Collections_Menus_LISTE_DES_TABLES_ListBox1.SelectedIndex <> -1) Then

            'MsgBox(Collections_Menus_LISTE_DES_TABLES_ListBox1.SelectedIndex.ToString & " - Oui, j'ai sélectionné une donnée")

            '' Récupération du nom de la table sélectionnée dans le listbox
            Dim nom_de_la_table As String = Nothing
            nom_de_la_table = Collections_Menus_LISTE_DES_TABLES_ListBox1.SelectedItem.ToString
            index_du_tableau_à_lire = Collections_Menus_LISTE_DES_TABLES_ListBox1.SelectedIndex

            DataGridView1.Show() 'Affiche le DATAGRIDVIEW1

            Try
                AfficheTable(nom_de_la_table, DataGridView1)   'Fonction d'affichage de la base
                Tableau_à_lire = nom_de_la_table

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DataGridView1.Hide()
            End Try



        Else
            'aucun éléments de la collection de l'objet du listbox n'a été sélectionné
            'MsgBox(Collections_Menus_LISTE_DES_TABLES_ListBox1.SelectedIndex.ToString & " - Non, je n'ai pas sélectionné une donnée")

        End If


    End Sub

    Private Sub Parametres_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'cache le formulaire parametres plutot que de le fermer 
        'plutot que de le fermer car sinon on perd les infos chargé
        'au lancement du logiciel
        Me.Hide()
        Form_Demarrage.Show() 'On montre 'Form_Demarrage'

    End Sub



    Private Sub Parametres_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            Me.Visible = False
            Attente_Chargement.Visible = True ' on affiche l'attente de chargement

            'on recharge tout les menus et les tableaux
            Form_Demarrage.Chargement_Collections_Menus(sender, Form_Demarrage.connexion, "Sauter_les_combobox")  'ReChargement des collections des menus 
            Form_Demarrage.chargements_constructeurs(Form_Demarrage.connexion)     'Rechargement de la table constructeurs
            Me.Chaine_connexion.Text = Form_Demarrage.connexion                    'Rechargement de la chaîne de connexion

            Attente_Chargement.Visible = False ' On n'affiche plus 
            Me.Visible = True

        Catch ex As Exception

            ' MsgBox("""" & sender.ToString & """ : " & vbCrLf & ex.Message, MsgBoxStyle.Critical) ', "Error")
            MsgBox(ex.Message, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub Activation_Modification_Objet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Activation_Modification_Objet.CheckedChanged
        If (Activation_Modification_Objet.Checked = True) Then
            Form_Demarrage.Modification_Objet = True
        Else
            Form_Demarrage.Modification_Objet = False
        End If
    End Sub

    Private Sub Titre_Connexion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Titre_Connexion.Click
        If (Form_Demarrage.FonctionModificationObjet(Titre_Connexion, sender.ToString) = True) Then
            Exit Sub
        End If
    End Sub

    Private Sub Sélection_Connexion_Vérouillé_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sélection_Connexion_Vérouillé.CheckedChanged
        If (Form_Demarrage.FonctionModificationObjet(Sélection_Connexion_Vérouillé, sender.ToString) = True) Then

        End If
    End Sub

    Private Sub Label47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label47.Click
        If (Form_Demarrage.FonctionModificationObjet(Label47, sender.ToString) = True) Then

        End If
    End Sub

    Private Sub Label57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label57.Click
        If (Form_Demarrage.FonctionModificationObjet(Label57, sender.ToString) = True) Then

        End If
    End Sub

    Private Sub Collections_Menus_LISTE_DES_TABLES_Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Collections_Menus_LISTE_DES_TABLES_Label2.Click
        If (Form_Demarrage.FonctionModificationObjet(Collections_Menus_LISTE_DES_TABLES_Label2, sender.ToString) = True) Then

        End If
    End Sub

    Private Sub Activation_debugage_Script_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Activation_debugage_Script.CheckedChanged
        If (Activation_debugage_Script.Checked = True) Then
            My.Settings.Debugage_SCRIPT = True

        Else
            My.Settings.Debugage_SCRIPT = False
            Form_Demarrage.WebBrowser1.Visible = False
            Form_Demarrage.WebBrowser2.Visible = False
        End If
    End Sub
End Class