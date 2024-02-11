'
'
'SAPI5 Sous programme de IDEALTAKE
'----------------------------------
'
'
'DESCRIPTION :
'
'Synthèse vocale pour le logiciel IDEALTAKE
'Ce logiciel a été developpé en  VISUAL BASIC 2010.
'
'
'Copyright (C) 2011 VIEIL Thierry Ce programme est un logiciel libre ;
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


Imports SpeechLib
Imports System.IO

Public Class Formulaire_SAPI

    Public Texte_à_dire_01 As String = ""

    Dim VoiceObj_SAPI5 As New SpeechLib.SpVoice             'Objet SpVoice Utilise SAPI 5.4
    Dim VoicesToken_SAPI5 As New SpeechLib.SpObjectToken    'Objet SpObjectToken Utilise SAPI 5.4

    Dim integer_index_Voice_SAPI5 As Integer = 0
    Dim integer_index_AudioOutput_SAPI5 As Integer = 0

    Public langue_logiciel As Integer = 3
    Dim LOAD_Terminée As Boolean = False

    'Liste contenant les instructions
    'Du fichier d'échange
    Public liste_des_instructions As New List(Of String)
    Public liste_des_instructions_TEMPO As New List(Of String)


    Function dire_un_texte(ByVal Texte_à_dire As String)

        'Me.TopLevel = False

        On Error GoTo Erreur_dire_un_texte
        VoiceObj_SAPI5.Speak(Texte_à_dire, SpeechVoiceSpeakFlags.SVSFDefault)

        Return 1

Erreur_dire_un_texte:
        'End
        Return 0

    End Function

    Private Sub idSpeakText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles idSpeakText.Click
        On Error GoTo Erreur_idSpeakText
        VoiceObj_SAPI5.Speak(idTextBox.Text, 1)

Erreur_idSpeakText:
        If Err.Number Then ShowErrMsg("idSpeakText_Click")
    End Sub

    Private Sub idsVoices_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles idsVoices.SelectedIndexChanged
        On Error GoTo EH

        If (idsVoices.Items.Count > -1 And LOAD_Terminée = True) Then

            'Paramètrage de l'objet  voix par la sélection du nom
            VoiceObj_SAPI5.Voice = VoiceObj_SAPI5.GetVoices().Item(idsVoices.SelectedIndex)
            integer_index_Voice_SAPI5 = idsVoices.SelectedIndex
            My.Settings.Index_Moteur_Synthèse_Vocale_SAPI5 = integer_index_Voice_SAPI5

            'MsgBox("integer_index_Voice_SAPI5 = " & My.Settings.Index_Moteur_Synthèse_Vocale_SAPI5)
        Else

            'MsgBox("Please select a voice from the list box.")
        End If

EH:
        If Err.Number Then ShowErrMsg("idsVoices_SelectedIndexChanged")

    End Sub

    Private Sub idsAudioOutput_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles idsAudioOutput.SelectedIndexChanged

        If (LOAD_Terminée = True) Then

            On Error GoTo Erreur_idsAudioOutput
            VoiceObj_SAPI5.AudioOutput = VoiceObj_SAPI5.GetAudioOutputs().Item(idsAudioOutput.SelectedIndex)
            My.Settings.Audio_Output_Device_SAPI5 = idsAudioOutput.SelectedIndex

        End If

Erreur_idsAudioOutput:
        If Err.Number Then ShowErrMsg("idsAudioOutput_SelectedIndexChanged")
    End Sub

    Private Sub Button_Rate_plus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Rate_plus.Click


        If (VoiceObj_SAPI5.Rate < 10) Then
            VoiceObj_SAPI5.Rate = VoiceObj_SAPI5.Rate + 1
        End If

        Label_Valeur_vitesse.Text = VoiceObj_SAPI5.Rate.ToString
        My.Settings.Index_Vitesse_Voix = VoiceObj_SAPI5.Rate

    End Sub

    Private Sub Button_Rate_moin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Rate_moin.Click

        If (VoiceObj_SAPI5.Rate > -10) Then
            VoiceObj_SAPI5.Rate = VoiceObj_SAPI5.Rate - 1
        End If

        Label_Valeur_vitesse.Text = VoiceObj_SAPI5.Rate.ToString
        My.Settings.Index_Vitesse_Voix = VoiceObj_SAPI5.Rate

    End Sub

    Private Sub Button_Volume_plus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Volume_plus.Click

        If (VoiceObj_SAPI5.Volume < 100) Then
            VoiceObj_SAPI5.Volume = VoiceObj_SAPI5.Volume + 10
        End If

        Label_Valeur_volume.Text = VoiceObj_SAPI5.Volume.ToString
        My.Settings.Index_Volume_Voix = VoiceObj_SAPI5.Volume

    End Sub

    Private Sub Button_Volume_moin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Volume_moin.Click

        If (VoiceObj_SAPI5.Volume > 9) Then
            VoiceObj_SAPI5.Volume = VoiceObj_SAPI5.Volume - 10
        End If

        Label_Valeur_volume.Text = VoiceObj_SAPI5.Volume.ToString
        My.Settings.Index_Volume_Voix = VoiceObj_SAPI5.Volume

    End Sub


    Private Sub ShowErrMsg(ByVal localisation_erreur As String)
        'fonction de message d'erreur

        ' Déclaration d'un identifiant :
        Dim message_erreur As String = ""
        Dim Titre As String = ""

        message_erreur = "Desc: " & Err.Description & vbNewLine
        message_erreur = message_erreur & "Err #: " & Err.Number



        Select Case langue_logiciel
            Case 0
                Titre = "Irrtum"
                message_erreur = "Ein Irrtum hat sich ereignet : " & vbNewLine & message_erreur
            Case 1
                Titre = "Error"
                message_erreur = "An error occurred : " & vbNewLine & message_erreur

            Case 2
                Titre = "error"
                message_erreur = "Un error se produjo : " & vbNewLine & message_erreur

            Case 3
                Titre = "Erreur"
                message_erreur = "Une erreur s'est produite : " & vbNewLine & message_erreur

            Case 4
                Titre = "Errore"
                message_erreur = "Un errore si è prodursi : " & vbNewLine & message_erreur

        End Select



        MsgBox(message_erreur, vbExclamation, Titre & " - " & localisation_erreur)


    End Sub

    Private Sub Formulaire_SAPI_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Dim path_app As String = My.Application.Info.DirectoryPath
        'on écrit dans commandes_SAPI.txt que le logiciel s'arréte ETAT_DE_FONCTIONNEMENT_SAPI = Non
        modification_état_ETAT_DE_FONCTIONNEMENT_SAPI(False, path_app)

    End Sub



    Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        'on va mettre les textes Titre, label, boutons
        'Dans la langue en cours
        Traduction_du_formulaire(langue_logiciel)

        'on remplit les combobox des voix et des périphériques de sorties
        InitializeControls(idsVoices, idsAudioOutput)


        'Sélectionne le nom du moteur
        If (My.Settings.Index_Moteur_Synthèse_Vocale_SAPI5 >= 0) Then

            Me.VoiceObj_SAPI5.Voice = Me.VoiceObj_SAPI5.GetVoices().Item(My.Settings.Index_Moteur_Synthèse_Vocale_SAPI5)
            integer_index_Voice_SAPI5 = My.Settings.Index_Moteur_Synthèse_Vocale_SAPI5
            idsVoices.SelectedIndex = integer_index_Voice_SAPI5

        End If

        'Sélectionne la sortie audio
        If (My.Settings.Audio_Output_Device_SAPI5 > 0) Then

            Me.VoiceObj_SAPI5.AudioOutput = Me.VoiceObj_SAPI5.GetAudioOutputs().Item(My.Settings.Audio_Output_Device_SAPI5)
            integer_index_AudioOutput_SAPI5 = My.Settings.Audio_Output_Device_SAPI5
            idsAudioOutput.SelectedIndex = integer_index_AudioOutput_SAPI5

        End If

        Label_Valeur_vitesse.Text = My.Settings.Index_Vitesse_Voix.ToString
        VoiceObj_SAPI5.Rate = My.Settings.Index_Vitesse_Voix


        Label_Valeur_volume.Text = My.Settings.Index_Volume_Voix.ToString
        VoiceObj_SAPI5.Volume = My.Settings.Index_Volume_Voix



        'On va lire un paramètre transmit au logiciel par des arguments en ligne de commande
        If (My.Application.CommandLineArgs.Count > 0) Then
            Dim index00 As Integer = 0
            For index00 = 0 To (My.Application.CommandLineArgs.Count - 1)
                'MsgBox(My.Application.CommandLineArgs.Item(index00))
                Texte_à_dire_01 = My.Application.CommandLineArgs.Item(index00)
                ' MsgBox(Texte_à_dire_01)
            Next index00
        End If

        If Texte_à_dire_01.ToString.Length > 0 Then

            dire_un_texte(Texte_à_dire_01)

        End If

        Dim path_app As String = My.Application.Info.DirectoryPath
        'on écrit dans commandes_SAPI.txt que le logiciel a démarré ETAT_DE_FONCTIONNEMENT_SAPI = Oui
        modification_état_ETAT_DE_FONCTIONNEMENT_SAPI(True, path_app)

        LOAD_Terminée = True
    End Sub


    Function Traduction_du_formulaire(ByVal Langue As Integer)


        'titre formulaire
        Me.Text = Texte_à_changer(0, langue_logiciel).ToString

        'texte à dire
        idTextBox.Text = Texte_à_changer(1, langue_logiciel).ToString

        'Texte_du_bouton
        idSpeakText.Text = Texte_à_changer(2, langue_logiciel).ToString

        'label_de_la_voix
        Label_Voice.Text = Texte_à_changer(3, langue_logiciel).ToString

        'label_sortie_audio
        Label_Audio_Output.Text = Texte_à_changer(4, langue_logiciel).ToString

        'label_vitesse_de_la_voix
        Label_Rate.Text = Texte_à_changer(5, langue_logiciel).ToString

        'label_du_volume
        Label_Volume.Text = Texte_à_changer(6, langue_logiciel).ToString



        Return True

    End Function

    Public Function Texte_à_changer(ByVal ID_Message As Integer, ByVal Langue As Integer)
        Dim Message As String = Nothing

        'Valeur de Langue
        ' 0 = Allemand
        ' 1 = Anglais
        ' 2 = Espagnole
        ' 3 = Français
        ' 4 = Italien


        Dim Table1(,) As String = { _
        {"AppSpeechSynthesis (SAPI 5.4)", "AppSpeechSynthesis (SAPI 5.4)", "AppSpeechSynthesis (SAPI 5.4)", "AppSpeechSynthesis (SAPI 5.4)", "AppSpeechSynthesis (SAPI 5.4)"}, _
        {"Gehen Sie den Text hinein, den Sie wünschen, hier zu sagen", "Enter text you wish spoken here", "Entre el texto que usted desea decir aquí", "Entrez le texte que vous souhaitez dire ici", "Introducete il testo che augurate dire qui"}, _
        {"Zu sagendem Text", "Speak Text", "Texto que hay que decir", "Texte à dire", "Testo a dire"}, _
        {"Stimme", "Voice", "Voz", "Voix", "Voce"}, _
        {"Audio herausgenommen", "Peripheral of audio output", "Periférico de salida audio", "Périphérique de sortie audio", "Periferico di uscita audio"}, _
        {"Geschwindigkeit :", "Rate :", "Velocidad :", "Vitesse : ", "velocità :"}, _
        {"Band :", "Volume :", "Volumen :", "Volume :", "Volume :"}}

        ''        {"", "", "", "", ""}, _

        Message = Table1(ID_Message, Langue)


        Return Message
    End Function


    Public Function InitializeControls(ByRef voix As ComboBox, ByRef sortie_audio As ComboBox)

        voix.Items.Clear()
        sortie_audio.Items.Clear()

        On Error GoTo Erreur_InitializeControls

        Dim strVoice As String



        'Pour chaque Token on retourne la voix par GetVoices
        For Each Me.VoicesToken_SAPI5 In VoiceObj_SAPI5.GetVoices
            strVoice = VoicesToken_SAPI5.GetDescription     ' Le nom du token
            voix.Items.Add(strVoice)                        ' Ajouter au combobox
        Next

        If voix.Items.Count > -1 Then
            voix.SelectedIndex = 0
            integer_index_Voice_SAPI5 = My.Settings.Index_Moteur_Synthèse_Vocale_SAPI5
            voix.SelectedIndex = integer_index_Voice_SAPI5

        End If


        'MsgBox("integer_index_Voice_SAPI5 = " & integer_index_Voice_SAPI5)

        Dim strAudio As String
        Dim strCurrentAudio As String



        VoicesToken_SAPI5 = VoiceObj_SAPI5.AudioOutput                'Token de la sortie audio courante
        strCurrentAudio = VoicesToken_SAPI5.GetDescription            'Obtient la description from token

        'Montre toutes les sorties disponibles
        'Celle qui est sélectionnée est en cours d'utilisation

        For Each Me.VoicesToken_SAPI5 In VoiceObj_SAPI5.GetAudioOutputs

            strAudio = VoicesToken_SAPI5.GetDescription             'Obtient la description du Token
            sortie_audio.Items.Add(strAudio)
        Next

        If sortie_audio.Items.Count > -1 Then
            sortie_audio.SelectedIndex = 0
            integer_index_AudioOutput_SAPI5 = My.Settings.Audio_Output_Device_SAPI5
            sortie_audio.SelectedIndex = integer_index_AudioOutput_SAPI5
        End If

        Label_Valeur_vitesse.Text = VoiceObj_SAPI5.Rate.ToString
        Label_Valeur_volume.Text = VoiceObj_SAPI5.Volume.ToString

Erreur_InitializeControls:
        If Err.Number Then ShowErrMsg("InitializeControls()")
        Return True
    End Function


    Private Sub ShowAudioOutputs()

        'Fonction de chargement du combobox idsAudioOutput

        On Error GoTo Erreur_ShowAudio_Outputs

        Dim strAudio As String
        Dim strCurrentAudio As String

        idsAudioOutput.Items.Clear()

        VoicesToken_SAPI5 = VoiceObj_SAPI5.AudioOutput                'Token de la sortie audio courante
        strCurrentAudio = VoicesToken_SAPI5.GetDescription            'Get description from token

        'Montre toutes les sorties disponibles

        For Each Me.VoicesToken_SAPI5 In VoiceObj_SAPI5.GetAudioOutputs

            strAudio = VoicesToken_SAPI5.GetDescription             'Obtient la description du token
            idsAudioOutput.Items.Add(strAudio)

        Next

Erreur_ShowAudio_Outputs:
        If Err.Number Then ShowErrMsg("ShowAudioOutputs()")
    End Sub


    Private Sub Timer_Lecture_instructions_Synthèse_vocale_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_Lecture_instructions_Synthèse_vocale.Tick



        Dim chemin_locale_application As String = ""
        chemin_locale_application = My.Application.Info.DirectoryPath

        'Dim liste_des_instructions As New List(Of String)
        'Dim liste_des_instructions_TEMPO As New List(Of String)


        'si date_de_dernière_modification_du_fichier_sauvegardée contient - alors on stocke celle trouvé dans le fichier
        'If date_de_dernière_modification_du_fichier_sauvegardée.Length < 2 Then My.Settings.Date_modification_commandes_SAPI5 = date_de_dernière_modification_du_fichier


        If (IO.File.Exists(chemin_locale_application & "\" & "commandes_SAPI5.txt")) Then 'test la présence du fichier de commandes


            'ETAPE 0
            'DEVONS NOUS REELLEMENT LIRE commandes_SAPI5.txt
            '

            Dim date_de_dernière_modification_du_fichier As String = ""
            Dim date_de_dernière_modification_du_fichier_sauvegardée As String = ""

            Dim fileData As FileInfo = My.Computer.FileSystem.GetFileInfo(My.Application.Info.DirectoryPath & "\commandes_SAPI5.txt")

            date_de_dernière_modification_du_fichier = fileData.LastWriteTime.Date.ToString & "-" & fileData.LastWriteTime.Hour.ToString & ":" & fileData.LastWriteTime.Minute.ToString & ":" & fileData.LastWriteTime.Second.ToString
            date_de_dernière_modification_du_fichier_sauvegardée = My.Settings.Date_modification_commandes_SAPI5

            If (date_de_dernière_modification_du_fichier_sauvegardée = date_de_dernière_modification_du_fichier) Then
                'nous n'avons pas à lire le fichier
                Exit Sub
            Else
                'nous devons lire le fichier
                'on stocke la dernière date à jour
                My.Settings.Date_modification_commandes_SAPI5 = date_de_dernière_modification_du_fichier

            End If

            'ETAPE 1 
            'LECTURE DU FICHIER commandes_SAPI5.txt
            '

            lecture_du_fichier_de_commandes_SAPI5(liste_des_instructions, chemin_locale_application)
            lecture_du_fichier_de_commandes_SAPI5(liste_des_instructions_TEMPO, chemin_locale_application)




            'ETAPE 2
            'INTERPRETATION DES COMMANDES
            '
            Interprétation_des_commandes(liste_des_instructions, liste_des_instructions_TEMPO, chemin_locale_application)
        End If

    End Sub


    Function lecture_du_fichier_de_commandes_SAPI5(ByRef Liste_des_commandes As List(Of String), ByVal chemin_locale_app As String)

        Liste_des_commandes.Clear()

        Try
            ' Open file.txt with the Using statement.
            Using LECTEUR_DE_FLUX As StreamReader = New StreamReader(chemin_locale_app & "\" & "commandes_SAPI5.txt")
                ' Store contents in this String.
                Dim ligne As String

                ' Read first line.
                ligne = LECTEUR_DE_FLUX.ReadLine

                ' Loop over each line in file, While list is Not Nothing.
                Do While (Not ligne Is Nothing)
                    ' Add this line to list.
                    Liste_des_commandes.Add(ligne)
                    ' Display to console.
                    'Console.WriteLine(ligne)
                    ' Read in the next line.
                    ligne = LECTEUR_DE_FLUX.ReadLine
                Loop
            End Using

        Catch message_erreur As Exception

            ShowErrMsg("lecture_du_fichier_de_commandes_SAPI5()") ' affiche un message d'erreur dans la langue en cours

            Return False 'une erreur s'est produite

        End Try

        Return True 'Le fichier de commande a été lu sans problème

    End Function

    Function modification_état_EXECUTION_DE_LA_SYNTHESE_VOCALE(ByVal ACTIF_OUI_NON As Boolean, ByVal chemin_locale As String)

        Dim Lignes_String As String = ""

        Try

            ' ETAPE 1
            'Lecture du fichier de commande
            lecture_du_fichier_de_commandes_SAPI5(liste_des_instructions, chemin_locale)

            For Each ligne In liste_des_instructions

                If (InStr(ligne.ToUpper, "EXECUTION_DE_LA_SYNTHESE_VOCALE") > 0) Then

                    If (ACTIF_OUI_NON = True) Then
                        ligne = "EXECUTION_DE_LA_SYNTHESE_VOCALE = Oui"
                    End If

                    If (ACTIF_OUI_NON = False) Then
                        ligne = "EXECUTION_DE_LA_SYNTHESE_VOCALE = Non"
                    End If

                End If

                'on stocke les lignes à réécrire dans Lignes_String
                Lignes_String = Lignes_String & vbCrLf & ligne

            Next

            'on enléve le 1er vbcrlf
            Lignes_String = Lignes_String.Substring(2)

            'ETAPE 2
            'écriture du nouveau fichier


            Dim Flux_écriture As StreamWriter
            Dim Flux_fichier As FileStream

            If File.Exists(chemin_locale & "\commandes_SAPI5.txt") Then

                Dim chemin_complet_commandes_SAPI As String = chemin_locale & "\commandes_SAPI5.txt"

                Flux_fichier = File.Create(chemin_complet_commandes_SAPI)
                Flux_écriture = New StreamWriter(Flux_fichier)
                Flux_écriture.WriteLine(Lignes_String)
                Flux_écriture.Close()

            End If


            Return True


        Catch ex As Exception
            ShowErrMsg("modification_état_EXECUTION_DE_LA_SYNTHESE_VOCALE()")
            Return False
        End Try


    End Function


    Function modification_état_ETAT_DE_FONCTIONNEMENT_SAPI(ByVal ACTIF_OUI_NON As Boolean, ByVal chemin_locale As String)

        Dim Lignes_String As String = ""

        Try

            ' ETAPE 1
            'Lecture du fichier de commande
            lecture_du_fichier_de_commandes_SAPI5(liste_des_instructions, chemin_locale)

            For Each ligne In liste_des_instructions

                If (InStr(ligne.ToUpper, "ETAT_DE_FONCTIONNEMENT_SAPI") > 0) Then

                    If (ACTIF_OUI_NON = True) Then
                        ligne = "ETAT_DE_FONCTIONNEMENT_SAPI = Oui"
                    End If

                    If (ACTIF_OUI_NON = False) Then
                        ligne = "ETAT_DE_FONCTIONNEMENT_SAPI = Non"
                    End If

                End If

                'on stocke les lignes à réécrire dans Lignes_String
                Lignes_String = Lignes_String & vbCrLf & ligne

            Next

            'on enléve le 1er vbcrlf
            Lignes_String = Lignes_String.Substring(2)

            'ETAPE 2
            'écriture du nouveau fichier


            Dim Flux_écriture As StreamWriter
            Dim Flux_fichier As FileStream

            If File.Exists(chemin_locale & "\commandes_SAPI5.txt") Then

                Dim chemin_complet_commandes_SAPI As String = chemin_locale & "\commandes_SAPI5.txt"

                Flux_fichier = File.Create(chemin_complet_commandes_SAPI)
                Flux_écriture = New StreamWriter(Flux_fichier)
                Flux_écriture.WriteLine(Lignes_String)
                Flux_écriture.Close()

            End If


            Return True


        Catch ex As Exception
            ShowErrMsg("modification_état_ETAT_DE_FONCTIONNEMENT_SAPI()")
            Return False
        End Try


    End Function

    Function effacer_TEXTE_A_DIRE(ByVal chemin_locale As String)

        Dim Lignes_String As String = ""

        Try

            ' ETAPE 1
            'Lecture du fichier de commande
            lecture_du_fichier_de_commandes_SAPI5(liste_des_instructions, chemin_locale)

            For Each ligne In liste_des_instructions

                If (InStr(ligne.ToUpper, "TEXTE_A_DIRE") > 0) Then

                    'on efface le texte à dire    TEXTE_A_DIRE = ""
                    ligne = "TEXTE_A_DIRE = """""
                End If

                'on stocke les lignes à réécrire dans Lignes_String
                Lignes_String = Lignes_String & vbCrLf & ligne

            Next


            'on enléve le 1er vbcrlf
            Lignes_String = Lignes_String.Substring(2)

            'ETAPE 2
            'écriture du nouveau fichier


            Dim Flux_écriture As StreamWriter
            Dim Flux_fichier As FileStream

            If File.Exists(chemin_locale & "\commandes_SAPI5.txt") Then

                Dim chemin_complet_commandes_SAPI As String = chemin_locale & "\commandes_SAPI5.txt"

                Flux_fichier = File.Create(chemin_complet_commandes_SAPI)
                Flux_écriture = New StreamWriter(Flux_fichier)
                Flux_écriture.WriteLine(Lignes_String)
                Flux_écriture.Close()

            End If


            Return True


        Catch ex As Exception
            ShowErrMsg("effacer_TEXTE_A_DIRE()")
            Return False
        End Try


    End Function

    Function consultation_TEMPS_DU_TIMER(ByVal chemin_locale As String)

        Dim Lignes_String As String = ""

        Try

            ' ETAPE 1
            'Lecture du fichier de commande
            lecture_du_fichier_de_commandes_SAPI5(liste_des_instructions, chemin_locale)

            For Each ligne In liste_des_instructions

                If (InStr(ligne.ToUpper, "TEMPS_DU_TIMER") > 0) Then
                    Dim valeur_timer As Integer = 800

                    ligne = ligne.Replace("TEMPS_DU_TIMER", "")         'on supprime  TEMPS_DU_TIMER
                    ligne = ligne.Replace("=", "")                     'on supprime  =
                    ligne = ligne.Replace(" ", "")                     'on supprime les espaces

                    valeur_timer = CInt(ligne)
                    Return valeur_timer


                End If


                

            Next



        Catch ex As Exception
            ShowErrMsg("Consultation TEMPS_DU_TIMER()")
        End Try

        Return False
    End Function

    Function Interprétation_des_commandes(ByRef liste_des_commandes As List(Of String), ByRef liste_des_commandes_TEMPORAIRE As List(Of String), ByVal chemin_locale As String, Optional ByVal réentrance As Boolean = False)

        Try


            'STRUCTURE FICHIER commandes_SAPI5.txt
            '
            'LANGUE = Allemand/Anglais/Espagnole/Francais/Italien
            '
            'VISIBILITE_DU_LOGICIEL = OUI/NON
            '
            'TEXTE_A_DIRE = "bla bla"
            '
            'EXECUTION_DE_LA_SYNTHESE_VOCALE = OUI/NON    (EN COURS/ARRET)
            '
            'ETAT_DE_FONCTIONNEMENT_SAPI   = OUI/NON

            '
            'TEMPS_DU_TIMER = 800

            Dim Str_LANGUE As String = ""
            Dim Str_VISIBILITE_DU_LOGICIEL As String = ""
            Dim Str_TEXTE_A_DIRE As String = ""
            Dim Str_EXECUTION_DE_LA_SYNTHESE_VOCALE As String = ""
            Dim Str_ETAT_DE_FONCTIONNEMENT_SAPI As String = ""
            Dim Str_TEMPS_DU_TIMER As String = ""
            'Dim liste_des_commandes_TEMPO As List(Of String)





            For Each Ligne As String In liste_des_commandes_TEMPORAIRE



                'LANGUE
                If (InStr(Ligne.ToUpper, "LANGUE") > 0) Then
                    Str_LANGUE = Ligne
                    Str_LANGUE = Str_LANGUE.ToUpper                 'en majuscule
                    Str_LANGUE = Str_LANGUE.Replace("LANGUE", "")   'on enléve le mot clef
                    Str_LANGUE = Str_LANGUE.Replace("=", "")        'on enléve le =
                    Str_LANGUE = Str_LANGUE.Replace(" ", "")        'on supprime les espaces

                    'Valeur de Langue
                    ' 0 = Allemand
                    ' 1 = Anglais
                    ' 2 = Espagnole
                    ' 3 = Français
                    ' 4 = Italien

                    Dim langue_tempo As Integer = langue_logiciel

                    'Détection de la valeur de LANGUE
                    If ((InStr(Str_LANGUE.ToUpper, "Allemand".ToUpper) > 0) Or (InStr(Str_LANGUE.ToUpper, "0".ToUpper) > 0)) Then

                        langue_tempo = 0 'Langue Allemande

                    End If

                    If ((InStr(Str_LANGUE.ToUpper, "Anglais".ToUpper) > 0) Or (InStr(Str_LANGUE.ToUpper, "1".ToUpper) > 0)) Then

                        langue_tempo = 1 'Langue Anglaise

                    End If


                    If ((InStr(Str_LANGUE.ToUpper, "Espagnole".ToUpper) > 0) Or (InStr(Str_LANGUE.ToUpper, "2".ToUpper) > 0)) Then

                        langue_tempo = 2 'Langue Espagnole

                    End If


                    If ((InStr(Str_LANGUE.ToUpper, "Francais".ToUpper) > 0) Or (InStr(Str_LANGUE.ToUpper, "3".ToUpper) > 0)) Then

                        langue_tempo = 3 'Langue Française

                    End If


                    If ((InStr(Str_LANGUE.ToUpper, "Italien".ToUpper) > 0) Or (InStr(Str_LANGUE.ToUpper, "4".ToUpper) > 0)) Then

                        langue_tempo = 4 'Langue Italienne

                    End If

                    If (langue_tempo <> langue_logiciel) Then 'évite de recharger les textes du formulaire
                        langue_logiciel = langue_tempo
                        Traduction_du_formulaire(langue_logiciel)
                        Me.Refresh()
                    End If


                End If ' fin LANGUE



                If (InStr(Ligne.ToUpper, "VISIBILITE_DU_LOGICIEL") > 0) Then

                    Str_VISIBILITE_DU_LOGICIEL = Ligne
                    Str_VISIBILITE_DU_LOGICIEL = Str_VISIBILITE_DU_LOGICIEL.ToUpper                                 'en majuscule
                    Str_VISIBILITE_DU_LOGICIEL = Str_VISIBILITE_DU_LOGICIEL.Replace("VISIBILITE_DU_LOGICIEL", "")   'on enléve le mot clef
                    Str_VISIBILITE_DU_LOGICIEL = Str_VISIBILITE_DU_LOGICIEL.Replace("=", "")                        'on enléve le =
                    Str_VISIBILITE_DU_LOGICIEL = Str_VISIBILITE_DU_LOGICIEL.Replace(" ", "")                        'on supprime les espaces

                    If (InStr(Str_VISIBILITE_DU_LOGICIEL.ToUpper, "Oui".ToUpper) > 0) Then
                        Me.TopLevel = True
                    End If

                    If (InStr(Str_VISIBILITE_DU_LOGICIEL.ToUpper, "Non".ToUpper) > 0) Then
                        Me.TopLevel = False
                    End If

                    Me.Refresh()

                End If 'fin VISIBILITE_DU_LOGICIEL



                If (InStr(Ligne.ToUpper, "TEXTE_A_DIRE") > 0) Then
                    Str_TEXTE_A_DIRE = Ligne
                    Str_TEXTE_A_DIRE = Str_TEXTE_A_DIRE.ToUpper                           'en majuscule
                    Str_TEXTE_A_DIRE = Str_TEXTE_A_DIRE.Replace("TEXTE_A_DIRE", "")       'on enléve le mot clef
                    Str_TEXTE_A_DIRE = Str_TEXTE_A_DIRE.Replace("=", "")                  'on enléve le =
                    'Str_TEXTE_A_DIRE = Str_TEXTE_A_DIRE.Replace(" ", "")                  'on enléve les espaces
                    Str_TEXTE_A_DIRE = Str_TEXTE_A_DIRE.Replace("""", "")                 'on enléve les guillemets

                    'If Str_TEXTE_A_DIRE.Length = 0 Then GoTo fin_texte_à_dire
                    If Str_TEXTE_A_DIRE = "" Then GoTo fin_texte_à_dire
                    If Str_TEXTE_A_DIRE = " " Then GoTo fin_texte_à_dire
                    If Str_TEXTE_A_DIRE = "  " Then GoTo fin_texte_à_dire
                    If Str_TEXTE_A_DIRE = "   " Then GoTo fin_texte_à_dire
                    If Str_TEXTE_A_DIRE = "    " Then GoTo fin_texte_à_dire
                    If Str_TEXTE_A_DIRE = "      " Then GoTo fin_texte_à_dire

                    If (consultation_état_EXECUTION_DE_LA_SYNTHESE_VOCALE(chemin_locale) = True) Then
                        'True = la synthèse vocale est en cours d'utilisation.
                        'réentrance dans la fonction   Interprétation_des_commandes()
                        Interprétation_des_commandes(liste_des_instructions, liste_des_instructions_TEMPO, chemin_locale, True)

                    Else

                        'False =  la synthèse vocale est disponible.

                        'on rend la synthèse vocale non disponible.
                        modification_état_EXECUTION_DE_LA_SYNTHESE_VOCALE(True, chemin_locale)

                        'on dit ce que l'on devait dire
                        dire_un_texte(Str_TEXTE_A_DIRE)

                        'Il faut vérifier l'état de la synthèse
                        'est ce que la synthèse a terminé de parler?
                        While True
                            If (VoiceObj_SAPI5.WaitUntilDone(10000) = True) Then
                                'indication sonore la synthèse vocale a terminé de parler
                                'My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
                                Exit While
                            Else
                            End If
                        End While

                        'on efface le TEXTE_A_DIRE dans le fichier commandes_SAPI5.txt
                        'TEXTE_A_DIRE = ""
                        effacer_TEXTE_A_DIRE(chemin_locale)


                        'on rend la synthèse vocale de nouveau disponible.
                        modification_état_EXECUTION_DE_LA_SYNTHESE_VOCALE(False, chemin_locale)

                    End If



fin_texte_à_dire:
                End If 'fin TEXTE_A_DIRE



                'If (InStr(Ligne.ToUpper, "EXECUTION_DE_LA_SYNTHESE_VOCALE") > 0) Then

                'End If 'fin EXECUTION_DE_LA_SYNTHESE_VOCALE



                If ((InStr(Ligne.ToUpper, "ETAT_DE_FONCTIONNEMENT_SAPI") > 0) And _
                    (InStr(Ligne.ToUpper, "Non".ToUpper) > 0)) Then

                    'on enregistre dans commandes_SAPI5.txt que le logiciel va s'arréter
                    modification_état_ETAT_DE_FONCTIONNEMENT_SAPI(False, chemin_locale)
                    End


                End If 'fin EXECUTION_DE_LA_SYNTHESE_VOCALE



                If (InStr(Ligne.ToUpper, "TEMPS_DU_TIMER") > 0) Then
                    'TEMPS_DU_TIMER
                    'consultation de la valeur du timer
                    'dans le fichier commandes_SAPI5.txt
                    Timer_Lecture_instructions_Synthèse_vocale.Interval = consultation_TEMPS_DU_TIMER(chemin_locale)
                    'MsgBox(Timer_Lecture_instructions_Synthèse_vocale.Interval.ToString)
                End If 'fin TEMPS_DU_TIMER






            Next

        Catch ex As Exception
            ShowErrMsg("Interprétation_des_commandes()")
        End Try

        Return True

    End Function


    Function consultation_état_EXECUTION_DE_LA_SYNTHESE_VOCALE(ByVal chemin_locale As String)

        Dim Lignes_String As String = ""

        Try

            ' ETAPE 1
            'Lecture du fichier de commande
            lecture_du_fichier_de_commandes_SAPI5(liste_des_instructions, chemin_locale)

            For Each ligne In liste_des_instructions

                If ((InStr(ligne.ToUpper, "EXECUTION_DE_LA_SYNTHESE_VOCALE") > 0) And _
                    (InStr(ligne.ToUpper, "Oui".ToUpper) > 0)) Then

                    Return True


                End If


                If ((InStr(ligne.ToUpper, "EXECUTION_DE_LA_SYNTHESE_VOCALE") > 0) And _
                    (InStr(ligne.ToUpper, "Non".ToUpper) > 0)) Then

                    Return False

                End If

            Next



        Catch ex As Exception
            ShowErrMsg("consultation_état_EXECUTION_DE_LA_SYNTHESE_VOCALE")
        End Try

        Return False
    End Function


End Class
