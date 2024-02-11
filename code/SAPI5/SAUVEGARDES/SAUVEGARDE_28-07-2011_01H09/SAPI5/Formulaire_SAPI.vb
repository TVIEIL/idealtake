


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

    Function dire_un_texte(ByVal Texte_à_dire As String)

        'Me.TopLevel = False

        On Error GoTo Erreur_dire_un_texte
        VoiceObj_SAPI5.Speak(Texte_à_dire, 1)

        Return 1

Erreur_dire_un_texte:
        'End
        Return 0

    End Function

    Private Sub idSpeakText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles idSpeakText.Click
        On Error GoTo Erreur_idSpeakText
        VoiceObj_SAPI5.Speak(idTextBox.Text, 1)

Erreur_idSpeakText:
        If Err.Number Then ShowErrMsg()
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
        If Err.Number Then ShowErrMsg()

    End Sub

    Private Sub idsAudioOutput_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles idsAudioOutput.SelectedIndexChanged

        If (LOAD_Terminée = True) Then

            On Error GoTo Erreur_idsAudioOutput
            VoiceObj_SAPI5.AudioOutput = VoiceObj_SAPI5.GetAudioOutputs().Item(idsAudioOutput.SelectedIndex)
            My.Settings.Audio_Output_Device_SAPI5 = idsAudioOutput.SelectedIndex

        End If

Erreur_idsAudioOutput:
        If Err.Number Then ShowErrMsg()
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


    Private Sub ShowErrMsg()
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



        MsgBox(message_erreur, vbExclamation, Titre)


    End Sub



    Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


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

        InitializeControls(idsVoices, idsAudioOutput)


        'Sélectionne le nom du moteur
        If (My.Settings.Index_Moteur_Synthèse_Vocale_SAPI5 > 0) Then

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



        'comment lire un paramètre transmit au logiciel par des arguments en ligne de commande
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

        LOAD_Terminée = True
    End Sub


    Public Function Texte_à_changer(ByVal ID_Message As Integer, ByVal Langue As Integer)
        Dim Message As String = Nothing

        'Valeur de Langue
        ' 0 = Allemand
        ' 1 = Anglais
        ' 2 = Espagnole
        ' 3 = Français
        ' 4 = Italien


        Dim Table1(,) As String = { _
        {"Test SAPI 5.4", "Test SAPI 5.4", "Prueba SAPI 5.4", "Test SAPI 5.4", "Test SAPI 5.4"}, _
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
        If Err.Number Then ShowErrMsg()
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
        If Err.Number Then ShowErrMsg()
    End Sub


    Private Sub Timer_Lecture_instructions_Synthèse_vocale_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_Lecture_instructions_Synthèse_vocale.Tick

        'STRUCTURE FICHIER commandes_SAPI5.txt
        '
        'VISIBILITE = OUI/NON
        'TEXTE = "bla bla"
        'EXECUTION = ARRETER/EN COURS
        'TEMPS TIMER = 800

        Dim chemin_locale_app As String = ""
        chemin_locale_app = My.Application.Info.DirectoryPath

        Dim liste_des_instructions As New List(Of String)
        If (IO.File.Exists(chemin_locale_app & "\" & "commandes_SAPI5.txt")) Then
            ' Open file.txt with the Using statement.
            Using LECTEUR_DE_FLUX As StreamReader = New StreamReader(chemin_locale_app & "\" & "commandes_SAPI5.txt")
                ' Store contents in this String.
                Dim ligne As String

                ' Read first line.
                ligne = LECTEUR_DE_FLUX.ReadLine

                ' Loop over each line in file, While list is Not Nothing.
                Do While (Not ligne Is Nothing)
                    ' Add this line to list.
                    liste_des_instructions.Add(ligne)
                    ' Display to console.
                    'Console.WriteLine(ligne)
                    ' Read in the next line.
                    ligne = LECTEUR_DE_FLUX.ReadLine
                Loop
            End Using

        End If

    End Sub
End Class
