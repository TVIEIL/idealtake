Imports SpeechLib
Imports System.IO
Imports System.Windows


Class MainWindow


    Dim altString As String = ""

    'Création de l'objet de Contexte de reconnaissance
    Dim WithEvents MonContexteDeReconnaissance As New SpSharedRecoContextClass
    'Création d'un objet grammaire
    Dim MaGrammaire As ISpeechRecoGrammar
    Dim ListeCommandesVocales As New List(Of String)
    Dim TmpText As String = ""

    Dim LangueLogiciel As Integer = 3 'Français
    Public Liste_des_commandes_SAPI5 As New List(Of String)

    Dim DiagnosticsProcessAppSpeechSynthesis As System.Diagnostics.Process
    Dim ProcessIdAppSpeechSynthesis As Integer = 0

    Private Sub MainWindow_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Closed

        If DiagnosticsProcessAppSpeechSynthesis.HasExited Then

            ' Le processus de AppSpeechSynthesis
            ' Est déjà terminé on fait rien

        Else

            ' à l'Arrêt de Reconnaissance Vocale
            ' On arrête AppSpeechSynthesis

            DiagnosticsProcessAppSpeechSynthesis.Kill()

        End If


    End Sub



    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        'Initialisation de la grammaire à partir du contexte de reconnaissance
        MaGrammaire = MonContexteDeReconnaissance.CreateGrammar(0)
        'Utilisation d'une grammaire statique
        MaGrammaire.DictationLoad("", SpeechLoadOption.SLOStatic)



        If File.Exists(My.Application.Info.DirectoryPath & "\AppSpeechSynthesis.exe") Then

            Try
                ' Call the Process.Start 


                DiagnosticsProcessAppSpeechSynthesis = System.Diagnostics.Process.Start("AppSpeechSynthesis.exe")
                ProcessIdAppSpeechSynthesis = DiagnosticsProcessAppSpeechSynthesis.Id



            Catch ex As Exception

                MessageBox.Show("Impossible d'executer AppSpeechSynthesis.exe.")
            End Try

        End If


    End Sub

    Private Sub ButtonActivation_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ButtonActivation.Click

        'Activation de ma grammaire
        MaGrammaire.DictationSetState(SpeechRuleState.SGDSActive)

        '----------------------------------

        Dim EtatDeLaRegle1 As ISpeechGrammarRuleState = Nothing


        'Création d'une première Règle
        Dim MaRègleOui As ISpeechGrammarRule
        MaRègleOui = MaGrammaire.Rules.Add("RegleOui", SpeechRuleAttributes.SRATopLevel Or SpeechRuleAttributes.SRADefaultToActive, 1)

        Dim proper1 As Object = ""

        'ci dessous le mot pour lequel cette règle sera déclenchée
        MaRègleOui.InitialState.AddWordTransition(EtatDeLaRegle1, "oui", " ", SpeechGrammarWordType.SGLexical, "OUI", 0, proper1, 1.0)


        '----------------------------------

        Dim EtatDeLaRegle2 As ISpeechGrammarRuleState = Nothing

        'Création d'une deuxième Règle
        Dim MaRègleNon As ISpeechGrammarRule
        MaRègleNon = MaGrammaire.Rules.Add("RegleNon", SpeechRuleAttributes.SRATopLevel Or SpeechRuleAttributes.SRADefaultToActive, 2)


        Dim proper2 As Object = ""

        'ci dessous le mot pour lequel cette règle sera déclenchée
        MaRègleOui.InitialState.AddWordTransition(EtatDeLaRegle2, "non", " ", SpeechGrammarWordType.SGLexical, "NON", 0, proper2, 1.0)

        'Ecoute commence
        'MaGrammaire.CmdSetRuleIdState(0, SpeechRuleState.SGDSActive)



        'monRecoContext.Recognition += new_ISpeechRecoContextEvents_RecognitionEventHandler(HandlerReco);

        'Create a recognition context
        'MonContexteDeReconnaissance = CType(recognizer.CreateRecoContext(), SpInProcRecoContext)

        'Wire up events to handlers
        'AddHandler MonContexteDeReconnaissance.Recognition, AddressOf RecoContext_Recognition   'Recognition

        AddHandler MonContexteDeReconnaissance.Hypothesis, AddressOf OnHypothesis              'OnHypothesis
        AddHandler MonContexteDeReconnaissance.Recognition, AddressOf OnRecognition            'OnRecognition
        AddHandler MonContexteDeReconnaissance.EndStream, AddressOf OnAudioStreamEnd           'OnAudioStreamEnd


        ButtonActivation.IsEnabled = False

    End Sub


    Private Sub RecoContext_Recognition(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult) 'Handles MonContexteDeReconnaissance.Recognition
        ' On ajoute un nouveau texte après le dernier mot du TextBox on ajoute un espace
        ' à la fin puis on ajoute nouvelle ligne et retour chariot


        TmpText = ""
        TmpText = Result.PhraseInfo.GetText

        'TextBoxReconnaissance.AppendText(Result.PhraseInfo.GetText & " " & vbCrLf)

        If Not String.IsNullOrEmpty(TmpText) Then
            ListeCommandesVocales.Add(TmpText)
            TextBoxDernierTexte.Text = Result.PhraseInfo.GetText
        Else
            TextBoxDernierTexte.Text = ""
        End If

        If (CheckBox1.IsChecked) Then
            ' on va faire répéter le texte reconnu par la synthèse vocale
            modification_commandes_SAPI5_TEXTE_A_DIRE(TmpText, My.Application.Info.DirectoryPath)
            TextBoxReconnaissance.Text = TextBoxDernierTexte.Text & vbCrLf & TextBoxReconnaissance.Text
        End If


    End Sub



    'This will be called many times, as data is analyzed
    Sub OnHypothesis(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal Result As ISpeechRecoResult)
        SyncLock Me
            Dim info As ISpeechPhraseInfo = Result.PhraseInfo
            'You could, of course, store this for further processing / analysis
            Dim el As ISpeechPhraseElement
            For Each el In info.Elements
                Debug.WriteLine(el.DisplayText)
            Next
            Debug.WriteLine("--Hypothesis over--")
        End SyncLock
    End Sub

    'This will be called once, after entire file is analyzed
    Private Sub OnRecognition(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechRecognitionType, ByVal Result As ISpeechRecoResult)
        Dim phraseInfo As ISpeechPhraseInfo = Result.PhraseInfo
        'The best guess at the completely recognized text
        Dim s As String = phraseInfo.GetText(0, -1, True)
        'RichTextBox1.AppendText(s)


        TmpText = ""
        TmpText = s


        'TextBoxReconnaissance.AppendText(Result.PhraseInfo.GetText & " " & vbCrLf)

        If Not String.IsNullOrEmpty(TmpText) Then
            ListeCommandesVocales.Add(TmpText)
            TextBoxDernierTexte.Text = Result.PhraseInfo.GetText
        Else
            TextBoxDernierTexte.Text = ""
        End If



        'Or you could look at alternates. Here, request up to 10 alternates from index position 0 considering all elements (-1)


        Try
            Dim alternate As ISpeechPhraseAlternate
            For Each alternate In Result.Alternates(10, 10, -1)
                Dim altResult As ISpeechRecoResult = alternate.RecoResult
                Dim altInfo As ISpeechPhraseInfo = altResult.PhraseInfo
                altString = altInfo.GetText(0, -1, True)
                Debug.WriteLine(altString)
            Next

        Catch ce As System.ArgumentException
        End Try

        'If (CheckBox1.IsChecked And altString <> "" Or 
        If (CheckBox1.IsChecked And TextBoxDernierTexte.Text <> "") Then
            ' on va faire répéter le texte reconnu par la synthèse vocale
            modification_commandes_SAPI5_TEXTE_A_DIRE(TmpText, My.Application.Info.DirectoryPath)
            TextBoxReconnaissance.Text = TmpText & vbCrLf & TextBoxReconnaissance.Text
        End If

    End Sub

    Private Sub OnAudioStreamEnd(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal someBool As Boolean)
        'Stop recognition
        MaGrammaire.DictationSetState(SpeechRuleState.SGDSInactive)
        'Unload the active dictation topic
        MaGrammaire.DictationUnload()
    End Sub



    Function modification_commandes_SAPI5_TEXTE_A_DIRE(ByVal Texte_à_dire As String, ByVal chemin_locale As String)

        'Modification 
        'TEXTE_A_DIRE

        Dim Lignes_String As String = ""
        Dim Ligne As String = ""

        Try

            ' ETAPE 1
            'Lecture du fichier de commande
            lecture_du_fichier_de_commandes_SAPI5(Liste_des_commandes_SAPI5, chemin_locale)

            For Each Ligne In Liste_des_commandes_SAPI5

                If (InStr(Ligne.ToUpper, "TEXTE_A_DIRE") > 0) Then

                    Ligne = "TEXTE_A_DIRE = """ & Texte_à_dire & """"


                End If

                'on stocke les lignes à réécrire dans Lignes_String
                Lignes_String = Lignes_String & vbCrLf & Ligne

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
            ShowErrMsg("modification_commandes_SAPI5_ETAT_DE_FONCTIONNEMENT_SAPI()")
            Return False
        End Try


    End Function


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

    Private Sub ShowErrMsg(ByVal Localisation_Erreur As String)
        'fonction de message d'erreur

        ' Déclaration d'un identifiant :
        Dim message_erreur As String = ""
        Dim Titre As String = ""

        message_erreur = "Desc: " & Err.Description & vbNewLine
        message_erreur = message_erreur & "Err #: " & Err.Number



        Select Case LangueLogiciel
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



        MsgBox(message_erreur, vbExclamation, Titre & " - " & Localisation_Erreur)


    End Sub

    Private Sub Label1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs)



    End Sub



End Class
