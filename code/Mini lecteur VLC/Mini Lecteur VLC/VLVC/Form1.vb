Imports AXVLC
Imports System.IO




Public Class Form1

    Public WithEvents objVlc As VLCPlugin
    Private Voptions() As Object
    Dim PathMultimediaFile As String = ""
    Dim LangueLogiciel As Integer = 3
    'Valeur de Langue
    ' 0 = Allemand
    ' 1 = Anglais
    ' 2 = Espagnole
    ' 3 = Français
    ' 4 = Italien

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        If objVlc Is Nothing Then
            objVlc = New AXVLC.VLCPlugin
        End If

        Dim testLectureFichier As Boolean = False

        testLectureFichier = LectureFichierConfig()

        'si on a pas pu charger les 6 paramètres on doit arrêter le Form_Load
        If testLectureFichier = False Then Exit Sub

        Initialiser_langue_logiciel()

        objVlc.addTarget(PathMultimediaFile, Voptions, 1, -666)
        objVlc.Volume = 100
        objVlc.play()
    End Sub

    Private Sub Play_Click(sender As System.Object, e As System.EventArgs) Handles ButtonPlay.Click
        objVlc.Volume = 100
        objVlc.play()
    End Sub

    Private Sub Pause_Click(sender As System.Object, e As System.EventArgs) Handles ButtonPause.Click

        objVlc.pause()

    End Sub

    Private Sub ButtonStop_Click(sender As System.Object, e As System.EventArgs) Handles ButtonStop.Click

        objVlc.stop()

    End Sub

    Private Sub ButtonMute_Click(sender As System.Object, e As System.EventArgs) Handles ButtonMute.Click

        If (objVlc.Volume = 0) Then
            objVlc.Volume = 100
        Else
            objVlc.Volume = 0
        End If

    End Sub

    Private Sub ButtonFullscreen_Click(sender As System.Object, e As System.EventArgs)

        objVlc.fullscreen()

    End Sub


    Private Function LectureFichierConfig()

        Dim chemin_MiniVLCPlayer_CONF As String = ""
        Dim Lignes_String As String = ""
        Dim Ligne_flux As FileStream
        Dim Ligne_ecriture As StreamWriter
        Dim TestInterprétation As Boolean = False

        chemin_MiniVLCPlayer_CONF = My.Application.Info.DirectoryPath & "\MiniVLCPlayer.TXT"
        If Not File.Exists(chemin_MiniVLCPlayer_CONF) Then

            Ligne_flux = File.Create(chemin_MiniVLCPlayer_CONF)
            Ligne_ecriture = New StreamWriter(Ligne_flux)
            Ligne_ecriture.WriteLine(Lignes_String)
            Ligne_ecriture.Close()
        Else


            ' Nous avons besoin de lire dans une liste.
            Dim ListTransfertConfiguration As New List(Of String)

            ' ouverture du fichier txt avec l'état Using.
            Using r As StreamReader = New StreamReader(chemin_MiniVLCPlayer_CONF)
                'Stockage du contenu dans une chaîne.
                Dim line As String

                'Lecture de la première ligne 
                line = r.ReadLine

                ' Boucle sur chaque ligne dans le fichier,
                ' Tant que list n'est pas Nothing.
                Do While (Not line Is Nothing)
                    'Ajouter cette ligne à list
                    ListTransfertConfiguration.Add(line)
                    ' Afficher à la console.
                    'Console.WriteLine(line)

                    ' Lecture de la ligne suivante.
                    line = r.ReadLine
                Loop
                r.Close()
            End Using

            TestInterprétation = Interprétation_MiniVLCPlayer_CONF(ListTransfertConfiguration)


        End If


        Return TestInterprétation

    End Function


    Public Function Interprétation_MiniVLCPlayer_CONF(ByVal liste As List(Of String))


        If liste.Count = 0 Then Return True

        'On recherche des symboles

        '[Path]       => chemin du fichier multimédia
        '[Langue]     => LangueLogiciel


        'test si tout les paramètres ont bien été lus  (score =6)
        ' en cas d'echec ne pas poursuivre le form_load
        Dim score As Integer = 0

        Try
            Dim T As Integer = 0
            Dim ligne_en_test_STR As String = ""


            For T = 0 To (liste.Count - 1)

                ligne_en_test_STR = ""

                ligne_en_test_STR = liste.Item(T)


                If (ligne_en_test_STR.Contains("[Path]")) Then PathMultimediaFile = ligne_en_test_STR.Replace("[Path]", "") : score += 1
                If (ligne_en_test_STR.Contains("[Langue]")) Then LangueLogiciel = CInt(ligne_en_test_STR.Replace("[Langue]", "")) : score += 1



            Next

        Catch ex As Exception
            MsgBox("Interprétation_MiniVLCPlayer_CONF() -  " & ex.Message, MsgBoxStyle.Critical, "- Error -")
            Return False
        End Try

        If score = 2 Then
            Return True
        Else
            Return False
        End If

    End Function


    Private Function Initialiser_langue_logiciel()

        Dim Langue As Integer = LangueLogiciel
        'Langue = My.Settings.Langue_logiciel
        ListeObjectATraduire()

        'Valeur de Langue
        ' 0 = Allemand
        ' 1 = Anglais
        ' 2 = Espagnole
        ' 3 = Français
        ' 4 = Italien

        Select Case Langue
            Case 0
                ToolStripStatusLabel1.Text = " - Deutsche Sprache -"

            Case 1
                ToolStripStatusLabel1.Text = " - English language -"

            Case 2
                ToolStripStatusLabel1.Text = " - Lengua Españoles -"

            Case 3
                ToolStripStatusLabel1.Text = " - Langue Française -"

            Case 4
                ToolStripStatusLabel1.Text = " - Lingua Italiano -"

            Case Else

        End Select


        Return Langue
    End Function





    Public Function ListeObjectATraduire()

        Me.Text = ObjAModifier(0, LangueLogiciel)
        ButtonPlay.Text = ObjAModifier(1, LangueLogiciel)
        ButtonPause.Text = ObjAModifier(2, LangueLogiciel)
        ButtonStop.Text = ObjAModifier(3, LangueLogiciel)
        ButtonMute.Text = ObjAModifier(4, LangueLogiciel)

        Return True

    End Function

    Private Function ObjAModifier(ByVal IDMessage As Integer, ByVal Langue As Integer)
        Dim Message As String = Nothing

        'Valeur de Langue
        ' 0 = Allemand
        ' 1 = Anglais
        ' 2 = Espagnole
        ' 3 = Français
        ' 4 = Italien

        Dim Table1(,) As String = { _
        {"Mini VLC Player", "Mini VLC Player", "Mini VLC Player", "Mini VLC Player", "Mini VLC Player"}, _
        {"Abspielen", "Play", "Tocar", "Lecture", "Ascolta"}, _
        {"Pause", "Pause", "Pausa", "Pause", "Pausa"}, _
        {"Stoppen", "Stop", "Alto", "Stop", "Stop"}, _
        {"Stumm", "Mute", "Mudo", "Muet", "Mute"} _
        }

        ''        {"", "", "", "", ""}, _

        Message = Table1(IDMessage, Langue)


        Return Message
    End Function


End Class
