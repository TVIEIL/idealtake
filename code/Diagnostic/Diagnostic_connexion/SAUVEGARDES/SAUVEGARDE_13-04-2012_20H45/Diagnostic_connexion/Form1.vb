Imports System.Net.NetworkInformation
Imports System.Net
Imports MySql.Data.MySqlClient
Imports Microsoft.Win32
Imports System.ComponentModel

Public Class Form_Diagnostic

    Dim langue_logiciel As Integer = 0
    Dim Diagnostique_Partie1 As String = ""
    Dim Diagnostique_Partie2 As String = ""
    Dim Nom_Connexion As String = ""
    Dim Adresse_Serveur As String = ""
    Dim Nom_de_la_base As String = ""
    Dim connexion As String = ""
    Dim Date_Recherche As New Date
    Dim Résultat_tracert As String = ""
    Dim Infos_diagnostic As String = ""

    Dim Diagnostic_OK As Boolean = False

    Public Structure Argument_objet_Diagnostic

        Public Nom_Connexion As String
        Public Adresse_Serveur As String
        Public chaine_de_connexion As String
        Public Nom_de_la_base_de_données As String
        Public langue_logiciel As Integer


    End Structure

    Private Sub Form_Diagnostic_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        My.Settings.ProgressBar_Diagnostic = 0
        My.Settings.Arret_Timer = False

    End Sub

    'RichTextBox1.SelectionStart = RichTextBox1.Find("are")
    'using the Find method to find the text "are" and setting it's
    'return property to SelectionStart which selects the text to format
    'Dim ifont As New Font(RichTextBox1.Font, FontStyle.Italic)
    'creating a new font object to set the font style
    'RichTextBox1.SelectionFont = ifont
    'assigning the value selected from the RichTextBox the font style
    'RichTextBox1.SelectionStart = RichTextBox1.Find("working")
    'Dim bfont As New Font(RichTextBox1.Font, FontStyle.Bold)
    'RichTextBox1.SelectionFont = bfont



    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Date_Recherche = Date.Now
        langue_logiciel = 3

        Nom_Connexion = "TEST03"
        Adresse_Serveur = "127.0.0.1"
        Nom_de_la_base = "TEST03"
        connexion = "Persist Security Info=true;database=TEST03;server=127.0.0.1;user id=root;pwd=FML6A42V1B"





        Dim argument_Diagnostic01 As New Argument_objet_Diagnostic

        argument_Diagnostic01.Nom_Connexion = Nom_Connexion
        argument_Diagnostic01.Adresse_Serveur = Adresse_Serveur
        argument_Diagnostic01.Nom_de_la_base_de_données = Nom_de_la_base
        argument_Diagnostic01.chaine_de_connexion = connexion

        If Not BackgroundWorker_Diagnostic.IsBusy = True Then


            BackgroundWorker_Diagnostic.RunWorkerAsync(argument_Diagnostic01)

        End If



    End Sub

    Private Function Diagnostic_de_la_connexion(ByVal Argument As Argument_objet_Diagnostic)

        'Diagnostic de la connexion
        Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(93, langue_logiciel) & " : " & Nom_Connexion & " - " & Date_Recherche & vbCrLf & vbCrLf
        If My.Computer.Network.IsAvailable Then

            'L'ordinateur est relié à un réseau local et/ou Internet
            Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(94, langue_logiciel) & "." & vbCrLf


        Else
            'L'ordinateur n'est pas relié à un réseau local et/ou Internet
            Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(95, langue_logiciel) & "." & vbCrLf
            'Exit Sub   empéche le retour à l'aide si une erreure se produit (2/05/2008)
        End If


        My.Settings.ProgressBar_Diagnostic += 1


        'IPGlobalProperties computerProperties = IPGlobalProperties.GetIPGlobalProperties();
        'NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
        'Console.WriteLine("Interface information for {0}.{1}     ", 

        Dim computerProperties As IPGlobalProperties
        computerProperties = IPGlobalProperties.GetIPGlobalProperties

        Dim nics As NetworkInterface()
        nics = NetworkInterface.GetAllNetworkInterfaces

        'Nombre de carte(s) Réseau(x)
        Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(96, langue_logiciel) & " : " & nics.Length.ToString & vbCrLf

        For Each carte As NetworkInterface In nics

            'Carte Réseau
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & " " & Message_utilisateur(97, langue_logiciel) & " : " & carte.Id.ToString & "  -  " & carte.Name.ToString & vbCrLf

            'Infos_diagnostic = Infos_diagnostic & " " & " " & carte.GetIPProperties.ToString & vbCrLf

            'Description de la carte
            Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(98, langue_logiciel) & " : " & carte.Description.ToString & vbCrLf

            'Type d'interface
            Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(99, langue_logiciel) & " : " & carte.NetworkInterfaceType.ToString & vbCrLf

            'Vitesse
            Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(100, langue_logiciel) & " : " & (carte.Speed / 1000000).ToString & " Mbits/s" & vbCrLf

            'Etat
            Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(101, langue_logiciel) & " : " & carte.OperationalStatus.ToString.Replace("Up", "Actif") & vbCrLf

            ' Only display informatin for interfaces that support IPv4.
            If carte.Supports(NetworkInterfaceComponent.IPv4) = False Then
                GoTo ContinueForEach1
            End If
            'Console.WriteLine()
            ''Console.WriteLine(carte.Description)
            ' Underline the description.
            ''Console.WriteLine(String.Empty.PadLeft(carte.Description.Length, "="c))
            Dim adapterProperties As IPInterfaceProperties = carte.GetIPProperties()
            ' Try to get the IPv4 interface properties.
            Dim p1 As IPv4InterfaceProperties = adapterProperties.GetIPv4Properties()

            If p1 Is Nothing Then
                'Pas d'information IPv4 disponible pour cette interface
                Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(102, langue_logiciel) & "." & vbCrLf
                GoTo ContinueForEach1
            End If
            ' Display the IPv4 specific data.
            Infos_diagnostic = Infos_diagnostic & " " & " Index ............................. : " & p1.Index.ToString & vbCrLf
            Infos_diagnostic = Infos_diagnostic & " " & " MTU ............................... : " & p1.Mtu.ToString & vbCrLf
            Infos_diagnostic = Infos_diagnostic & " " & " APIPA active....................... : " & p1.IsAutomaticPrivateAddressingActive.ToString & vbCrLf
            Infos_diagnostic = Infos_diagnostic & " " & " APIPA enabled...................... : " & p1.IsAutomaticPrivateAddressingEnabled.ToString & vbCrLf
            Infos_diagnostic = Infos_diagnostic & " " & " Forwarding enabled................. : " & p1.IsForwardingEnabled.ToString & vbCrLf
            Infos_diagnostic = Infos_diagnostic & " " & " Utilise WINS....................... : " & p1.UsesWins.ToString & vbCrLf
            Infos_diagnostic = Infos_diagnostic & " " & " Utilise DHCP....................... : " & p1.IsDhcpEnabled.ToString & vbCrLf
ContinueForEach1:


            Infos_diagnostic = Infos_diagnostic & " " & " DHCP : "

            Dim nombre_de_DHCP As Integer = 0
            nombre_de_DHCP = carte.GetIPProperties.DhcpServerAddresses.Count
            Dim indice_DHCP As Integer = 1

            If nombre_de_DHCP = 0 Then
                GoTo ligne_11290
            End If

            For indice_DHCP = 1 To nombre_de_DHCP
                Dim DHCP00 As String = ""

                DHCP00 = carte.GetIPProperties.DhcpServerAddresses.Item(indice_DHCP - 1).ToString

                Infos_diagnostic = Infos_diagnostic & "  " & DHCP00

            Next
ligne_11290:
            Infos_diagnostic = Infos_diagnostic & vbCrLf

            'Passerelle
            Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(103, langue_logiciel) & " : "

            Dim nombre_de_passerelle As Integer = 0
            nombre_de_passerelle = carte.GetIPProperties.GatewayAddresses.Count
            Dim indice_passerelle As Integer = 1

            If nombre_de_passerelle = 0 Then
                GoTo ligne_11310
            End If

            For indice_passerelle = 1 To nombre_de_passerelle
                Dim passerelle00 As String = ""

                passerelle00 = carte.GetIPProperties.GatewayAddresses.Item(indice_passerelle - 1).Address.ToString
                Infos_diagnostic = Infos_diagnostic & "  " & passerelle00

            Next
ligne_11310:
            Infos_diagnostic = Infos_diagnostic & vbCrLf



            Infos_diagnostic = Infos_diagnostic & " " & " DNS : "

            Dim nombre_de_DNS As Integer = 0
            nombre_de_DNS = carte.GetIPProperties.DnsAddresses.Count
            Dim indice_DNS As Integer = 1

            If nombre_de_DNS = 0 Then
                GoTo ligne_11332
            End If

            For indice_DNS = 1 To nombre_de_DNS
                Dim DNS00 As String = ""

                DNS00 = carte.GetIPProperties.DnsAddresses.Item(indice_DNS - 1).ToString
                Infos_diagnostic = Infos_diagnostic & "  " & DNS00

            Next
ligne_11332:
            Infos_diagnostic = Infos_diagnostic & vbCrLf



            Infos_diagnostic = Infos_diagnostic & " " & " MAC : "
            'Console.Write("  Physical address ........................ : ");
            'PhysicalAddress address = adapter.GetPhysicalAddress();
            Dim Adresse_physique As PhysicalAddress
            Adresse_physique = carte.GetPhysicalAddress

            Dim octets As Byte()
            octets = Adresse_physique.GetAddressBytes
            Dim I As Integer = 0
            Dim mac_string = ""

            For I = 0 To (octets.Length - 1)
                mac_string = mac_string & octets(I).ToString("x2").ToUpper

                If (I < octets.Length - 1) Then
                    mac_string = mac_string & "-"
                End If

            Next

            Infos_diagnostic = Infos_diagnostic & mac_string & vbCrLf

            Dim Octets_Envoyées As String = ""
            Octets_Envoyées = carte.GetIPv4Statistics.BytesSent.ToString

            'Emission de données IPv4
            Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(104, langue_logiciel) & " : " & Octets_Envoyées & " octets" & vbCrLf

            Dim Octets_Réçues As String = ""
            Octets_Réçues = carte.GetIPv4Statistics.BytesReceived.ToString
            'Réception de données IPv4
            Infos_diagnostic = Infos_diagnostic & " " & " " & Message_utilisateur(105, langue_logiciel) & " : " & Octets_Réçues & " octets" & vbCrLf


        Next

        Infos_diagnostic = Infos_diagnostic & vbCrLf & vbCrLf



        My.Settings.ProgressBar_Diagnostic += 1


        Dim key As RegistryKey
        Dim Nom_Serveur As Object
        key = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\windows NT\CurrentVersion\Winlogon")
        Nom_Serveur = key.GetValue("DefaultDomainName")

        'Marche plus avec Windows 7
        'Infos_diagnostic = Infos_diagnostic & vbCrLf & " Nom de votre ordinateur : " & Nom_Serveur & vbCrLf

        'Nom de votre ordinateur
        Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(106, langue_logiciel) & " : " & Environment.MachineName.ToString() & vbCrLf

        My.Settings.ProgressBar_Diagnostic += 1

        'Adresse(s) IP de cette machine
        Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(107, langue_logiciel) & " : " & vbCrLf

        My.Settings.ProgressBar_Diagnostic += 1

        'Quelques informations utiles
        ' lblOsVersion.Text = Environment.OSVersion.ToString
        ' lblUser.Text = Environment.UserName.ToString
        ' lblRuntime.Text = Environment.Version.ToString
        ' lblUptime.Text = Mid((Environment.TickCount / 3600000), 1, 5) & " :Hours"
        ' lblSystemRoot.Text = Environment.SystemDirectory.ToString
        ' lblCompName.Text = Environment.MachineName.ToString
        ' lblCurrentDirectory.Text = Environment.CurrentDirectory.ToString
        ' lblUserDomainName.Text = Environment.UserDomainName.ToString
        ' lblWorkingSet.Text = Environment.WorkingSet.ToString
        ' lblCommandLine.Text = Environment.CommandLine.ToString
        ' 
        ' For i As Integer = 0 To Environment.GetLogicalDrives.Length - 1
        ' 
        '  lblGetLogicalDrives.Text += Environment.GetLogicalDrives(i)
        ' 
        '  lblGetLogicalDrives.Text += " "
        ' 
        ' Next


        ' Nom de la machine
        Dim NomMachine As String = Dns.GetHostName

        ' Récupération de la liste des IP de la machine
        Dim InfoIps As IPHostEntry = Dns.GetHostEntry(NomMachine)
        Dim MesIp As IPAddress() = InfoIps.AddressList


        For Each CurrentIp As IPAddress In MesIp
            Infos_diagnostic = Infos_diagnostic & " " & CurrentIp.ToString & vbCrLf

            'Infos_diagnostic = Infos_diagnostic & " Broadcast : " & IPAddress.Broadcast.ToString & vbCrLf
            ' Console.WriteLine("Ip : {0}", CurrentIp.ToString)
        Next

        My.Settings.ProgressBar_Diagnostic += 1

        ' Pinguer le serveur
        Dim monPing As Ping = New Ping
        Try
            Dim Reply As PingReply = monPing.Send(Adresse_Serveur)
            'Console.WriteLine("Statut du ping : {0}", Reply.Status)
            'Ping du Serveur distant
            Infos_diagnostic = Infos_diagnostic & vbCrLf & "  " & Message_utilisateur(108, langue_logiciel) & " : """ & Adresse_Serveur & """; Statut : " & Reply.Status.ToString & vbCrLf

        Catch ex As Exception
            'Ping du Serveur distant
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(108, langue_logiciel) & " : """ & Adresse_Serveur & " : """ & Adresse_Serveur & """ " + ex.Message & vbCrLf & vbCrLf

        End Try

        My.Settings.ProgressBar_Diagnostic += 1

        'Recherche en cours
        Infos_diagnostic = Infos_diagnostic & vbCrLf & vbCrLf & " Tracert " & Message_utilisateur(109, langue_logiciel) & "..."
        Diagnostique_Partie1 = Infos_diagnostic  'avant Tracert
        Try
            Tâche_de_fond.RunWorkerAsync(Adresse_Serveur)
        Catch ex As Exception
            'arrêt de la tâche
            Tâche_de_fond.CancelAsync()
            'Réinitialisation des variables
            Diagnostique_Partie1 = ""
            Résultat_tracert = ""
            Diagnostique_Partie2 = ""
        End Try

        My.Settings.ProgressBar_Diagnostic += 1

        Dim reader As MySqlDataReader
        'reader = Nothing
        Dim ConnStr As String = Me.connexion
        'substituer dans connexion le nom de la base par 'mysql'
        ConnStr = ConnStr.Replace(Nom_de_la_base, "mysql")

        'Chaîne de connexion au serveur
        Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(110, langue_logiciel) & " : """ & ConnStr & """ " & vbCrLf
        Diagnostique_Partie2 = vbCrLf & " " & Message_utilisateur(110, langue_logiciel) & " : """ & ConnStr & """ " & vbCrLf
        'MsgBox(ConnStr)
        'UCase()




        Try
            Dim conn As MySqlConnection = New MySqlConnection(ConnStr)


            conn.Open()

            '' Version de MySql Server
            ''select version();
            Dim reader001 As MySqlDataReader
            Dim MySql_Version As String = ""

            'Version du logiciel Serveur MySql
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(111, langue_logiciel)
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & Message_utilisateur(111, langue_logiciel)

            'Requête
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : select version();" & vbCrLf
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : select version();" & vbCrLf
            Dim cmd001 As New MySqlCommand("select version();", conn)
            reader001 = cmd001.ExecuteReader()

            While (reader001.Read())
                MySql_Version = reader001.GetString(0)
                Infos_diagnostic = Infos_diagnostic & " MySql : " & MySql_Version & vbCrLf
                Diagnostique_Partie2 = Diagnostique_Partie2 & " MySql : " & MySql_Version & vbCrLf
            End While

            reader001.Close()
            cmd001.Dispose()

            My.Settings.ProgressBar_Diagnostic += 1

            '' Visualistaion des utilisateurs connectés
            ''SHOW FULL PROCESSLIST
            Dim reader002 As MySqlDataReader
            Dim MySql_users_connected As String = ""

            'Visualisation des utilisateurs connectés
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(112, langue_logiciel)
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(112, langue_logiciel)
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & "  : SHOW FULL PROCESSLIST;" & vbCrLf
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : SHOW FULL PROCESSLIST;" & vbCrLf
            Dim cmd002 As New MySqlCommand("SHOW FULL PROCESSLIST;", conn)
            reader002 = cmd002.ExecuteReader()

            Infos_diagnostic = Infos_diagnostic & " (ID, USER, HOST)" & vbCrLf
            Diagnostique_Partie2 = Diagnostique_Partie2 & " (ID, USER, HOST)" & vbCrLf

            While (reader002.Read())
                MySql_users_connected = " " & reader002.GetString(0) _
                                       & "  " & reader002.GetString(1) _
                                       & "  " & reader002.GetString(2)
                Infos_diagnostic = Infos_diagnostic & MySql_users_connected & vbCrLf
                Diagnostique_Partie2 = Diagnostique_Partie2 & MySql_users_connected & vbCrLf
            End While

            reader002.Close()
            cmd002.Dispose()

            My.Settings.ProgressBar_Diagnostic += 1

            ''Visualisation des bases présentent

            Dim i As Integer = 1
            Dim base_name As String = ""

            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""SHOW DATABASES;"" "
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""SHOW DATABASES;"" "

            Dim cmd As New MySqlCommand("SHOW DATABASES;", conn)
            Dim diag_data_base_list(1000, 10)  ''tableau contenant la liste des base d'un serveur 


            reader = cmd.ExecuteReader()
            'Base(s) présente(nt)
            Infos_diagnostic = Infos_diagnostic & vbCrLf & vbCrLf & " " & Message_utilisateur(114, langue_logiciel) & " : " & vbCrLf & vbCrLf
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & vbCrLf & " " & Message_utilisateur(114, langue_logiciel) & " : " & vbCrLf & vbCrLf

            While (reader.Read())
                base_name = reader.GetString(0)
                Infos_diagnostic = Infos_diagnostic & " -> " & base_name & vbCrLf
                Diagnostique_Partie2 = Diagnostique_Partie2 & " -> " & base_name & vbCrLf
                'on enregistre dans un tableau la liste des bases trouvées sur ce serveur
                diag_data_base_list(i, 1) = base_name 'on stocke l'information dans une table
                i += 1
            End While

            My.Settings.ProgressBar_Diagnostic += 1

            Dim x As Integer = 0
            Dim Base_Name_tableau As String = ""
            Dim Nom_de_la_base As String = ""
            Dim pas_de_base As Boolean = False

            For x = 1 To 1000
                Base_Name_tableau = diag_data_base_list(x, 1)
                Base_Name_tableau = UCase(Base_Name_tableau)

                Nom_de_la_base = UCase(Nom_de_la_base.ToString)

                If (Base_Name_tableau = Nom_de_la_base) Then

                    'Dim Retour = (MsgBox("La base  : """ & Nom_de_la_base & """ Existe Bien. Etes vous sûr de vouloir la recréer?", MsgBoxStyle.YesNo, "Attention !"))
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(10, langue_logiciel) & " """ & Nom_de_la_base & """ " & Message_utilisateur(115, langue_logiciel) & "." & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(10, langue_logiciel) & " """ & Nom_de_la_base & """ " & Message_utilisateur(115, langue_logiciel) & "." & vbCrLf
                    pas_de_base = False


                    '' Combien y a t'il d'enregistrements dans la table menu_verrouillage ---> TABLE_A 13-07-2008
                    Dim reader5 As MySqlDataReader
                    Dim nbre_enr5 As String = ""

                    reader5 = Nothing
                    Dim conn5 As MySqlConnection
                    conn5 = New MySqlConnection(connexion)

                    conn5.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_A;"""   'menu_verrouillage
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_A;"""   'menu_verrouillage
                    Dim cmd_Sql5 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".TABLE_A;", conn5)
                    reader5 = cmd_Sql5.ExecuteReader()
                    reader5.Read()
                    nbre_enr5 = reader5.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'TABLE_A'(menu_verrouillage) : " & nbre_enr5 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'TABLE_A'(menu_verrouillage) : " & nbre_enr5 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn5.Close()


                    My.Settings.ProgressBar_Diagnostic += 1

                    '' Combien y a t'il d'enregistrements dans la table Table_B
                    Dim reader2 As MySqlDataReader
                    Dim nbre_enr2 As String = ""

                    reader2 = Nothing
                    Dim conn2 As MySqlConnection
                    conn2 = New MySqlConnection(connexion)

                    conn2.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_B;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_B;"""

                    Dim cmd_Sql2 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".TABLE_B;", conn2)
                    reader2 = cmd_Sql2.ExecuteReader()
                    reader2.Read()
                    nbre_enr2 = reader2.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'TABLE_B'(menu_audio_video) : " & nbre_enr2 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'TABLE_B'(menu_audio_video) : " & nbre_enr2 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn2.Close()

                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table TABLE_C
                    Dim reader1 As MySqlDataReader
                    Dim nbre_enr As String = ""

                    reader1 = Nothing
                    Dim conn1 As MySqlConnection
                    conn1 = New MySqlConnection(connexion)

                    conn1.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_C;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_C;"""
                    Dim cmd_Sql1 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".TABLE_C;", conn1)
                    reader1 = cmd_Sql1.ExecuteReader()
                    reader1.Read()
                    nbre_enr = reader1.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'TABLE_C' : " & nbre_enr & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'TABLE_C' : " & nbre_enr & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn1.Close()

                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table TABLE_D
                    Dim reader4 As MySqlDataReader
                    Dim nbre_enr4 As String = ""

                    reader4 = Nothing
                    Dim conn4 As MySqlConnection
                    conn4 = New MySqlConnection(connexion)

                    conn4.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_D;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_D;"""
                    Dim cmd_Sql4 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".TABLE_D;", conn4)
                    reader4 = cmd_Sql4.ExecuteReader()
                    reader4.Read()
                    nbre_enr4 = reader4.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'TABLE_D' : " & nbre_enr4 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'TABLE_D' : " & nbre_enr4 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn4.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table TABLE_E (menu_controle_acces)
                    Dim reader3 As MySqlDataReader
                    Dim nbre_enr3 As String = ""

                    reader3 = Nothing
                    Dim conn3 As MySqlConnection
                    conn3 = New MySqlConnection(connexion)

                    conn3.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_E;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_E;"""
                    Dim cmd_Sql3 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".TABLE_E;", conn3)
                    reader3 = cmd_Sql3.ExecuteReader()
                    reader3.Read()
                    nbre_enr3 = reader3.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'TABLE_E' : " & nbre_enr3 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'TABLE_E' : " & nbre_enr3 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn3.Close()

                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table niveau_acces
                    Dim reader6 As MySqlDataReader
                    Dim nbre_enr6 As String = ""

                    reader6 = Nothing
                    Dim conn6 As MySqlConnection
                    conn6 = New MySqlConnection(connexion)

                    conn6.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".niveau_acces;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".niveau_acces;"""
                    Dim cmd_Sql6 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".niveau_acces;", conn6)
                    reader6 = cmd_Sql6.ExecuteReader()
                    reader6.Read()
                    nbre_enr6 = reader6.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'niveau_acces' : " & nbre_enr6 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'niveau_acces' : " & nbre_enr6 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn6.Close()


                    My.Settings.ProgressBar_Diagnostic += 1

                    '' Combien y a t'il d'enregistrements dans la table enregistrements(successeur de la table schéma)
                    ''19/04/2009
                    Dim reader19042009 As MySqlDataReader
                    Dim nbre_enr19042009 As String = ""

                    reader19042009 = Nothing
                    Dim conn19042009 As MySqlConnection
                    conn19042009 = New MySqlConnection(connexion)

                    conn19042009.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".enregistrements;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".enregistrements;"""
                    Dim cmd_Sql19042009 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".enregistrements;", conn19042009)
                    reader19042009 = cmd_Sql19042009.ExecuteReader()
                    reader19042009.Read()
                    nbre_enr19042009 = reader19042009.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'enregistrements' : " & nbre_enr19042009 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'enregistrements' : " & nbre_enr19042009 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn19042009.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table schema
                    Dim reader7 As MySqlDataReader
                    Dim nbre_enr7 As String = ""

                    reader7 = Nothing
                    Dim conn7 As MySqlConnection
                    conn7 = New MySqlConnection(connexion)

                    conn7.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".schema;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".schema;"""
                    Dim cmd_Sql7 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".schema;", conn7)
                    reader7 = cmd_Sql7.ExecuteReader()
                    reader7.Read()
                    nbre_enr7 = reader7.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'schema' : " & nbre_enr7 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'schema' : " & nbre_enr7 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    conn7.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table constructeurs
                    Dim reader8 As MySqlDataReader
                    Dim nbre_enr8 As String = ""

                    reader8 = Nothing
                    Dim conn8 As MySqlConnection
                    conn8 = New MySqlConnection(connexion)

                    conn8.Open()
                    Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".CONSTRUCTEURS;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".CONSTRUCTEURS;"""
                    Dim cmd_Sql8 As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".CONSTRUCTEURS;", conn8)
                    reader8 = cmd_Sql8.ExecuteReader()
                    reader8.Read()
                    nbre_enr8 = reader8.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'CONSTRUCTEURS' : " & nbre_enr8 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'CONSTRUCTEURS' : " & nbre_enr8 & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    conn8.Close()

                    My.Settings.ProgressBar_Diagnostic += 1

                    '' Combien y a t'il d'enregistrements dans la table TYPES_DOCUMENTS
                    Dim reader100508_d As MySqlDataReader
                    Dim nbre_enr100508_d As String = ""

                    reader100508_d = Nothing
                    Dim conn100508_d As MySqlConnection
                    conn100508_d = New MySqlConnection(connexion)

                    conn100508_d.Open()
                    Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".types_documents;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".types_documents;"""
                    Dim cmd_Sql100508_d As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".types_documents;", conn100508_d)
                    reader100508_d = cmd_Sql100508_d.ExecuteReader()
                    reader100508_d.Read()
                    nbre_enr100508_d = reader100508_d.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'types_documents' : " & nbre_enr100508_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'types_documents' : " & nbre_enr100508_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    conn100508_d.Close()

                    My.Settings.ProgressBar_Diagnostic += 1

                    '' Combien y a t'il d'enregistrements dans la table DOC_URL
                    Dim reader310508_d As MySqlDataReader
                    Dim nbre_enr310508_d As String = ""

                    reader310508_d = Nothing
                    Dim conn310508_d As MySqlConnection
                    conn310508_d = New MySqlConnection(connexion)

                    conn310508_d.Open()
                    Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".doc_url;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".doc_url;"""
                    Dim cmd_Sql310508_d As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".doc_url;", conn310508_d)
                    reader310508_d = cmd_Sql310508_d.ExecuteReader()
                    reader310508_d.Read()
                    nbre_enr310508_d = reader310508_d.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'doc_url' : " & nbre_enr310508_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'doc_url' : " & nbre_enr310508_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    conn310508_d.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table categories
                    Dim reader14062008_d As MySqlDataReader
                    Dim nbre_enr14062008_d As String = ""

                    reader14062008_d = Nothing
                    Dim conn14062008_d As MySqlConnection
                    conn14062008_d = New MySqlConnection(connexion)

                    conn14062008_d.Open()
                    Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".categories;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".categories;"""
                    Dim cmd_Sql14062008_d As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".categories;", conn14062008_d)
                    reader14062008_d = cmd_Sql14062008_d.ExecuteReader()
                    reader14062008_d.Read()
                    nbre_enr14062008_d = reader14062008_d.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'categories' : " & nbre_enr14062008_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'categories' : " & nbre_enr14062008_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf
                    conn14062008_d.Close()


                    '' Combien y a t'il d'enregistrements dans la table ref_constructeurs
                    Dim reader20062008_d As MySqlDataReader
                    Dim nbre_enr20062008_d As String = ""

                    reader20062008_d = Nothing
                    Dim conn20062008_d As MySqlConnection
                    conn20062008_d = New MySqlConnection(connexion)

                    conn20062008_d.Open()
                    Infos_diagnostic = Infos_diagnostic & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".ref_constructeurs;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".ref_constructeurs;"""
                    Dim cmd_Sql20062008_d As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".ref_constructeurs;", conn20062008_d)
                    reader20062008_d = cmd_Sql20062008_d.ExecuteReader()
                    reader20062008_d.Read()
                    nbre_enr20062008_d = reader20062008_d.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'ref_constructeurs' : " & nbre_enr20062008_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'ref_constructeurs' : " & nbre_enr20062008_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn20062008_d.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table Table_XAML
                    Dim reader_19_02_2009_d As MySqlDataReader
                    Dim nbre_enr_19_02_2009_d As String = ""

                    reader_19_02_2009_d = Nothing
                    Dim conn_19_02_2009_d As MySqlConnection
                    conn_19_02_2009_d = New MySqlConnection(connexion)

                    conn_19_02_2009_d.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_XAML;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".TABLE_XAML;"""
                    Dim cmd_Sql_19_02_2009_d As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".TABLE_XAML;", conn_19_02_2009_d)
                    reader_19_02_2009_d = cmd_Sql_19_02_2009_d.ExecuteReader()
                    reader_19_02_2009_d.Read()
                    nbre_enr_19_02_2009_d = reader_19_02_2009_d.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'TABLE_XAML' : " & nbre_enr_19_02_2009_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'TABLE_XAML' : " & nbre_enr_19_02_2009_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_19_02_2009_d.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table messages_utilisateur
                    Dim reader_17_04_2009_d As MySqlDataReader
                    Dim nbre_enr_17_04_2009_d As String = ""

                    reader_17_04_2009_d = Nothing
                    Dim conn_17_04_2009_d As MySqlConnection
                    conn_17_04_2009_d = New MySqlConnection(connexion)

                    conn_17_04_2009_d.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".messages_utilisateur;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".messages_utilisateur;"""
                    Dim cmd_Sql_17_04_2009_d As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".messages_utilisateur;", conn_17_04_2009_d)
                    reader_17_04_2009_d = cmd_Sql_17_04_2009_d.ExecuteReader()
                    reader_17_04_2009_d.Read()
                    nbre_enr_17_04_2009_d = reader_17_04_2009_d.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'messages_utilisateur' : " & nbre_enr_17_04_2009_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'messages_utilisateur' : " & nbre_enr_17_04_2009_d & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_17_04_2009_d.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'chrono_idref'
                    Dim reader_17_05_2009a As MySqlDataReader
                    Dim nbre_enr_17_05_2009a As String = ""

                    reader_17_05_2009a = Nothing
                    Dim conn_17_05_2009a As MySqlConnection
                    conn_17_05_2009a = New MySqlConnection(connexion)

                    conn_17_05_2009a.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".chrono_idref;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".chrono_idref;"""
                    Dim cmd_Sql_17_05_2009a As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".chrono_idref;", conn_17_05_2009a)
                    reader_17_05_2009a = cmd_Sql_17_05_2009a.ExecuteReader()
                    reader_17_05_2009a.Read()
                    nbre_enr_17_05_2009a = reader_17_05_2009a.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'chrono_idref' : " & nbre_enr_17_05_2009a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'chrono_idref' : " & nbre_enr_17_05_2009a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_17_05_2009a.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'documents'
                    Dim reader_06_06_2009a As MySqlDataReader
                    Dim nbre_enr_06_06_2009a As String = ""

                    reader_06_06_2009a = Nothing
                    Dim conn_06_06_2009a As MySqlConnection
                    conn_06_06_2009a = New MySqlConnection(connexion)

                    conn_06_06_2009a.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".documents;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".documents;"""
                    Dim cmd_Sql_06_06_2009a As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".documents;", conn_06_06_2009a)
                    reader_06_06_2009a = cmd_Sql_06_06_2009a.ExecuteReader()
                    reader_06_06_2009a.Read()
                    nbre_enr_06_06_2009a = reader_06_06_2009a.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'documents' : " & nbre_enr_06_06_2009a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'documents' : " & nbre_enr_06_06_2009a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_06_06_2009a.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'elements_winform'
                    Dim reader_11_07_2009a As MySqlDataReader
                    Dim nbre_enr_11_07_2009a As String = ""

                    reader_11_07_2009a = Nothing
                    Dim conn_11_07_2009a As MySqlConnection
                    conn_11_07_2009a = New MySqlConnection(connexion)

                    conn_11_07_2009a.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".elements_winform;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".elements_winform;"""
                    Dim cmd_Sql_11_07_2009a As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".elements_winform;", conn_11_07_2009a)
                    reader_11_07_2009a = cmd_Sql_11_07_2009a.ExecuteReader()
                    reader_11_07_2009a.Read()
                    nbre_enr_11_07_2009a = reader_11_07_2009a.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'elements_winform' : " & nbre_enr_11_07_2009a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'elements_winform' : " & nbre_enr_11_07_2009a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_11_07_2009a.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine_zero'
                    Dim reader_13_05_2010e As MySqlDataReader
                    Dim nbre_enr_13_05_2010e As String = ""

                    reader_13_05_2010e = Nothing
                    Dim conn_13_05_2010e As MySqlConnection
                    conn_13_05_2010e = New MySqlConnection(connexion)

                    conn_13_05_2010e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".racine_zero;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".racine_zero;"""
                    Dim cmd_Sql_13_05_2010e As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".racine_zero;", conn_13_05_2010e)
                    reader_13_05_2010e = cmd_Sql_13_05_2010e.ExecuteReader()
                    reader_13_05_2010e.Read()
                    nbre_enr_13_05_2010e = reader_13_05_2010e.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine_zero' : " & nbre_enr_13_05_2010e & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine_zero' : " & nbre_enr_13_05_2010e & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010e.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine1'
                    Dim reader_13_05_2010f As MySqlDataReader
                    Dim nbre_enr_13_05_2010f As String = ""

                    reader_13_05_2010f = Nothing
                    Dim conn_13_05_2010f As MySqlConnection
                    conn_13_05_2010f = New MySqlConnection(connexion)

                    conn_13_05_2010f.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine1`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine1`;"""
                    Dim cmd_Sql_13_05_2010f As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine1`;", conn_13_05_2010f)
                    reader_13_05_2010f = cmd_Sql_13_05_2010f.ExecuteReader()
                    reader_13_05_2010f.Read()
                    nbre_enr_13_05_2010f = reader_13_05_2010f.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine1' : " & nbre_enr_13_05_2010f & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine1' : " & nbre_enr_13_05_2010f & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010f.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine2'
                    Dim reader_13_05_2010g As MySqlDataReader
                    Dim nbre_enr_13_05_2010g As String = ""

                    reader_13_05_2010g = Nothing
                    Dim conn_13_05_2010g As MySqlConnection
                    conn_13_05_2010g = New MySqlConnection(connexion)

                    conn_13_05_2010g.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine2`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine2`;"""
                    Dim cmd_Sql_13_05_2010g As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine2`;", conn_13_05_2010g)
                    reader_13_05_2010g = cmd_Sql_13_05_2010g.ExecuteReader()
                    reader_13_05_2010g.Read()
                    nbre_enr_13_05_2010g = reader_13_05_2010g.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine2' : " & nbre_enr_13_05_2010g & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine2' : " & nbre_enr_13_05_2010g & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010g.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine3'
                    Dim reader_13_05_2010h As MySqlDataReader
                    Dim nbre_enr_13_05_2010h As String = ""

                    reader_13_05_2010h = Nothing
                    Dim conn_13_05_2010h As MySqlConnection
                    conn_13_05_2010h = New MySqlConnection(connexion)

                    conn_13_05_2010h.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine3`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine3`;"""
                    Dim cmd_Sql_13_05_2010h As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine3`;", conn_13_05_2010h)
                    reader_13_05_2010h = cmd_Sql_13_05_2010h.ExecuteReader()
                    reader_13_05_2010h.Read()
                    nbre_enr_13_05_2010h = reader_13_05_2010h.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine3' : " & nbre_enr_13_05_2010h & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine3' : " & nbre_enr_13_05_2010h & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010h.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine4'
                    Dim reader_13_05_2010i As MySqlDataReader
                    Dim nbre_enr_13_05_2010i As String = ""

                    reader_13_05_2010i = Nothing
                    Dim conn_13_05_2010i As MySqlConnection
                    conn_13_05_2010i = New MySqlConnection(connexion)

                    conn_13_05_2010i.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine4`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine4`;"""
                    Dim cmd_Sql_13_05_2010i As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine4`;", conn_13_05_2010i)
                    reader_13_05_2010i = cmd_Sql_13_05_2010i.ExecuteReader()
                    reader_13_05_2010i.Read()
                    nbre_enr_13_05_2010i = reader_13_05_2010i.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine4' : " & nbre_enr_13_05_2010i & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine4' : " & nbre_enr_13_05_2010i & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010i.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine5'
                    Dim reader_13_05_2010j As MySqlDataReader
                    Dim nbre_enr_13_05_2010j As String = ""

                    reader_13_05_2010j = Nothing
                    Dim conn_13_05_2010j As MySqlConnection
                    conn_13_05_2010j = New MySqlConnection(connexion)

                    conn_13_05_2010j.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine5`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine5`;"""
                    Dim cmd_Sql_13_05_2010j As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine5`;", conn_13_05_2010j)
                    reader_13_05_2010j = cmd_Sql_13_05_2010j.ExecuteReader()
                    reader_13_05_2010j.Read()
                    nbre_enr_13_05_2010j = reader_13_05_2010j.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine5' : " & nbre_enr_13_05_2010j & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine5' : " & nbre_enr_13_05_2010j & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010j.Close()


                    My.Settings.ProgressBar_Diagnostic += 1



                    '' Combien y a t'il d'enregistrements dans la table 'racine6'
                    Dim reader_13_05_2010k As MySqlDataReader
                    Dim nbre_enr_13_05_2010k As String = ""

                    reader_13_05_2010k = Nothing
                    Dim conn_13_05_2010k As MySqlConnection
                    conn_13_05_2010k = New MySqlConnection(connexion)

                    conn_13_05_2010k.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine6`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine6`;"""
                    Dim cmd_Sql_13_05_2010k As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine6`;", conn_13_05_2010k)
                    reader_13_05_2010k = cmd_Sql_13_05_2010k.ExecuteReader()
                    reader_13_05_2010k.Read()
                    nbre_enr_13_05_2010k = reader_13_05_2010k.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine6' : " & nbre_enr_13_05_2010k & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine6' : " & nbre_enr_13_05_2010k & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010k.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine7'
                    Dim reader_13_05_2010l As MySqlDataReader
                    Dim nbre_enr_13_05_2010l As String = ""

                    reader_13_05_2010l = Nothing
                    Dim conn_13_05_2010l As MySqlConnection
                    conn_13_05_2010l = New MySqlConnection(connexion)

                    conn_13_05_2010l.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine7`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine7`;"""
                    Dim cmd_Sql_13_05_2010l As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine7`;", conn_13_05_2010l)
                    reader_13_05_2010l = cmd_Sql_13_05_2010l.ExecuteReader()
                    reader_13_05_2010l.Read()
                    nbre_enr_13_05_2010l = reader_13_05_2010l.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine7' : " & nbre_enr_13_05_2010l & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine7' : " & nbre_enr_13_05_2010l & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010l.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine8'
                    Dim reader_13_05_2010m As MySqlDataReader
                    Dim nbre_enr_13_05_2010m As String = ""

                    reader_13_05_2010m = Nothing
                    Dim conn_13_05_2010m As MySqlConnection
                    conn_13_05_2010m = New MySqlConnection(connexion)

                    conn_13_05_2010m.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine8`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine8`;"""
                    Dim cmd_Sql_13_05_2010m As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine8`;", conn_13_05_2010m)
                    reader_13_05_2010m = cmd_Sql_13_05_2010m.ExecuteReader()
                    reader_13_05_2010m.Read()
                    nbre_enr_13_05_2010m = reader_13_05_2010m.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine8' : " & nbre_enr_13_05_2010m & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine8' : " & nbre_enr_13_05_2010m & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010m.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'racine9'
                    Dim reader_13_05_2010n As MySqlDataReader
                    Dim nbre_enr_13_05_2010n As String = ""

                    reader_13_05_2010n = Nothing
                    Dim conn_13_05_2010n As MySqlConnection
                    conn_13_05_2010n = New MySqlConnection(connexion)

                    conn_13_05_2010n.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine9`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`racine9`;"""
                    Dim cmd_Sql_13_05_2010n As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`racine9`;", conn_13_05_2010n)
                    reader_13_05_2010n = cmd_Sql_13_05_2010n.ExecuteReader()
                    reader_13_05_2010n.Read()
                    nbre_enr_13_05_2010n = reader_13_05_2010n.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'racine9' : " & nbre_enr_13_05_2010n & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'racine9' : " & nbre_enr_13_05_2010n & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_13_05_2010n.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'elements_HTML'
                    Dim reader_20_08_2010a As MySqlDataReader
                    Dim nbre_enr_20_08_2010a As String = ""

                    reader_20_08_2010a = Nothing
                    Dim conn_20_08_2010a As MySqlConnection
                    conn_20_08_2010a = New MySqlConnection(connexion)

                    conn_20_08_2010a.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`elements_HTML`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`elements_HTML`;"""
                    Dim cmd_Sql_20_08_2010a As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`elements_HTML`;", conn_20_08_2010a)
                    reader_20_08_2010a = cmd_Sql_20_08_2010a.ExecuteReader()
                    reader_20_08_2010a.Read()
                    nbre_enr_20_08_2010a = reader_20_08_2010a.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'elements_HTML' : " & nbre_enr_20_08_2010a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'elements_HTML' : " & nbre_enr_20_08_2010a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_20_08_2010a.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'HISTORY_EVENTS'
                    Dim reader_20_08_2010b As MySqlDataReader
                    Dim nbre_enr_20_08_2010b As String = ""

                    reader_20_08_2010b = Nothing
                    Dim conn_20_08_2010b As MySqlConnection
                    conn_20_08_2010b = New MySqlConnection(connexion)

                    conn_20_08_2010b.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`HISTORY_EVENTS`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`HISTORY_EVENTS`;"""
                    Dim cmd_Sql_20_08_2010b As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`HISTORY_EVENTS`;", conn_20_08_2010b)
                    reader_20_08_2010b = cmd_Sql_20_08_2010b.ExecuteReader()
                    reader_20_08_2010b.Read()
                    nbre_enr_20_08_2010b = reader_20_08_2010b.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'HISTORY_EVENTS' : " & nbre_enr_20_08_2010b & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'HISTORY_EVENTS' : " & nbre_enr_20_08_2010b & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_20_08_2010b.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'Users'
                    Dim reader_20_08_2010c As MySqlDataReader
                    Dim nbre_enr_20_08_2010c As String = ""

                    reader_20_08_2010c = Nothing
                    Dim conn_20_08_2010c As MySqlConnection
                    conn_20_08_2010c = New MySqlConnection(connexion)

                    conn_20_08_2010c.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`Users`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`Users`;"""
                    Dim cmd_Sql_20_08_2010c As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`Users`;", conn_20_08_2010c)
                    reader_20_08_2010c = cmd_Sql_20_08_2010c.ExecuteReader()
                    reader_20_08_2010c.Read()
                    nbre_enr_20_08_2010c = reader_20_08_2010c.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'Users' : " & nbre_enr_20_08_2010c & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'Users' : " & nbre_enr_20_08_2010c & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_20_08_2010c.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'Questions'
                    Dim reader_05_09_2010a As MySqlDataReader
                    Dim nbre_enr_05_09_2010a As String = ""

                    reader_05_09_2010a = Nothing
                    Dim conn_05_09_2010a As MySqlConnection
                    conn_05_09_2010a = New MySqlConnection(connexion)

                    conn_05_09_2010a.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`Questions`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`Questions`;"""
                    Dim cmd_Sql_05_09_2010a As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`Questions`;", conn_05_09_2010a)
                    reader_05_09_2010a = cmd_Sql_05_09_2010a.ExecuteReader()
                    reader_05_09_2010a.Read()
                    nbre_enr_05_09_2010a = reader_05_09_2010a.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'Questions' : " & nbre_enr_05_09_2010a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'Questions' : " & nbre_enr_05_09_2010a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf
                    conn_05_09_2010a.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    '' Combien y a t'il d'enregistrements dans la table 'VERSIONS_TABLES'
                    Dim reader_30_04_2011a As MySqlDataReader
                    Dim nbre_enr_30_04_2011a As String = ""

                    reader_30_04_2011a = Nothing
                    Dim conn_30_04_2011a As MySqlConnection
                    conn_30_04_2011a = New MySqlConnection(connexion)

                    conn_30_04_2011a.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`VERSIONS_TABLES`;"""
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""select count(*) from " & Nom_de_la_base & ".`VERSIONS_TABLES`;"""
                    Dim cmd_Sql_30_04_2011a As New MySqlCommand("select count(*) from " & Nom_de_la_base & ".`VERSIONS_TABLES`;", conn_30_04_2011a)
                    reader_30_04_2011a = cmd_Sql_30_04_2011a.ExecuteReader()
                    reader_30_04_2011a.Read()
                    nbre_enr_30_04_2011a = reader_30_04_2011a.GetString(0)
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & "  -> Table 'VERSIONS_TABLES' : " & nbre_enr_30_04_2011a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & "  -> Table 'VERSIONS_TABLES' : " & nbre_enr_30_04_2011a & " " & Message_utilisateur(116, langue_logiciel) & vbCrLf & vbCrLf & vbCrLf
                    conn_30_04_2011a.Close()

                    Dim tempo1, tempo2, tempo3, tempo4, tempo5, tempo6
                    Dim result1 As String = ""



                    My.Settings.ProgressBar_Diagnostic += 1

                    ''Description de la table menu_verrouillage TABLE_A
                    Dim reader13 As MySqlDataReader
                    result1 = ""

                    reader13 = Nothing
                    Dim conn13 As MySqlConnection
                    conn13 = New MySqlConnection(connexion)

                    conn13.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_A;""" & vbCrLf  'menu_verrouillage
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_A;""" & vbCrLf  'menu_verrouillage
                    Dim cmd_Sql13 As New MySqlCommand("describe  " & Nom_de_la_base & ".TABLE_A;", conn13)
                    reader13 = cmd_Sql13.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader13.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader13.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader13.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader13.Read()

                        tempo1 = reader13.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader13.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader13.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn13.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table TABLE_B
                    Dim reader10 As MySqlDataReader
                    result1 = ""

                    reader10 = Nothing
                    Dim conn10 As MySqlConnection
                    conn10 = New MySqlConnection(connexion)

                    conn10.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_B;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_B;""" & vbCrLf
                    Dim cmd_Sql10 As New MySqlCommand("describe  " & Nom_de_la_base & ".TABLE_B;", conn10)
                    reader10 = cmd_Sql10.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader10.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader10.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader10.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf


                    While reader10.Read()

                        tempo1 = reader10.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader10.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader10.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn10.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table TABLE_C
                    Dim reader9 As MySqlDataReader


                    reader9 = Nothing
                    Dim conn9 As MySqlConnection
                    conn9 = New MySqlConnection(connexion)

                    conn9.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_C;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_C;""" & vbCrLf
                    Dim cmd_Sql9 As New MySqlCommand("describe  " & Nom_de_la_base & ".TABLE_C;", conn9)
                    reader9 = cmd_Sql9.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader9.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader9.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader9.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader9.Read()

                        tempo1 = reader9.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader9.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader9.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn9.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table TABLE_D (menu_serie)
                    Dim reader12 As MySqlDataReader
                    result1 = ""

                    reader12 = Nothing
                    Dim conn12 As MySqlConnection
                    conn12 = New MySqlConnection(connexion)

                    conn12.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_D;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_D;""" & vbCrLf
                    Dim cmd_Sql12 As New MySqlCommand("describe  " & Nom_de_la_base & ".TABLE_D;", conn12)
                    reader12 = cmd_Sql12.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader12.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader12.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader12.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader12.Read()

                        tempo1 = reader12.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader12.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader12.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn12.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table TABLE_E (menu_controle_acces)
                    Dim reader11 As MySqlDataReader
                    result1 = ""

                    reader11 = Nothing
                    Dim conn11 As MySqlConnection
                    conn11 = New MySqlConnection(connexion)

                    conn11.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_E;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_E;""" & vbCrLf
                    Dim cmd_Sql11 As New MySqlCommand("describe  " & Nom_de_la_base & ".TABLE_E;", conn11)
                    reader11 = cmd_Sql11.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader11.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader11.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader11.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader11.Read()

                        tempo1 = reader11.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader11.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader11.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn11.Close()



                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table niveau_acces
                    Dim reader14 As MySqlDataReader
                    result1 = ""

                    reader14 = Nothing
                    Dim conn14 As MySqlConnection
                    conn14 = New MySqlConnection(connexion)

                    conn14.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".niveau_acces;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".niveau_acces;""" & vbCrLf
                    Dim cmd_Sql14 As New MySqlCommand("describe  " & Nom_de_la_base & ".niveau_acces;", conn14)
                    reader14 = cmd_Sql14.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader14.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader14.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader14.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader14.Read()

                        tempo1 = reader14.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader14.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader14.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn14.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table enregistrements
                    Dim reader19042009a As MySqlDataReader
                    result1 = ""

                    reader19042009a = Nothing
                    Dim conn19042009a As MySqlConnection
                    conn19042009a = New MySqlConnection(connexion)

                    conn19042009a.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".enregistrements;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".enregistrements;""" & vbCrLf
                    Dim cmd_Sql19042009a As New MySqlCommand("describe  " & Nom_de_la_base & ".enregistrements;", conn19042009a)
                    reader19042009a = cmd_Sql19042009a.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader19042009a.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader19042009a.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader19042009a.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader19042009a.Read()

                        tempo1 = reader19042009a.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader19042009a.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader19042009a.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn19042009a.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table schema
                    Dim reader15 As MySqlDataReader
                    result1 = ""

                    reader15 = Nothing
                    Dim conn15 As MySqlConnection
                    conn15 = New MySqlConnection(connexion)

                    conn15.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".schema;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".schema;""" & vbCrLf
                    Dim cmd_Sql15 As New MySqlCommand("describe  " & Nom_de_la_base & ".schema;", conn15)
                    reader15 = cmd_Sql15.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader15.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader15.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader15.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader15.Read()

                        tempo1 = reader15.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader15.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader15.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn15.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table constructeurs
                    Dim reader16 As MySqlDataReader
                    result1 = ""

                    reader16 = Nothing
                    Dim conn16 As MySqlConnection
                    conn16 = New MySqlConnection(connexion)

                    conn16.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".CONSTRUCTEURS;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".CONSTRUCTEURS;""" & vbCrLf
                    Dim cmd_Sql16 As New MySqlCommand("describe  " & Nom_de_la_base & ".CONSTRUCTEURS;", conn16)
                    reader16 = cmd_Sql16.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader16.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader16.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader16.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader16.Read()

                        tempo1 = reader16.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader16.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader16.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn16.Close()

                    Infos_diagnostic = Infos_diagnostic & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table types_documents
                    Dim reader100508_e As MySqlDataReader
                    result1 = ""

                    reader100508_e = Nothing
                    Dim conn100508_e As MySqlConnection
                    conn100508_e = New MySqlConnection(connexion)

                    conn100508_e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".types_documents;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".types_documents;""" & vbCrLf
                    Dim cmd_Sql100508_e As New MySqlCommand("describe  " & Nom_de_la_base & ".types_documents;", conn100508_e)
                    reader100508_e = cmd_Sql100508_e.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader100508_e.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader100508_e.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader100508_e.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader100508_e.Read()

                        tempo1 = reader100508_e.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader100508_e.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader100508_e.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn100508_e.Close()

                    Infos_diagnostic = Infos_diagnostic & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table doc_url
                    Dim reader310508_e As MySqlDataReader
                    result1 = ""

                    reader310508_e = Nothing
                    Dim conn310508_e As MySqlConnection
                    conn310508_e = New MySqlConnection(connexion)

                    conn310508_e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".doc_url;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".doc_url;""" & vbCrLf
                    Dim cmd_Sql310508_e As New MySqlCommand("describe  " & Nom_de_la_base & ".doc_url;", conn310508_e)
                    reader310508_e = cmd_Sql310508_e.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader310508_e.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader310508_e.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader310508_e.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader310508_e.Read()

                        tempo1 = reader310508_e.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader310508_e.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader310508_e.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn310508_e.Close()

                    Infos_diagnostic = Infos_diagnostic & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table categories
                    Dim reader14062008_e As MySqlDataReader
                    result1 = ""

                    reader14062008_e = Nothing
                    Dim conn14062008_e As MySqlConnection
                    conn14062008_e = New MySqlConnection(connexion)

                    conn14062008_e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".categories;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".categories;""" & vbCrLf
                    Dim cmd_Sql14062008_e As New MySqlCommand("describe  " & Nom_de_la_base & ".categories;", conn14062008_e)
                    reader14062008_e = cmd_Sql14062008_e.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader14062008_e.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader14062008_e.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader14062008_e.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader14062008_e.Read()

                        tempo1 = reader14062008_e.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader14062008_e.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader14062008_e.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn14062008_e.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table ref_constructeurs
                    Dim reader20062008_e As MySqlDataReader
                    result1 = ""

                    reader20062008_e = Nothing
                    Dim conn20062008_e As MySqlConnection
                    conn20062008_e = New MySqlConnection(connexion)

                    conn20062008_e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".ref_constructeurs;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".ref_constructeurs;""" & vbCrLf
                    Dim cmd_Sql20062008_e As New MySqlCommand("describe  " & Nom_de_la_base & ".ref_constructeurs;", conn20062008_e)
                    reader20062008_e = cmd_Sql20062008_e.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader20062008_e.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader20062008_e.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader20062008_e.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader20062008_e.Read()

                        tempo1 = reader20062008_e.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader20062008_e.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader20062008_e.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn20062008_e.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `TABLE_XAML`
                    Dim reader_19_02_2009_e As MySqlDataReader
                    result1 = ""

                    reader_19_02_2009_e = Nothing
                    Dim conn_19_02_2009_e As MySqlConnection
                    conn_19_02_2009_e = New MySqlConnection(connexion)

                    conn_19_02_2009_e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_XAML;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".TABLE_XAML;""" & vbCrLf
                    Dim cmd_Sql_19_02_2009_e As New MySqlCommand("describe  " & Nom_de_la_base & ".TABLE_XAML;", conn_19_02_2009_e)
                    reader_19_02_2009_e = cmd_Sql_19_02_2009_e.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader_19_02_2009_e.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_19_02_2009_e.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_19_02_2009_e.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_19_02_2009_e.Read()

                        tempo1 = reader_19_02_2009_e.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_19_02_2009_e.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_19_02_2009_e.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_19_02_2009_e.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `messages_utilisateur`
                    Dim reader_17_04_2009_e As MySqlDataReader
                    result1 = ""

                    reader_17_04_2009_e = Nothing
                    Dim conn_17_04_2009_e As MySqlConnection
                    conn_17_04_2009_e = New MySqlConnection(connexion)

                    conn_17_04_2009_e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".messages_utilisateur;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".messages_utilisateur;""" & vbCrLf
                    Dim cmd_Sql_17_04_2009_e As New MySqlCommand("describe  " & Nom_de_la_base & ".messages_utilisateur;", conn_17_04_2009_e)
                    reader_17_04_2009_e = cmd_Sql_17_04_2009_e.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader_17_04_2009_e.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_17_04_2009_e.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_17_04_2009_e.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_17_04_2009_e.Read()

                        tempo1 = reader_17_04_2009_e.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_17_04_2009_e.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_17_04_2009_e.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_17_04_2009_e.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `chrono_idref`
                    Dim reader_17_05_2009_b As MySqlDataReader
                    result1 = ""

                    reader_17_05_2009_b = Nothing
                    Dim conn_17_05_2009_b As MySqlConnection
                    conn_17_05_2009_b = New MySqlConnection(connexion)

                    conn_17_05_2009_b.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".chrono_idref;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".chrono_idref;""" & vbCrLf
                    Dim cmd_Sql_17_05_2009_b As New MySqlCommand("describe  " & Nom_de_la_base & ".chrono_idref;", conn_17_05_2009_b)
                    reader_17_05_2009_b = cmd_Sql_17_05_2009_b.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader_17_05_2009_b.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_17_05_2009_b.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_17_05_2009_b.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_17_05_2009_b.Read()

                        tempo1 = reader_17_05_2009_b.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_17_05_2009_b.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_17_05_2009_b.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_17_05_2009_b.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `documents`
                    Dim reader_06_06_2009_b As MySqlDataReader
                    result1 = ""

                    reader_06_06_2009_b = Nothing
                    Dim conn_06_06_2009_b As MySqlConnection
                    conn_06_06_2009_b = New MySqlConnection(connexion)

                    conn_06_06_2009_b.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".documents;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".documents;""" & vbCrLf
                    Dim cmd_Sql_06_06_2009_b As New MySqlCommand("describe  " & Nom_de_la_base & ".documents;", conn_06_06_2009_b)
                    reader_06_06_2009_b = cmd_Sql_06_06_2009_b.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader_06_06_2009_b.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_06_06_2009_b.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_06_06_2009_b.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_06_06_2009_b.Read()

                        tempo1 = reader_06_06_2009_b.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_06_06_2009_b.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_06_06_2009_b.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_06_06_2009_b.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `elements_winform`
                    Dim reader_11_07_2009_b As MySqlDataReader
                    result1 = ""

                    reader_11_07_2009_b = Nothing
                    Dim conn_11_07_2009_b As MySqlConnection
                    conn_11_07_2009_b = New MySqlConnection(connexion)

                    conn_11_07_2009_b.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".elements_winform;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".elements_winform;""" & vbCrLf
                    Dim cmd_Sql_11_07_2009_b As New MySqlCommand("describe  " & Nom_de_la_base & ".elements_winform;", conn_11_07_2009_b)
                    reader_11_07_2009_b = cmd_Sql_11_07_2009_b.ExecuteReader()


                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "



                    tempo4 = reader_11_07_2009_b.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_11_07_2009_b.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_11_07_2009_b.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_11_07_2009_b.Read()

                        tempo1 = reader_11_07_2009_b.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_11_07_2009_b.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_11_07_2009_b.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_11_07_2009_b.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine_zero`
                    Dim reader_13_05_2010o As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010o = Nothing
                    Dim conn_13_05_2010o As MySqlConnection
                    conn_13_05_2010o = New MySqlConnection(connexion)

                    conn_13_05_2010o.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".racine_zero;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".racine_zero;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010o As New MySqlCommand("describe  " & Nom_de_la_base & ".racine_zero;", conn_13_05_2010o)
                    reader_13_05_2010o = cmd_Sql_13_05_2010o.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010o.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010o.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010o.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010o.Read()

                        tempo1 = reader_13_05_2010o.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010o.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010o.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010o.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine1`
                    Dim reader_13_05_2010p As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010p = Nothing
                    Dim conn_13_05_2010p As MySqlConnection
                    conn_13_05_2010p = New MySqlConnection(connexion)

                    conn_13_05_2010p.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine1`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine1`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010p As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine1`;", conn_13_05_2010p)
                    reader_13_05_2010p = cmd_Sql_13_05_2010p.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010p.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010p.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010p.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010p.Read()

                        tempo1 = reader_13_05_2010p.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010p.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010p.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010p.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine2`
                    Dim reader_13_05_2010q As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010q = Nothing
                    Dim conn_13_05_2010q As MySqlConnection
                    conn_13_05_2010q = New MySqlConnection(connexion)

                    conn_13_05_2010q.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine2`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine2`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010q As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine2`;", conn_13_05_2010q)
                    reader_13_05_2010q = cmd_Sql_13_05_2010q.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010q.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010q.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010q.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010q.Read()

                        tempo1 = reader_13_05_2010q.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010q.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010q.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010q.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine3`
                    Dim reader_13_05_2010r As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010r = Nothing
                    Dim conn_13_05_2010r As MySqlConnection
                    conn_13_05_2010r = New MySqlConnection(connexion)

                    conn_13_05_2010r.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine3`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine3`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010r As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine3`;", conn_13_05_2010r)
                    reader_13_05_2010r = cmd_Sql_13_05_2010r.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010r.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010r.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010r.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010r.Read()

                        tempo1 = reader_13_05_2010r.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010r.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010r.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010r.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine4`
                    Dim reader_13_05_2010s As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010s = Nothing
                    Dim conn_13_05_2010s As MySqlConnection
                    conn_13_05_2010s = New MySqlConnection(connexion)

                    conn_13_05_2010s.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine4`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine4`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010s As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine4`;", conn_13_05_2010s)
                    reader_13_05_2010s = cmd_Sql_13_05_2010s.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010s.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010s.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010s.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010s.Read()

                        tempo1 = reader_13_05_2010s.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010s.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010s.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010s.Close()



                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine5`
                    Dim reader_13_05_2010t As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010t = Nothing
                    Dim conn_13_05_2010t As MySqlConnection
                    conn_13_05_2010t = New MySqlConnection(connexion)

                    conn_13_05_2010t.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine5`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine5`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010t As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine5`;", conn_13_05_2010t)
                    reader_13_05_2010t = cmd_Sql_13_05_2010t.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010t.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010t.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010t.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010t.Read()

                        tempo1 = reader_13_05_2010t.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010t.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010t.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010t.Close()


                    My.Settings.ProgressBar_Diagnostic += 1



                    ''Description de la table `racine6`
                    Dim reader_13_05_2010u As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010u = Nothing
                    Dim conn_13_05_2010u As MySqlConnection
                    conn_13_05_2010u = New MySqlConnection(connexion)

                    conn_13_05_2010u.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine6`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine6`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010u As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine6`;", conn_13_05_2010u)
                    reader_13_05_2010u = cmd_Sql_13_05_2010u.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010u.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010u.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010u.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010u.Read()

                        tempo1 = reader_13_05_2010u.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010u.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010u.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010u.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine7`
                    Dim reader_13_05_2010v As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010v = Nothing
                    Dim conn_13_05_2010v As MySqlConnection
                    conn_13_05_2010v = New MySqlConnection(connexion)

                    conn_13_05_2010v.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine7`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine7`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010v As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine7`;", conn_13_05_2010v)
                    reader_13_05_2010v = cmd_Sql_13_05_2010v.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010v.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010v.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010v.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010v.Read()

                        tempo1 = reader_13_05_2010v.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010v.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010v.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010v.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine8`
                    Dim reader_13_05_2010w As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010w = Nothing
                    Dim conn_13_05_2010w As MySqlConnection
                    conn_13_05_2010w = New MySqlConnection(connexion)

                    conn_13_05_2010w.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine8`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine8`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010w As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine8`;", conn_13_05_2010w)
                    reader_13_05_2010w = cmd_Sql_13_05_2010w.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010w.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010w.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010w.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010w.Read()

                        tempo1 = reader_13_05_2010w.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010w.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010w.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010w.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `racine9`
                    Dim reader_13_05_2010x As MySqlDataReader
                    result1 = ""

                    reader_13_05_2010x = Nothing
                    Dim conn_13_05_2010x As MySqlConnection
                    conn_13_05_2010x = New MySqlConnection(connexion)

                    conn_13_05_2010x.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine9`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`racine9`;""" & vbCrLf
                    Dim cmd_Sql_13_05_2010x As New MySqlCommand("describe  " & Nom_de_la_base & ".`racine9`;", conn_13_05_2010x)
                    reader_13_05_2010x = cmd_Sql_13_05_2010x.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_13_05_2010x.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_13_05_2010x.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_13_05_2010x.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_13_05_2010x.Read()

                        tempo1 = reader_13_05_2010x.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_13_05_2010x.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_13_05_2010x.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_13_05_2010x.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `elements_HTML`
                    Dim reader_20_08_2010e As MySqlDataReader
                    result1 = ""

                    reader_20_08_2010e = Nothing
                    Dim conn_20_08_2010e As MySqlConnection
                    conn_20_08_2010e = New MySqlConnection(connexion)

                    conn_20_08_2010e.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`elements_HTML`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`elements_HTML`;""" & vbCrLf
                    Dim cmd_Sql_20_08_2010e As New MySqlCommand("describe  " & Nom_de_la_base & ".`elements_HTML`;", conn_20_08_2010e)
                    reader_20_08_2010e = cmd_Sql_20_08_2010e.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_20_08_2010e.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_20_08_2010e.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_20_08_2010e.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_20_08_2010e.Read()

                        tempo1 = reader_20_08_2010e.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_20_08_2010e.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_20_08_2010e.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_20_08_2010e.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `HISTORY_EVENTS`
                    Dim reader_20_08_2010f As MySqlDataReader
                    result1 = ""

                    reader_20_08_2010f = Nothing
                    Dim conn_20_08_2010f As MySqlConnection
                    conn_20_08_2010f = New MySqlConnection(connexion)

                    conn_20_08_2010f.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`HISTORY_EVENTS`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`HISTORY_EVENTS`;""" & vbCrLf
                    Dim cmd_Sql_20_08_2010f As New MySqlCommand("describe  " & Nom_de_la_base & ".`HISTORY_EVENTS`;", conn_20_08_2010f)
                    reader_20_08_2010f = cmd_Sql_20_08_2010f.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_20_08_2010f.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_20_08_2010f.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_20_08_2010f.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_20_08_2010f.Read()

                        tempo1 = reader_20_08_2010f.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_20_08_2010f.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_20_08_2010f.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_20_08_2010f.Close()


                    My.Settings.ProgressBar_Diagnostic += 1



                    ''Description de la table `Users`
                    Dim reader_20_08_2010g As MySqlDataReader
                    result1 = ""

                    reader_20_08_2010g = Nothing
                    Dim conn_20_08_2010g As MySqlConnection
                    conn_20_08_2010g = New MySqlConnection(connexion)

                    conn_20_08_2010g.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`Users`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`Users`;""" & vbCrLf
                    Dim cmd_Sql_20_08_2010g As New MySqlCommand("describe  " & Nom_de_la_base & ".`Users`;", conn_20_08_2010g)
                    reader_20_08_2010g = cmd_Sql_20_08_2010g.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_20_08_2010g.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_20_08_2010g.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_20_08_2010g.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_20_08_2010g.Read()

                        tempo1 = reader_20_08_2010g.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_20_08_2010g.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_20_08_2010g.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_20_08_2010g.Close()


                    My.Settings.ProgressBar_Diagnostic += 1


                    ''Description de la table `Questions`
                    Dim reader_05_09_2010b As MySqlDataReader
                    result1 = ""

                    reader_05_09_2010b = Nothing
                    Dim conn_05_09_2010b As MySqlConnection
                    conn_05_09_2010b = New MySqlConnection(connexion)

                    conn_05_09_2010b.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`Questions`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`Questions`;""" & vbCrLf
                    Dim cmd_Sql_05_09_2010b As New MySqlCommand("describe  " & Nom_de_la_base & ".`Questions`;", conn_05_09_2010b)
                    reader_05_09_2010b = cmd_Sql_05_09_2010b.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_05_09_2010b.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_05_09_2010b.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_05_09_2010b.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_05_09_2010b.Read()

                        tempo1 = reader_05_09_2010b.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_05_09_2010b.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_05_09_2010b.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_05_09_2010b.Close()


                    My.Settings.ProgressBar_Diagnostic += 1



                    ''Description de la table `VERSIONS_TABLES`
                    Dim reader_30_04_2011b As MySqlDataReader
                    result1 = ""

                    reader_30_04_2011b = Nothing
                    Dim conn_30_04_2011b As MySqlConnection
                    conn_30_04_2011b = New MySqlConnection(connexion)

                    conn_30_04_2011b.Open()
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`VERSIONS_TABLES`;""" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(113, langue_logiciel) & " : ""describe " & Nom_de_la_base & ".`VERSIONS_TABLES`;""" & vbCrLf
                    Dim cmd_Sql_30_04_2011b As New MySqlCommand("describe  " & Nom_de_la_base & ".`VERSIONS_TABLES`;", conn_30_04_2011b)
                    reader_30_04_2011b = cmd_Sql_30_04_2011b.ExecuteReader()

                    tempo1 = " "
                    tempo2 = " "
                    tempo3 = " "
                    tempo4 = " "
                    tempo5 = " "
                    tempo6 = " "


                    tempo4 = reader_30_04_2011b.GetName(0).ToString
                    While (tempo4.Length < 40)
                        tempo4 = tempo4 & " "
                    End While


                    tempo5 = reader_30_04_2011b.GetName(1).ToString
                    While (tempo5.Length < 25)
                        tempo5 = tempo5 & " "
                    End While


                    tempo6 = reader_30_04_2011b.GetName(2).ToString
                    While (tempo6.Length < 16)
                        tempo6 = tempo6 & " "
                    End While

                    Infos_diagnostic = Infos_diagnostic & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & " " & tempo4 & tempo5 & tempo6 & vbCrLf & " ===============================================" & vbCrLf

                    While reader_30_04_2011b.Read()

                        tempo1 = reader_30_04_2011b.GetString(0).ToString

                        While (tempo1.Length < 40)
                            tempo1 = tempo1 & " "
                        End While


                        tempo2 = reader_30_04_2011b.GetString(1).ToString

                        While (tempo2.Length < 25)
                            tempo2 = tempo2 & " "
                        End While


                        tempo3 = reader_30_04_2011b.GetString(2).ToString

                        While (tempo3.Length < 16)
                            tempo3 = tempo3 & " "
                        End While


                        result1 = "  " & LTrim(tempo1) & LTrim(tempo2) & LTrim(tempo3)
                        tempo1 = " "
                        tempo2 = " "
                        tempo3 = " "

                        Infos_diagnostic = Infos_diagnostic & result1 & vbCrLf
                        Diagnostique_Partie2 = Diagnostique_Partie2 & result1 & vbCrLf
                    End While

                    conn_30_04_2011b.Close()




                    'Le diagnostic de cette connexion est un succés
                    Infos_diagnostic = Infos_diagnostic & vbCrLf & vbCrLf & _
                                         " " & Message_utilisateur(117, langue_logiciel) & "." & vbCrLf
                    Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & vbCrLf & _
                                         " " & Message_utilisateur(117, langue_logiciel) & "." & vbCrLf

                    My.Settings.Arret_Timer = True

                    My.Settings.ProgressBar_Diagnostic += 1


                    Return True

                End If
            Next


        Catch ex As MySqlException
            'Il y a eu une erreur
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(120, langue_logiciel) & " : " + ex.Message & vbCrLf
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(120, langue_logiciel) & " : " + ex.Message & vbCrLf

            My.Settings.Arret_Timer = True

            Return False
            '
            'Valeur de Langue
            ' 0 = Allemand
            ' 1 = Anglais
            ' 2 = Espagnole
            ' 3 = Français
            ' 4 = Italien

            ' on va faire parler la synthèse vocale e.g : 'Le diagnostique de cette connexion a échoué'
            'MsgBox(lecture_du_resultat)
            'If (My.Settings.Activation_Synthese_Vocale = True) Then

            'Dim phrase_à_dire As String = ""



            ' Me.xvoice.Select(My.Settings.Index_Moteur_Synthèse_Vocale)


            'Début SAPI 5.4
            'Select Case Me.langue_logiciel

            'Case 0
            'Allemand
            'phrase_à_dire = "Der Diagnostische dieser Verbindung ist ein Misserfolg"

            '   Case 1
            'Americain Anglais
            'phrase_à_dire = "The diagnostic of this connection is a failure"

            '   Case 2
            'Espagnole
            'phrase_à_dire = "El diagnostico de esta conexion es un fracaso"

            '    Case 3
            'Français
            'phrase_à_dire = "Le diagnostique de cette connexion a aichouai"

            '   Case 4
            'Italien
            'phrase_à_dire = "Il diagnostico di questa connessione e un insuccesso"

            '   Case Else

            ' End Select



            'modification_commandes_SAPI5_TEXTE_A_DIRE(phrase_à_dire, My.Application.Info.DirectoryPath)

            'Fin SAPI 5.4

            'My.Settings.Texte_à_dire = phrase_à_dire
            'Me.Parle_Click(lecture_vocale_du_resultat, a)
            'End If


        Catch ey As System.ArgumentException
            'Il y a eu une erreur : 
            Infos_diagnostic = Infos_diagnostic & vbCrLf & " " & Message_utilisateur(120, langue_logiciel) & " : " + ey.Message & vbCrLf
            Diagnostique_Partie2 = Diagnostique_Partie2 & vbCrLf & " " & Message_utilisateur(120, langue_logiciel) & " : " + ey.Message & vbCrLf

            My.Settings.Arret_Timer = True

            Return False
            '
            'Valeur de Langue
            ' 0 = Allemand
            ' 1 = Anglais
            ' 2 = Espagnole
            ' 3 = Français
            ' 4 = Italien



            'My.Settings.Texte_à_dire = phrase_à_dire
            'xvoice.Select(My.Settings.Index_Moteur_Synthèse_Vocale)
            'Me.Parle_Click(lecture_vocale_du_resultat, a)

        Finally
            'désactivé le 10/12/2010
            'If Not reader Is Nothing Then reader.Close()

        End Try

        My.Settings.Arret_Timer = True

        Return False
    End Function

    Private Function Message_utilisateur(ByVal ID_Message As Integer, ByVal Langue As Integer)
        Dim Message As String = Nothing

        'Valeur de Langue
        ' 0 = Allemand
        ' 1 = Anglais
        ' 2 = Espagnole
        ' 3 = Français
        ' 4 = Italien

        'Valeur de ID_Message
        'Numéro de message allant de 0 à 18
        '' 0, 'ERREUR', 'Fehler im Laufe von Verbindung  Dienstprogramm : ', 'Error during connection to the server : ', 'Error durante comunicación Servidor : ', 'Erreur lors de la connexion au serveur : ', 'Errore durante connessione Calcolatore : '
        '' 1, 'ERREUR', 'Unmöglichkeit zu  schaffen  Daten Basis : ', 'Unable to create data base : ', 'Imposibilitado crear base de datos : ', 'Impossible de créer la base de donnée : ', 'Incapace creare base dati : '
        '' 2, 'ERREUR', 'Unmöglichkeit  zu  nutzen Basis : ', 'Unable to use data base : ', 'Imposibilitado usar base  : ', 'Impossible d'utiliser la base : ', 'Incapace uso base : '
        '' 3, 'ERREUR', 'Unmöglichkeit  zu Löschtaste  tabelle ', 'Unable to delete table ', 'Imposibilitado suprimir índice  ', 'Impossible de supprimer la table ', 'Incapace cancellare classifica '
        '' 4, 'ERREUR', 'Unmöglichkeit  zu schaffen  tabelle ', 'Unable to create table ', 'Imposibilitado crear índice  ', 'Impossible de créer la table ', 'Incapace creare classifica '
        '' 5, 'ERREUR', 'Unmöglichkeit  zu einfügen daten zu tabelle ', 'Unable to insert data in table ', 'Imposibilitado insertar datos índice', 'Impossible d'insérer des informations dans la table', 'Incapace inserire classifica '
        '' 6, 'INFORMATION', 'Löschtaste prozedur : ', 'Delete procedure : ', 'Suprimir procedimiento : ', 'Suppression de la procédure : ', 'Cancellare procedura : '
        '' 7, 'INFORMATION', ', ob existieren.', ', if it exist.', ', si existir.', ', si elle existe.', ', se esistere.'
        '' 8, 'ERREUR', 'Unmöglichkeit  zu Löschtaste prozedur ', 'Unable to delete procedure ', 'Imposibilitado suprimir procedimiento  ', 'Impossible de supprimer la procédure ', 'Incapace cancellare procedura '
        '' 9, 'INFORMATION', 'Einfügung prozedur ', 'Insertion of procedure ', 'Inserción procedimiento  ', 'Insertion de la procédure : ', 'Inserzione procedura : '
        '' 10, 'INFORMATION', 'Daten Basis : ', 'Data base : ', 'Base  : ', 'La base  : ', 'Base : '
        '' 11, 'INFORMATION', ' Vorhandensein schon. Willst du neu schaffen?', ' Exist already. Do you want to create it again? ', ' Existir ya. Querer crear de nuevo?', ' Existe déjà. Etes vous sûr de vouloir la recréer?', ' Esistere già. desiderare ricreare ancora?'
        '' 12, 'TITRE_AVERTISSEMENT', 'Achtung!', 'Warning!', 'Atención!', 'Attention !', 'Attento!'
        '' 13, 'INFORMATION', 'Kreation   Daten Basis : ', 'Creation of data base : ', 'Creación  base de datos : ', 'Création de la base : ', 'Creazione base dati : '
        '' 14, 'INFORMATION', 'Nutzung  Daten Basis : ', 'Use of data base : ', 'Usar base de datos :', 'Utilisation de la base : ', 'Usare  base dati : '
        '' 15, 'INFORMATION', 'Löschtaste  tabelle : ', 'Delete table : ', 'Suprimir índice  : ', 'Suppression de la table :  ', 'Cancellare classifica : '
        '' 16, 'INFORMATION', 'Kreation  tabelle  : ', 'Creation of table : ', 'Creación  índice : ', 'Création de la table : ', 'Creazione classifica  : '
        '' 17, 'INFORMATION', 'Einfügung  tabelle  : ', 'Insertion in table :  ', 'Inserción  índice : ', 'Insertion dans la table : ', 'Inserzione classifica : '
        '' 18, 'INFORMATION', 'Informationen Versäumnisurteil.', ' Informations by default.', ' Información a falta.', ' d'informations par défaut.', ' Notifica in contumacia.'
        '' 19, 'INFORMATION', 'Löschtaste Daten Basis : ', 'Delete database : ', 'Suprimir base de datos : ', 'Suppression de la base : ', 'Cancellare  base dati : '
        '' 20, 'ERREUR', 'Unmöglichkeit  zu Löschtaste Daten Basis : ', 'Unable to delete database : ', 'Imposibilitado suprimir base de datos : ', 'Impossible de supprimer la base : ', 'Incapace cancellare base dati : '
        '' 21, 'INFORMATION', 'Löschtaste Verbindung : ', 'Do you really want to delete connection : ', 'Suprimir conexión : ', 'Êtes vous sûr de vouloir supprimer la connexion : ', 'Cancellare connessione : '
        '' 22, 'INFORMATION', 'Zur Information', 'Information', 'Información', 'Information', 'Informazione'
        '' 23, 'INFORMATION', 'Verbindung ', 'Connection ', 'Conexión ', 'Connexion ', 'Connessione '
        '' 24, 'INFORMATION', 'löschtaste', 'deleted', 'suprimir', 'supprimée', 'cancellare'
        '' 25, 'INFORMATION', 'Titel verbindung : ', 'The title of your connection has : ', 'Título Conexión : ', 'Le nom de votre connexion à une longueur de :', 'Intitolare connessione  : '
        '' 26, 'INFORMATION', 'charakters', 'characters', 'tipo.', 'caractères.', 'carattere.'
        '' 27, 'INFORMATION', '3 charakters < Titel verbindung  < 29 charakters', '3 characters < Title of your connexion < 29 characters', '3 tipo < Título Conexión < 29 tipo', 'Le titre de la connexion doit avoir entre 3 et 29 caractères.', '3 carattere < Intitolare connessione  < 29 carattere'
        '' 28, 'INFORMATION', 'Falsch IP adresse formen', 'Bad format IP address', 'Podrido formato dirigir IP', 'L'adresse IP ou le nom d'hote est trop court (e.g : 2.5.1.9 7 ou www.mon_serveur_mysql.fr).', 'Marcio formato indirizzare IP '
        '' 29, 'INFORMATION', 'Benutzer Name zu klein (3-29 Charakters)', 'User name too small (3-29 characters)', 'Usuario denominación también pequeño (3-29 tipo)', 'Le nom de l'utilisateur doit avoir entre 3 et 29 caractères.', 'Utente nome troppo piccolo (3-29 carattere)'
        '' 30, 'INFORMATION', 'Paßwort zu klein (4 Charakters Minimum)', 'Password too small (4 characters minimum)', 'Contraseña también pequeño (4 tipo mínimo)', 'Le mot de passe est trop court (4 caractères minimum).', 'Parola d'ordine troppo piccolo (4 Carattere minimo)'
        '' 31, 'INFORMATION', 'Konfirmieren paßwort zu klein (4 Charakters Minimum)', 'Confirmed password too small (4 characters minimum)', 'Contraseña también pequeño (4 tipo mínimo)', 'Le mot de passe confirmé est trop court (4 caractères minimum).', 'Parola d'ordine corroborare troppo piccolo (4 Carattere minimo)'
        '' 32, 'INFORMATION', 'Konfirmieren paßwort nicht gleichgestellter paßwort ', 'Confirmed password not equal password', 'Contraseña también no igual contraseña también ', 'La confirmation du mot de passe est différente du mot de passe.', 'Parola d'ordine corroborare non uguale parola d'ordine'
        '' 33, 'INFORMATION', 'Daten Basis name länge : ', 'Data base name length :', 'Base denominación largo : ', 'Le nom de votre base à une longueur de : ', 'Base nome lunghezza : '
        '' 34, 'INFORMATION', ' charakters.', ' characters.', ' tipo.', ' caractères.', ' carattere.'
        '' 35, 'INFORMATION', '3  charakters < Daten Basis name länge < 29  charakters', '3  characters < Data base name length< 29  characters', '3 tipo < Base denominación < 29 tipo', 'Le nom de la base doit avoir entre 3 et 29 caractères.', '3 carattere < Base nome < 29 carattere'
        '' 36, 'INFORMATION', 'Weg elektrisches schema ist zu klein.', 'Path for electrical diagram is too short.', 'Curso eléctrico plan también pequeño.', 'Le nom du chemin des schémas est trop court.', 'Corso elettrico schema troppo piccolo.'
        '' 37, 'INFORMATION', '(4 charakters minimum  e.g : c:\SCHEMAS oder http://www.site.fr/data/schemas)', '(4 characters minimum  e.g : c:\SCHEMAS or http://www.site.fr/data/schemas)', '(4 tipo mínimo  e.g : c:\SCHEMAS, http://www.site.fr/data/schemas)', '(4 caractères minimum e.g : c:\SCHEMAS ou http://www.site.fr/data/schemas)', '(4 Carattere minimo  e.g : c:\SCHEMAS, http://www.site.fr/data/schemas)'
        '' 38, 'INFORMATION', 'Weg synopsis schema ist zu klein.', 'Path for synopsis diagram  is too short.', 'Curso sinopsis plan también pequeño.', 'Le nom du chemin des synoptiques est trop court.', 'Corso sinossi schema troppo piccolo.'
        '' 39, 'INFORMATION', '(4 charakters minimum  e.g : c:\SYNOPTIQUES oder  http://www.site.fr/data/synoptiques)', '(4 characters minimum e.g : c:\SYNOPTIQUES or  http://www.site.fr/data/synoptiques)', '(4 tipo mínimo e.g : c:\SYNOPTIQUES ,  http://www.site.fr/data/synoptiques)', '(4 caractères minimum e.g : c:\SYNOPTIQUES ou  http://www.site.fr/data/synoptiques)', '(4 Carattere minimo  e.g : c:\SYNOPTIQUES oder  http://www.site.fr/data/synoptiques)'
        '' 40, 'INFORMATION', 'Weg vermerk ist zu klein.', 'Path for instructions  is too short.', 'Curso conciso también pequeño.', 'Le  nom du chemin des notices est trop court.', 'Corso conciso troppo piccolo.'
        '' 41, 'INFORMATION', '(4 charakters minimum  e.g : c:\NOTICES oder http://www.site.fr/data/notices)', '(4 characters minimum e.g : c:\NOTICES or http://www.site.fr/data/notices)', '(4 tipo mínimo e.g : c:\NOTICES, http://www.site.fr/data/notices)', '(4 caractères minimum e.g : c:\NOTICES ou http://www.site.fr/data/notices)', '(4 Carattere minimo e.g : c:\NOTICES, http://www.site.fr/data/notices)'
        '' 42, 'INFORMATION', 'Weg beschreibung produzieren ist zu klein.', 'Path for product description  is too short.', 'Curso hoja producir también pequeño.', 'Le  nom du chemin des fiches produits est trop court.', 'Corso foglio prodotto troppo piccolo.'
        '' 43, 'INFORMATION', '(4 charakters minimum  e.g : c:\FICHES_PRODUITS oder  http://www.site.fr/data/fiches_produits)', '(4 characters minimum e.g : c:\FICHES_PRODUITS or  http://www.site.fr/data/fiches_produits)', '(4 tipo mínimo e.g : c:\FICHES_PRODUITS,  http://www.site.fr/data/fiches_produits)', '(4 caractères minimum e.g : c:\FICHES_PRODUITS ou http://www.site.fr/data/fiches_produits)', '(4 Carattere minimo e.g : c:\FICHES_PRODUITS,  http://www.site.fr/data/fiches_produits)'
        '' 44, 'INFORMATION', 'Weg photographie  ist zu klein.', 'Path for picture  is too short.', 'Curso foto también pequeño.', 'Le  nom du chemin des photos des produits est trop court.', 'Corso foto troppo piccolo.'
        '' 45, 'INFORMATION', '(4 charakters minimum  e.g : c:\PHOTOS_PRODUITS oder http://www.site.fr/data/Photos_produits)', '(4 characters minimum  e.g : c:\PHOTOS_PRODUITS or http://www.site.fr/data/Photos_produits)', '(4 tipo mínimo  e.g : c:\PHOTOS_PRODUITS,  http://www.site.fr/data/Photos_produits)', '(4 caractères minimum e.g : c:\PHOTOS_PRODUITS ou http://www.site.fr/data/Photos_produits)', '(4 Carattere minimo  e.g : c:\PHOTOS_PRODUITS, http://www.site.fr/data/Photos_produits)'

        '' 46, 'INFORMATION', 'Deutsch Beschreibung', 'German Description', 'Descripción Alemán', 'Description Allemande', 'Descrizione  Tedesco'
        '' 47, 'INFORMATION', 'Engländer Beschreibung', 'English Description', 'Descripción Ingleses', 'Description Anglaise', 'Descrizione Inglese'
        '' 48, 'INFORMATION', 'Spanisch Beschreibung', 'Spanish Description', 'Descripción Españoles', 'Description Espagnole', 'Descrizione Spagnolo'
        '' 49, 'INFORMATION', 'Franzose Beschreibung', 'French Description', 'Descripción Franceses', 'Description Française', 'Descrizione Francese'
        '' 50, 'INFORMATION', 'Italiener Beschreibung', 'Italian Description', 'Descripción Italianos', 'Description Italienne', 'Descrizione Italiano'
        '' 51, 'INFORMATION', 'Euer', 'Your ', 'Tu', 'Votre', 'Tua'
        '' 52, 'INFORMATION', '3 charakters < beschreibung verbindung  < 29 charakters', '3 characters < Connection description < 29 characters', '3 tipo < Conexión descripción < 29 tipo', 'La description de la connexion doit avoir entre 3 et 29 caractères.', '3 carattere < Connessione descrizione  < 29 carattere'

        Dim Table1(,) As String = { _
        {"Fehler im Laufe von Verbindung  Dienstprogramm : ", "Error during connection to the server : ", "Error durante comunicación Servidor : ", "Erreur lors de la connexion au serveur : ", "Errore durante connessione Calcolatore : "}, _
        {"Unmöglichkeit zu  schaffen  Daten Basis : ", "Unable to create data base : ", "Imposibilitado crear base de datos : ", "Impossible de créer la base de donnée : ", "Incapace creare base dati : "}, _
        {"Unmöglichkeit  zu  nutzen Basis : ", "Unable to use data base : ", "Imposibilitado usar base  : ", "Impossible d'utiliser la base : ", "Incapace uso base : "}, _
        {"Unmöglichkeit  zu Löschtaste  tabelle ", "Unable to delete table ", "Imposibilitado suprimir índice  ", "Impossible de supprimer la table ", "Incapace cancellare classifica "}, _
        {"Unmöglichkeit  zu schaffen  tabelle ", "Unable to create table ", "Imposibilitado crear índice  ", "Impossible de créer la table ", "Incapace creare classifica "}, _
        {"Unmöglichkeit  zu einfügen daten zu tabelle ", "Unable to insert data in table ", "Imposibilitado insertar datos índice", "Impossible d'insérer des informations dans la table", "Incapace inserire classifica "}, _
        {"Löschtaste prozedur : ", "Delete procedure : ", "Suprimir procedimiento : ", "Suppression de la procédure : ", "Cancellare procedura : "}, _
        {", ob existieren.", ", if it exist.", ", si existir.", ", si elle existe.", ", se esiste."}, _
        {"Unmöglichkeit  zu Löschtaste prozedur ", "Unable to delete procedure ", "Imposibilitado suprimir procedimiento  ", "Impossible de supprimer la procédure ", "Incapace cancellare procedura "}, _
        {"Einfügung prozedur ", "Insertion of procedure ", "Inserción procedimiento  ", "Insertion de la procédure : ", "Inserzione procedura : "}, _
        {"Daten Basis : ", "Data base : ", "Base  : ", "La base  : ", "Base : "}, _
        {" Vorhandensein schon. Willst du neu schaffen?", " Exist already. Do you want to create it again? ", " Existir ya. Querer crear de nuevo?", " Existe déjà. Etes vous sûr de vouloir la recréer?", " già esistente. Siete sicuri di volerlo  ricreare?"}, _
        {"Vorsicht!", "Warning!", "Atención!", "Attention !", "Attento!"}, _
        {"Kreation   Daten Basis : ", "Creation of data base : ", "Creación  base de datos : ", "Création de la base : ", "Creazione base dati : "}, _
        {"Nutzung  Daten Basis : ", "Use of data base : ", "Usar base de datos :", "Utilisation de la base : ", "Usare  base dati : "}, _
        {"Löschtaste  tabelle : ", "Delete table : ", "Suprimir índice  : ", "Suppression de la table :  ", "Cancellare classifica : "}, _
        {"Kreation  tabelle  : ", "Creation of table : ", "Creación  índice : ", "Création de la table : ", "Creazione classifica  : "}, _
        {"Einfügung  tabelle  : ", "Insertion in table :  ", "Inserción  índice : ", "Insertion dans la table : ", "Inserzione classifica : "}, _
        {"Informationen Versäumnisurteil.", " Informations by default.", " Información a falta.", " d'informations par défaut.", " Notifica per diffetto."}, _
        {"Löschtaste Daten Basis : ", "Delete database : ", "Suprimir base de datos : ", "Suppression de la base : ", "Cancellare  base dati : "}, _
        {"Unmöglichkeit  zu Löschtaste Daten Basis : ", "Unable to delete database : ", "Imposibilitado suprimir base de datos : ", "Impossible de supprimer la base : ", "Incapace cancellare base dati : "}, _
        {"Löschtaste Verbindung : ", "Do you really want to delete connection : ", "Suprimir conexión : ", "Êtes vous sûr de vouloir supprimer la connexion : ", "Cancellare connessione : "}, _
        {"Zur Information", "Information", "Información", "Information", "Informazione"}, _
        {"Verbindung ", "Connection ", "Conexión ", "Connexion ", "Connessione "}, _
        {"löschtaste", "deleted", "suprimir", "supprimée", "cancellare"}, _
        {"Titel verbindung : ", "The title of your connection has : ", "Título Conexión : ", "Le nom de votre connexion à une longueur de : ", "Intitolare connessione  : "}, _
        {" charakters.", " characters.", " tipo.", " caractères.", " carattere."}, _
        {"3 charakters < Titel verbindung  < 29 charakters", "3 characters < Title of your connexion < 29 characters", "3 tipo < Título Conexión < 29 tipo", "Le titre de la connexion doit avoir entre 3 et 29 caractères.", "3 carattere < Intitolare connessione  < 29 carattere"}, _
        {"Falsch IP adresse formen", "Bad format IP address", "Podrido formato dirigir IP", "L'adresse IP ou le nom d'hote est trop court (e.g : 2.5.1.9 7 ou www.mon_serveur_mysql.fr).", "Marcio formato indirizzare IP "}, _
        {"Benutzer Name zu klein (3-29 Charakters)", "User name too small (3-29 characters)", "Usuario denominación también pequeño (3-29 tipo)", "Le nom de l'utilisateur doit avoir entre 3 et 29 caractères.", "Utente nome troppo piccolo (3-29 carattere)"}, _
        {"Paßwort zu klein (4 Charakters Minimum)", "Password too small (4 characters minimum)", "Contraseña también pequeño (4 tipo mínimo)", "Le mot de passe est trop court (4 caractères minimum).", "Parola d'ordine troppo piccolo (4 Carattere minimo)"}, _
        {"Konfirmieren paßwort zu klein (4 Charakters Minimum)", "Confirmed password too small (4 characters minimum)", "Contraseña también pequeño (4 tipo mínimo)", "Le mot de passe confirmé est trop court (4 caractères minimum).", "Parola d'ordine corroborare troppo piccolo (4 Carattere minimo)"}, _
        {"Konfirmieren paßwort nicht gleichgestellter paßwort.", "Confirmed password not equal password.", "Contraseña también no igual contraseña también.", "La confirmation du mot de passe est différente du mot de passe.", "Parola d'ordine corroborare non uguale parola d'ordine."}, _
        {"Daten Basis name länge : ", "Data base name length : ", "Base denominación largo : ", "Le nom de votre base à une longueur de : ", "Base nome lunghezza : "}, _
        {" charakters.", " characters.", " tipo.", " caractères.", " carattere."}, _
        {"3  charakters < Länge von dem Namen ihres Datenbasis < 29  charakters", "3  characters < Data base name length < 29  characters", "3 tipo < Base denominación largo < 29 tipo", "Le nom de la base doit avoir entre 3 et 29 caractères.", "3 carattere < Base nome < 29 carattere"}, _
        {"Der Weg des elektrischen Schemas ist zu klein.", "Path for electrical diagram is too short.", "Curso eléctrico plan también pequeño.", "Le nom du chemin des schémas est trop court.", "Corso elettrico schema troppo piccolo."}, _
        {"(4 charakters minimum e.g : c:\SCHEMAS oder http://www.site.fr/data/schemas)", "(4 characters minimum e.g : c:\SCHEMAS or http://www.site.fr/data/schemas)", "(4 tipo mínimo  e.g : c:\SCHEMAS, http://www.site.fr/data/schemas)", "(4 caractères minimum e.g : c:\SCHEMAS ou http://www.site.fr/data/schemas)", "(4 Carattere minimo  e.g : c:\SCHEMAS, http://www.site.fr/data/schemas)"}, _
        {"Der Weg des synoptischen Schema ist zu klein.", "Path for synopsis diagram  is too short.", "Curso sinopsis plan también pequeño.", "Le nom du chemin des synoptiques est trop court.", "Corso sinossi schema troppo piccolo."}, _
        {"(4 charakters minimum  e.g : c:\SYNOPTIQUES oder  http://www.site.fr/data/synoptiques)", "(4 characters minimum e.g : c:\SYNOPTIQUES or  http://www.site.fr/data/synoptiques)", "(4 tipo mínimo e.g : c:\SYNOPTIQUES ,  http://www.site.fr/data/synoptiques)", "(4 caractères minimum e.g : c:\SYNOPTIQUES ou  http://www.site.fr/data/synoptiques)", "(4 Carattere minimo  e.g : c:\SYNOPTIQUES,  http://www.site.fr/data/synoptiques)"}, _
        {"Weg vermerk ist zu klein.", "Path for instructions  is too short.", "Curso conciso también pequeño.", "Le  nom du chemin des notices est trop court.", "Corso conciso troppo piccolo."}, _
        {"(4 charakters minimum  e.g : c:\NOTICES oder http://www.site.fr/data/notices)", "(4 characters minimum e.g : c:\NOTICES or http://www.site.fr/data/notices)", "(4 tipo mínimo e.g : c:\NOTICES, http://www.site.fr/data/notices)", "(4 caractères minimum e.g : c:\NOTICES ou http://www.site.fr/data/notices)", "(4 Carattere minimo e.g : c:\NOTICES, http://www.site.fr/data/notices)"}, _
        {"Der Weg von dem Beschreibungprodukt ist zu klein.", "Path for product description  is too short.", "Curso hoja producir también pequeño.", "Le  nom du chemin des fiches produits est trop court.", "Corso foglio prodotto troppo piccolo."}, _
        {"(4 charakters minimum  e.g : c:\FICHES_PRODUITS oder  http://www.site.fr/data/fiches_produits)", "(4 characters minimum e.g : c:\FICHES_PRODUITS or  http://www.site.fr/data/fiches_produits)", "(4 tipo mínimo e.g : c:\FICHES_PRODUITS,  http://www.site.fr/data/fiches_produits)", "(4 caractères minimum e.g : c:\FICHES_PRODUITS ou http://www.site.fr/data/fiches_produits)", "(4 Carattere minimo e.g : c:\FICHES_PRODUITS,  http://www.site.fr/data/fiches_produits)"}, _
        {"Der Weg von dem Foto ist zu klein.", "Path for picture  is too short.", "Curso foto también pequeño.", "Le  nom du chemin des photos des produits est trop court.", "Corso foto troppo piccolo."}, _
        {"(4 charakters minimum  e.g : c:\PHOTOS_PRODUITS oder http://www.site.fr/data/Photos_produits)", "(4 characters minimum  e.g : c:\PHOTOS_PRODUITS or http://www.site.fr/data/Photos_produits)", "(4 tipo mínimo  e.g : c:\PHOTOS_PRODUITS,  http://www.site.fr/data/Photos_produits)", "(4 caractères minimum e.g : c:\PHOTOS_PRODUITS ou http://www.site.fr/data/Photos_produits)", "(4 Carattere minimo  e.g : c:\PHOTOS_PRODUITS, http://www.site.fr/data/Photos_produits)"}, _
        {"Deutsche Beschreibung", "German Description", "German Description", "Description Allemande", "Descrizione  Tedesco"}, _
        {"Englische Beschreibung", "English Description", "Descripción Ingleses", "Description Anglaise", "Descrizione Inglese"}, _
        {"Spanische Beschreibung", "Spanish Description", "Descripción Españoles", "Description Espagnole", "Descrizione Spagnolo"}, _
        {"Französische Beschreibung", "French Description", "Descripción Franceses", "Description Française", "Descrizione Francese"}, _
        {"Italienische Beschreibung", "Italian Description", "Descripción Italianos", "Description Italienne", "Descrizione Italiano"}, _
        {"Ihre", "Your ", "Tu ", "Votre ", "Tua "}, _
        {"3 charakters < Verbindungsbeschreibung < 29 charakters", "3 characters < Connection description < 29 characters", "3 tipo < Conexión descripción < 29 tipo", "La description de la connexion doit avoir entre 3 et 29 caractères.", "3 carattere < Connessione descrizione  < 29 carattere"}, _
        {"Verfügbar", "Available", "Disponible", "Disponible", "Disponibile"}, _
        {"Unbekannte Datenbasis", "Unknown Database", "Base de datos desconocida", "Base de données inconnue", "Banca dati sconosciuta"}, _
        {"Nicht verfügbar", "Unavailable", "No disponible", "Non disponible", "Non disponibile"}, _
        {"Name des executable", "Name of the executable", "Nombre del executable", "Nom de l'executable", "Nome dell'eseguibile"}, _
        {"Einrichtungsweg", "Path of installation", "Camino de instalación", "Chemin d'installation", "Strada di installazione"}, _
        {"Menge körperlichen Gedächtnisses die im Zusammenhang des Prozesses", "Quantity of physical memory map in the context of the process", "Cantidad de memoria física al contexto del proceso", "Quantité de mémoire physique mappée au contexte du processus", "Quantità di memoria fisico mappée al contesto"}, _
        {"Stecker MySql", "Connector MySql", "Conectador MySql", "Connecteur MySql", "Connettore MySql"}, _
        {"Name der Datei", "File Name", "Nombre del fichero", "Nom du fichier", "Nome dello schedario"}, _
        {"Name der Gesellschaft", "Name of the company", "Nombre de la sociedad", "Nom de la société", "Nome della società"}, _
        {"Beschreibung der Datei", "File Description", "Descripción del fichero", "Description du fichier", "Descrizione dello schedario"}, _
        {"Version der Datei", "File Version", "Versión del fichero", "Version du fichier", "Versione dello schedario"}, _
        {"Sprache der Datei", "File Language", "Lengua del fichero", "Langue du fichier", "Lingua dello schedario"}, _
        {"Name des Produktes der Datei", "File Product Name", "Nombre del producto del fichero", "Nom du produit du fichier", "Nome del prodotto dello schedario"}, _
        {"Version des Produktes", "File Product Version", "Versión del Producto", "Version du Produit", "Versione del Prodotto"}, _
        {"Vollständiger Name des Betriebssystems", "Complete name of the operating system", "Nombre completo del sistema operativo", "Nom complet du système d'exploitation", "Nome completo del sistema di sfruttamento"}, _
        {"Platforme Computer", "Computer Platforme", "Platforme ordenador", "Platforme ordinateur", "Platforme computer"}, _
        {"Kultur", "Installed Culture", "Cultura ", "Culture installlée", "Cultura "}, _
        {"Informationen Prozessor", "Information Processor", "Informaciones Procesador", "Informations Processeur", "Notizie Processore"}, _
        {"Informationen Karte Mutter", "Information  mother Card", "Informaciones Tarjeta madre", "Informations Carte mère", "Notizie piastra principale"}, _
        {"Erbauer", "Manufacturer", "Constructor", "Constructeur", "Costruttore"}, _
        {"Name", "Name", "Nombre", "Nom", "Nome"}, _
        {"Beschreibung", "Description", "Descripción", "Description", "Descrizione"}, _
        {"Totales körperliches Gedächtnis", "Total physical memory", "Memoria física total", "Mémoire physique totale", "Memoria fisica totale"}, _
        {"Verfügbares körperliches Gedächtnis", "Available physical memory", "Memoria física disponible", "Mémoire physique disponible", "Memoria fisica disponibile"}, _
        {"Totales virtuelles Gedächtnis", "Total virtual memory", "Memoria virtual total", "Mémoire virtuelle  totale", "Memoria virtuale totale"}, _
        {"Verfügbares virtuelles Gedächtnis", "Available virtual memory", "Memoria virtual disponible", "Mémoire virtuelle  disponible", "Memoria virtuale disponibile"}, _
        {"Vollständiger Name des Benutzers der aktuellen Session", "Complete name of the user of the current session", "Nombre completo del usuario de la sesión corriente", "Nom complet de l'utilisateur de la session en cours", "Nome completo dell'utente della sessione in corso"}, _
        {"Dieser Benutzer gehört zu den Gruppen", "This user belongs to the groups", "Este usuario pertenece a los grupos", "Cet utilisateur appartient aux groupes", "Questo utente appartiene ai gruppi"}, _
        {"Kontooperateur", "Operator of account", "Operador de cuenta", "Opérateur de compte", "Operatore di conto"}, _
        {"Ja", "Yes", "Sí", "Oui", "Sì"}, _
        {"Nein", "No", "No", "Non", "No"}, _
        {"Verwalter", "Administrator", "Administrador", "Administrateur", "Amministratore"}, _
        {"Schutzoperateur", "Operator of saving", "Operador de salvaguardia", "Opérateur de sauvegarde", "Operatore di salvaguardia"}, _
        {"Gast", "Guest", "Invitado", "Invité", "Invitato"}, _
        {"Benutzer, zu können", "User to be able to", "Usuario poder", "Utilisateur avec pouvoir", "Utente con potere"}, _
        {"Druckoperateur", "Operator of printing", "Operador de impresión", "Opérateur d'impression", "Operatore di impressione"}, _
        {"Replicator", "Replicator", "Replicator", "Réplicateur", "Replicator"}, _
        {"Operateur System", "Operator system", "Operador sistema", "Opérateur système", "Operatore sistema"}, _
        {"Benutzer", "User", "Usuario", "Utilisateur", "Utente"}, _
        {"Recycling profiliert, des firewall örtlich", "Recovery of profiles local of the firewall", "Recuperación de perfila local del cortafuego", "Récupération du profile locale du firewall", "Recupero dello profilo locale del firewall"}, _
        {"Diagnostik der Verbindung", "Diagnosis of the connection", "Diagnóstico de la conexión", "Diagnostic de la connexion", "Diagnosi della connessione"}, _
        {"Der Computer ist mit einem örtlichen und / oder Internet-Netz verbunden", "The computer is connected with a local area network and\or Internet", "El ordenador es unido a una red de área local y\o Internet", "L'ordinateur est relié à un réseau local et/ou Internet", "Il computer è collegato locale et/ou Internet ad una rete"}, _
        {"Der Computer ist mit einem örtlichen und / oder Internet-Netz nicht verbunden", "The computer is not connected with a local area network and\or Internet", "El ordenador no es unido a una red de área local y\o Internet", "L'ordinateur n'est pas relié à un réseau local et/ou Internet", "Il computer non è collegato locale et/ou Internet ad una rete"}, _
        {"Anzahl von der Netzwerkkarte", "Number of Network(s) card(s)", "Número de tarjeta Red", "Nombre de carte(s) Réseau(x) ", "Numero di carta Rete"}, _
        {"Netzwerkkarte", "Network card", "Tarjeta Red", "Carte Réseau", "Carta Rete"}, _
        {"Beschreibung der Karte", "Description of the card", "Descripción de la tarjeta", "Description de la carte", "Descrizione della carta"}, _
        {"Typ von Interface", "Type of interface", "Tipo de interfaz", "Type d'interface", "Tipo di interfaccia"}, _
        {"Geschwindigkeit", "Speed", "Velocidad", "Vitesse", "Velocità"}, _
        {"Status", "Status", "Estatuto", "Etat", "Statuto"}, _
        {"Kein verfügbare Information IPv4 für dieses Interface", "No available information IPv4 for this interface", "Ninguna información IPv4 disponible para esta interfaz", "Pas d'information IPv4 disponible pour cette interface", "Non di notizia IPv4 disponibile per questa interfaccia"}, _
        {"Fußgängerbrücke", "Gateway", "Pasarela", "Passerelle", "Passerella"}, _
        {"Datensendung IPv4", "Emission of data IPv4", "Emisión de dato IPv4", "Emission de données IPv4", "Emissione del dati IPv4"}, _
        {"Datenempfang IPv4", "Reception of data IPv4", "Recepción de dato IPv4", "Réception de données IPv4", "Ricevimento del dati IPv4"}, _
        {"Name Ihres Computers", "Name of your computer", "Nombre de su ordenador", "Nom de votre ordinateur", "Nome del vostro computer"}, _
        {"Schicke IP dieser Maschine", "IP of this machine", "IP de esta máquina", "Adresse(s) IP de cette machine", " IP di questa macchina"}, _
        {"Ping des entfernten Servers", "Ping of the Remote server", "Ping del Servidor distante", "Ping du Serveur distant", "Ping del Server distante"}, _
        {"Aktuelle Suche", "Search at present", "Investigación actualmente", "Recherche en cours", "ricerca attualmente"}, _
        {"Kette von Verbindung mit dem Server", "Chain of connection to the server", "Cadena de conexión al servidor", "Chaîne de connexion au serveur", "Catena di connessione al server"}, _
        {"Version der Software Server MySql", "Version of the server software MySql", "Versión del software Camarero MySql", "Version du logiciel Serveur MySql", "Versione del di software Cameriere MySql"}, _
        {"Visualisation von angeschlossenen Benutzern", "Display of the connected users", "Visualización de los usuarios conectados", "Visualisation des utilisateurs connectés", "Visualizzazione degli utenti connessi"}, _
        {"Antrag", "Request", "Demanda", "Requête", "Richiesta"}, _
        {"Gefundene Datenbasis", "Found database(s)", "Base de datos encontrada", "Base(s) présente(nt)", "Banca dati trovata"}, _
        {"Existiere auf diesem Server gut", "Exist on this server", "Existe sobre este servidor", "existe bien sur ce serveur", "esisti su questo server"}, _
        {"Registrierung", "Recording", "Registro", "enregistrement", "registrazione"}, _
        {"Das Diagnostische dieser Verbindung hat fertiggebracht", "The diagnostic of this connection made a success", "Diagnostico de esta conexion consiguio", "Le diagnostic de cette connexion est un succés", "Il diagnostico di questa connessione e riuscito"}, _
        {"existiere nicht auf diesem Server", "Exist not on this server", "Existe no sobre este servidor", "n'existe pas sur ce serveur", "non esistere su questo server"}, _
        {"Die Verbindung beseitigen, sie rekonstruieren, das Feld anstreichend, leere Basis schaffen.", "Delete the connection, recreate it by marking the compartment create empty base.", "Suprimir la conexión, recrearlo punteando el compartimiento crear base vacía.", "Supprimer la connexion, la recréer en cochant la case créer base vide.", "Sopprimere la connessione, ricreare segnandola la capanna creare base vuoto."}, _
        {"Es hat einen Irrtum gegeben", "There was an error", "Hubo un error", "Il y a eu une erreur", "C'è stato un errore"}, _
        {"Die Diagnostik dieser Verbindung ist gescheitert", "The diagnosis of this connection failed", "El diagnóstico de esta conexión fue suspendido", "Le diagnostic de cette connexion a échoué", "La diagnosi di questa connessione è fallita"}, _
        {"Stärken sich Sie der Text der Hilfe bitte?", "Would you like to restore the text of the help?", "¿ Por favor, alimente el texto de la ayuda?", "Voulez vous restaurer le texte de l'aide?", "Volete restaurare il testo dell'aiuto?"}, _
        {"Anschlag Hilft", "Display Helps", "Fijación Ayuda", "Affichage Aide", "Affissione Aiuta"}, _
        {"Sie sollen einen Kode von Dokument tippen, um eine Suche anzufangen.", "You have to type a code of document to begin research.", "Usted debe golpear un código de documento para empezar una búsqueda.", "Vous devez taper un code de document pour débuter une recherche.", "Dovete battere un codice di documento per esordire una ricerca."}, _
        {"Sie sollen Ihre Auswahl machen.", "You have to make your selection.", "Usted debe hacer su selección.", "Vous devez faire votre sélection.", "Dovete fare la vostra selezione."}, _
        {"Es gibt keinen gegenwärtig angezeigten Bericht", "There is no report shown at present", "No hay informe fijado actualmente", "Il n'y a pas de rapport affiché actuellement", "Non c'è di rapporto affisso attualmente"}, _
        {"Der Name des Berichtes für das Retten ist nicht informiert", "The name of the Report to be saved is not informed", "El nombre del Informe que hay que salvar no es informado", "le nom du Rapport à sauver n'est pas renseigné", "il nome del Rapporto a salvare non è informato"}, _
        {"Der Name des Berichtes für das Retten enthält mehr als 20 Zeichen.", "The name of the Report to be saved contains more than 20 characters.", "El nombre del Informe que hay que salvar contiene más de 20 carácteres.", "Le nom du Rapport à sauver contient plus de 20 caractères.", "Il nome del Rapporto a salvare contiene più di 20 caratteri."}, _
        {"Ihr Beziehungsname enthält :", "Your name of report contains: ", "Su nombre de informe contiene: ", "Votre nom de rapport contient : ", "Il vostro nome di rapporto contiene : "}, _
        {"Der deutliche Bericht gegenwärtig ist schon gerettet.", "The visible Report at present is already saved.", "El Informe visible actualmente es ya salvado.", "Le Rapport visible actuellement est déjà sauvé.", "Il Rapporto visibile è salvato attualmente già."}, _
        {"Gesicherter Bericht", "Protected report", "Informe Salvaguardado", "Rapport Sauvegardé", "Rapporto Salvaguardato"}, _
        {"Ein Irrtum hat sich ereignet.", "An error occurred.", "Un error se produjo.", "Une erreur s'est produite.", "Un errore si è prodursi."}, _
        {"Es gibt ein Auswahlproblem von Drucker.", "There is a problem of selection of printer.", "Hay un problema de selección de impresora.", "il y a un problème de sélection d'imprimante.", "c'è un problema di selezione di stampante."}, _
        {"Unmöglich, den Bericht zu laden.", "Impossible to load the report.", "Imposible cargar el informe.", "Impossible de charger le rapport.", "Impossibile di incaricare il rapporto."}, _
        {"Wünschen Sie, sich der Bericht umzubringen?", "Wish to commit suicide the report?", "¿ Desee suicidarse el informe?", "Désirez vous supprimer le rapport?", "Desiderate sopprimere il rapporto?"}, _
        {"Beseitigter Bericht", "Deleted report", "Informe suprimido", "Rapport supprimé", "Rapporto soppresso"}, _
        {"Unmöglich, den Bericht zu beseitigen.", "Impossible to delete  the Report.", "Imposible suprimir el Informe.", "Impossible de supprimer le Rapport.", "Impossibile di sopprimere il Rapporto."}, _
        {"Server", "Server", "Servidor", "Serveur", "Server"}, _
        {"Statistik von Typ von Papieren", "Statistics by Document type", "Estadística por Tipo de documentos", "Statistique par Type de documents", "Statistico per Tipo di documenti"}, _
        {"Papiere", "Documents", "Documentos", "Documents", "Documenti"}, _
        {"Statistik von Kode Registrierung", "Statistics by Code recording", "Estadística por Código registro", "Statistique par Code enregistrement", "Statistico per Codice registrazione"}, _
        {"Verschlüssle Registrierung", "Code Recording", "Codifica Registro", "Code Enregistrement", "Codice Registrazione"}, _
        {"Restliche Registrierungen", "Remaining recordings", "Registros restantes", "Enregistrements restants", "Registrazioni restante"}, _
        {"Datum", "Date", "Fecha", "Date", "Data"}, _
        {"Recht des Benutzers, das kontaktiert das Verwalter der Software nicht erkannt ist.", "Right of the user not recognized contacted the administrator of the software.", "Derecho del usuario no reconocido puesto en contacto el administrador del software.", "Droit de l'utilisateur non reconnu contacté l'administrateur du logiciel", "Diritto dell'utente non riconosciuto contattato l'amministratore del software."}, _
        {"Nicht gültiger Benutzer", "Not valid user", "Usuario no válido", "Utilisateur non valide", "Utente non valido"}, _
        {"Nicht gültiges Kennwort", "Not valid password", "Contraseña no válida", "Mot de passe non valide", "Parola di ordine non valido"}, _
        {"Der Charakter _ ist im Kennwort verboten", "The character _ is forbidden in the password", "El carácter _ es prohibido en la contraseña", "le caractère _ est interdit dans le mot de passe", "il carattere _ è vietato nella parola di ordine"}, _
        {"Der Raumcharakter ist im Kennwort verboten", "The character of space is forbidden in the password", "El carácter de espacio es prohibido en la contraseña", "le caractère d'espace est interdit dans le mot de passe", "il carattere di spazio è vietato nella parola di ordine"}, _
        {"Der Charakter ' ist im Kennwort verboten", "The character ' is forbidden in the password", "El carácter ' es prohibido en la contraseña", "le caractère ' est interdit dans le mot de passe", "il carattere ' è vietato nella parola di ordine"}, _
        {"Das Kennwort soll 4 Zeichen mindestens sein", "The minimal length of a password is 4 characters", "La contraseña debe ser de 4 carácteres por lo menos", "La longueur minimale d'un  mot de passe est de 4 caractères.", "la parola di ordine deve essere di 4 caratteri al minimo"}, _
        {"Das Kennwort soll 15 maximaler Zeichen sein", "The maximal length of a password is 15 characters", "La contraseña debe ser de 15 carácteres máximo", "La longueur maximale d'un  mot de passe est de 15 caractères.", "la parola di ordine deve essere di 15 caratteri massimo"}, _
        {"Änderung des Kennwortes, die mit succes realisiert ist.", "Modification of the password realized with succes.", "Modificación de la palabra de paso realizada con succes.", "Modification du mot de passe réalisée avec succes.", "Modifica della parola di ordine realizzata con successo."}, _
        {"Sie müssen in einer Tätigkeit auswählen: tragen Sie bei, Modifizieren Sie, oder Löschen.", "You have to select in an action: add, Modify or Delete.", "Seleccione una acción: aumentar, Modificar o Suprimir.", "Vous devez sélectionner au moin une action : Ajouter, Modifier ou Supprimer.", "Selezionate un'azione:  Aggiungere, Modificare o Sopprimere."}, _
        {"Die besonderen oder betonten Zeichen sind im Namen von Benutzer verboten.", "The special or stressed characters are forbidden in user's name", "Los carácteres especiales o acentuados son prohibidos en el nombre de usuario.", "Les caractères spéciaux ou accentués sont interdit dans le  nom d'utilisateur.", "I caratteri speciali o accentuati sono vietati nel nome di utente."}, _
        {"Der Name des Benutzers kann 15 Zeichen. nicht überschreiten", "The name of the user cannot exceed 15 characters.", "nombre del usuario no puede sobrepasar 15 carácteres.", "Le nom de l'utilisateur ne peut dépasser 15 caractères.", "Il nome dell'utente non può superare 15 caratteri."}, _
        {"Der Raumcharakter ist im Namen von Benutzer verboten.", "The character of space is forbidden in user's name.", "El carácter de espacio es prohibido en el nombre de usuario.", "Le caractère d'espace est interdit dans le  nom d'utilisateur.", "Il carattere di spazio è vietato nel nome di utente."}, _
        {"Der Name des Benutzers ist 4 Zeichen mindestens.", "The name of the user is 4 characters at least.", "El nombre del usuario es de 4 carácteres por lo menos.", "Le nom de l'utilisateur est de 4 caractères au minimum.", "Il nome dell'utente è di 4 caratteri al minimo."}, _
        {"Die maximale Anzahl vom Benutzer ist, Unmöglichkeit erreicht, dieses Konto hinzuzufügen.", "The maximal number of user is reached, impossibility to add this account.", "El número máximo de usuario padece, imposibilidad de añadir esta cuenta.", "Le nombre maximal d'utilisateur est atteint, impossibilité de rajouter ce compte.", "Il numero massimale di utente è raggiunto, impossibilità di aggiungere questo conto."}, _
        {"Dieser Name von Benutzer existiert, Unmöglichkeit schon, ihn hinzuzufügen.", "This user's name already exists, impossibility to add it.", "Este nombre de usuario ya existe, imposibilidad de añadirlo.", "Ce nom d'utilisateur existe déjà, impossibilité de le rajouter.", "Questo nome di utente esiste già, impossibilità di aggiungerlo."}, _
        {"Sie sollen einen Typ von Benutzer auswählen.", "You must have selected a Type of user.", "Usted debe seleccionar un Tipo de usuario.", "Vous devez avoir sélectionné le Type d'utilisateur soit Administrateur ou Gestionnaire.", "Dovete selezionare un Tipo di utente."}, _
        {"Sie sollen den zu ändernden Benutzer auswählen.", "You have to select the user to modify.", "Usted debe seleccionar al usuario que hay que modificar.", "Vous devez sélectionner l'utilisateur à modifier.", "Dovete selezionare l'utente a modificare."}, _
        {"Der zu ändernde Benutzer existiert nicht.", "The user to modify does not exist.", "El usuario que hay que modificar no existe.", "L'utilisateur à modifier n'existe pas.", "L'utente a modificare non esistere."}, _
        {"Sie können das letzte Verwalter nicht ändern.", "You cannot modify the last administrator.", "Usted no puede modificar al último administrador.", "Vous ne pouvez pas modifier le dernier gestionnaire en administrateur.", "Non potete modificare l'ultimo amministratore."}, _
        {"Änderung des Benutzers, die mit succes realisiert ist.", "Modification of the user realized with succes.", "Cambiar el usuario realiza con éxito.", "Modification de l'utilisateur réalisée avec succes.", "Cambiare l'utente con successo."}, _
        {"Der Benutzer ist nicht vorhanden.", "The user does not delete.", "El usuario no existe.", "L'utilisateur à supprimer n'existe pas.", "L'utente non esiste."}, _
        {"Sie können nicht Ihr Konto löschen.", "You can not delete your account.", "No se puede eliminar su cuenta.", "Vous ne pouvez pas supprimer votre compte.", "Tu non puoi cancellare il tuo account."}, _
        {"Entfernen Sie den Benutzer erfolgreich durchgeführt.", "Remove the user performed successfully.", "Quite el usuario realiza con éxito.", "Suppression de l'utilisateur réalisée avec succes.", "Rimuovere l'utente eseguita con successo."}, _
        {"Beurkundung", "Authentification", "Autentificación", "Authentification", "Autenticazione"}, _
        {"Mit succes hinzugefügter Benutzer.", "User added with succes.", "Usuario aumentado con succes.", "Utilisateur ajouté avec succes.", "Utente aggiunto con successo."}, _
        {"Der Charakter _ ist im Namen von Benutzer verboten.", "The character _ is forbidden in user's name.", "El carácter _ es prohibido en el nombre de usuario.", "Le caractère _ est interdit dans le nom d'utilisateur.", "Il carattere _ è vietato nel nome di utente."}, _
        {"Der Raumcharakter ist im Namen von Benutzer verboten.", "The character of space is forbidden in user's name.", "El carácter de espacio es prohibido en el nombre de usuario.", "Le caractère d'espace est interdit dans le  nom d'utilisateur.", "Il carattere di spazio è vietato nel nome di utente."}, _
        {"Der Charakter ' ist im Namen von Benutzer verboten.", "The character ' is forbidden in user's name.", "El carácter ' es prohibido en el nombre de usuario.", "Le caractère ' est interdit dans le nom d'utilisateur.", "Il carattere ' è vietato nel nome di utente."}, _
        {"Die minimale Länge namens Benutzer ist 4 Zeichen.", "The minimal length of user's name is 4 characters.", "La longitud mínima del nombre de usuario es de 4 carácteres.", "La longueur minimale du nom d'utilisateur est de 4 caractères.", "La lunghezza minimale del nome di utente è di 4 caratteri."}, _
        {"Sie können das einzige Konto Verwalter nicht beseitigen.", "You cannot delete the only account Administrator", "Usted no puede suprimir la sola cuenta Administrador.", "Vous ne pouvez pas supprimer le seul compte Administrateur.", "Non potete sopprimere solo il conto Amministratore."}, _
        {"Aktualisiert Element", "Updated Element", "Actualizada Element", "Mise à jour Element", "Collocamento aggiornato Elemento"}, _
        {"Wünschen Sie, Ihnen den Tisch jetzt zu aktualisieren?", "Wish to update you the table now?", "¿ Desee actualizarle la mesa ahora?", "Désirez vous mettre à jour la table maintenant?", "Desiderate mettere aggiornati il conto adesso?"}, _
        {"Änderung", "Modification", "Modificación", "Modification", "Modifica"}, _
        {"Beendete Aktualisierung", "Ended update", "Actualización acabada", "Mise à jour terminée", "Collocamento aggiornato finito"}, _
        {"Irrtum Stimmsynthese", "Error Speech synthesis", "Error Síntesis Vocal", "Erreur Synthèse Vocale", "Errore Sintesi Vocale"}, _
        {"Sie haben keine Software von Stimmsynthese eingerichtet", "You did not install the software of speech synthesis", "Usted no instaló el software de síntesis vocal", "Vous n'avez pas installé le logiciel de synthèse vocale", "Non avete installato il software di sintesi vocale"}, _
        {"Test SAPI 5.4", "Test SAPI 5.4", "Prueba SAPI 5.4", "Test SAPI 5.4", "Test SAPI 5.4"}, _
        {"Gehen Sie den Text hinein, den Sie wünschen, hier zu sagen", "Enter text you wish spoken here", "Entre el texto que usted desea decir aquí", "Entrez le texte que vous souhaitez dire ici", "Introducete il testo che augurate dire qui"}, _
        {"Zu sagendem Text", "Speak Text", "Texto que hay que decir", "Texte à dire", "Testo a dire"}, _
        {"Stimme", "Voice", "Voz", "Voix", "Voce"}, _
        {"Audio herausgenommen", "Peripheral of audio output", "Periférico de salida audio", "Périphérique de sortie audio", "Periferico di uscita audio"}, _
        {"Geschwindigkeit :", "Rate :", "Velocidad :", "Vitesse : ", "velocità :"}, _
        {"Band :", "Volume :", "Volumen :", "Volume :", "Volume :"}, _
        {"Server-Setup", "Setting server", "Configuración del servidor", "Paramétrage serveur", "Configurazione dei server"}, _
        {"Registrierung abgeschlossen", "Registration Completed", "Registro completo", "Enregistrement Terminé", "Registrazione completata"}, _
        {"Erstellen Ansicht MySQL :", "Create view MySQL :", "Crear una vista de MySQL :", "Création de la vue MySQL :", "Creare vista MySQL :"}, _
        {"Kann Ansicht MySQL erstellen :", "Unable to create view MySQL :", "No se puede crear MySQL vista", "Impossible de Créer la vue MySQL :", "Impossibile creare MySQL vista :"}, _
        {"Erfolgreiche Sicherung", "Successful backup", "Copia de seguridad correcta", "Sauvegarde Réussie", "Backup riuscito"}, _
        {"Basis nicht verfügbar ist, aktualisieren Sie die Status der Basis in der Registerkarte Hilfe.", "Base not available, refresh the status of the base in the Help tab.", "Base no está disponible, actualizar el estado de la base en la ficha Ayuda.", "Base non disponible, rafraichir l'état de la base dans l'onglet Aide.", "Base non è disponibile, aggiornare lo stato della base nella scheda Guida."}, _
        {"Xml-Datei nicht gefunden", "Xml file not found", "Xml no se encuentra", "Fichier xml non trouvé", "Xml file non trovato"}, _
        {"Restaurierung abgeschlossen", "Restoration Completed", "Restauración completada", "Restauration Terminée", "Completato il restauro"}, _
        {"Diese Datei ist nicht die Backup-Datei.", "This file is not the backup file.", "Este archivo no es el archivo de copia de seguridad.", "Ce fichier n'est pas le fichier de Sauvegarde.", "Questo file non è file di backup."}, _
        {"Backup-Datei nicht gefunden", "Backup file not found", "Archivo de copia de seguridad que no se encuentran", "Fichier de sauvegarde introuvable", "File di backup non trovato"}, _
        {"Basis nicht verfügbar ist, aktualisieren Sie die Status der Basis in der Registerkarte Hilfe", "Base not available, refresh the status of the base in the Help tab", "Base no está disponible, actualizar el estado de la base en la pestaña Ayuda", "Base non disponible, rafraichir l'état de la base dans l'onglet Aide", "Base non è disponibile, aggiornare lo stato della base nella scheda Guida"}, _
        {"Wählen Sie File Backup", "Choose File Backup", "Seleccione Copia de seguridad de archivos", "Sélectionnez le Fichier de sauvegarde", "Scegliete File Backup"}}

        ''        {"", "", "", "", ""}, _

        Message = Table1(ID_Message, Langue)


        Return Message
    End Function


    Private Sub Tâche_de_fond_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles Tâche_de_fond.DoWork
        Dim travail As BackgroundWorker = CType(sender, BackgroundWorker)
        e.Result = Recherche(e.Argument, travail, e)
    End Sub

    Function Recherche(ByVal URL_à_Tracer As String, ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As String

        ' start NT tracert command 
        Dim ExeFileName As String
        Dim Arguments As String
        ExeFileName = "tracert"
        Arguments = " -h 20 " & URL_à_Tracer
        Dim s As String
        Dim p As New Process


        Dim key As RegistryKey
        Dim Chemin_OS As String

        key = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\windows NT\CurrentVersion")
        Chemin_OS = key.GetValue("PathName").ToString

        With p.StartInfo
            .WorkingDirectory = Chemin_OS & "\system32"
            .FileName = ExeFileName
            .Arguments = Arguments
            .UseShellExecute = False
            .RedirectStandardError = True
            .RedirectStandardInput = True
            .RedirectStandardOutput = True
            .WindowStyle = ProcessWindowStyle.Hidden
            .CreateNoWindow = True
        End With
        Try
            p.Start()
            s = p.StandardOutput.ReadToEnd()
            p.WaitForExit()
        Catch ex As Exception
            s = ex.ToString
        End Try
        p.Dispose()

        Résultat_tracert = s.ToString()
        Résultat_tracert = Résultat_tracert.Replace("ÿ", " ")
        Résultat_tracert = Résultat_tracert.Replace("‚", "e")
        Résultat_tracert = Résultat_tracert.Replace("Š", "e")
        'Résultat_tracert = Résultat_tracert.Replace("“", "o")

        Return Résultat_tracert

    End Function

    Private Sub Tâche_de_fond_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Tâche_de_fond.RunWorkerCompleted
        Infos_diagnostic = Diagnostique_Partie1 & vbCrLf & Résultat_tracert & vbCrLf & Diagnostique_Partie2
        Diagnostique_Partie1 = ""
        Résultat_tracert = ""
        Diagnostique_Partie2 = ""
    End Sub

    Private Sub BackgroundWorker_Diagnostic_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_Diagnostic.DoWork

        Diagnostic_OK = Diagnostic_de_la_connexion(e.Argument)

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        RichTextBox_Diagnostic.Text = Infos_diagnostic

        ProgressBar_Diagnostic.Value = My.Settings.ProgressBar_Diagnostic

        If (My.Settings.Arret_Timer = True) Then
            Timer1.Stop()
            Timer1.Dispose()
            My.Settings.ProgressBar_Diagnostic = 0
        End If

    End Sub
End Class
