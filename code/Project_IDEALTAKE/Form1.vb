Imports System.IO
Imports NetFwTypeLib
Imports System.Net

Public Class Form1

    Const MyKey As String = "1234"
    Const MyIV As String = "5678"


    Dim ChaineDeConnexion As String = ""
    Dim NomDeLaBase As String = ""
    Dim LangueLogiciel As Integer = 3
    'Valeur de Langue
    ' 0 = Allemand
    ' 1 = Anglais
    ' 2 = Espagnole
    ' 3 = Français
    ' 4 = Italien

    Dim Auteur As String = ""
    Dim Chrono As String = ""
    Dim Numéro_Enregistrement As String = ""


    'Liste des types de documents dans la base de données
    'Utilisation dans ComboBoxTypeDocumentMaintenance
    ' et dans ComboBoxTypeDocument
    Dim ListOfTypesDocuments As New List(Of ClasseTypesDocuments)

    'Liste des documents lié au chrono
    'Utilisation dans ListBoxDocuments
    Dim ListeDesDocumentsPourLeChrono As New List(Of ClasseDocuments)

    'Création d'un nouveau document
    Dim NouveauDocument As New ClasseDocuments


    Dim BooleanAllemand As Boolean = False
    Dim BooleanAnglais As Boolean = False
    Dim BooleanEspagnol As Boolean = False
    Dim BooleanFrançais As Boolean = False
    Dim BooleanItalien As Boolean = False

    Dim NouvauDocumentInfos As New Document

    Public Structure RetourLog

        Public Erreur As Boolean
        Public Infos_requêtes As String
        Public RetourLog As Integer


    End Structure

    Public Structure RetourChargementListeDocuments

        Public Erreur As Boolean
        Public Infos_requêtes As String
        Public RetourChargementListeDocuments As Integer

    End Structure

    Public Structure RetourSupprimeDocument

        Public Erreur As Boolean
        Public Infos_requêtes As String
        Public RetourSupprimeDocument As Integer

    End Structure

    Public Structure RetourMaintenanceTransfert

        Public Erreur As Boolean
        Public Infos_requêtes As String
        Public RetourMaintenanceTransfert As Integer

    End Structure


    Public Structure RetourAjoutDocument

        Public Erreur As Boolean
        Public Infos_requêtes As String
        Public RetourAjoutDocument As Integer

    End Structure


    Public Structure Argument_ChargementTypeDocument

        ' paramètres pour RequêteSQLChargementTypeDocument()
        Public Auteur As String
        Public chaine_de_connexion As String
        Public Nom_de_la_base_de_données As String
        Public langue_logiciel As Integer

    End Structure

    Public Structure Argument_ListeDocuments

        ' paramètres pour RequêteSQLChargementListeDocuments()

        Public chaine_de_connexion As String
        Public Nom_de_la_base_de_données As String
        Public langue_logiciel As Integer
        Public Chrono As String
        Public Numéro_Enregistrement As String

    End Structure

    Public Structure Argument_MaintenanceTransfert
        ' paramètres pour RequêteSQLMaintenanceTransfert()

        Public Auteur As String
        Public chaine_de_connexion As String
        Public Nom_de_la_base_de_données As String
        Public langue_logiciel As Integer
        Public IDTypeDocument As Integer  'Provient de 'type_documents'
        Public TypeDocument As Integer    'Provient de 'Transfert'  
        Public Destination As String
        Public TypeTransfert As Integer
        Public Utilisateur As String
        Public MotDePasse As String

    End Structure

    Public Structure Argument_SupprimeDocument
        ' paramètres pour RequêteSQLSupprimeDocument()

        Public Auteur As String
        Public chaine_de_connexion As String
        Public Nom_de_la_base_de_données As String
        Public langue_logiciel As Integer
        Public Chrono As String
        Public NomDuLien As String


    End Structure


    Public Structure Argument_AjoutDocument
        ' paramètres pour RequêteSQLAjoutDocument()

        Public Auteur As String
        Public chaine_de_connexion As String
        Public Nom_de_la_base_de_données As String
        Public langue_logiciel As Integer
        Public Chrono As String

    End Structure

    Public Structure RetourChargementTypeDocument

        Public Erreur As Boolean
        Public Infos_requêtes As String
        Public RetourChargementTypeDocument As Integer

    End Structure


    Public Structure TypeDocument

        'TypeTransfert 
        ' 2 = Local
        ' 1 = FTP

        'Provient de la table 'types_documents'
        Dim IDTypeDocument As String
        Dim DescriptionTypeDocument As String

        'Provient de la table 'Transfert'
        Dim TypeDocument As String
        Dim TypeTransfert As Integer
        Dim Destination As String
        Dim User As String
        Dim MotDePasse As String


    End Structure





    Private Structure Document

        Dim NumEnregistrement As String
        Dim Chrono As String

        Dim NomTypeDocument As String
        Dim NuméroTypeDocument As String
        Dim TypeTransfert As Integer
        Dim Destination As String
        Dim Utilisateur As String
        Dim MotDePasse As String

        Dim NomDuLienFichierALL As String
        Dim NomDuLienFichierANG As String
        Dim NomDuLienFichierESP As String
        Dim NomDuLienFichierFRA As String
        Dim NomDuLienFichierITA As String

        Dim OrigineFichierALL As String
        Dim NomFichierALL As String
        Dim TailleFichierALL As String

        Dim OrigineFichierANG As String
        Dim NomFichierANG As String
        Dim TailleFichierANG As String

        Dim OrigineFichierESP As String
        Dim NomFichierESP As String
        Dim TailleFichierESP As String

        Dim OrigineFichierFRA As String
        Dim NomFichierFRA As String
        Dim TailleFichierFRA As String

        Dim OrigineFichierITA As String
        Dim NomFichierITA As String
        Dim TailleFichierITA As String

    End Structure


    Private Function ConvertStringToBase64(ByVal StringToConvert As String)

        'byte[] byt = System.Text.Encoding.UTF8.GetBytes(strOriginal);
        '// convert the byte array to a Base64 string
        'strModified = Convert.ToBase64String(byt);

        Dim TableByte As Byte() = System.Text.Encoding.UTF8.GetBytes(StringToConvert)
        Return Convert.ToBase64String(TableByte)

    End Function


    Private Function ConvertBase64ToString(ByVal BASE64ToConvert As String)

        'System.Text.Encoding.UTF8.GetBytes(StringToConvert)
        'Dim TableByte As Byte() = Convert.FromBase64String(BASE64ToConvert)
        'System.Text.UTF8Encoding enc = new System.Text.UTF8Encoding();
        'str = enc.GetString(dBytes);
        Dim TableByte As Byte() = System.Convert.FromBase64String(BASE64ToConvert)
        Dim EncodeTable As System.Text.UTF8Encoding = New System.Text.UTF8Encoding()
        Return EncodeTable.GetString(TableByte)


    End Function


    Private Function MessageUtilisateur(ByVal IDMessage As Integer, ByVal Langue As Integer)
        Dim Message As String = Nothing

        'Valeur de Langue
        ' 0 = Allemand
        ' 1 = Anglais
        ' 2 = Espagnole
        ' 3 = Français
        ' 4 = Italien

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
        {"Wählen Sie File Backup", "Choose File Backup", "Seleccione Copia de seguridad de archivos", "Sélectionnez le Fichier de sauvegarde", "Scegliete File Backup"}, _
        {"Im Bau", "Under construction", "Bajo construcción", "En cours de construction", "In costruzione"},
        {"Einfache Suche", "Simple Search", "Búsqueda simple", "Recherche Simplifiée", "Ricerca Semplice"}, _
        {"Erweiterte Suche", "Advanced Search", "Búsqueda Avanzada", "Recherche Avancée", "Ricerca Avanzata"}, _
        {"Geführte Suche", "Guided Search", "Búsqueda Guiada", "Recherche Guidée", "Ricerca Guidata"}, _
        {"Schließung der Seite", "Closure of the page", "El cierre de la página", "Fermeture de la page", "Chiusura della pagina"}, _
        {"Maximale Größe des Bildes 500 KB", "Maximum size of your image 500 KB", "El tamaño máximo de la imagen de 500 KB", "Taille Maximale de votre image 500 Ko", "La dimensione massima della tua immagine 500 KB"}, _
        {"Datumsformat Ungültig", "Format of the date invalid", "Formato de la fecha no válida", "Format de la date non valide", "Data formato non valido"}, _
        {"Forschung Abgeschlossene", "Research Completed", "Investigación realizada", "Recherche Terminée", "Completato ricerca"}, _
        {"Suche im Gange...", "Search in progress...", "Búsqueda en curso...", "Recherche en cours...", "Ricerca in corso..."}, _
        {"Bitte geben Sie Ihren Benutzernamen aus mindestens 6 Zeichen.", "Please enter your user name at least 6 characters.", "Por favor, introduzca su nombre de usuario por lo menos 6 caracteres.", "Veuillez saisir votre nom d'utilisateur 6 caractères minimum.", "Inserisci il tuo nome utente di almeno 6 caratteri."}, _
        {"Bitte geben Sie Ihr Passwort mindestens 6 Zeichen.", "Please enter your password at least 6 characters.", "Por favor, introduzca su contraseña de 6 caracteres como mínimo.", "Veuillez saisir votre mot de passe 6 caractères minimum.", "Inserisci il tuo password di almeno 6 caratteri."}, _
        {"Bestätigung des Kennworts unterscheidet sich von der Passwort.", "Confirmation of the password is different from the password.", "La confirmación de la contraseña es diferente de la contraseña.", "La confirmation du mot de passe est différente du mot de passe.", "La conferma della password è diversa dalla password."}, _
        {"Die E-Mail-Adresse ist nicht gültig.", "The email address is not valid.", "La dirección de correo electrónico no es válida.", "L'adresse e-mail n'est pas valide.", "L'indirizzo email non è valido."}, _
        {"Bitte geben Sie den Bestätigungscode Passwort mindestens 6 Zeichen.", "Please type the confirmation password at least 6 characters.", "Por favor, escriba la contraseña de confirmación de al menos 6 caracteres.", "Veuillez saisir la confirmation du mot de passe 6 caractères minimum.", "Si prega di digitare la password di conferma di almeno 6 caratteri."}, _
        {"Die Antwort muss mindestens 4 Zeichen.", "The answer must contain a minimum of 4 characters.", "La respuesta debe contener un mínimo de 4 caracteres.", "La réponse doit contenir 4 caractères au minimum.", "La risposta deve contenere un minimo di 4 caratteri."}, _
        {"Aufnahme läuft...", "Recording in progress...", "Grabación en progreso...", "Enregistrement en cours...", "Registrazione in corso..."}, _
        {"Ihr Passwort ist zu lang, um 45 Zeichen nicht überschreiten.", "Your password is too long to exceed 45 characters.", "Su contraseña es demasiado larga para ser superior a 45 caracteres.", "Votre mot de passe est trop long 45 caractères Maximum.", "La tua password è troppo lungo per superare i 45 caratteri."}, _
        {"Ihre E-Mail ist zu lang, zu 45 Zeichen nicht überschreiten.", "Your email is too long to exceed 45 characters.", "Su correo electrónico es demasiado larga para ser superior a 45 caracteres.", "Votre email est trop longue 45 caractères Maximum.", "Vostro email è troppo lungo per superare i 45 caratteri."}, _
        {"Die Antwort sollte 256 Zeichen enthalten.", "The answer should contain 256 characters.", "La respuesta debe contener 256 caracteres.", "La réponse doit contenir 256 caractères au maximum.", "La risposta dovrebbe contenere 256 caratteri."}, _
        {"Ihr Benutzername ist zu lang, zu 45 Zeichen nicht überschreiten.", "Your username is too long to exceed 45 characters.", "Su nombre de usuario es demasiado largo para superar los 45 caracteres.", "Votre nom d'utilisateur est trop long 45 caractères Maximum.", "Il nome utente è troppo lungo per superare i 45 caratteri."}, _
        {"Hinzufügen von Benutzer-Abgeschlossene", "Adding User Completed", "Adición de usuarios terminada", "Ajout Utilisateur Terminé", "Aggiunta utente ha completato"}, _
        {"Failure Bestehende Anwender", "Failure Existing user", "El incumplimiento de usuario existentes", "Echec Utilisateur existant", "La mancata utente esistente"}, _
        {"Der Benutzername oder das eingegebene Passwort ist falsch.", "The username or password you entered is incorrect.", "El nombre de usuario o contraseña es incorrecta.", "Le nom d'utilisateur ou le mot de passe saisi est incorrect.", "Il nome utente o la password che hai inserito non è corretto."}, _
        {"Erfolgreiche Authentifizierung", "Successful authentication", "La autenticación exitosa", "Authentification réussie", "l'autenticazione di successo"}, _
        {"Laufende Prüfung...", "Checking during...", "La auditoría en curso...", "Vérification en cours...", "In corso controllo..."}, _
        {"Der Benutzername oder das eingegebene Passwort ist falsch.", "The username or password you entered is incorrect.", "El nombre de usuario o contraseña es incorrecta.", "Le nom d'utilisateur ou le mot de passe saisi est incorrect.", "Il nome utente o la password che hai inserito non è corretto."}, _
        {"Erfolgreiche Authentifizierung", "Successful authentication", "La autenticación exitosa", "Authentification réussie", "l'autenticazione di successo"}, _
        {"Laufende Prüfung...", "Checking during...", "La auditoría en curso...", "Vérification en cours...", "In corso controllo..."}, _
        {"Vollständige Wiederherstellung Benutzer", "Complete Recovery User", "Recuperación completa del usuario", "Récupération Utilisateur Terminée", "Recupero User Complete"}, _
        {"Benutzer nicht gefunden", "User not found", "Usuario no encontrado", "Utilisateur non trouvé", "Utente non trovato"}, _
        {"Löschen Ihres Kontos", "Deleting your account", "Al eliminar su cuenta", "Suppression de votre compte", "Eliminare il tuo account"}, _
        {"Sind Sie sicher, dass Sie dieses Konto löschen?", "Are you sure you want to delete this account?", "¿Está seguro que desea eliminar esta cuenta?", "Êtes vous sûr de vouloir supprimer ce compte?", "Sei sicuro di voler cancellare questo account?"}, _
        {"Konto Löschen erfolgreich ist.", "Account deletion is successful.", "Eliminación de la cuenta es correcta.", "La suppression du compte s'est terminée correctement.", "La cancellazione dell'account è successo."}, _
        {"Das Konto wurde nicht gefunden.", "The account was not found.", "La cuenta no se ha encontrado.", "Le compte n'a pas été trouvé.", "L'account non è stato trovato."}, _
        {"Abbrechen des Löschvorgangs.", "Canceling the delete operation.", "Cancelación de la operación de eliminación.", "Annulation de l'opération de suppression.", "Annullamento l'operazione di eliminazione."}, _
        {"User Verifizierung Abgeschlossene.", "User Verification Completed.", "Verificación de usuarios terminada.", "Vérification  Utilisateur Terminée.", "Verifica utente ha completato."}, _
        {"Ihr Konto existiert.", "Your Account exists.", "Su cuenta existe.", "Votre Compte existe.", "Il tuo account esiste."}, _
        {"Ihr Konto ist nicht vorhanden.", "Your account does not exist.", "Su cuenta no existe.", "Votre Compte n'existe pas.", "Il tuo account non esiste."}, _
        {"Sie sind der Autor des Eintrags wird die Entfernung zugelassen.", "You are the author of the record, the removal is authorized.", "Usted es el autor de la grabación, la eliminación está autorizado.", "Vous êtes l'auteur de l'enregistrement, la suppression est autorisée.", "Tu sei l'autore del record, la rimozione è autorizzato."}, _
        {"Du bist nicht der Autor der Aufnahme wird das Löschen nicht erlaubt.", "You are not the author of the recording, deleting is not allowed.", "Usted no es el autor de la grabación, la eliminación no está permitido.", "Vous n'êtes pas l'auteur de l'enregistrement, la suppression n'est pas autorisée.", "Tu non sei l'autore della registrazione, la cancellazione non è consentita."}, _
        {"Aktuelle Löschung...", "Current deletion...", "Eliminación actual...", "Suppression en cours...", "La cancellazione di corrente..."}, _
        {"Erfolgreiche Entfernung.", "Successful removal.", "El éxito de la eliminación.", "Suppression Réussie.", "La rimozione di successo."}, _
        {"Gescheiterte Abschaffung.", "Failed Deletion.", "Supresión Sida suspendido.", "Suppression Echouée.", "Soppressione Fallita."}, _
        {"Registrierungs-Nummer 2 Zeichen Minimum.", "Registration number 2 characters minimum.", "Número de registro mínimo de 2 caracteres.", "Numéro d'enregistrement 2 caractères minimum.", "Numero di registrazione minimo di 2 caratteri."}, _
        {"Registrierungs-Nummer bis zu 50 Zeichen.", "Registration number up to 50 characters.", "Número de registro de hasta 50 caracteres.", "Numéro d'enregistrement 50 caractères maximum.", "Numero di registrazione fino a 50 caratteri"}, _
        {"Hersteller-Feld Problem.", "Manufacturer field problem.", "Fabricante de problemas de campo.", "Problème champ Constructeur.", "Produttore campo di problema."}, _
        {"Problem auf dem Access Level.", "Problem on the Access Level.", "Problemas en el nivel de acceso.", "Problème sur le Niveau d'accès.", "Problema del livello di accesso."}, _
        {"Author-Feld Problem.", "Author field problem.", "Autor de problemas de campo.", "Problème champ Auteur.", "Campo Autore problema."}, _
        {"Deutsch Titel bis zu 45 Zeichen.", "German title up to 45 characters.", "Título Alemán de hasta 45 caracteres.", "Titre Allemand  maximum 45 caractères.", "Titolo Tedesco fino a 45 caratteri."}, _
        {"Deutsch Untertitel  bis zu 45 Zeichen.", "German Subtitle up to 45 characters.", "Subtítulo Alemán de hasta 45 caracteres.", "Sous Titre Allemand  maximum 45 caractères.", "Sottotitolo Tedesco fino a 45 caratteri."}, _
        {"Deutsch Kommentar Maximal 1024 Zeichen.", "German Comment Maximum 1024 characters.", "Alemanes Comentario Máximo de 1024 caracteres.", "Commentaire  Allemand  maximum 1024 caractères.", "Commento tedeschi Massimo 1024 caratteri."}, _
        {"Englisch Titel bis zu 45 Zeichen.", "English title up to 45 characters.", "Título Inglés de hasta 45 caracteres.", "Titre Anglais  maximum 45 caractères.", "Titolo Inglese fino a 45 caratteri."}, _
        {"Englisch Untertitel  bis zu 45 Zeichen.", "English Subtitle up to 45 characters.", "Subtítulo Inglés de hasta 45 caracteres.", "Sous Titre Anglais  maximum 45 caractères.", "Sottotitolo Inglese fino a 45 caratteri."}, _
        {"Englisch Kommentar Maximal 1024 Zeichen.", "English Comment Maximum 1024 characters.", "Inglés Comentario Máximo de 1024 caracteres.", "Commentaire  Anglais  maximum 1024 caractères.", "Commento Inglese Massimo 1024 caratteri."}, _
        {"Spanisch Titel bis zu 45 Zeichen.", "Spanish title up to 45 characters.", "Título Español de hasta 45 caracteres.", "Titre Espagnol  maximum 45 caractères.", "Titolo spagnolo fino a 45 caratteri."}, _
        {"Spanisch Untertitel  bis zu 45 Zeichen.", "Spanish Subtitle up to 45 characters.", "Subtítulo Español de hasta 45 caracteres.", "Sous Titre Espagnol  maximum 45 caractères.", "Sottotitolo spagnolo fino a 45 caratteri."}, _
        {"Spanisch Kommentar Maximal 1024 Zeichen.", "Spanish Comment Maximum 1024 characters.", "Español Comentario Máximo de 1024 caracteres.", "Commentaire  Espagnol  maximum 1024 caractères.", "Commento spagnolo Massimo 1024 caratteri."}, _
        {"Französisch Titel bis zu 45 Zeichen.", "French title up to 45 characters.", "Título Francés de hasta 45 caracteres.", "Titre Français  maximum 45 caractères.", "Titolo Francese fino a 45 caratteri."}, _
        {"Französisch Untertitel  bis zu 45 Zeichen.", "French Subtitle up to 45 characters.", "Subtítulo Francés de hasta 45 caracteres.", "Sous Titre Français  maximum 45 caractères.", "Sottotitolo Francese fino a 45 caratteri."}, _
        {"Französisch Kommentar Maximal 1024 Zeichen.", "French Comment Maximum 1024 characters.", "Francés Comentario Máximo de 1024 caracteres.", "Commentaire  Français  maximum 1024 caractères.", "Commento Francese Massimo 1024 caratteri."}, _
        {"Italiano Titel bis zu 45 Zeichen.", "Italiano title up to 45 characters.", "Título Italiano de hasta 45 caracteres.", "Titre Italien  maximum 45 caractères.", "Titolo Italiano fino a 45 caratteri."}, _
        {"Italiano Untertitel  bis zu 45 Zeichen.", "Italiano Subtitle up to 45 characters.", "Subtítulo Italiano de hasta 45 caracteres.", "Sous Titre Italien  maximum 45 caractères.", "Sottotitolo Italiano fino a 45 caratteri."}, _
        {"Italiano Kommentar Maximal 1024 Zeichen.", "Italiano Comment Maximum 1024 characters.", "Italiano Comentario Máximo de 1024 caracteres.", "Commentaire  Italien  maximum 1024 caractères.", "Commento Italiano Massimo 1024 caratteri."}, _
        {"Feld-Problem ""PARAM_A""", "Field problem ""PARAM_A""", "Campo problema ""PARAM_A""", "Problème champ ""PARAM_A""", "Campo di problema ""PARAM_A"""}, _
        {"Feld-Problem ""PARAM_B""", "Field problem ""PARAM_B""", "Campo problema ""PARAM_B""", "Problème champ ""PARAM_B""", "Campo di problema ""PARAM_B"""}, _
        {"Feld-Problem ""PARAM_C""", "Field problem ""PARAM_C""", "Campo problema ""PARAM_C""", "Problème champ ""PARAM_C""", "Campo di problema ""PARAM_C"""}, _
        {"Feld-Problem ""PARAM_D""", "Field problem ""PARAM_D""", "Campo problema ""PARAM_D""", "Problème champ ""PARAM_D""", "Campo di problema ""PARAM_D"""}, _
        {"Feld-Problem ""PARAM_E""", "Field problem ""PARAM_E""", "Campo problema ""PARAM_E""", "Problème champ ""PARAM_E""", "Campo di problema ""PARAM_E"""}, _
        {"Feld-Problem ""PARAM_F""", "Field problem ""PARAM_F""", "Campo problema ""PARAM_F""", "Problème champ ""PARAM_F""", "Campo di problema ""PARAM_F"""}, _
        {"Feld-Problem ""PARAM_G""", "Field problem ""PARAM_G""", "Campo problema ""PARAM_G""", "Problème champ ""PARAM_G""", "Campo di problema ""PARAM_G"""}, _
        {"Feld-Problem ""PARAM_H""", "Field problem ""PARAM_H""", "Campo problema ""PARAM_H""", "Problème champ ""PARAM_H""", "Campo di problema ""PARAM_H"""}, _
        {"Feld-Problem ""PARAM_I""", "Field problem ""PARAM_I""", "Campo problema ""PARAM_I""", "Problème champ ""PARAM_I""", "Campo di problema ""PARAM_I"""}, _
        {"Feld-Problem ""STANDARD""", "Field problem ""STANDARD""", "Campo problema ""STANDARD""", "Problème champ ""STANDARD""", "Campo di problema ""STANDARD"""}, _
        {"Diese Nummer ist bereits vorhanden, werden mit Registrierung.", "This registration number already exists, registration canceled.", "Este número de registro ya existe, la inscripción cancelada.", "Ce numéro d'enregistrement existe déjà, enregistrement annulé.", "Questo numero di registrazione esiste già, la registrazione annullata."}, _
        {"Erstellen Sie einen Datensatz, bevor Sie eine Option.", "Create a record before using an option.", "Crear un registro antes de utilizar una opción.", "Créez un enregistrement avant d'utiliser une option.", "Creare un record prima di utilizzare l'opzione."}, _
        {"Aktualisieren Sie den aktuellen Datensatz...", "Update the current record...", "Actualizar el registro actual...", "Mise à jour de l'enregistrement en cours...", "Aggiornare il record corrente..."}, _
        {"Hinzufügen Option Symbolbaum Fertig.", "Adding option Symbol Tree Done.", "Adición de la opción del árbol símbolo de Hecho.", "Ajout option Symbole Arbre Terminé.", "Aggiunta opzione di Symbol Albero Fatto."}, _
        {"Die Referenzliste ist leer.", "The reference list is empty.", "La lista de referencia está vacía.", "La liste de référence est vide.", "La lista di riferimento è vuoto."}, _
        {"Hinzufügen von Links Chrono, Referenzen Fertig.", "Adding links Chrono, References Done.", "Adición de vínculos Referencias, Chrono Hecho.", "Ajout des liens Chrono, Références Terminé.", "L'aggiunta di collegamenti riferimenti, Chrono Done."}, _
        {"Entfernen von Links Chrono, Referenzen Fertig.", "Removing links Chrono, References Done.", "Eliminación de enlaces Referencias, Chrono Hecho.", "Suppression des liens Chrono, Références Terminé.", "Rimozione link Riferimenti, Chrono Fatto."}, _
        {"Laden...", "Loading...", "Cargando...", "Chargement en cours...", "Caricamento..."}, _
        {"Daten werden geladen Arten Dokumentation Fertig.", "Loading data types documentation Done.", "Carga de datos de tipos de documentación de Hecho.", "Chargement des données des types documents Terminé.", "Caricamento dei dati tipi di documentazione Fatto."}, _
        {"Zielpfad zu kurz (mindestens 3 Zeichen).", "Destination path too short (3 characters minimum).", "Destino demasiado corta trayectoria (mínimo 3 caracteres).", "Chemin de Destination trop court (3 caractères minimum).", "Destinazione percorso troppo breve (minimo 3 caratteri)."}, _
        {"Zielpfad zu lang (250 Zeichen).", "Destination path too long (250 characters).", "Destino demasiado largo recorrido (250 caracteres).", "Chemin de Destination trop long (250 caractères maximum).", "Destinazione percorso troppo lungo (250 caratteri)."}, _
        {"Benutzername zu kurz (mindestens 3 Zeichen).", "Username too short (3 characters minimum).", "Nombre de usuario demasiado corto (mínimo 3 caracteres).", "Nom d'utilisateur trop court (3 caractères minimum).", "Nome utente Password troppo breve (minimo 3 caratteri)."}, _
        {"Benutzername zu lang (20 Zeichen).", "Username too long (20 characters).", "Nombre de usuario demasiado largo (20 caracteres).", "Nom d'utilisateur trop long (20 caractères maximum).", "Nome troppo lungo (20 caratteri)."}, _
        {"Passwort zu kurz (mindestens 3 Zeichen).", "Password too short (3 characters minimum).", "Contraseña demasiado corta (mínimo 3 caracteres).", "Mot de passe trop court (3 caractères minimum).", "Password troppo breve (minimo 3 caratteri)."}, _
        {"Passwort zu lang (20 Zeichen).", "Password too long (20 characters).", "Contraseña demasiado largo (20 caracteres).", "Mot de passe trop long (20 caractères maximum).", "Password troppo lunga (20 caratteri)."}, _
        {"Einlegen von Dokumenten-Liste.", "Loading Documents List Done.", "Lista de Carga de documentos en Hecho.", "Chargement Liste Documents Terminé.", "Completato Lista Caricamento dei documenti."}, _
        {"Aktualisiert 'Transfert' Fertig.", "Updated 'Transfert' Done.", "Actualizado 'Transfert' Hecho.", "Mise à jour 'Transfert' Terminé.", "Aggiornato 'Transfert' Fatto."}, _
        {"Möchten Sie dieses Dokument löschen?", "Would you delete this document?", "¿Le eliminar este documento?", "Voulez vous Supprimer ce document?", "Vuoi eliminare questo documento?"}, _
        {"Dokument löschen", "Delete Document", "Eliminar Documento", "Supprimer Document", "Elimina Documento"}, _
        {"Sie müssen ein Dokument.", "You must select a document.", "Debe seleccionar un documento.", "Vous devez sélectionner un document.", "È necessario selezionare un documento."}, _
        {"Aktualisiert 'Transfert' im Gange...", "Updated 'Transfert' in progress...", "Actualización 'Transfert' en el progreso...", "Mise à jour 'Transfert' en cours...", "Aggiornato il 'Transfert' in progress..."}, _
        {"Clearing-Dokument unter...", "Clearing Document under...", "Eliminación de documentos en...", "Effacement Document en cours...", "Cancellazione del documento in..."}, _
        {"Entfernen Sie das fertige Dokument.", "Remove the finished document.", "Retire el documento final.", "Suppression du document terminé.", "Rimuovere il documento finito."}, _
        {"Erfolgreich hochgeladenen Datei.", "File uploaded successfully.", "Archivo subido con éxito.", "Téléchargement réussi.", "File caricato con successo."}, _
        {"Download fehlgeschlagen.", "Download failed.", "Download fehlgeschlagen.", "Téléchargement échoué.", "Download non riuscito."}, _
        {"Das Ziel wird in dem Dokument nicht für den Transfer-Typ angegeben.", "The destination is not specified in the document for the transfer type", "El destino no se especifica en el documento para el tipo de transferencia.", "La Destination n'est pas renseigné dans le types document pour le transfert.", "La destinazione non è specificato nel documento per il tipo di trasferimento."},
        {"Der Link-Name in der deutschen fehlt.", "The link name in German is absent.", "El nombre del enlace en alemán está ausente.", "Le nom du lien en allemand est absent.", "Il nome del collegamento in tedesco è assente."}, _
        {"Der Link-Name in englischer Sprache fehlt.", "The link name in English is absent.", "El nombre del enlace en Inglés está ausente.", "Le nom du lien en anglais est absent.", "Il nome del link in inglese è assente."}, _
        {"Der Link-Name in Spanisch ist nicht vorhanden.", "The link name in Spanish is absent.", "El nombre del enlace en español está ausente.", "Le nom du lien en espagnol est absent.", "Il nome del collegamento in spagnolo è assente."}, _
        {"Der Link-Name auf Französisch fehlt.", "The link name in French is absent.", "El nombre del enlace en francés está ausente.", "Le nom du lien en Français est absent.", "Il nome del collegamento in francese è assente."}, _
        {"El nombre del enlace en italiano está ausente.", "The link name in Italian is absent.", "El nombre del enlace en italiano está ausente.", "Le nom du lien en Italien est absent.", "Il nome del link in italiano è assente."}, _
        {"Der Dateiname ist in deutscher Sprache fehlt.", "The file name is absent in German.", "El nombre del archivo no está presente en Alemania.", "Le nom du fichier en allemand est absent.", "Il nome del file è assente in tedesco."}, _
        {"Der Dateiname fehlt in englischer Sprache.", "The file name is absent in English.", "El nombre del archivo no está presente en Inglés.", "Le nom du fichier en anglais est absent.", "Il nome del file è assente in lingua inglese."}, _
        {"Der Name der Datei in spanischer Sprache fehlt.", "The file name in Spanish is absent.", "El nombre del archivo en español está ausente.", "Le nom du fichier en espagnol est absent.", "Il nome del file in lingua spagnola è assente."}, _
        {"Der Name der Datei in Französisch ist nicht vorhanden.", "The file name in French is absent.", "El nombre del archivo en francés está ausente.", "Le nom du fichier en Français est absent.", "Il nome del file in lingua francese è assente."}, _
        {"Der Name der Datei in Italienisch ist nicht vorhanden.", "The file name in Italian is absent.", "El nombre del archivo en italiano está ausente.", "Le nom du fichier en Italien est absent.", "Il nome del file in lingua italiana è assente."}, _
        {"Keine Datei in deutscher Sprache ausgewählt.", "No file selected in German.", "No existe el fichero seleccionado en alemán.", "Pas de fichier sélectionné en allemand.", "Nessun file selezionato in tedesco."}, _
        {"Keine Datei in englischer Sprache ausgewählt.", "No file selected in English.", "No existe el fichero seleccionado en Inglés.", "Pas de fichier sélectionné  en anglais.", "Nessun file selezionato in inglese."}, _
        {"Keine Datei in spanischer Sprache ausgewählt.", "No file selected in Spanish.", "No existe el fichero seleccionado en español.", "Pas de fichier sélectionné en espagnol.", "Nessun file selezionato in spagnolo."}, _
        {"Keine Datei in Französisch gewählt.", "No file selected in French.", "No existe el fichero seleccionado en francés.", "Pas de fichier sélectionné en Français.", "Nessun file selezionato in francese."}, _
        {"Keine Datei auf Italienisch gewählt.", "No file selected in Italian.", "No existe el fichero seleccionado en italiano.", "Pas de fichier sélectionné en Italien.", "Nessun file selezionato in italiano."}, _
        {"Die ausgewählte Datei nicht vorhanden ist in deutscher Sprache.", "The selected file is absent in German.", "El archivo seleccionado no existe en alemán.", "Le fichier sélectionné en allemand est absent.", "Il file selezionato è assente in tedesco."}, _
        {"Die ausgewählte Datei nicht vorhanden ist in englischer Sprache.", "The selected file is absent in English.", "El archivo seleccionado no existe en Inglés.", "Le fichier sélectionné  en anglais est absent.", "Il file selezionato è assente in lingua inglese."}, _
        {"Die ausgewählte Datei ist nicht vorhanden in Spanisch.", "The selected file is not present in Spanish.", "El archivo seleccionado no está presente en español.", "Le fichier sélectionné en espagnol est absent.", "Il file selezionato non è presente in spagnolo."}, _
        {"Die ausgewählte Datei ist nicht vorhanden in Französisch.", "The selected file is not present in French.", "El archivo seleccionado no está presente en Francia.", "Le fichier sélectionné en Français est absent.", "Il file selezionato non è presente in francese."}, _
        {"Die ausgewählte Datei nicht vorhanden ist in italienischer Sprache.", "The selected file is absent in Italian.", "El archivo seleccionado no se da en italiano.", "Le fichier sélectionné en Italien est absent.", "Il file selezionato è assente in italiano."}, _
        {"Hinzufügen von Dokument-Fertig.", "Adding Document Done.", "Adición de Documento Hecho.", "Ajout document Terminé.", "Aggiunta di documenti Fatto."}}

        ''        {"", "", "", "", ""}, _

        Message = Table1(IDMessage, Langue)


        Return Message
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
        {"Transfer-Idealtake", "Transfer Idealtake", "Transferencia Idealtake", "Transfert Idealtake", "Transfer Idealtake"}, _
        {"Der Name der Datenbank", "Name of the database", "Nombre de la base de datos", "Nom de la base de données", "Nome del database"}, _
        {"Registry Number", "Registry Number", "Numero del Registro di sistema", "Numéro d'enregistrement", "Numero del Registro di sistema"}, _
        {"Anzahl Chrono", "Number Chrono", "Número de Chrono", "Numéro de Chrono", "Numero di Chrono"}, _
        {"Document Type", "Document Type", "Tipo de documento", "Type Document", "Tipo di documento"}, _
        {"Ziel", "Destination", "Destino", "Destination", "Destinazione"}, _
        {"Lokale Reiseziel", "Local destination", "Destino local", "Destination Local", "Destinazione locale"}, _
        {"Benutzer", "User", "Usuario", "Utilisateur", "Utente"}, _
        {"FTP", "FTP", "FTP", "FTP", "FTP"}, _
        {"Lokal", "Local", "Local", "Local", "Locale"}, _
        {"Kennwort", "Password", "Contraseña", "Mot de passe", "Password"}, _
        {"Aktualisiert", "Updated", "Actualizado", "Mise à jour", "Aggiornato"}, _
        {"Liste von Dokumenten, Aufzeichnung", "List of documents recording", "Lista de grabación documentos", "Liste des documents de l'enregistrement", "Elenco di registrazione documenti"}, _
        {"Document Type", "Document Type", "Tipo de documento", "Type Document", "Tipo di documento"}, _
        {"Wählen Sie Datei", "Select File", "Seleccione Archivo", "Sélectionner le Fichier", "Selezionare File"}, _
        {"Link Name Datei", "Link Name File", "Enlace de nombre de archivo", "Nom du lien du Fichier", "Link con il nome del file"}, _
        {"Dateiname", "File Name", "Nombre de archivo", "Nom du Fichier", "Nome del file"}, _
        {"Löschen Sie das ausgewählte Dokument", "Delete the selected document", "Eliminar el documento seleccionado", "Effacer le document sélectionné", "Eliminare il documento selezionato"}, _
        {"übertragen", "Transfer", "Transferir", "Transfert", "Trasferire"}, _
        {"Neues Dokument", "New Document", "Nuevo Documento", "Nouveau Document", "Nuovo Documento"}
        }

        ''        {"", "", "", "", ""}, _

        Message = Table1(IDMessage, Langue)


        Return Message
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
        LabelDataBaseName.Text = ObjAModifier(1, LangueLogiciel)
        LabelEnregistrement.Text = ObjAModifier(2, LangueLogiciel)
        LabelChrono.Text = ObjAModifier(3, LangueLogiciel)
        LabelTypeDocumentMaintenance.Text = ObjAModifier(4, LangueLogiciel)
        LabelDestinationMaintenance.Text = ObjAModifier(5, LangueLogiciel)
        ButtonDestinationLocalMaintenance.Text = ObjAModifier(6, LangueLogiciel)
        LabelUserMaintenance.Text = ObjAModifier(7, LangueLogiciel)
        RadioButtonFTPMaintenance.Text = ObjAModifier(8, LangueLogiciel)
        RadioButtonTransfertLocalMaintenance.Text = ObjAModifier(9, LangueLogiciel)
        LabelPasswordMaintenance.Text = ObjAModifier(10, LangueLogiciel)
        ButtonMiseàJourMaintenance.Text = ObjAModifier(11, LangueLogiciel)

        LabelListeDesDocuments.Text = ObjAModifier(12, LangueLogiciel)
        LabelTypeDocDocument.Text = ObjAModifier(13, LangueLogiciel)
        ButtonSelectFile.Text = ObjAModifier(14, LangueLogiciel)
        LabelNomLienDuFichier.Text = ObjAModifier(15, LangueLogiciel)
        LabelNomFichier.Text = ObjAModifier(16, LangueLogiciel)
        ButtonEffaceUnDocument.Text = ObjAModifier(17, LangueLogiciel)
        ButtonAddDocument.Text = ObjAModifier(18, LangueLogiciel)
        ButtonNouveauDocument.Text = ObjAModifier(19, LangueLogiciel)

        Return True

    End Function



    Private Sub PictureBoxGermanFlag_Click(sender As System.Object, e As System.EventArgs) Handles PictureBoxGermanFlag.Click
        PictureBoxGermanFlag.BorderStyle = BorderStyle.Fixed3D
        PictureBoxEnglishgFlag.BorderStyle = BorderStyle.None
        PictureBoxSpanishFlag.BorderStyle = BorderStyle.None
        PictureBoxFrenchFlag.BorderStyle = BorderStyle.None
        PictureBoxItalianFlag.BorderStyle = BorderStyle.None

        Me.BooleanAllemand = True
        Me.BooleanAnglais = False
        Me.BooleanEspagnol = False
        Me.BooleanFrançais = False
        Me.BooleanItalien = False


        ButtonSelectFile.Enabled = True
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = True
        ButtonAddDocument.Enabled = True


        'LabelNomFichierOrigine.Text = "---"
        'TextBoxNomDuLienDuFichier.Enabled = False
        'TextBoxNomDuLienDuFichier.Text = ""
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = False
        'TextBoxNomFichier.Text = ""
        'TextBoxNomFichier.Enabled = True



        If Me.BooleanAllemand = True Then ' Allemand

            If NouveauDocument.OrigineFichierALL.Length > 0 Then
                LabelNomFichierOrigine.Text = " - Ausgewählt - "
            Else
                LabelNomFichierOrigine.Text = "---"
            End If


            TextBoxNomDuLienDuFichier.Enabled = False
            TextBoxNomDuLienDuFichier.Text = NouveauDocument.NomDuLienFichierALL
            TextBoxNomDuLienDuFichier.Enabled = True


            TextBoxNomFichier.Enabled = False
            TextBoxNomFichier.Text = NouveauDocument.NomFichierALL
            TextBoxNomFichier.Enabled = True

        End If






    End Sub

    Private Sub PictureBoxEnglishgFlag_Click(sender As System.Object, e As System.EventArgs) Handles PictureBoxEnglishgFlag.Click
        PictureBoxGermanFlag.BorderStyle = BorderStyle.None
        PictureBoxEnglishgFlag.BorderStyle = BorderStyle.Fixed3D
        PictureBoxSpanishFlag.BorderStyle = BorderStyle.None
        PictureBoxFrenchFlag.BorderStyle = BorderStyle.None
        PictureBoxItalianFlag.BorderStyle = BorderStyle.None

        Me.BooleanAllemand = False
        Me.BooleanAnglais = True
        Me.BooleanEspagnol = False
        Me.BooleanFrançais = False
        Me.BooleanItalien = False


        ButtonSelectFile.Enabled = True
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = True
        ButtonAddDocument.Enabled = True


        'LabelNomFichierOrigine.Text = "---"
        'TextBoxNomDuLienDuFichier.Enabled = False
        'TextBoxNomDuLienDuFichier.Text = ""
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = False
        'TextBoxNomFichier.Text = ""
        'TextBoxNomFichier.Enabled = True


        If Me.BooleanAnglais = True Then ' Anglais

            If NouveauDocument.OrigineFichierANG.Length > 0 Then
                LabelNomFichierOrigine.Text = " - Selected - "
            Else
                LabelNomFichierOrigine.Text = "---"
            End If


            TextBoxNomDuLienDuFichier.Enabled = False
            TextBoxNomDuLienDuFichier.Text = NouveauDocument.NomDuLienFichierANG
            TextBoxNomDuLienDuFichier.Enabled = True


            TextBoxNomFichier.Enabled = False
            TextBoxNomFichier.Text = NouveauDocument.NomFichierANG
            TextBoxNomFichier.Enabled = True

        End If

    End Sub

    Private Sub PictureBoxSpanishFlag_Click(sender As System.Object, e As System.EventArgs) Handles PictureBoxSpanishFlag.Click
        PictureBoxGermanFlag.BorderStyle = BorderStyle.None
        PictureBoxEnglishgFlag.BorderStyle = BorderStyle.None
        PictureBoxSpanishFlag.BorderStyle = BorderStyle.Fixed3D
        PictureBoxFrenchFlag.BorderStyle = BorderStyle.None
        PictureBoxItalianFlag.BorderStyle = BorderStyle.None

        Me.BooleanAllemand = False
        Me.BooleanAnglais = False
        Me.BooleanEspagnol = True
        Me.BooleanFrançais = False
        Me.BooleanItalien = False


        ButtonSelectFile.Enabled = True
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = True
        ButtonAddDocument.Enabled = True


        'LabelNomFichierOrigine.Text = "---"
        'TextBoxNomDuLienDuFichier.Enabled = False
        'TextBoxNomDuLienDuFichier.Text = ""
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = False
        'TextBoxNomFichier.Text = ""
        'TextBoxNomFichier.Enabled = True


        If Me.BooleanEspagnol = True Then ' Espagnol

            If NouveauDocument.OrigineFichierESP.Length > 0 Then
                LabelNomFichierOrigine.Text = " - Seleccionado - "
            Else
                LabelNomFichierOrigine.Text = "---"
            End If

            TextBoxNomDuLienDuFichier.Enabled = False
            TextBoxNomDuLienDuFichier.Text = NouveauDocument.NomDuLienFichierESP
            TextBoxNomDuLienDuFichier.Enabled = True


            TextBoxNomFichier.Enabled = False
            TextBoxNomFichier.Text = NouveauDocument.NomFichierESP
            TextBoxNomFichier.Enabled = True

        End If

    End Sub

    Private Sub PictureBoxFrenchFlag_Click(sender As System.Object, e As System.EventArgs) Handles PictureBoxFrenchFlag.Click
        PictureBoxGermanFlag.BorderStyle = BorderStyle.None
        PictureBoxEnglishgFlag.BorderStyle = BorderStyle.None
        PictureBoxSpanishFlag.BorderStyle = BorderStyle.None
        PictureBoxFrenchFlag.BorderStyle = BorderStyle.Fixed3D
        PictureBoxItalianFlag.BorderStyle = BorderStyle.None

        Me.BooleanAllemand = False
        Me.BooleanAnglais = False
        Me.BooleanEspagnol = False
        Me.BooleanFrançais = True
        Me.BooleanItalien = False


        ButtonSelectFile.Enabled = True
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = True
        ButtonAddDocument.Enabled = True


        'LabelNomFichierOrigine.Text = "---"
        'TextBoxNomDuLienDuFichier.Enabled = False
        'TextBoxNomDuLienDuFichier.Text = ""
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = False
        'TextBoxNomFichier.Text = ""
        'TextBoxNomFichier.Enabled = True


        If Me.BooleanFrançais = True Then ' Français

            If NouveauDocument.OrigineFichierFRA.Length > 0 Then
                LabelNomFichierOrigine.Text = " - Selectionné - "
            Else
                LabelNomFichierOrigine.Text = "---"
            End If

            TextBoxNomDuLienDuFichier.Enabled = False
            TextBoxNomDuLienDuFichier.Text = NouveauDocument.NomDuLienFichierFRA
            TextBoxNomDuLienDuFichier.Enabled = True


            TextBoxNomFichier.Enabled = False
            TextBoxNomFichier.Text = NouveauDocument.NomFichierFRA
            TextBoxNomFichier.Enabled = True

        End If

    End Sub

    Private Sub PictureBoxItalianFlag_Click(sender As System.Object, e As System.EventArgs) Handles PictureBoxItalianFlag.Click
        PictureBoxGermanFlag.BorderStyle = BorderStyle.None
        PictureBoxEnglishgFlag.BorderStyle = BorderStyle.None
        PictureBoxSpanishFlag.BorderStyle = BorderStyle.None
        PictureBoxFrenchFlag.BorderStyle = BorderStyle.None
        PictureBoxItalianFlag.BorderStyle = BorderStyle.Fixed3D

        Me.BooleanAllemand = False
        Me.BooleanAnglais = False
        Me.BooleanEspagnol = False
        Me.BooleanFrançais = False
        Me.BooleanItalien = True


        ButtonSelectFile.Enabled = True
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = True
        ButtonAddDocument.Enabled = True


        'LabelNomFichierOrigine.Text = "---"
        'TextBoxNomDuLienDuFichier.Enabled = False
        'TextBoxNomDuLienDuFichier.Text = ""
        'TextBoxNomDuLienDuFichier.Enabled = True
        'TextBoxNomFichier.Enabled = False
        'TextBoxNomFichier.Text = ""
        'TextBoxNomFichier.Enabled = True

        If Me.BooleanItalien = True Then ' Italien

            If NouveauDocument.OrigineFichierITA.Length > 0 Then
                LabelNomFichierOrigine.Text = " - Selezionato - "
            Else
                LabelNomFichierOrigine.Text = "---"
            End If


            TextBoxNomDuLienDuFichier.Enabled = False
            TextBoxNomDuLienDuFichier.Text = NouveauDocument.NomDuLienFichierITA
            TextBoxNomDuLienDuFichier.Enabled = True


            TextBoxNomFichier.Enabled = False
            TextBoxNomFichier.Text = NouveauDocument.NomFichierITA
            TextBoxNomFichier.Enabled = True


        End If

    End Sub


    Private Function RequêteSQLChargementTypeDocument(ByVal Auteur As String, ByVal ChaineDeConnexion As String, ByVal NomDeLaBaseDeDonnées As String, ByVal LangueDuSoft As Integer)

        Dim Retour As New RetourChargementTypeDocument
        Dim Infos_sur_la_requête_ChargementTypeDocument As String = ""
        Infos_sur_la_requête_ChargementTypeDocument += "Connexion : " & ChaineDeConnexion & vbCrLf

        Try



            'on va effacer la liste des types documents
            ListOfTypesDocuments.Clear()


            Dim Requête_ChargementTypeDocument1 As String = ""

            Requête_ChargementTypeDocument1 += "SELECT INDICE_AUTO,Description_francais,Description_anglais,Description_allemand,Description_espagnole,Description_italien " & _
                                                " FROM " & NomDeLaBaseDeDonnées & ".types_documents;"

            Infos_sur_la_requête_ChargementTypeDocument += vbCrLf & " RequêteSQLChargementTypeDocument()" & vbCrLf
            Infos_sur_la_requête_ChargementTypeDocument += Requête_ChargementTypeDocument1 & vbCrLf

            Dim Mysql_connection As New MySql.Data.MySqlClient.MySqlConnection(ChaineDeConnexion)
            Mysql_connection.Open()
            Dim myCommand1 As New MySql.Data.MySqlClient.MySqlCommand(Requête_ChargementTypeDocument1, Mysql_connection)

            ' Structure TypeDocument
            'Provient de la table 'types_documents'
            'Dim IDTypeDocument As String
            'Dim DescriptionTypeDocument As String

            'Provient de la table 'Transfert'
            'Dim TypeDocument As String
            'Dim TypeTransfert As Integer
            'Dim Destination As String
            'Dim User As String
            'Dim MotDePasse As String


            Dim myReader As MySql.Data.MySqlClient.MySqlDataReader
            Dim StrIDTypeDocument As String = ""
            Dim StrDescriptionTypeDocument As String = ""

            myReader = myCommand1.ExecuteReader()

            While myReader.Read()
                Infos_sur_la_requête_ChargementTypeDocument += myReader.GetString(0) & vbCrLf


                StrIDTypeDocument = myReader.GetString("INDICE_AUTO")


                Select Case LangueDuSoft
                    Case 0 'Allemand
                        StrDescriptionTypeDocument = myReader.GetString("Description_allemand")
                    Case 1 'Anglais
                        StrDescriptionTypeDocument = myReader.GetString("Description_anglais")
                    Case 2 'Espagnol
                        StrDescriptionTypeDocument = myReader.GetString("Description_espagnole")
                    Case 3 'Français
                        StrDescriptionTypeDocument = myReader.GetString("Description_francais")
                    Case 4 'Italien
                        StrDescriptionTypeDocument = myReader.GetString("Description_italien")
                End Select


                Dim item01 As New ClasseTypesDocuments
                item01.IDTypeDocument = StrIDTypeDocument
                item01.DescriptionTypeDocument = StrDescriptionTypeDocument

                item01.TypeDocument = "0"
                item01.TypeTransfert = 0
                item01.User = ""
                item01.MotDePasse = ""
                item01.Destination = ""

                ListOfTypesDocuments.Add(item01)



            End While

            myReader.Close()

            Infos_sur_la_requête_ChargementTypeDocument += vbCrLf & " RequêteSQLChargementTypeDocument() - Chargement Infos table 'types_documents' OK." & vbCrLf

            Dim Requête_ChargementTypeDocument2 As String = ""


            If (ListOfTypesDocuments.Count = 0) Then GoTo Fin


            Dim I01 As Integer = 0
            Dim myCommand2 As New MySql.Data.MySqlClient.MySqlCommand
            myCommand2.Connection = Mysql_connection



            For I01 = 0 To (ListOfTypesDocuments.Count - 1)

                Requête_ChargementTypeDocument2 = "SELECT TypeDocument,TypeTransfert,Destination,User,Password " & _
                                        " FROM " & NomDeLaBaseDeDonnées & ".Transfert WHERE TypeDocument = " & ListOfTypesDocuments.Item(I01).IDTypeDocument & ";"
                myCommand2.CommandText = Requête_ChargementTypeDocument2
                Dim myReader2 As MySql.Data.MySqlClient.MySqlDataReader = Nothing
                myReader2 = myCommand2.ExecuteReader()

                If myReader2.HasRows = False Then
                    GoTo line1
                End If



                While myReader2.Read()

                    'If (myReader2.IsDBNull("TypeDocument") = True) Then Exit While
                    ListOfTypesDocuments.Item(I01).TypeDocument = myReader2.GetString("TypeDocument")
                    ListOfTypesDocuments.Item(I01).TypeTransfert = myReader2.GetInt32("TypeTransfert")
                    ListOfTypesDocuments.Item(I01).Destination = myReader2.GetString("Destination")
                    ListOfTypesDocuments.Item(I01).User = myReader2.GetString("User")
                    ListOfTypesDocuments.Item(I01).MotDePasse = myReader2.GetString("Password")

                End While

line1:          myReader2.Close()

            Next


            Infos_sur_la_requête_ChargementTypeDocument += vbCrLf & " RequêteSQLChargementTypeDocument() - Chargement Infos table 'Transfert' OK." & vbCrLf

fin:        Mysql_connection.Close()
            Infos_sur_la_requête_ChargementTypeDocument += vbCrLf & "Déconnexion." & vbCrLf

        Catch ex As MySql.Data.MySqlClient.MySqlException
            'MessageBox.Show(ex.ErrorCode & " : " & ex.Message, "Requêtes_SQL_Recherche_simplifiée() - Error - ")
            Infos_sur_la_requête_ChargementTypeDocument += " : " & ex.Message & " RequêteSQLChargementTypeDocument() - Error - " & vbCrLf

            'test si la connection mysql est toujours ouverte
            'If Mysql_connection.State = ConnectionState.Open Then
            'si oui on la ferme
            'Mysql_connection.Close()
            'End If

            GoTo erreur_fin
        End Try
        Retour.Erreur = False
        Retour.Infos_requêtes = Infos_sur_la_requête_ChargementTypeDocument

        Return Retour

erreur_fin:
        Retour.Erreur = True
        Retour.Infos_requêtes = Infos_sur_la_requête_ChargementTypeDocument
        Return Retour
    End Function

    Private Function RequêteSQLAjoutDocument(ByVal Chrono As String, ByVal ChaineDeConnexion As String, ByVal NomDeLaBaseDeDonnées As String, _
                                                    ByVal LangueDuSoft As Integer, ByVal Auteur As String)


        Dim Crypto1 = New CryptoUtil(MyKey, MyIV)
        Dim Retour As New RetourAjoutDocument
        Dim Infos_sur_la_requête_AjoutDocument As String = ""
        Dim UtilisateurCrypté As String = ""
        Dim MotDePasseCrypté As String = ""
        Dim Requête_AjoutDocument As String = ""

        Infos_sur_la_requête_AjoutDocument += "Connexion : " & ChaineDeConnexion & vbCrLf

        Try

            ' INSERT INTO table_name ( field1, field2,...fieldN )
            ' VALUES                 ( value1, value2,...valueN );


            Requête_AjoutDocument = "INSERT INTO " & NomDeLaBaseDeDonnées & ".documents (Chrono,Nom_Lien_francais,Nom_Lien_anglais,Nom_Lien_allemand,Nom_Lien_espagnole," & _
                                    "Nom_Lien_italien,Type_Doc," & _
                                    "Nom_Fichier_francais,Nom_Fichier_anglais,Nom_Fichier_allemand,Nom_Fichier_espagnole,Nom_Fichier_italien," & _
                                    "Taille_fichier_Fr,Taille_fichier_anglais,Taille_fichier_allemand,Taille_fichier_espagnole,Taille_fichier_italien" & _
                                    ") VALUES " & _
                                    " (" & NouveauDocument.Chrono & ",'" & NouveauDocument.NomDuLienFichierFRA & "','" & NouveauDocument.NomDuLienFichierANG & "','" & NouveauDocument.NomDuLienFichierALL & "','" & NouveauDocument.NomDuLienFichierESP & "'," & _
                                    "'" & NouveauDocument.NomDuLienFichierITA & "'," & NouveauDocument.NuméroTypeDocument & "," & _
                                    "'" & NouveauDocument.NomFichierFRA & "','" & NouveauDocument.NomFichierANG & "','" & NouveauDocument.NomFichierALL & "','" & NouveauDocument.NomFichierESP & "','" & NouveauDocument.NomFichierITA & "'," & _
                                    "'" & NouveauDocument.TailleFichierFRA & "','" & NouveauDocument.TailleFichierANG & "','" & NouveauDocument.TailleFichierALL & "','" & NouveauDocument.TailleFichierESP & "','" & NouveauDocument.TailleFichierITA & "'" & _
                                    ");"

            Infos_sur_la_requête_AjoutDocument += vbCrLf & " RequêteSQLAjoutDocument() - INSERT -" & vbCrLf
            Infos_sur_la_requête_AjoutDocument += Requête_AjoutDocument & vbCrLf

            Dim Mysql_connection As New MySql.Data.MySqlClient.MySqlConnection(ChaineDeConnexion)
            Mysql_connection.Open()

            Dim myCommand As New MySql.Data.MySqlClient.MySqlCommand(Requête_AjoutDocument, Mysql_connection)
            Dim NbreLignaAjoutée As Integer = 0
            NbreLignaAjoutée = myCommand.ExecuteNonQuery()


            Infos_sur_la_requête_AjoutDocument += vbCrLf & " RequêteSQLAjoutDocument() - INSERT OK -" & CStr(NbreLignaAjoutée) & vbCrLf



            Dim User_ID_Utilisateur As Integer = 0
            Dim RequêteRécupérationUserID As String = ""
            Dim hostNameOrAddress As String = "localhost"
            Dim returnValue As IPHostEntry
            Dim AdresseIP As String = ""
            Dim i0 As Integer = 0
            Dim RetourEcritureLog As New RetourLog

            returnValue = Dns.GetHostEntry(hostNameOrAddress)

            ' Nom de la machine
            Dim NomMachine As String = Dns.GetHostName

            ' Récupération de la liste des IP de la machine
            Dim InfoIps As IPHostEntry = Dns.GetHostEntry(NomMachine)
            Dim MesIp As IPAddress() = InfoIps.AddressList


            For Each CurrentIp As IPAddress In MesIp
                AdresseIP += CurrentIp.ToString & "; "
            Next

            If AdresseIP.Length > 254 Then AdresseIP = AdresseIP.Substring(0, 254)


            RequêteRécupérationUserID = "SELECT User_ID FROM " & NomDeLaBaseDeDonnées & ".Users WHERE  User_Name = '" & Auteur & "';"
            Infos_sur_la_requête_AjoutDocument += vbCrLf & " RequêteSQLAjoutDocument() - Requête RécupérationUserID" & vbCrLf
            Infos_sur_la_requête_AjoutDocument += RequêteRécupérationUserID & vbCrLf

            Dim myCommand4 As New MySql.Data.MySqlClient.MySqlCommand(RequêteRécupérationUserID, Mysql_connection)
            Dim myReader1 As MySql.Data.MySqlClient.MySqlDataReader
            myReader1 = myCommand4.ExecuteReader()


            myReader1.Read()
            User_ID_Utilisateur = myReader1.GetInt16(0)

            myReader1.Close()

            'Phase 3 - On créé un Log de la création du document dans la table documents


            RetourEcritureLog = RequêteSQLLog(User_ID_Utilisateur, AdresseIP, returnValue.HostName, My.Computer.Info.OSFullName, _
                                              "Add", "documents", Chrono, Requête_AjoutDocument, "1", ChaineDeConnexion, NomDeLaBaseDeDonnées)


            myCommand4.Dispose()



fin:        Mysql_connection.Close()
            Infos_sur_la_requête_AjoutDocument += vbCrLf & "Déconnexion." & vbCrLf


            If (NbreLignaAjoutée > 0) Then

                ' on a réussie à enregistrer le document dans la table 'documents'
                ' On poursuit avec le transfert des fichiers vers le serveur soit en FTP ou en local
                ' si on a 5 fois le même fichier on le transmet une seule fois.

                If NouveauDocument.TypeTransfert = 1 Then
                    'Transfert par FTP

                    Dim NomUtilisateurDécrypté As String = ""
                    NomUtilisateurDécrypté = ConvertBase64ToString(NouveauDocument.Utilisateur)
                    NomUtilisateurDécrypté = Crypto1.DecryptText(NomUtilisateurDécrypté)

                    Dim MotDePasseDécrypté As String = ""
                    MotDePasseDécrypté = ConvertBase64ToString(NouveauDocument.MotDePasse)
                    MotDePasseDécrypté = Crypto1.DecryptText(MotDePasseDécrypté)

                    Dim RetourFTP As Boolean = False

                    If (NouveauDocument.OrigineFichierALL = NouveauDocument.OrigineFichierANG And _
                        NouveauDocument.OrigineFichierANG = NouveauDocument.OrigineFichierESP And _
                        NouveauDocument.OrigineFichierESP = NouveauDocument.OrigineFichierFRA And _
                        NouveauDocument.OrigineFichierFRA = NouveauDocument.OrigineFichierITA) Then

                        'Tout les fichiers sont les même on transfert une seule fois 
                        RetourFTP = UploadFileByFTP(NouveauDocument.Destination, NouveauDocument.NomFichierFRA, NouveauDocument.OrigineFichierFRA, NouveauDocument.NomFichierFRA, _
                                                    NomUtilisateurDécrypté, MotDePasseDécrypté)
                        If RetourFTP = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                            End Select


                        End If

                    Else
                        'Tout les fichiers sont différents on transfert une cinq fois 

                        'FTP --- ALLEMAND ---
                        RetourFTP = UploadFileByFTP(NouveauDocument.Destination, NouveauDocument.NomFichierALL, NouveauDocument.OrigineFichierALL, NouveauDocument.NomFichierALL, _
                                                    NomUtilisateurDécrypté, MotDePasseDécrypté)
                        If RetourFTP = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (FTP)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (FTP)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (FTP)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (FTP)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (FTP)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                            End Select

                        End If


                        'FTP --- ANGLAIS ---
                        RetourFTP = UploadFileByFTP(NouveauDocument.Destination, NouveauDocument.NomFichierANG, NouveauDocument.OrigineFichierANG, NouveauDocument.NomFichierANG, _
                            NomUtilisateurDécrypté, MotDePasseDécrypté)
                        If RetourFTP = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (FTP)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (FTP)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (FTP)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (FTP)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (FTP)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                            End Select

                        End If


                        'FTP --- ESPAGNOL ---
                        RetourFTP = UploadFileByFTP(NouveauDocument.Destination, NouveauDocument.NomFichierESP, NouveauDocument.OrigineFichierESP, NouveauDocument.NomFichierESP, _
                            NomUtilisateurDécrypté, MotDePasseDécrypté)
                        If RetourFTP = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (FTP)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (FTP)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (FTP)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (FTP)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (FTP)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                            End Select

                        End If



                        'FTP --- FRANCAIS ---
                        RetourFTP = UploadFileByFTP(NouveauDocument.Destination, NouveauDocument.NomFichierFRA, NouveauDocument.OrigineFichierFRA, NouveauDocument.NomFichierFRA, _
                                  NomUtilisateurDécrypté, MotDePasseDécrypté)
                        If RetourFTP = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (FTP)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                            End Select

                        End If



                        'FTP --- ITALIEN ---
                        RetourFTP = UploadFileByFTP(NouveauDocument.Destination, NouveauDocument.NomFichierITA, NouveauDocument.OrigineFichierITA, NouveauDocument.NomFichierITA, _
                                  NomUtilisateurDécrypté, MotDePasseDécrypté)
                        If RetourFTP = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (FTP)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (FTP)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (FTP)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (FTP)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (FTP)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                            End Select

                        End If


                End If
                End If

                If NouveauDocument.TypeTransfert = 2 Then
                    'Transfert en local
                    Dim RetourCopieLocale As Boolean = False

                    If (NouveauDocument.OrigineFichierALL = NouveauDocument.OrigineFichierANG And _
                        NouveauDocument.OrigineFichierANG = NouveauDocument.OrigineFichierESP And _
                        NouveauDocument.OrigineFichierESP = NouveauDocument.OrigineFichierFRA And _
                        NouveauDocument.OrigineFichierFRA = NouveauDocument.OrigineFichierITA) Then

                        If File.Exists(NouveauDocument.Destination & NouveauDocument.NomFichierFRA) = True Then
                            System.IO.File.Delete(NouveauDocument.Destination & NouveauDocument.NomFichierFRA)
                        End If

                        'les cinq fichiers sont identiques on va en copier qu'un seul

                        'si le fichier existe déjà on le supprime
                        RetourCopieLocale = UploadLocalFile(NouveauDocument.OrigineFichierFRA, NouveauDocument.Destination, NouveauDocument.NomFichierFRA)

                        If RetourCopieLocale = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                            End Select


                        End If

                    Else
                        'les cinq fichiers sont différents on va en copier 5 fois 

                        ' --- COPIE LOCALE ALLEMAND ---

                        If File.Exists(NouveauDocument.Destination & NouveauDocument.NomFichierALL) = True Then
                            System.IO.File.Delete(NouveauDocument.Destination & NouveauDocument.NomFichierALL)
                        End If

                        RetourCopieLocale = UploadLocalFile(NouveauDocument.OrigineFichierALL, NouveauDocument.Destination, NouveauDocument.NomFichierALL)

                        If RetourCopieLocale = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (Local)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (Local)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (Local)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (Local)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (Local)" & vbCrLf & NouveauDocument.NomFichierALL, MsgBoxStyle.Critical)
                            End Select
                        End If


                        ' --- COPIE LOCALE ANGLAIS ---


                        If File.Exists(NouveauDocument.Destination & NouveauDocument.NomFichierANG) = True Then
                            System.IO.File.Delete(NouveauDocument.Destination & NouveauDocument.NomFichierANG)
                        End If

                        RetourCopieLocale = UploadLocalFile(NouveauDocument.OrigineFichierANG, NouveauDocument.Destination, NouveauDocument.NomFichierANG)

                        If RetourCopieLocale = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (Local)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (Local)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (Local)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (Local)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (Local)" & vbCrLf & NouveauDocument.NomFichierANG, MsgBoxStyle.Critical)
                            End Select
                        End If


                        ' --- COPIE LOCALE ESPAGNOL ---


                        If File.Exists(NouveauDocument.Destination & NouveauDocument.NomFichierESP) = True Then
                            System.IO.File.Delete(NouveauDocument.Destination & NouveauDocument.NomFichierESP)
                        End If

                        RetourCopieLocale = UploadLocalFile(NouveauDocument.OrigineFichierESP, NouveauDocument.Destination, NouveauDocument.NomFichierESP)

                        If RetourCopieLocale = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (Local)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (Local)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (Local)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (Local)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (Local)" & vbCrLf & NouveauDocument.NomFichierESP, MsgBoxStyle.Critical)
                            End Select
                        End If


                        ' --- COPIE LOCALE FRANCAIS ---


                        If File.Exists(NouveauDocument.Destination & NouveauDocument.NomFichierFRA) = True Then
                            System.IO.File.Delete(NouveauDocument.Destination & NouveauDocument.NomFichierFRA)
                        End If

                        RetourCopieLocale = UploadLocalFile(NouveauDocument.OrigineFichierFRA, NouveauDocument.Destination, NouveauDocument.NomFichierFRA)

                        If RetourCopieLocale = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (Local)" & vbCrLf & NouveauDocument.NomFichierFRA, MsgBoxStyle.Critical)
                            End Select
                        End If


                        ' --- COPIE LOCALE ITALIEN ---


                        If File.Exists(NouveauDocument.Destination & NouveauDocument.NomFichierITA) = True Then
                            System.IO.File.Delete(NouveauDocument.Destination & NouveauDocument.NomFichierITA)
                        End If

                        RetourCopieLocale = UploadLocalFile(NouveauDocument.OrigineFichierITA, NouveauDocument.Destination, NouveauDocument.NomFichierITA)

                        If RetourCopieLocale = False Then

                            Select Case LangueDuSoft
                                Case 0 ' Allemand
                                    MsgBox("Ein Fehler ist aufgetreten während der Übertragung. (Local)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 1 ' Anglais
                                    MsgBox("An error occurred during transfer. (Local)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 2 ' Espagnol
                                    MsgBox("Se produjo un error durante la transferencia. (Local)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 3 ' Français
                                    MsgBox("Une erreur s'est produite durant le transfert. (Local)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                                Case 4 ' Italien
                                    MsgBox("È verificato un errore durante il trasferimento. (Local)" & vbCrLf & NouveauDocument.NomFichierITA, MsgBoxStyle.Critical)
                            End Select
                        End If



                    End If

                End If


            End If


        Catch ex As MySql.Data.MySqlClient.MySqlException
            'MessageBox.Show(ex.ErrorCode & " : " & ex.Message, "Requêtes_SQL_Recherche_simplifiée() - Error - ")
            Infos_sur_la_requête_AjoutDocument += " : " & ex.Message & " RequêteSQLAjoutDocument() - Error - " & vbCrLf

            'test si la connection mysql est toujours ouverte
            'If Mysql_connection.State = ConnectionState.Open Then
            'si oui on la ferme
            'Mysql_connection.Close()
            'End If

            GoTo erreur_fin
        End Try
        Retour.Erreur = False
        Retour.Infos_requêtes = Infos_sur_la_requête_AjoutDocument

        Return Retour

erreur_fin:
        Retour.Erreur = True
        Retour.Infos_requêtes = Infos_sur_la_requête_AjoutDocument
        Return Retour

    End Function


    Private Function RequêteSQLMaintenanceTransfert(ByVal IDTypeDocument As Integer, ByVal TypeDocument As Integer, ByVal TypeTransfert As Integer, ByVal Destination As String, _
                                                    ByVal Utilisateur As String, ByVal MotDePasse As String, ByVal ChaineDeConnexion As String, ByVal NomDeLaBaseDeDonnées As String, _
                                                    ByVal LangueDuSoft As Integer, ByVal Auteur As String)

        Dim Crypto1 = New CryptoUtil(MyKey, MyIV)
        Dim Retour As New RetourMaintenanceTransfert
        Dim Infos_sur_la_requête_MaintenanceTransfert As String = ""
        Dim UtilisateurCrypté As String = ""
        Dim MotDePasseCrypté As String = ""
        Dim Requête_MaintenanceTransfert As String = ""

        Infos_sur_la_requête_MaintenanceTransfert += "Connexion : " & ChaineDeConnexion & vbCrLf

        Try

            UtilisateurCrypté = Crypto1.EncryptText(Utilisateur)
            UtilisateurCrypté = ConvertStringToBase64(UtilisateurCrypté)

            MotDePasseCrypté = Crypto1.EncryptText(MotDePasse)
            MotDePasseCrypté = ConvertStringToBase64(MotDePasseCrypté)

            Destination = Destination.Replace("\", "\\")

            Dim Mysql_connection As New MySql.Data.MySqlClient.MySqlConnection(ChaineDeConnexion)
            Mysql_connection.Open()

            If TypeDocument = 0 Then

                ' INSERT INTO table_name ( field1, field2,...fieldN )
                ' VALUES                 ( value1, value2,...valueN );
                Requête_MaintenanceTransfert = "INSERT INTO " & NomDeLaBaseDeDonnées & ".Transfert (TypeDocument,TypeTransfert,Destination,User,Password) VALUES " & _
                                               " (" & IDTypeDocument & "," & TypeTransfert & ",'" & Destination & "','" & UtilisateurCrypté & "','" & MotDePasseCrypté & "');"


                Infos_sur_la_requête_MaintenanceTransfert += vbCrLf & " RequêteSQLMaintenanceTransfert() - INSERT -" & vbCrLf
                Infos_sur_la_requête_MaintenanceTransfert += Requête_MaintenanceTransfert & vbCrLf

                Dim myCommand As New MySql.Data.MySqlClient.MySqlCommand(Requête_MaintenanceTransfert, Mysql_connection)
                Dim NbreLignaAjoutée As Integer = 0
                NbreLignaAjoutée = myCommand.ExecuteNonQuery()

                myCommand.Dispose()


            Else
                'Update

                'Active les mises à jour
                'SET SQL_SAFE_UPDATES=0;
                Dim requete_Activation_mise_à_jour As String = "SET SQL_SAFE_UPDATES=0;"


                Dim myCommand2 As New MySql.Data.MySqlClient.MySqlCommand(requete_Activation_mise_à_jour, Mysql_connection)
                Dim NbreLignaAjoutée2 As Integer = 0
                NbreLignaAjoutée2 = myCommand2.ExecuteNonQuery()

                myCommand2.Dispose()

                'UPDATE [LOW_PRIORITY] [IGNORE] tbl_name
                'SET col_name1=expr1 [, col_name2=expr2 ...]
                '[WHERE where_definition]
                '[ORDER BY ...]
                '[LIMIT row_count]

                Requête_MaintenanceTransfert = "UPDATE " & NomDeLaBaseDeDonnées & ".Transfert SET TypeTransfert = " & CStr(TypeTransfert) & _
                                               ", Destination =  '" & Destination & "'" & _
                                               ", User = '" & UtilisateurCrypté & "'" & _
                                               ", Password = '" & MotDePasseCrypté & "'" & _
                                               " WHERE TypeDocument= " & TypeDocument & ";"

                Infos_sur_la_requête_MaintenanceTransfert += vbCrLf & " RequêteSQLMaintenanceTransfert() - UPDATE -" & vbCrLf
                Infos_sur_la_requête_MaintenanceTransfert += Requête_MaintenanceTransfert & vbCrLf


                Dim myCommand3 As New MySql.Data.MySqlClient.MySqlCommand(Requête_MaintenanceTransfert, Mysql_connection)
                Dim NbreLignaAjoutée3 As Integer = 0
                NbreLignaAjoutée3 = myCommand3.ExecuteNonQuery()

                myCommand3.Dispose()

            End If


            Dim User_ID_Utilisateur As Integer = 0
            Dim RequêteRécupérationUserID As String = ""
            Dim hostNameOrAddress As String = "localhost"
            Dim returnValue As IPHostEntry
            Dim AdresseIP As String = ""
            Dim i0 As Integer = 0
            Dim RetourEcritureLog As New RetourLog

            returnValue = Dns.GetHostEntry(hostNameOrAddress)

            ' Nom de la machine
            Dim NomMachine As String = Dns.GetHostName

            ' Récupération de la liste des IP de la machine
            Dim InfoIps As IPHostEntry = Dns.GetHostEntry(NomMachine)
            Dim MesIp As IPAddress() = InfoIps.AddressList


            For Each CurrentIp As IPAddress In MesIp
                AdresseIP += CurrentIp.ToString & "; "
            Next

            If AdresseIP.Length > 254 Then AdresseIP = AdresseIP.Substring(0, 254)


            RequêteRécupérationUserID = "SELECT User_ID FROM " & NomDeLaBaseDeDonnées & ".Users WHERE  User_Name = '" & Auteur & "';"
            Infos_sur_la_requête_MaintenanceTransfert += vbCrLf & " RequêteSQLMaintenanceTransfert() - Requête RécupérationUserID" & vbCrLf
            Infos_sur_la_requête_MaintenanceTransfert += RequêteRécupérationUserID & vbCrLf

            Dim myCommand4 As New MySql.Data.MySqlClient.MySqlCommand(RequêteRécupérationUserID, Mysql_connection)
            Dim myReader1 As MySql.Data.MySqlClient.MySqlDataReader
            myReader1 = myCommand4.ExecuteReader()


            myReader1.Read()
            User_ID_Utilisateur = myReader1.GetInt16(0)

            myReader1.Close()

            'Phase 3 - On créé un Log de la création ou mise à jour de la table Transfert

            Dim Action As String = ""

            If TypeDocument = 0 Then
                Action = "Add"
            End If

            If TypeDocument > 0 Then
                Action = "Modify"
            End If

            RetourEcritureLog = RequêteSQLLog(User_ID_Utilisateur, AdresseIP, returnValue.HostName, My.Computer.Info.OSFullName, _
                                              Action, "Transfert", Chrono, Requête_MaintenanceTransfert, "1", ChaineDeConnexion, NomDeLaBaseDeDonnées)


            myCommand4.Dispose()



fin:        Mysql_connection.Close()
            Infos_sur_la_requête_MaintenanceTransfert += vbCrLf & "Déconnexion." & vbCrLf

        Catch ex As MySql.Data.MySqlClient.MySqlException
            'MessageBox.Show(ex.ErrorCode & " : " & ex.Message, "Requêtes_SQL_Recherche_simplifiée() - Error - ")
            Infos_sur_la_requête_MaintenanceTransfert += " : " & ex.Message & " RequêteSQLMaintenanceTransfert() - Error - " & vbCrLf

            'test si la connection mysql est toujours ouverte
            'If Mysql_connection.State = ConnectionState.Open Then
            'si oui on la ferme
            'Mysql_connection.Close()
            'End If

            GoTo erreur_fin
        End Try
        Retour.Erreur = False
        Retour.Infos_requêtes = Infos_sur_la_requête_MaintenanceTransfert

        Return Retour

erreur_fin:
        Retour.Erreur = True
        Retour.Infos_requêtes = Infos_sur_la_requête_MaintenanceTransfert
        Return Retour

    End Function


    Private Function RequêteSQLChargementListeDocuments(ByVal Chrono As String, ByVal ChaineDeConnexion As String, ByVal NomDeLaBaseDeDonnées As String, ByVal LangueDuSoft As Integer)

        Dim Crypto1 = New CryptoUtil(MyKey, MyIV)
        Dim Retour As New RetourChargementListeDocuments
        Dim Infos_sur_la_requête_ChargementListeDocuments As String = ""

        Infos_sur_la_requête_ChargementListeDocuments += "Connexion : " & ChaineDeConnexion & vbCrLf

        Try


            'Liste des documents lié au chrono
            'Utilisation dans ListBoxDocuments
            ' On  réinitialise
            ListeDesDocumentsPourLeChrono.Clear()


            Dim Requête_ChargementListeDocuments As String = ""

            Requête_ChargementListeDocuments = "SELECT Nom_Lien_francais,Nom_Lien_anglais,Nom_Lien_allemand,Nom_Lien_espagnole,Nom_Lien_italien,Type_Doc," & _
                                                "Nom_Fichier_francais,Nom_Fichier_anglais,Nom_Fichier_allemand,Nom_Fichier_espagnole,Nom_Fichier_italien," & _
                                                "Taille_fichier_Fr,Taille_fichier_anglais,Taille_fichier_allemand,Taille_fichier_espagnole,Taille_fichier_italien" & _
                                                " FROM " & NomDeLaBaseDeDonnées & ".documents  WHERE Chrono = " & Chrono & ";"

            Infos_sur_la_requête_ChargementListeDocuments += vbCrLf & " RequêteSQLChargementListeDocuments()" & vbCrLf
            Infos_sur_la_requête_ChargementListeDocuments += Requête_ChargementListeDocuments & vbCrLf

            Dim Mysql_connection As New MySql.Data.MySqlClient.MySqlConnection(ChaineDeConnexion)
            Mysql_connection.Open()
            Dim myCommand1 As New MySql.Data.MySqlClient.MySqlCommand(Requête_ChargementListeDocuments, Mysql_connection)

            Dim myReader As MySql.Data.MySqlClient.MySqlDataReader
            Dim StrNom_Lien_francais As String = ""
            Dim StrNom_Lien_anglais As String = ""
            Dim StrNom_Lien_allemand As String = ""
            Dim StrNom_Lien_espagnole As String = ""
            Dim StrNom_Lien_italien As String = ""
            Dim IntType_Doc As Integer = 0
            Dim StrNom_Fichier_francais As String = ""
            Dim StrNom_Fichier_anglais As String = ""
            Dim StrNom_Fichier_allemand As String = ""
            Dim StrNom_Fichier_espagnole As String = ""
            Dim StrNom_Fichier_italien As String = ""
            Dim StrTaille_fichier_Fr As String = ""
            Dim StrTaille_fichier_anglais As String = ""
            Dim StrTaille_fichier_allemand As String = ""
            Dim StrTaille_fichier_espagnole As String = ""
            Dim StrTaille_fichier_italien As String = ""



            myReader = myCommand1.ExecuteReader()

            While myReader.Read()

                Infos_sur_la_requête_ChargementListeDocuments += myReader.GetString("Nom_Lien_francais") & vbCrLf


                If myReader.IsDBNull(0) Then StrNom_Lien_francais = "" Else StrNom_Lien_francais = myReader.GetString("Nom_Lien_francais")
                If myReader.IsDBNull(1) Then StrNom_Lien_anglais = "" Else StrNom_Lien_anglais = myReader.GetString("Nom_Lien_anglais")
                If myReader.IsDBNull(2) Then StrNom_Lien_allemand = "" Else StrNom_Lien_allemand = myReader.GetString("Nom_Lien_allemand")
                If myReader.IsDBNull(3) Then StrNom_Lien_espagnole = "" Else StrNom_Lien_espagnole = myReader.GetString("Nom_Lien_espagnole")
                If myReader.IsDBNull(4) Then StrNom_Lien_italien = "" Else StrNom_Lien_italien = myReader.GetString("Nom_Lien_italien")

                IntType_Doc = myReader.GetUInt32("Type_Doc")

                If myReader.IsDBNull(6) Then StrNom_Fichier_francais = "" Else StrNom_Fichier_francais = myReader.GetString("Nom_Fichier_francais")
                If myReader.IsDBNull(7) Then StrNom_Fichier_anglais = "" Else StrNom_Fichier_anglais = myReader.GetString("Nom_Fichier_anglais")
                If myReader.IsDBNull(8) Then StrNom_Fichier_allemand = "" Else StrNom_Fichier_allemand = myReader.GetString("Nom_Fichier_allemand")
                If myReader.IsDBNull(9) Then StrNom_Fichier_espagnole = "" Else StrNom_Fichier_espagnole = myReader.GetString("Nom_Fichier_espagnole")
                If myReader.IsDBNull(10) Then StrNom_Fichier_italien = "" Else StrNom_Fichier_italien = myReader.GetString("Nom_Fichier_italien")

                If myReader.IsDBNull(11) Then StrTaille_fichier_Fr = "" Else StrTaille_fichier_Fr = myReader.GetString("Taille_fichier_Fr")
                If myReader.IsDBNull(12) Then StrTaille_fichier_anglais = "" Else StrTaille_fichier_anglais = myReader.GetString("Taille_fichier_anglais")
                If myReader.IsDBNull(13) Then StrTaille_fichier_allemand = "" Else StrTaille_fichier_allemand = myReader.GetString("Taille_fichier_allemand")
                If myReader.IsDBNull(14) Then StrTaille_fichier_espagnole = "" Else StrTaille_fichier_espagnole = myReader.GetString("Taille_fichier_espagnole")
                If myReader.IsDBNull(15) Then StrTaille_fichier_italien = "" Else StrTaille_fichier_italien = myReader.GetString("Taille_fichier_italien")

                Dim item01 As New ClasseDocuments
                item01.Chrono = Chrono

                item01.NomDuLienFichierFRA = StrNom_Lien_francais
                item01.NomDuLienFichierANG = StrNom_Lien_anglais
                item01.NomDuLienFichierALL = StrNom_Lien_allemand
                item01.NomDuLienFichierESP = StrNom_Lien_espagnole
                item01.NomDuLienFichierITA = StrNom_Fichier_italien

                item01.NuméroTypeDocument = CStr(IntType_Doc)

                item01.NomFichierFRA = StrNom_Fichier_francais
                item01.NomFichierANG = StrNom_Fichier_anglais
                item01.NomFichierALL = StrNom_Fichier_allemand
                item01.NomFichierESP = StrNom_Fichier_espagnole
                item01.NomFichierITA = StrNom_Fichier_italien

                item01.TailleFichierFRA = StrTaille_fichier_Fr
                item01.TailleFichierANG = StrTaille_fichier_anglais
                item01.TailleFichierALL = StrTaille_fichier_allemand
                item01.TailleFichierESP = StrTaille_fichier_espagnole
                item01.TailleFichierITA = StrTaille_fichier_italien

                item01.Destination = ""
                item01.TypeTransfert = 0
                item01.Utilisateur = ""
                item01.MotDePasse = ""
                item01.OrigineFichierALL = ""
                item01.OrigineFichierANG = ""
                item01.OrigineFichierESP = ""
                item01.OrigineFichierFRA = ""
                item01.OrigineFichierITA = ""

                ListeDesDocumentsPourLeChrono.Add(item01)

            End While

            myReader.Close()

            Infos_sur_la_requête_ChargementListeDocuments += vbCrLf & " RequêteSQLChargementListeDocuments() - Chargement Infos table 'documents' OK." & vbCrLf

            'Il faut compléter chaque ligne de ListeDesDocumentsPourLeChrono  par les informations de  ListOfTypesDocuments 

            If ListeDesDocumentsPourLeChrono.Count = 0 Then GoTo fin


            Dim I05 As Integer = 0
            Dim I06 As Integer = 0
            Dim TempUserName As String = ""
            Dim TempMotDePasse As String = ""

            For I05 = 0 To (ListeDesDocumentsPourLeChrono.Count - 1)

                For I06 = 0 To (ListOfTypesDocuments.Count - 1)
                    If (ListeDesDocumentsPourLeChrono.Item(I05).NuméroTypeDocument = ListOfTypesDocuments.Item(I06).IDTypeDocument) Then

                        ListeDesDocumentsPourLeChrono.Item(I05).TypeTransfert = ListOfTypesDocuments.Item(I06).TypeTransfert
                        ListeDesDocumentsPourLeChrono.Item(I05).Destination = ListOfTypesDocuments.Item(I06).Destination

                        If (ListOfTypesDocuments.Item(I06).User.Length = 0) Then
                            ListeDesDocumentsPourLeChrono.Item(I05).Utilisateur = ""
                        Else
                            'User n'est pas vide
                            TempUserName = ConvertBase64ToString(ListOfTypesDocuments.Item(I06).User)
                            ListeDesDocumentsPourLeChrono.Item(I05).Utilisateur = Crypto1.DecryptText(TempUserName)
                            TempUserName = ""
                        End If


                        If (ListOfTypesDocuments.Item(I06).MotDePasse.Length = 0) Then
                            ListeDesDocumentsPourLeChrono.Item(I05).MotDePasse = ""
                        Else
                            'MotDePasse n'est pas vide
                            TempMotDePasse = ConvertBase64ToString(ListOfTypesDocuments.Item(I06).MotDePasse)
                            ListeDesDocumentsPourLeChrono.Item(I05).MotDePasse = Crypto1.DecryptText(TempMotDePasse)
                            TempMotDePasse = ""
                        End If

                        ListeDesDocumentsPourLeChrono.Item(I05).NomTypeDocument = ListOfTypesDocuments.Item(I06).DescriptionTypeDocument

                    End If

                Next

            Next

            Infos_sur_la_requête_ChargementListeDocuments += vbCrLf & " RequêteSQLChargementListeDocuments() - Chargement Infos de Transfert grace à 'ListOfTypesDocuments' OK." & vbCrLf



fin:        Mysql_connection.Close()
            Infos_sur_la_requête_ChargementListeDocuments += vbCrLf & "Déconnexion." & vbCrLf

        Catch ex As MySql.Data.MySqlClient.MySqlException
            'MessageBox.Show(ex.ErrorCode & " : " & ex.Message, "Requêtes_SQL_Recherche_simplifiée() - Error - ")
            Infos_sur_la_requête_ChargementListeDocuments += " : " & ex.Message & " RequêteSQLChargementListeDocuments() - Error - " & vbCrLf

            'test si la connection mysql est toujours ouverte
            'If Mysql_connection.State = ConnectionState.Open Then
            'si oui on la ferme
            'Mysql_connection.Close()
            'End If

            GoTo erreur_fin
        End Try
        Retour.Erreur = False
        Retour.Infos_requêtes = Infos_sur_la_requête_ChargementListeDocuments

        Return Retour

erreur_fin:
        Retour.Erreur = True
        Retour.Infos_requêtes = Infos_sur_la_requête_ChargementListeDocuments
        Return Retour

    End Function

    Public Function RequêteSQLSupprimeDocument(ByVal Chrono As String, ByVal Description_Lien As String, ByVal ChaineDeConnexion As String, ByVal NomDeLaBaseDeDonnées As String, ByVal Auteur As String, ByVal LangueDuSoft As Integer)

        Dim Retour As New RetourSupprimeDocument
        Dim Infos_sur_la_requête_RequêteSQLSupprimeDocument As String = ""

        Infos_sur_la_requête_RequêteSQLSupprimeDocument += "Connexion : " & ChaineDeConnexion & vbCrLf

        Try

            ' CHAMPS Description_Lien ==>  Nom_Lien_francais,Nom_Lien_anglais,Nom_Lien_allemand,Nom_Lien_espagnole,Nom_Lien_italien

            Dim Champ_Description As String = ""

            Select Case LangueDuSoft
                Case 0 'Allemand
                    Champ_Description = "Nom_Lien_allemand"
                Case 1 'Anglais
                    Champ_Description = "Nom_Lien_anglais"
                Case 2 'Espagnol
                    Champ_Description = "Nom_Lien_espagnole"
                Case 3 'Français
                    Champ_Description = "Nom_Lien_francais"
                Case 4 'Italien
                    Champ_Description = "Nom_Lien_italien"
            End Select


            ' Syntaxe Delete
            ' DELETE [LOW_PRIORITY] [QUICK]   
            ' FROM(table_name)
            ' [WHERE conditions] [ORDER BY ...] [LIMIT rows]

            Dim Requête_SupprimeDocument As String = ""

            Requête_SupprimeDocument = "DELETE  FROM " & NomDeLaBaseDeDonnées & ".documents  WHERE Chrono = " & Chrono & " AND " & Champ_Description & " = '" & Description_Lien & "';"

            Infos_sur_la_requête_RequêteSQLSupprimeDocument += vbCrLf & " RequêteSQLSupprimeDocument()" & vbCrLf
            Infos_sur_la_requête_RequêteSQLSupprimeDocument += Infos_sur_la_requête_RequêteSQLSupprimeDocument & vbCrLf

            Dim Mysql_connection As New MySql.Data.MySqlClient.MySqlConnection(ChaineDeConnexion)
            Mysql_connection.Open()
            Dim myCommand1 As New MySql.Data.MySqlClient.MySqlCommand(Requête_SupprimeDocument, Mysql_connection)

            Dim DocSupprime As Integer = 0
            DocSupprime = myCommand1.ExecuteNonQuery

            myCommand1.Dispose()

            Infos_sur_la_requête_RequêteSQLSupprimeDocument += vbCrLf & " RequêteSQLSupprimeDocument() Chrono : " & CStr(Chrono) & " OK." & vbCrLf

            Dim User_ID_Utilisateur As Integer = 0
            Dim RequêteRécupérationUserID As String = ""
            Dim hostNameOrAddress As String = "localhost"
            Dim returnValue As IPHostEntry
            Dim AdresseIP As String = ""
            Dim i0 As Integer = 0
            Dim RetourEcritureLog As New RetourLog

            returnValue = Dns.GetHostEntry(hostNameOrAddress)

            ' Nom de la machine
            Dim NomMachine As String = Dns.GetHostName

            ' Récupération de la liste des IP de la machine
            Dim InfoIps As IPHostEntry = Dns.GetHostEntry(NomMachine)
            Dim MesIp As IPAddress() = InfoIps.AddressList


            For Each CurrentIp As IPAddress In MesIp
                AdresseIP += CurrentIp.ToString & "; "
            Next

            If AdresseIP.Length > 254 Then AdresseIP = AdresseIP.Substring(0, 254)


            RequêteRécupérationUserID = "SELECT User_ID FROM " & NomDeLaBaseDeDonnées & ".Users WHERE  User_Name = '" & Auteur & "';"
            Infos_sur_la_requête_RequêteSQLSupprimeDocument += vbCrLf & " RequêteSQLMaintenanceTransfert() - Requête RécupérationUserID" & vbCrLf
            Infos_sur_la_requête_RequêteSQLSupprimeDocument += RequêteRécupérationUserID & vbCrLf

            Dim myCommand4 As New MySql.Data.MySqlClient.MySqlCommand(RequêteRécupérationUserID, Mysql_connection)
            Dim myReader1 As MySql.Data.MySqlClient.MySqlDataReader
            myReader1 = myCommand4.ExecuteReader()


            myReader1.Read()
            User_ID_Utilisateur = myReader1.GetInt16(0)

            myReader1.Close()

            'Phase 3 - On créé un Log de la suppression du document de la table 'documents'



            RetourEcritureLog = RequêteSQLLog(User_ID_Utilisateur, AdresseIP, returnValue.HostName, My.Computer.Info.OSFullName, _
                                              "Delete", "documents", Chrono, Requête_SupprimeDocument, "1", ChaineDeConnexion, NomDeLaBaseDeDonnées)


            myCommand4.Dispose()



fin:        Mysql_connection.Close()
            Infos_sur_la_requête_RequêteSQLSupprimeDocument += vbCrLf & "Déconnexion." & vbCrLf

        Catch ex As MySql.Data.MySqlClient.MySqlException
            'MessageBox.Show(ex.ErrorCode & " : " & ex.Message, "Requêtes_SQL_Recherche_simplifiée() - Error - ")
            Infos_sur_la_requête_RequêteSQLSupprimeDocument += " : " & ex.Message & " RequêteSQLSupprimeDocument() - Error - " & vbCrLf

            'test si la connection mysql est toujours ouverte
            'If Mysql_connection.State = ConnectionState.Open Then
            'si oui on la ferme
            'Mysql_connection.Close()
            'End If

            GoTo erreur_fin
        End Try
        Retour.Erreur = False
        Retour.Infos_requêtes = Infos_sur_la_requête_RequêteSQLSupprimeDocument

        Return Retour

erreur_fin:
        Retour.Erreur = True
        Retour.Infos_requêtes = Infos_sur_la_requête_RequêteSQLSupprimeDocument
        Return Retour

    End Function

    Private Function RequêteSQLLog(ByVal UserID As Integer, ByVal IPAdress As String, ByVal NameOfComputer As String, ByVal OperatingSystem As String, ByVal Action As String, _
                                                 ByVal TableName As String, ByVal Chrono As String, ByVal RequêteSQL As String, ByVal ResultAction As String, _
                                                 ByVal ChaineDeConnexion As String, ByVal NomDeLaBaseDeDonnées As String)


        ' INSERT INTO table_name ( field1, field2,...fieldN )
        ' VALUES                 ( value1, value2,...valueN );


        Dim Retour As New RetourLog
        Dim Infos_sur_la_requête_Log As String = ""
        Dim EnumResultAction As String = ""

        If (ResultAction = "1") Then EnumResultAction = "Success"
        If (ResultAction = "0") Then EnumResultAction = "Failed"

        Infos_sur_la_requête_Log += "Connexion : " & ChaineDeConnexion & vbCrLf

        Try

            Dim Mysql_connection As New MySql.Data.MySqlClient.MySqlConnection(ChaineDeConnexion)
            Mysql_connection.Open()


            'Ici on enregistre le log
            Dim Requête_Log As String = ""

            Dim ModifiedSQLString As String = RequêteSQL
            ModifiedSQLString = ModifiedSQLString.Replace("'", " ")
            ModifiedSQLString = ModifiedSQLString.Replace("'", "\'")
            ModifiedSQLString = ModifiedSQLString.Replace("""", " \""")

            Requête_Log = "INSERT INTO " & NomDeLaBaseDeDonnées & ".HISTORY_EVENTS (User_ID,IP,Name_of_Computer,Operating_System,Action,Table_Name,ID_Record,SQL_String,Result_Action)  VALUES " & _
                                                                   "(" & UserID & ",'" & IPAdress & "','" & NameOfComputer & "','" & OperatingSystem & "','" & Action & "','" & TableName & "'," & _
                                                                    "'" & Chrono & "','" & ModifiedSQLString & "','" & EnumResultAction & "');"


            Dim Commande_insertion_LOG As New MySql.Data.MySqlClient.MySqlCommand
            Dim NombreDeLigneAffectée As Integer = 0

            'exécution de la requête
            Commande_insertion_LOG.CommandText = Requête_Log
            Commande_insertion_LOG.Connection = Mysql_connection

            Infos_sur_la_requête_Log += vbCrLf & "Insertion du Log" & vbCrLf
            Infos_sur_la_requête_Log += "Requête : " & Requête_Log & vbCrLf


            NombreDeLigneAffectée = Commande_insertion_LOG.ExecuteNonQuery() 'Execution de la requête
            Infos_sur_la_requête_Log += "Ajout du Log. OK" & vbCrLf


            Infos_sur_la_requête_Log += vbCrLf & " RequêteSQLLog()" & vbCrLf
            Infos_sur_la_requête_Log += Requête_Log & vbCrLf



            Commande_insertion_LOG.Dispose()


            'Si on réussit à enregistrer le nouveau utilisateur on retourne son Nom
            Retour.RetourLog = NombreDeLigneAffectée


fin:        Mysql_connection.Close()
            Infos_sur_la_requête_Log += vbCrLf & "Déconnexion." & vbCrLf

        Catch ex As MySql.Data.MySqlClient.MySqlException
            'MessageBox.Show(ex.ErrorCode & " : " & ex.Message, "Requêtes_SQL_Recherche_simplifiée() - Error - ")
            Infos_sur_la_requête_Log += " : " & ex.Message & " RequêteSQLLog() - Error - " & vbCrLf

            'test si la connection mysql est toujours ouverte
            'If Mysql_connection.State = ConnectionState.Open Then
            'si oui on la ferme
            'Mysql_connection.Close()
            'End If

            GoTo erreur_fin
        End Try
        Retour.Erreur = False
        Retour.Infos_requêtes = Infos_sur_la_requête_Log

        Return Retour

erreur_fin:
        Retour.Erreur = True
        Retour.Infos_requêtes = Infos_sur_la_requête_Log
        Return Retour

    End Function



    Private Function UploadFileByFTP(ByVal DestinationFTP As String, ByVal NomFichierDansFTP As String, ByVal CheminCompletDuFichierATransférer As String, ByVal NomFichierATransférer As String, _
                                     ByVal NomUtilisateur As String, ByVal MotDepasse As String)

        'Exemple de code 
        'http://www.developpez.net/forums/d949262/dotnet/langages/vb-net/upload-via-ftpwebrequest-progressbar/

        Try

            Dim RequestFTP As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create(DestinationFTP & NomFichierDansFTP), System.Net.FtpWebRequest)
            RequestFTP.Credentials = New System.Net.NetworkCredential(NomUtilisateur, MotDepasse)
            RequestFTP.Method = System.Net.WebRequestMethods.Ftp.UploadFile

            Dim File() As Byte = System.IO.File.ReadAllBytes(CheminCompletDuFichierATransférer)
            Dim UpFile As System.IO.Stream = RequestFTP.GetRequestStream()
            UpFile.Write(File, 0, File.Length)
            UpFile.Close()
            UpFile.Dispose()


            ' On affiche que le transfert s'est bien déroulé
            ToolStripStatusLabel1.Text = MessageUtilisateur(297, Me.LangueLogiciel) & " - """ & NomFichierATransférer & """ "

            Return True

        Catch ex As Exception

            ' On affiche que le transfert s'est mal passé
            ToolStripStatusLabel1.Text = MessageUtilisateur(298, Me.LangueLogiciel) & " - """ & NomFichierATransférer & """ "

            Return False

        End Try


    End Function

    Private Function UploadLocalFile(ByVal CheminCompletDuFichierATransférer As String, ByVal Destination As String, ByVal NomFichierDistant As String)



        Try

            'on copie le fichier source vers destination

            System.IO.File.Copy(CheminCompletDuFichierATransférer, Destination & NomFichierDistant)
            Return True

        Catch ex As Exception
            Return False
        End Try



    End Function


    Private Sub Form1_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        If File.Exists(My.Application.Info.DirectoryPath & "\TRANSFERT.TXT") = True Then
            System.IO.File.Delete(My.Application.Info.DirectoryPath & "\TRANSFERT.TXT")
        End If

    End Sub


    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Ajout d'une exception au Firewall
        AjoutFireWallException()

        Dim testLectureFichier As Boolean = False

        testLectureFichier = LectureFichierConfig()

        'si on a pas pu charger les 6 paramètres on doit arrêter le Form_Load
        If testLectureFichier = False Then Exit Sub

        Initialiser_langue_logiciel()


        'Affichage des données récupérées dans le fichier "TRANSFERT.TXT"
        '[NOM_DE_LA_BASE]           => NomDeLaBase
        LabelNomDeLaBase.Text = NomDeLaBase

        '[ENREGISTREMENT]           => Numéro_Enregistrement
        LabelNuméroEnregistrement.Text = Numéro_Enregistrement

        '[CHRONO]                   => Chrono
        LabelNuméroChrono.Text = Chrono


        Dim ArgumentChargementTypeDocument As New Argument_ChargementTypeDocument

        ' "Chargement TypeDocuments en cours..."
        ToolStripStatusLabel1.Text = MessageUtilisateur(281, LangueLogiciel)

        ArgumentChargementTypeDocument.Auteur = Me.Auteur
        ArgumentChargementTypeDocument.chaine_de_connexion = Me.ChaineDeConnexion
        ArgumentChargementTypeDocument.Nom_de_la_base_de_données = Me.NomDeLaBase
        ArgumentChargementTypeDocument.langue_logiciel = Me.LangueLogiciel

        'on désactive les combobox avant le chargement
        ComboBoxTypeDocumentMaintenance.Enabled = False
        ComboBoxTypeDocument.Enabled = False

        If Not BackgroundWorkerChargementTypeDocument.IsBusy = True Then

            TimerChargementTypeDocument.Start()
            BackgroundWorkerChargementTypeDocument.RunWorkerAsync(ArgumentChargementTypeDocument)

        End If

        ' chargement de ListBoxDocuments

        Dim ArgumentChargementListeDocuments As New Argument_ListeDocuments
        ArgumentChargementListeDocuments.chaine_de_connexion = Me.ChaineDeConnexion
        ArgumentChargementListeDocuments.Nom_de_la_base_de_données = Me.NomDeLaBase
        ArgumentChargementListeDocuments.langue_logiciel = Me.LangueLogiciel
        ArgumentChargementListeDocuments.Chrono = Me.Chrono


        If Not BackgroundWorkerChargementListeDocuments.IsBusy = True Then

            TimerChargementListeDocuments.Start()
            BackgroundWorkerChargementListeDocuments.RunWorkerAsync(ArgumentChargementListeDocuments)

        End If

    End Sub


    Private Function LectureFichierConfig()

        Dim chemin_TRANSFERT_CONF As String = ""
        Dim Lignes_String As String = ""
        Dim Ligne_flux As FileStream
        Dim Ligne_ecriture As StreamWriter
        Dim TestInterprétation As Boolean = False

        chemin_TRANSFERT_CONF = My.Application.Info.DirectoryPath & "\TRANSFERT.TXT"
        If Not File.Exists(chemin_TRANSFERT_CONF) Then

            Ligne_flux = File.Create(chemin_TRANSFERT_CONF)
            Ligne_ecriture = New StreamWriter(Ligne_flux)
            Ligne_ecriture.WriteLine(Lignes_String)
            Ligne_ecriture.Close()
        Else


            ' Nous avons besoin de lire dans une liste.
            Dim ListTransfertConfiguration As New List(Of String)

            ' ouverture du fichier txt avec l'état Using.
            Using r As StreamReader = New StreamReader(chemin_TRANSFERT_CONF)
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

            TestInterprétation = Interprétation_TRANSFERT_CONF(ListTransfertConfiguration)


        End If


        Return TestInterprétation

    End Function




    Public Function Interprétation_TRANSFERT_CONF(ByVal liste As List(Of String))


        If liste.Count = 0 Then Return True

        'On recherche des symboles

        '[CHAINE_CONNEXION]         => ChaineDeConnexion
        '[NOM_DE_LA_BASE]           => NomDeLaBase
        '[CHRONO]                   => Chrono
        '[ENREGISTREMENT]           => Numéro_Enregistrement
        '[AUTEUR]                   => Auteur
        '[LANGUE]                   => LangueLogiciel
        '[MAINTENANCE]              => GroupBoxMaintenance.visible = true

        'test si tout les paramètres ont bien été lus  (score =6)
        ' en cas d'echec ne pas poursuivre le form_load
        Dim score As Integer = 0

        Try
            Dim T As Integer = 0
            Dim ligne_en_test_STR As String = ""


            For T = 0 To (liste.Count - 1)

                ligne_en_test_STR = ""

                ligne_en_test_STR = liste.Item(T)


                If (ligne_en_test_STR.Contains("[CHAINE_CONNEXION]")) Then ChaineDeConnexion = ligne_en_test_STR.Replace("[CHAINE_CONNEXION]", "") : score += 1
                If (ligne_en_test_STR.Contains("[NOM_DE_LA_BASE]")) Then NomDeLaBase = ligne_en_test_STR.Replace("[NOM_DE_LA_BASE]", "") : score += 1
                If (ligne_en_test_STR.Contains("[ENREGISTREMENT]")) Then Numéro_Enregistrement = ligne_en_test_STR.Replace("[ENREGISTREMENT]", "") : score += 1
                If (ligne_en_test_STR.Contains("[CHRONO]")) Then Chrono = ligne_en_test_STR.Replace("[CHRONO]", "") : score += 1
                If (ligne_en_test_STR.Contains("[AUTEUR]")) Then Auteur = ligne_en_test_STR.Replace("[AUTEUR]", "") : score += 1
                If (ligne_en_test_STR.Contains("[LANGUE]")) Then LangueLogiciel = CInt(ligne_en_test_STR.Replace("[LANGUE]", "")) : score += 1
                If (ligne_en_test_STR.Contains("[MAINTENANCE]")) Then GroupBoxMaintenance.Visible = True


            Next

        Catch ex As Exception
            MsgBox("Interprétation_TRANSFERT_CONF() -  " & ex.Message, MsgBoxStyle.Critical, "- Error -")
            Return False
        End Try

        If score = 6 Then
            Return True
        Else
            Return False
        End If

    End Function



    Public Sub AjoutFireWallException()


        'Exemple trouvé sur http://www.johnkoerner.com/index.php?/archives/49-Creating-a-Firewall-Exception-in-VB.Net.html

        Try
            'Création de l'application que nous voulons ajouter à la liste des exceptions
            Dim appType As Type = Type.GetTypeFromProgID("HnetCfg.FwAuthorizedApplication")
            Dim app As INetFwAuthorizedApplication
            app = DirectCast(Activator.CreateInstance(appType), INetFwAuthorizedApplication)

            'Paramètrage des propriétées de l'application
            app.Name = "Transfert_Idealtake"
            app.ProcessImageFileName = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe"
            app.Enabled = True


            'Obtenir le gestionnaire du Firewall, afin d'obtenir la liste des applications autorisées
            Dim fwMgrType As Type = Type.GetTypeFromProgID("HnetCfg.FwMgr")
            Dim fwMgr As INetFwMgr
            fwMgr = DirectCast(Activator.CreateInstance(fwMgrType), INetFwMgr)

            ' Obtient la liste des applications autorisées par le gestionnaire du FireWall, Ainsi nous pouvons
            ' Ajouter notre application à la liste
            Dim apps As INetFwAuthorizedApplications
            apps = fwMgr.LocalPolicy.CurrentProfile.AuthorizedApplications
            apps.Add(app)

        Catch ErreurAjoutExceptionFirewall As Exception

            MsgBox("Sub AjoutFireWallException() - " & ErreurAjoutExceptionFirewall.Message, "Error")

        End Try

    End Sub

    Private Sub ButtonSelectFile_Click(sender As System.Object, e As System.EventArgs) Handles ButtonSelectFile.Click
        Dim RETOUR As New System.Windows.Forms.DialogResult
        Dim TempOrigine As String = ""

        OpenFileDialog.Reset()
        LabelNomFichierOrigine.Text = "---"

        TextBoxNomFichier.Enabled = False

        If OpenFileDialog.ShowDialog() = DialogResult.OK Then

            ' MsgBox(OpenFileDialog.FileName & "," & OpenFileDialog.SafeFileName)
            LabelNomFichierOrigine.Text = """" & OpenFileDialog.SafeFileName & """"
            TempOrigine = OpenFileDialog.FileName

            If Me.BooleanAllemand = True Then ' Allemand
                Me.NouveauDocument.NomFichierALL = OpenFileDialog.SafeFileName
                Me.NouveauDocument.OrigineFichierALL = TempOrigine
            End If

            If Me.BooleanAnglais = True Then ' Anglais
                Me.NouveauDocument.NomFichierANG = OpenFileDialog.SafeFileName
                Me.NouveauDocument.OrigineFichierANG = TempOrigine
            End If

            If Me.BooleanEspagnol = True Then ' Espagnol
                Me.NouveauDocument.NomFichierESP = OpenFileDialog.SafeFileName
                Me.NouveauDocument.OrigineFichierESP = TempOrigine
            End If

            If Me.BooleanFrançais = True Then ' Français
                Me.NouveauDocument.NomFichierFRA = OpenFileDialog.SafeFileName
                Me.NouveauDocument.OrigineFichierFRA = TempOrigine
            End If

            If Me.BooleanItalien = True Then ' Italien
                Me.NouveauDocument.NomFichierITA = OpenFileDialog.SafeFileName
                Me.NouveauDocument.OrigineFichierITA = TempOrigine
            End If

            TextBoxNomFichier.Text = OpenFileDialog.SafeFileName
            TextBoxNomFichier.Enabled = True

        End If


    End Sub

    Private Sub BackgroundWorkerChargementTypeDocument_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorkerChargementTypeDocument.DoWork



        Dim Arguments_SQL As New Argument_ChargementTypeDocument

        Arguments_SQL = e.Argument

        Dim Retour_Recherche As New RetourChargementTypeDocument

        Retour_Recherche = RequêteSQLChargementTypeDocument(Arguments_SQL.Auteur, Arguments_SQL.chaine_de_connexion, Arguments_SQL.Nom_de_la_base_de_données, Arguments_SQL.langue_logiciel)



        If (Retour_Recherche.Erreur = True) Then
            'Type_de_message
            '0 = pas de message
            '1 = chargement des données de types documents Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageChargementTypeDocument = 2  'Type_message = 2 => Erreur 
            My.Settings.MessageChargementTypeDocument = Retour_Recherche.Infos_requêtes
            Exit Sub

        End If



        If (Retour_Recherche.Erreur = False) Then
            'Type_de_message
            '0 = pas de message
            '1 = chargement des données de types documents Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageChargementTypeDocument = 1
            My.Settings.MessageChargementTypeDocument = ""
            Exit Sub

        End If






    End Sub


    Public Class ClasseTypesDocuments
        Implements IEquatable(Of ClasseTypesDocuments)

        ' Structure TypeDocument
        'Provient de la table 'types_documents'
        'Dim IDTypeDocument As String
        'Dim DescriptionTypeDocument As String

        'Provient de la table 'Transfert'
        'Dim TypeDocument As String
        'Dim TypeTransfert As Integer
        'Dim Destination As String
        'Dim User As String
        'Dim MotDePasse As String


        Public IDTypeDocument As String = ""
        Public DescriptionTypeDocument As String = ""
        Public TypeDocument As String = ""
        Public TypeTransfert As Integer = 0
        Public Destination As String = ""
        Public User As String = ""
        Public MotDePasse As String = " "


        Public Sub New( _
                      ByVal m_IDTypeDocument As String, _
                      ByVal m_DescriptionTypeDocument As String, _
                      ByVal m_TypeDocument As String, _
                      ByVal m_TypeTransfert As Integer, _
                      ByVal m_Destination As String, _
                      ByVal m_User As String, _
                      ByVal m_MotDePasse As String)

            IDTypeDocument = m_IDTypeDocument
            DescriptionTypeDocument = m_DescriptionTypeDocument
            TypeDocument = m_TypeDocument
            TypeTransfert = m_TypeTransfert
            Destination = m_Destination
            User = m_User
            MotDePasse = m_MotDePasse


        End Sub

        Sub New()
            ' TODO: Complete member initialization 
        End Sub

        Public Overloads Function Equals(ByVal other As ClasseTypesDocuments) _
      As Boolean Implements IEquatable(Of ClasseTypesDocuments).Equals

            If Me.IDTypeDocument = other.IDTypeDocument And Me.DescriptionTypeDocument = other.DescriptionTypeDocument And Me.TypeDocument = other.TypeDocument _
              And Me.TypeTransfert = other.TypeTransfert And Me.Destination = other.Destination And Me.User = other.User _
              And Me.MotDePasse = other.MotDePasse Then
                Return True
            Else
                Return False
            End If

        End Function

        Sub Add(ClasseTypesDocuments As ClasseTypesDocuments)
            Throw New NotImplementedException
        End Sub



    End Class

    Public Class ClasseDocuments
        Implements IEquatable(Of ClasseDocuments)




        Public NumEnregistrement As String = ""
        Public Chrono As String = ""

        Public NomTypeDocument As String = ""
        Public NuméroTypeDocument As String = ""
        Public TypeTransfert As Integer = 0
        Public Destination As String = ""
        Public Utilisateur As String = ""
        Public MotDePasse As String = ""

        Public NomDuLienFichierALL As String = ""
        Public NomDuLienFichierANG As String = ""
        Public NomDuLienFichierESP As String = ""
        Public NomDuLienFichierFRA As String = ""
        Public NomDuLienFichierITA As String = ""

        Public OrigineFichierALL As String = ""
        Public NomFichierALL As String = ""
        Public TailleFichierALL As String = ""

        Public OrigineFichierANG As String = ""
        Public NomFichierANG As String = ""
        Public TailleFichierANG As String = ""

        Public OrigineFichierESP As String = ""
        Public NomFichierESP As String = ""
        Public TailleFichierESP As String = ""

        Public OrigineFichierFRA As String = ""
        Public NomFichierFRA As String = ""
        Public TailleFichierFRA As String = ""

        Public OrigineFichierITA As String = ""
        Public NomFichierITA As String = ""
        Public TailleFichierITA As String = ""


        Public Sub New( _
                      ByVal m_NumEnregistrement As String, _
                      ByVal m_Chrono As String, _
                      ByVal m_NomTypeDocument As String, _
                      ByVal m_TypeTransfert As Integer, _
                      ByVal m_Destination As String, _
                      ByVal m_User As String, _
                      ByVal m_MotDePasse As String, _
                      ByVal m_NomDuLienFichierALL As String, _
                      ByVal m_NomDuLienFichierANG As String, _
                      ByVal m_NomDuLienFichierESP As String, _
                      ByVal m_NomDuLienFichierFRA As String, _
                      ByVal m_NomDuLienFichierITA As String, _
                      ByVal m_OrigineFichierALL As String, _
                      ByVal m_NomFichierALL As String, _
                      ByVal m_TailleFichierALL As String, _
                      ByVal m_OrigineFichierANG As String, _
                      ByVal m_NomFichierANG As String, _
                      ByVal m_TailleFichierANG As String, _
                      ByVal m_OrigineFichierESP As String, _
                      ByVal m_NomFichierESP As String, _
                      ByVal m_TailleFichierESP As String, _
                      ByVal m_OrigineFichierFRA As String, _
                      ByVal m_NomFichierFRA As String, _
                      ByVal m_TailleFichierFRA As String, _
                      ByVal m_OrigineFichierITA As String, _
                      ByVal m_NomFichierITA As String, _
                      ByVal m_TailleFichierITA As String)

            NumEnregistrement = m_NumEnregistrement
            Chrono = m_Chrono
            NomTypeDocument = m_NomTypeDocument
            TypeTransfert = m_TypeTransfert
            Destination = m_Destination
            Utilisateur = m_User
            MotDePasse = m_MotDePasse

            NomDuLienFichierALL = m_NomDuLienFichierALL
            NomDuLienFichierANG = m_NomDuLienFichierANG
            NomDuLienFichierESP = m_NomDuLienFichierESP
            NomDuLienFichierFRA = m_NomDuLienFichierFRA
            NomDuLienFichierITA = m_NomDuLienFichierITA

            OrigineFichierALL = m_OrigineFichierALL
            NomFichierALL = m_NomFichierALL
            TailleFichierALL = m_TailleFichierALL

            OrigineFichierANG = m_OrigineFichierANG
            NomFichierANG = m_NomFichierANG
            TailleFichierANG = m_TailleFichierANG

            OrigineFichierESP = m_OrigineFichierESP
            NomFichierESP = m_NomFichierESP
            TailleFichierESP = m_TailleFichierESP

            OrigineFichierFRA = m_OrigineFichierFRA
            NomFichierFRA = m_NomFichierFRA
            TailleFichierFRA = m_TailleFichierFRA

            OrigineFichierITA = m_OrigineFichierITA
            NomFichierITA = m_NomFichierITA
            TailleFichierITA = m_TailleFichierITA

        End Sub

        Sub New()
            ' TODO: Complete member initialization 
        End Sub

        Public Overloads Function Equals(ByVal other As ClasseDocuments) _
      As Boolean Implements IEquatable(Of ClasseDocuments).Equals

            If Me.NumEnregistrement = other.NumEnregistrement And Me.Chrono = other.Chrono And Me.NomTypeDocument = other.NomTypeDocument _
              And Me.TypeTransfert = other.TypeTransfert And Me.Destination = other.Destination And Me.Utilisateur = other.Utilisateur _
              And Me.MotDePasse = other.MotDePasse And Me.NomDuLienFichierALL = other.NomDuLienFichierALL And Me.NomDuLienFichierANG = other.NomDuLienFichierANG _
              And Me.NomDuLienFichierESP = other.NomDuLienFichierESP And Me.NomDuLienFichierFRA = other.NomDuLienFichierFRA And Me.NomDuLienFichierITA = other.NomDuLienFichierITA _
              And Me.OrigineFichierALL = other.OrigineFichierALL And Me.NomFichierALL = other.NomFichierALL And Me.TailleFichierALL = other.TailleFichierALL _
              And Me.OrigineFichierANG = other.OrigineFichierANG And Me.NomFichierANG = other.NomFichierANG And Me.TailleFichierANG = other.TailleFichierANG _
              And Me.OrigineFichierESP = other.OrigineFichierESP And Me.NomFichierESP = other.NomFichierESP And Me.TailleFichierESP = other.TailleFichierESP _
              And Me.OrigineFichierFRA = other.OrigineFichierFRA And Me.NomFichierFRA = other.NomFichierFRA And Me.TailleFichierFRA = other.TailleFichierFRA _
              And Me.OrigineFichierITA = other.OrigineFichierITA And Me.NomFichierITA = other.NomFichierITA And Me.TailleFichierITA = other.TailleFichierITA _
              Then
                Return True
            Else
                Return False
            End If

        End Function

        Sub Add(ClasseTypesDocuments As ClasseDocuments)
            Throw New NotImplementedException
        End Sub



    End Class


    Private Sub TimerChargementTypeDocument_Tick(sender As System.Object, e As System.EventArgs) Handles TimerChargementTypeDocument.Tick
        'Type_de_message
        '0 = pas de message
        '1 = chargement des données de types documents Terminé
        '2 = Erreur Requête


        If My.Settings.TypeMessageChargementTypeDocument = 1 Then
            '1 = chargement des données des types documents Terminé

            ' PHASE 1 - Indiquer la fin de chargement à l'utilisateur
            ToolStripStatusLabel1.Text = MessageUtilisateur(282, LangueLogiciel)
            Me.Refresh()
            'pour apercevoir le message 'Terminé'

            Dim pauseTime As Integer = 1000
            System.Threading.Thread.Sleep(pauseTime)

            'PHASE 2  - Indiquer le chargement des combobox puis les charger : ComboBoxTypeDocumentMaintenance, ComboBoxTypeDocument
            ToolStripStatusLabel1.Text = "---"

            Dim i02 As Integer = 0

            For i02 = 0 To (ListOfTypesDocuments.Count - 1)

                ' si le groupe de maintenance est visible alors on fait le chargement du combobox 'ComboBoxTypeDocumentMaintenance'
                If GroupBoxMaintenance.Visible = True Then ComboBoxTypeDocumentMaintenance.Items.Add(ListOfTypesDocuments.Item(i02).DescriptionTypeDocument)

                ComboBoxTypeDocument.Items.Add(ListOfTypesDocuments.Item(i02).DescriptionTypeDocument)
            Next

            ' si le groupe de maintenance est visible alors on positionne le combobox 'ComboBoxTypeDocumentMaintenance' à sa première ligne
            If GroupBoxMaintenance.Visible = True Then

                'on positionne uniquement si il ya au moin 1 item
                If ComboBoxTypeDocumentMaintenance.Items.Count > 0 Then ComboBoxTypeDocumentMaintenance.SelectedIndex = 0

            End If

            If ComboBoxTypeDocument.Items.Count > 0 Then ComboBoxTypeDocument.SelectedIndex = 0

            My.Settings.TypeMessageChargementTypeDocument = 0
            My.Settings.MessageChargementTypeDocument = ""
            TimerChargementTypeDocument.Stop()


            'on réactive les combobox après le chargement
            ComboBoxTypeDocumentMaintenance.Enabled = True
            ComboBoxTypeDocument.Enabled = True

            Exit Sub
        End If



        If My.Settings.TypeMessageChargementTypeDocument = 2 Then
            '2 = Erreur Requête
            MsgBox(My.Settings.MessageChargementTypeDocument, MsgBoxStyle.Critical, "- Error -")
            TimerChargementTypeDocument.Stop()
            My.Settings.TypeMessageChargementTypeDocument = 0
            My.Settings.MessageChargementTypeDocument = ""

            'on réactive les combobox après le chargement
            ComboBoxTypeDocumentMaintenance.Enabled = True
            ComboBoxTypeDocument.Enabled = True

        End If

    End Sub

    Private Sub ComboBoxTypeDocumentMaintenance_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxTypeDocumentMaintenance.SelectedIndexChanged


        Dim Crypto1 = New CryptoUtil(MyKey, MyIV)

        'Dim temp As String = ""
        'temp = Crypto1.EncryptText(TextBoxUserNameMaintenance.Text)
        'TextBoxPassword.Text = ConvertStringToBase64(temp)

        'ListOfTypesDocuments contient l'ensemble des moyens pour le transfert

        'Exit Sub

        TextBoxDestinationMaintenance.Text = ""
        TextBoxUserNameMaintenance.Text = ""
        TextBoxPassword.Text = ""
        RadioButtonFTPMaintenance.Checked = True
        RadioButtonTransfertLocalMaintenance.Checked = False

        If (ComboBoxTypeDocumentMaintenance.Items.Count = 0) Then Exit Sub

        Dim I03 As Integer = 0

        For I03 = 0 To (ListOfTypesDocuments.Count - 1)

            If (ComboBoxTypeDocumentMaintenance.SelectedItem = ListOfTypesDocuments.Item(I03).DescriptionTypeDocument) Then
                'MsgBox("Index of 'ListOfTypesDocuments' : " & CStr(I03) & vbNewLine &
                ' "Combobox SelectedItem : " & ComboBoxTypeDocumentMaintenance.SelectedItem)

                TextBoxDestinationMaintenance.Text = ListOfTypesDocuments.Item(I03).Destination


                Try
                    Dim TempUserName As String = ""
                    TempUserName = ListOfTypesDocuments.Item(I03).User

                    If TempUserName.Length > 20 Then
                        TempUserName = ConvertBase64ToString(TempUserName)
                        TextBoxUserNameMaintenance.Text = Crypto1.DecryptText(TempUserName)
                        TempUserName = ""
                    End If

                    Dim TempTextBoxPassword As String = ""
                    TempTextBoxPassword = ListOfTypesDocuments.Item(I03).MotDePasse

                    If TempTextBoxPassword.Length > 20 Then
                        TempTextBoxPassword = ConvertBase64ToString(TempTextBoxPassword)
                        TextBoxPassword.Text = Crypto1.DecryptText(TempTextBoxPassword)
                    End If

                Catch EX As System.Security.Cryptography.CryptographicException
                    MsgBox("ComboBoxTypeDocumentMaintenance_SelectedIndexChanged() " & EX.Message, MsgBoxStyle.Critical, "- Error -")
                End Try


                If (CInt(ListOfTypesDocuments.Item(I03).TypeTransfert) = 1) Then
                    'Transfert par FTP
                    RadioButtonFTPMaintenance.Checked = True
                    RadioButtonTransfertLocalMaintenance.Checked = False
                End If

                If (CInt(ListOfTypesDocuments.Item(I03).TypeTransfert) = 2) Then
                    'Transfert en local
                    RadioButtonFTPMaintenance.Checked = False
                    RadioButtonTransfertLocalMaintenance.Checked = True
                End If

            End If

        Next



    End Sub

    Private Sub ButtonDestinationLocalMaintenance_Click(sender As System.Object, e As System.EventArgs) Handles ButtonDestinationLocalMaintenance.Click
        FolderBrowserDialog.Reset()
        FolderBrowserDialog.ShowDialog()


        TextBoxDestinationMaintenance.Text = FolderBrowserDialog.SelectedPath

    End Sub

    Private Sub ButtonMiseàJourMaintenance_Click(sender As System.Object, e As System.EventArgs) Handles ButtonMiseàJourMaintenance.Click

        If (TextBoxDestinationMaintenance.Text.Length < 3) Then
            MsgBox(MessageUtilisateur(283, LangueLogiciel), MsgBoxStyle.Information)
            Exit Sub
        End If

        If (TextBoxDestinationMaintenance.Text.Length > 250) Then
            MsgBox(MessageUtilisateur(284, LangueLogiciel), MsgBoxStyle.Information)
            Exit Sub
        End If

        ' Nom d'utilisateur vide doit être possible
        'If (TextBoxUserNameMaintenance.Text.Length < 3) Then
        'MsgBox(MessageUtilisateur(285, LangueLogiciel), MsgBoxStyle.Information)
        'Exit Sub
        'End If

        If (TextBoxUserNameMaintenance.Text.Length > 20) Then
            MsgBox(MessageUtilisateur(286, LangueLogiciel), MsgBoxStyle.Information)
            Exit Sub
        End If

        ' Mot de passe vide doit être possible
        'If (TextBoxPassword.Text.Length < 3) Then
        'MsgBox(MessageUtilisateur(287, LangueLogiciel), MsgBoxStyle.Information)
        'Exit Sub
        'End If

        If (TextBoxPassword.Text.Length > 20) Then
            MsgBox(MessageUtilisateur(288, LangueLogiciel), MsgBoxStyle.Information)
            Exit Sub
        End If


        Dim Argument_Maintenance As New Argument_MaintenanceTransfert

        Argument_Maintenance.chaine_de_connexion = ChaineDeConnexion
        Argument_Maintenance.langue_logiciel = LangueLogiciel
        Argument_Maintenance.Nom_de_la_base_de_données = NomDeLaBase
        Argument_Maintenance.Destination = TextBoxDestinationMaintenance.Text

        If RadioButtonFTPMaintenance.Checked = True And RadioButtonTransfertLocalMaintenance.Checked = False Then
            'FTP
            Argument_Maintenance.TypeTransfert = 1
        End If

        If RadioButtonFTPMaintenance.Checked = False And RadioButtonTransfertLocalMaintenance.Checked = True Then
            'Local
            Argument_Maintenance.TypeTransfert = 2
        End If

        Argument_Maintenance.Utilisateur = TextBoxUserNameMaintenance.Text
        Argument_Maintenance.MotDePasse = TextBoxPassword.Text
        Argument_Maintenance.Auteur = Auteur

        Dim TexteCombo As String = ComboBoxTypeDocumentMaintenance.SelectedItem

        Dim I06 As Integer = 0

        For I06 = 0 To ListOfTypesDocuments.Count - 1

            'IDTypeDocument provient de 'types_documents'
            'TypeDocument provient de 'Transfert'

            If ListOfTypesDocuments.Item(I06).DescriptionTypeDocument = TexteCombo Then


                Argument_Maintenance.TypeDocument = ListOfTypesDocuments.Item(I06).TypeDocument
                'si TypeDocument = 0 c'est un SELECT qu'il faudra faire dans la requête
                'si TypeDocument > 0 c'est un INSERT qu'il faudra faire dans la requête

                Argument_Maintenance.IDTypeDocument = ListOfTypesDocuments.Item(I06).IDTypeDocument

            End If

        Next

        ToolStripStatusLabel1.Text = MessageUtilisateur(294, LangueLogiciel)

        If Not BackgroundWorker_MaintenanceTransfert.IsBusy = True Then

            TimerEtatMaintenanceTransfert.Start()
            BackgroundWorker_MaintenanceTransfert.RunWorkerAsync(Argument_Maintenance)

        End If


    End Sub

    Private Sub BackgroundWorkerChargementListeDocuments_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorkerChargementListeDocuments.DoWork


        Dim Arguments_SQL As New Argument_ListeDocuments

        Arguments_SQL = e.Argument

        Dim Retour_Recherche As New RetourChargementListeDocuments

        Retour_Recherche = RequêteSQLChargementListeDocuments(Arguments_SQL.Chrono, Arguments_SQL.chaine_de_connexion, Arguments_SQL.Nom_de_la_base_de_données, Arguments_SQL.langue_logiciel)



        If (Retour_Recherche.Erreur = True) Then
            'Type_de_message
            '0 = pas de message
            '1 = chargement Liste Documents Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageChargementListeDocuments = 2  'Type_message = 2 => Erreur 
            My.Settings.MessageChargementListeDocuments = Retour_Recherche.Infos_requêtes
            Exit Sub

        End If



        If (Retour_Recherche.Erreur = False) Then
            'Type_de_message
            '0 = pas de message
            '1 =chargement Liste Documents Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageChargementListeDocuments = 1
            My.Settings.MessageChargementListeDocuments = ""
            Exit Sub

        End If



    End Sub


    Private Sub TimerChargementListeDocuments_Tick(sender As System.Object, e As System.EventArgs) Handles TimerChargementListeDocuments.Tick

        'Type_de_message
        '0 = pas de message
        '1 = chargement Liste Documents Terminé
        '2 = Erreur Requête


        If My.Settings.TypeMessageChargementListeDocuments = 1 Then
            '1 = chargement Liste Documents Terminé

            ' PHASE 1 - Indiquer la fin de chargement à l'utilisateur
            ToolStripStatusLabel1.Text = MessageUtilisateur(289, LangueLogiciel)
            Me.Refresh()
            'pour apercevoir le message 'Terminé'

            Dim pauseTime As Integer = 1000
            System.Threading.Thread.Sleep(pauseTime)

            'PHASE 2  - Indiquer le chargement des combobox puis les charger : ListBoxDocuments
            ToolStripStatusLabel1.Text = "---"

            Dim i02 As Integer = 0

            If ListeDesDocumentsPourLeChrono.Count = 0 Then GoTo Fin0

            For i02 = 0 To (ListeDesDocumentsPourLeChrono.Count - 1)

                Select Case LangueLogiciel
                    Case 0 ' Allemand
                        ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierALL)
                    Case 1 ' Anglais
                        ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierANG)
                    Case 2 ' Espagnol
                        ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierESP)
                    Case 3 ' Français
                        ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierFRA)
                    Case 4 ' Italien
                        ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierITA)
                End Select




            Next



Fin0:       My.Settings.TypeMessageChargementListeDocuments = 0
            My.Settings.MessageChargementListeDocuments = ""
            TimerChargementListeDocuments.Stop()
            Exit Sub
        End If



        If My.Settings.TypeMessageChargementListeDocuments = 2 Then
            '2 = Erreur Requête


            MsgBox(My.Settings.MessageChargementListeDocuments, MsgBoxStyle.Critical, "- Error -")
            TimerChargementListeDocuments.Stop()
            My.Settings.TypeMessageChargementListeDocuments = 0
            My.Settings.MessageChargementListeDocuments = ""


        End If



    End Sub

    Private Sub BackgroundWorker_MaintenanceTransfert_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_MaintenanceTransfert.DoWork


        Dim Arguments_SQL As New Argument_MaintenanceTransfert

        Arguments_SQL = e.Argument

        Dim Retour_Recherche As New RetourMaintenanceTransfert

        Retour_Recherche = RequêteSQLMaintenanceTransfert(Arguments_SQL.IDTypeDocument, Arguments_SQL.TypeDocument, Arguments_SQL.TypeTransfert, Arguments_SQL.Destination, _
                                                          Arguments_SQL.Utilisateur, Arguments_SQL.MotDePasse, Arguments_SQL.chaine_de_connexion, _
                                                          Arguments_SQL.Nom_de_la_base_de_données, Arguments_SQL.langue_logiciel, Arguments_SQL.Auteur)



        If (Retour_Recherche.Erreur = True) Then
            'Type_de_message
            '0 = pas de message
            '1 = Mise à jour 'Transfert' Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageMaintenanceTransfert = 2  'Type_message = 2 => Erreur 
            My.Settings.MessageMaintenanceTransfert = Retour_Recherche.Infos_requêtes
            Exit Sub

        End If



        If (Retour_Recherche.Erreur = False) Then
            'Type_de_message
            '0 = pas de message
            '1 = Mise à jour 'Transfert' Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageMaintenanceTransfert = 1
            My.Settings.MessageMaintenanceTransfert = ""




            Dim ArgumentChargementTypeDocument As New Argument_ChargementTypeDocument

            ' "Chargement en cours..."
            ToolStripStatusLabel1.Text = MessageUtilisateur(281, LangueLogiciel)

            ArgumentChargementTypeDocument.Auteur = Me.Auteur
            ArgumentChargementTypeDocument.chaine_de_connexion = Me.ChaineDeConnexion
            ArgumentChargementTypeDocument.Nom_de_la_base_de_données = Me.NomDeLaBase
            ArgumentChargementTypeDocument.langue_logiciel = Me.LangueLogiciel

            If Not BackgroundWorkerChargementTypeDocument.IsBusy = True Then

                TimerChargementTypeDocument.Start()
                BackgroundWorkerChargementTypeDocument.RunWorkerAsync(ArgumentChargementTypeDocument)

            End If


            Exit Sub

        End If



    End Sub

    Private Sub TimerEtatMaintenanceTransfert_Tick(sender As System.Object, e As System.EventArgs) Handles TimerEtatMaintenanceTransfert.Tick
        'Type_de_message
        '0 = pas de message
        '1 = Mise à jour 'Transfert' Terminé
        '2 = Erreur Requête


        If My.Settings.TypeMessageMaintenanceTransfert = 1 Then
            '1 = Mise à jour 'Transfert' Terminé

            ' PHASE 1 - Indiquer la fin de Mise à jour 'Transfert' 
            ToolStripStatusLabel1.Text = MessageUtilisateur(290, LangueLogiciel)
            Me.Refresh()
            'pour apercevoir le message 'Terminé'

            Dim pauseTime As Integer = 1000
            System.Threading.Thread.Sleep(pauseTime)

            'PHASE 2  - Indiquer le chargement des combobox puis les charger : ListBoxDocuments
            ToolStripStatusLabel1.Text = "---"


            My.Settings.TypeMessageMaintenanceTransfert = 0
            My.Settings.MessageMaintenanceTransfert = ""
            TimerEtatMaintenanceTransfert.Stop()
            Exit Sub
        End If



        If My.Settings.TypeMessageMaintenanceTransfert = 2 Then
            '2 = Erreur Requête

            MsgBox(My.Settings.MessageMaintenanceTransfert, MsgBoxStyle.Critical, "- Error -")
            TimerEtatMaintenanceTransfert.Stop()
            My.Settings.TypeMessageMaintenanceTransfert = 0
            My.Settings.MessageMaintenanceTransfert = ""




        End If

    End Sub

    Private Sub ButtonEffaceUnDocument_Click(sender As System.Object, e As System.EventArgs) Handles ButtonEffaceUnDocument.Click

        Dim DocumentSélectionnéDansListeDoc As String = ""
        DocumentSélectionnéDansListeDoc = "" & CStr(ListBoxDocuments.SelectedItem)

        If DocumentSélectionnéDansListeDoc.Length = 0 Then MsgBox(MessageUtilisateur(293, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub




        'Affichage Aide
        'Voulez vous restaurer Supprimer ce document?
        'Dim Retour = (MsgBox(Ajout_Suppression_connexion_base.Message_utilisateur(291, Langue_logiciel).ToString, MsgBoxStyle.YesNo, Ajout_Suppression_connexion_base.Message_utilisateur(123, Langue_logiciel).ToString))

        Dim Boite_dialogue As New TransfertIdealtake.Message_Box_Question
        Boite_dialogue.Titre_Boite = MessageUtilisateur(292, LangueLogiciel).ToString
        Boite_dialogue.Question_Boite = MessageUtilisateur(291, LangueLogiciel).ToString
        Boite_dialogue.Oui_Boite = MessageUtilisateur(82, LangueLogiciel).ToString
        Boite_dialogue.Non_Boite = MessageUtilisateur(83, LangueLogiciel).ToString
        Boite_dialogue.ShowDialog()

        'If (Retour = vbYes) Then
        If (Boite_dialogue.Réponse_question = True) Then

            'Ici on va supprimer les informations dans la table 'documents'
            ' on va utiliser le numéro de chrono et 
            '  'Nom_Lien_francais' ou 'Nom_Lien_anglais' ou 'Nom_Lien_allemand'  ou 'Nom_Lien_espagnole' ou 'Nom_Lien_italien'

            ToolStripStatusLabel1.Text = MessageUtilisateur(295, LangueLogiciel)

            Dim Argument_SupprimeDocument As New Argument_SupprimeDocument

            Argument_SupprimeDocument.chaine_de_connexion = ChaineDeConnexion
            Argument_SupprimeDocument.langue_logiciel = LangueLogiciel
            Argument_SupprimeDocument.Nom_de_la_base_de_données = NomDeLaBase
            Argument_SupprimeDocument.Auteur = Me.Auteur
            Argument_SupprimeDocument.Chrono = Me.Chrono
            Argument_SupprimeDocument.NomDuLien = DocumentSélectionnéDansListeDoc

            'on désactive le bouton  d'effacement
            'a réactiver à la fin par TimerEtatSupprimeDocument

            ButtonEffaceUnDocument.Enabled = False

            If Not BackgroundWorker_SupprimeDocument.IsBusy = True Then

                TimerEtatSupprimeDocument.Start()
                BackgroundWorker_SupprimeDocument.RunWorkerAsync(Argument_SupprimeDocument)

            End If


        End If


    End Sub

    Private Sub ButtonAddDocument_Click(sender As System.Object, e As System.EventArgs) Handles ButtonAddDocument.Click


        ' on effectue Des contrôle sur         'NouveauDocument'

        If NouveauDocument.Chrono.Length = 0 Then Exit Sub

        ' La Destination n'est pas renseigné dans le types document pour le transfert
        If NouveauDocument.Destination.Length = 0 Then MsgBox(MessageUtilisateur(299, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub



        ' Le nom du lien en allemand est absent.
        If NouveauDocument.NomDuLienFichierALL.Length = 0 Then MsgBox(MessageUtilisateur(300, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du lien en anglais est absent.
        If NouveauDocument.NomDuLienFichierANG.Length = 0 Then MsgBox(MessageUtilisateur(301, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du lien en espagnol est absent.
        If NouveauDocument.NomDuLienFichierESP.Length = 0 Then MsgBox(MessageUtilisateur(302, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du lien en Français est absent.
        If NouveauDocument.NomDuLienFichierFRA.Length = 0 Then MsgBox(MessageUtilisateur(303, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du lien en Italien est absent.
        If NouveauDocument.NomDuLienFichierITA.Length = 0 Then MsgBox(MessageUtilisateur(304, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub



        ' Le nom du fichier en allemand est absent.
        If NouveauDocument.NomFichierALL.Length = 0 Then MsgBox(MessageUtilisateur(305, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du fichier en anglais est absent.
        If NouveauDocument.NomFichierANG.Length = 0 Then MsgBox(MessageUtilisateur(306, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du fichier en espagnol est absent.
        If NouveauDocument.NomFichierESP.Length = 0 Then MsgBox(MessageUtilisateur(307, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du fichier en Français est absent.
        If NouveauDocument.NomFichierFRA.Length = 0 Then MsgBox(MessageUtilisateur(308, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Le nom du fichier en Italien est absent.
        If NouveauDocument.NomFichierITA.Length = 0 Then MsgBox(MessageUtilisateur(309, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub



        ' Pas de fichier sélectionné en allemand.
        If NouveauDocument.OrigineFichierALL.Length = 0 Then MsgBox(MessageUtilisateur(310, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Pas de fichier sélectionné  en anglais.
        If NouveauDocument.OrigineFichierANG.Length = 0 Then MsgBox(MessageUtilisateur(311, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Pas de fichier sélectionné en espagnol.
        If NouveauDocument.OrigineFichierESP.Length = 0 Then MsgBox(MessageUtilisateur(312, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Pas de fichier sélectionné en Français.
        If NouveauDocument.OrigineFichierFRA.Length = 0 Then MsgBox(MessageUtilisateur(313, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        ' Pas de fichier sélectionné en Italien.
        If NouveauDocument.OrigineFichierITA.Length = 0 Then MsgBox(MessageUtilisateur(314, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub



        ' Le fichier sélectionné en allemand est absent.
        'MsgBox(File.Exists(NouveauDocument.OrigineFichierALL))
        If File.Exists(NouveauDocument.OrigineFichierALL) = False Then MsgBox(MessageUtilisateur(315, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        Dim infoFile01 As New FileInfo(NouveauDocument.OrigineFichierALL)

        ' Get length of the file.
        Dim LengthFile01 As Long = infoFile01.Length
        LengthFile01 = LengthFile01 / 1024 ' résultat en Ko
        'MsgBox(CStr(LengthFile01) & " Kb")
        NouveauDocument.TailleFichierALL = CStr(LengthFile01) & " Kb"

        ' Le fichier sélectionné  en anglais est absent.
        If File.Exists(NouveauDocument.OrigineFichierANG) = False Then MsgBox(MessageUtilisateur(316, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        infoFile01 = New FileInfo(NouveauDocument.OrigineFichierANG)

        ' Get length of the file.
        LengthFile01 = infoFile01.Length
        LengthFile01 = LengthFile01 / 1024 ' résultat en Ko
        'MsgBox(CStr(LengthFile01) & " Kb")
        NouveauDocument.TailleFichierANG = CStr(LengthFile01) & " Kb"

        ' Le fichier sélectionné en espagnol est absent.
        If File.Exists(NouveauDocument.OrigineFichierESP) = False Then MsgBox(MessageUtilisateur(317, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        infoFile01 = New FileInfo(NouveauDocument.OrigineFichierESP)

        ' Get length of the file.
        LengthFile01 = infoFile01.Length
        LengthFile01 = LengthFile01 / 1024 ' résultat en Ko
        'MsgBox(CStr(LengthFile01) & " Kb")
        NouveauDocument.TailleFichierESP = CStr(LengthFile01) & " Kb"


        ' Le fichier sélectionné en Français est absent.
        If File.Exists(NouveauDocument.OrigineFichierFRA) = False Then MsgBox(MessageUtilisateur(318, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        infoFile01 = New FileInfo(NouveauDocument.OrigineFichierFRA)

        ' Get length of the file.
        LengthFile01 = infoFile01.Length
        LengthFile01 = LengthFile01 / 1024 ' résultat en Ko
        'MsgBox(CStr(LengthFile01) & " Ko")
        NouveauDocument.TailleFichierFRA = CStr(LengthFile01) & " Ko"

        ' Le fichier sélectionné en Italien est absent.
        If File.Exists(NouveauDocument.OrigineFichierITA) = False Then MsgBox(MessageUtilisateur(319, LangueLogiciel), MsgBoxStyle.Information, "") : Exit Sub

        infoFile01 = New FileInfo(NouveauDocument.OrigineFichierITA)

        ' Get length of the file.
        LengthFile01 = infoFile01.Length
        LengthFile01 = LengthFile01 / 1024 ' résultat en Ko
        'MsgBox(CStr(LengthFile01) & " Kb")
        NouveauDocument.TailleFichierITA = CStr(LengthFile01) & " Kb"

        'Ici on va enregistrer dans la table documents notre vouveau enregistrement
        'Puis on va envoyer les 5 fichiers ou 1 fichier si ils sont tous les même



        ToolStripStatusLabel1.Text = MessageUtilisateur(190, LangueLogiciel)

        Dim Argument_AjouteUnDocument As New Argument_AjoutDocument

        Argument_AjouteUnDocument.chaine_de_connexion = ChaineDeConnexion
        Argument_AjouteUnDocument.langue_logiciel = LangueLogiciel
        Argument_AjouteUnDocument.Nom_de_la_base_de_données = NomDeLaBase
        Argument_AjouteUnDocument.Auteur = Me.Auteur
        Argument_AjouteUnDocument.Chrono = Me.Chrono


        If Not BackgroundWorker_AjoutDocument.IsBusy = True Then

            TimerEtatAjoutDocument.Start()
            BackgroundWorker_AjoutDocument.RunWorkerAsync(Argument_AjouteUnDocument)

        End If





        GroupBoxInfosDocument.Enabled = False
    End Sub

    Private Sub ButtonNouveauDocument_Click(sender As System.Object, e As System.EventArgs) Handles ButtonNouveauDocument.Click

        If ListOfTypesDocuments.Count = 0 Then Exit Sub

        ' On réinitialise le nouveau enregistrement
        NouveauDocument.Chrono = Chrono

        NouveauDocument.Destination = ""
        NouveauDocument.MotDePasse = ""
        NouveauDocument.NomDuLienFichierALL = ""
        NouveauDocument.NomDuLienFichierANG = ""
        NouveauDocument.NomDuLienFichierESP = ""
        NouveauDocument.NomDuLienFichierFRA = ""
        NouveauDocument.NomDuLienFichierITA = ""
        NouveauDocument.NomFichierALL = ""
        NouveauDocument.NomFichierANG = ""
        NouveauDocument.NomFichierESP = ""
        NouveauDocument.NomFichierFRA = ""
        NouveauDocument.NomFichierITA = ""
        NouveauDocument.NomTypeDocument = ""
        NouveauDocument.NumEnregistrement = ""
        NouveauDocument.NuméroTypeDocument = ""
        NouveauDocument.OrigineFichierALL = ""
        NouveauDocument.OrigineFichierANG = ""
        NouveauDocument.OrigineFichierESP = ""
        NouveauDocument.OrigineFichierFRA = ""
        NouveauDocument.OrigineFichierITA = ""
        NouveauDocument.TailleFichierALL = ""
        NouveauDocument.TailleFichierANG = ""
        NouveauDocument.TailleFichierESP = ""
        NouveauDocument.TailleFichierFRA = ""
        NouveauDocument.TailleFichierITA = ""
        NouveauDocument.TypeTransfert = 0
        NouveauDocument.Utilisateur = ""


        NouveauDocument.NomTypeDocument = ComboBoxTypeDocument.SelectedItem

        Dim i08 As Integer = 0

        For i08 = 0 To ListOfTypesDocuments.Count - 1

            If (ListOfTypesDocuments.Item(i08).DescriptionTypeDocument = NouveauDocument.NomTypeDocument) Then

                NouveauDocument.NuméroTypeDocument = ListOfTypesDocuments.Item(i08).TypeDocument
                NouveauDocument.TypeTransfert = ListOfTypesDocuments.Item(i08).TypeTransfert
                NouveauDocument.Destination = ListOfTypesDocuments.Item(i08).Destination
                NouveauDocument.Utilisateur = ListOfTypesDocuments.Item(i08).User
                NouveauDocument.MotDePasse = ListOfTypesDocuments.Item(i08).MotDePasse

            End If

        Next



        'on réinitialise le groupebox nouveau enregistrement

        PictureBoxGermanFlag.BorderStyle = BorderStyle.None
        PictureBoxEnglishgFlag.BorderStyle = BorderStyle.None
        PictureBoxSpanishFlag.BorderStyle = BorderStyle.None
        PictureBoxFrenchFlag.BorderStyle = BorderStyle.None
        PictureBoxItalianFlag.BorderStyle = BorderStyle.None

        Me.BooleanAllemand = False
        Me.BooleanAnglais = False
        Me.BooleanEspagnol = False
        Me.BooleanFrançais = False
        Me.BooleanItalien = False

        LabelNomFichierOrigine.Text = "---"
        TextBoxNomDuLienDuFichier.Text = ""
        TextBoxNomFichier.Text = ""

        ButtonSelectFile.Enabled = False
        TextBoxNomDuLienDuFichier.Enabled = False
        TextBoxNomFichier.Enabled = False
        ButtonAddDocument.Enabled = False

        GroupBoxInfosDocument.Enabled = True
    End Sub

    Private Sub BackgroundWorker_SupprimeDocument_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_SupprimeDocument.DoWork


        Dim Arguments_SQL As New Argument_SupprimeDocument

        Arguments_SQL = e.Argument

        Dim Retour_Recherche As New RetourSupprimeDocument

        Retour_Recherche = RequêteSQLSupprimeDocument(Arguments_SQL.Chrono, Arguments_SQL.NomDuLien, Arguments_SQL.chaine_de_connexion, Arguments_SQL.Nom_de_la_base_de_données, Arguments_SQL.Auteur, Arguments_SQL.langue_logiciel)



        If (Retour_Recherche.Erreur = True) Then
            'Type_de_message
            '0 = pas de message
            '1 = Suppression document Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageSupprimeDocument = 2  'Type_message = 2 => Erreur 
            My.Settings.MessageSupprimeDocument = Retour_Recherche.Infos_requêtes
            Exit Sub

        End If



        If (Retour_Recherche.Erreur = False) Then
            'Type_de_message
            '0 = pas de message
            '1 = Suppression document Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageSupprimeDocument = 1
            My.Settings.MessageSupprimeDocument = ""

            ' chargement de ListBoxDocuments

            Dim ArgumentChargementListeDocuments As New Argument_ListeDocuments
            ArgumentChargementListeDocuments.chaine_de_connexion = Me.ChaineDeConnexion
            ArgumentChargementListeDocuments.Nom_de_la_base_de_données = Me.NomDeLaBase
            ArgumentChargementListeDocuments.langue_logiciel = Me.LangueLogiciel
            ArgumentChargementListeDocuments.Chrono = Me.Chrono


            If Not BackgroundWorkerChargementListeDocuments.IsBusy = True Then

                TimerChargementListeDocuments.Start()
                BackgroundWorkerChargementListeDocuments.RunWorkerAsync(ArgumentChargementListeDocuments)

            End If


            Exit Sub

        End If

    End Sub

    Private Sub TimerEtatSupprimeDocument_Tick(sender As System.Object, e As System.EventArgs) Handles TimerEtatSupprimeDocument.Tick
        'Type_de_message
        '0 = pas de message
        '1 = Suppression document Terminé
        '2 = Erreur Requête



        If My.Settings.TypeMessageSupprimeDocument = 1 Then
            '1 = Suppression document Terminé

            ' PHASE 1 - Indiquer la fin de la suppression du documents dans 'documents' 
            ToolStripStatusLabel1.Text = MessageUtilisateur(296, LangueLogiciel)
            Me.Refresh()
            'pour apercevoir le message 'Terminé'

            Dim pauseTime As Integer = 1000
            System.Threading.Thread.Sleep(pauseTime)

            '
            ToolStripStatusLabel1.Text = "---"

            If (Not BackgroundWorkerChargementListeDocuments.IsBusy = True) Then

                ListBoxDocuments.Items.Clear()

                Dim i02 As Integer = 0

                If ListeDesDocumentsPourLeChrono.Count = 0 Then GoTo Fin0

                For i02 = 0 To (ListeDesDocumentsPourLeChrono.Count - 1)

                    Select Case LangueLogiciel
                        Case 0 ' Allemand
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierALL)
                        Case 1 ' Anglais
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierANG)
                        Case 2 ' Espagnol
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierESP)
                        Case 3 ' Français
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierFRA)
                        Case 4 ' Italien
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierITA)
                    End Select


                Next

                'On réactive le listeBox
                'et le bouton effacement document
Fin0:           ListBoxDocuments.Enabled = True
                ButtonEffaceUnDocument.Enabled = True

                My.Settings.TypeMessageSupprimeDocument = 0
                My.Settings.MessageSupprimeDocument = ""
                TimerEtatSupprimeDocument.Stop()
                Exit Sub

            End If

        End If



        If My.Settings.TypeMessageMaintenanceTransfert = 2 Then
            '2 = Erreur Requête

            MsgBox(My.Settings.MessageMaintenanceTransfert, MsgBoxStyle.Critical, "- Error -")
            TimerEtatSupprimeDocument.Stop()
            My.Settings.TypeMessageSupprimeDocument = 0
            My.Settings.MessageSupprimeDocument = ""

            'On réactive le listeBox
            'et le bouton effacement document
            ListBoxDocuments.Enabled = True
            ButtonEffaceUnDocument.Enabled = True


        End If


    End Sub

    Private Sub ComboBoxTypeDocument_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxTypeDocument.SelectedIndexChanged

        NouveauDocument.NomTypeDocument = ComboBoxTypeDocument.SelectedItem

        Dim i08 As Integer = 0

        For i08 = 0 To ListOfTypesDocuments.Count - 1

            If (ListOfTypesDocuments.Item(i08).DescriptionTypeDocument = NouveauDocument.NomTypeDocument) Then

                NouveauDocument.NuméroTypeDocument = ListOfTypesDocuments.Item(i08).TypeDocument
                NouveauDocument.TypeTransfert = ListOfTypesDocuments.Item(i08).TypeTransfert
                NouveauDocument.Destination = ListOfTypesDocuments.Item(i08).Destination
                NouveauDocument.Utilisateur = ListOfTypesDocuments.Item(i08).User
                NouveauDocument.MotDePasse = ListOfTypesDocuments.Item(i08).MotDePasse

            End If

        Next

    End Sub

    Private Sub TextBoxNomFichier_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBoxNomFichier.TextChanged


        If Me.BooleanAllemand = True Then ' Allemand
            Me.NouveauDocument.NomFichierALL = TextBoxNomFichier.Text

        End If

        If Me.BooleanAnglais = True Then ' Anglais
            Me.NouveauDocument.NomFichierANG = TextBoxNomFichier.Text

        End If

        If Me.BooleanEspagnol = True Then ' Espagnol
            Me.NouveauDocument.NomFichierESP = TextBoxNomFichier.Text

        End If

        If Me.BooleanFrançais = True Then ' Français
            Me.NouveauDocument.NomFichierFRA = TextBoxNomFichier.Text

        End If

        If Me.BooleanItalien = True Then ' Italien
            Me.NouveauDocument.NomFichierITA = TextBoxNomFichier.Text

        End If


    End Sub

    Private Sub TextBoxNomDuLienDuFichier_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBoxNomDuLienDuFichier.TextChanged

        If Me.BooleanAllemand = True Then ' Allemand
            Me.NouveauDocument.NomDuLienFichierALL = TextBoxNomDuLienDuFichier.Text

        End If

        If Me.BooleanAnglais = True Then ' Anglais
            Me.NouveauDocument.NomDuLienFichierANG = TextBoxNomDuLienDuFichier.Text

        End If

        If Me.BooleanEspagnol = True Then ' Espagnol
            Me.NouveauDocument.NomDuLienFichierESP = TextBoxNomDuLienDuFichier.Text

        End If

        If Me.BooleanFrançais = True Then ' Français
            Me.NouveauDocument.NomDuLienFichierFRA = TextBoxNomDuLienDuFichier.Text

        End If

        If Me.BooleanItalien = True Then ' Italien
            Me.NouveauDocument.NomDuLienFichierITA = TextBoxNomDuLienDuFichier.Text

        End If

    End Sub

    Private Sub BackgroundWorker_AjoutDocument_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_AjoutDocument.DoWork


        Dim Arguments_SQL As New Argument_AjoutDocument

        Arguments_SQL = e.Argument

        Dim Retour_Recherche As New RetourAjoutDocument

        Retour_Recherche = RequêteSQLAjoutDocument(Arguments_SQL.Chrono, Arguments_SQL.chaine_de_connexion, Arguments_SQL.Nom_de_la_base_de_données, Arguments_SQL.langue_logiciel, Arguments_SQL.Auteur)


        If (Retour_Recherche.Erreur = True) Then
            'Type_de_message
            '0 = pas de message
            '1 = Ajout document Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageAjoutDocument = 2  'Type_message = 2 => Erreur 
            My.Settings.MessageAjoutDocument = Retour_Recherche.Infos_requêtes
            Exit Sub

        End If



        If (Retour_Recherche.Erreur = False) Then
            'Type_de_message
            '0 = pas de message
            '1 = Ajout document Terminé
            '2 = Erreur Requête


            My.Settings.TypeMessageAjoutDocument = 1
            My.Settings.MessageAjoutDocument = ""


            ' chargement de ListBoxDocuments

            Dim ArgumentChargementListeDocuments As New Argument_ListeDocuments
            ArgumentChargementListeDocuments.chaine_de_connexion = Me.ChaineDeConnexion
            ArgumentChargementListeDocuments.Nom_de_la_base_de_données = Me.NomDeLaBase
            ArgumentChargementListeDocuments.langue_logiciel = Me.LangueLogiciel
            ArgumentChargementListeDocuments.Chrono = Me.Chrono


            If Not BackgroundWorkerChargementListeDocuments.IsBusy = True Then

                TimerChargementListeDocuments.Start()
                BackgroundWorkerChargementListeDocuments.RunWorkerAsync(ArgumentChargementListeDocuments)

            End If


            Exit Sub

        End If




    End Sub

    Private Sub TimerEtatAjoutDocument_Tick(sender As System.Object, e As System.EventArgs) Handles TimerEtatAjoutDocument.Tick
        'Type_de_message
        '0 = pas de message
        '1 = Ajout document Terminé
        '2 = Erreur Requête


        If My.Settings.TypeMessageAjoutDocument = 1 Then
            '1 = Ajout document Terminé

            ' PHASE 1 - Indiquer la fin de l' Ajout document Terminé
            ToolStripStatusLabel1.Text = MessageUtilisateur(320, LangueLogiciel)
            Me.Refresh()
            'pour apercevoir le message 'Terminé'

            Dim pauseTime As Integer = 1000
            System.Threading.Thread.Sleep(pauseTime)

            ToolStripStatusLabel1.Text = "---"


            If (Not BackgroundWorkerChargementListeDocuments.IsBusy = True) Then

                ListBoxDocuments.Enabled = False

                ListBoxDocuments.Items.Clear()

                Dim i02 As Integer = 0

                If ListeDesDocumentsPourLeChrono.Count = 0 Then GoTo Fin0

                For i02 = 0 To (ListeDesDocumentsPourLeChrono.Count - 1)

                    Select Case LangueLogiciel
                        Case 0 ' Allemand
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierALL)
                        Case 1 ' Anglais
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierANG)
                        Case 2 ' Espagnol
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierESP)
                        Case 3 ' Français
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierFRA)
                        Case 4 ' Italien
                            ListBoxDocuments.Items.Add(ListeDesDocumentsPourLeChrono.Item(i02).NomDuLienFichierITA)
                    End Select


                Next

                'On réactive le listeBox
                'et le bouton effacement document
Fin0:           ListBoxDocuments.Enabled = True
            End If



            My.Settings.TypeMessageAjoutDocument = 0
            My.Settings.MessageAjoutDocument = ""
            TimerEtatAjoutDocument.Stop()
            Exit Sub
        End If



        If My.Settings.TypeMessageAjoutDocument = 2 Then
            '2 = Erreur Requête

            MsgBox(My.Settings.MessageAjoutDocument, MsgBoxStyle.Critical, "- Error -")
            TimerEtatAjoutDocument.Stop()
            My.Settings.TypeMessageAjoutDocument = 0
            My.Settings.MessageAjoutDocument = ""

        End If

    End Sub
End Class
