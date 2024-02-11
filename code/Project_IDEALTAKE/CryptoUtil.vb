'IDEALTAKE
'---------
'
'
'DESCRIPTION :
'
'Moteur de recherche de documents pdf utilisant le SGBD MysQL.
'Ce logiciel a été developpé en VISUAL BASIC 2005, VISUAL BASIC 2008.
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

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

' Utilitaire de cryptographie
Public Class CryptoUtil


    Dim Rijndael As New RijndaelManaged   ' Fournit une implémentation de l'algorithme de Rijndael pour le cryptage et le décryptage des données

    Private m_Key() As Byte                     ' Clé de cryptage
    Private m_IV() As Byte                      ' Vecteur d'interruption

    ' Constructeur par défaut : ne fait rien
    Public Sub New()
    End Sub

    ' Constructeur qui initialise la clé de cryptage et le vecteur d'interruption.
    ' Les paramètres Key et IV représentent des tableaux de 16 octets
    Public Sub New(ByVal Key() As Byte, ByVal iv() As Byte)
        Me.Key = Key
        Me.IV = iv
    End Sub

    ' Autre constructeur qui initialise la clé de cryptage et le vecteur d'interruption.
    ' Les paramètres sKey et sIV représentent des chaînes de 8 caractères UNICODE.
    Public Sub New(ByVal sKey As String, ByVal sIV As String)
        Dim ue As New UnicodeEncoding   ' Servira à convertir les chaînes en tableaux d'octets

        If sKey.Length > 8 Then         ' Si la chaîne est trop longue, on la tronque
            sKey = sKey.Substring(0, 8)
        ElseIf sKey.Length < 8 Then     ' Si par contre la chaîne est trop courte, on la complète
            Dim n As Integer = 8 - sKey.Length

            For i As Integer = 0 To n - 1
                sKey &= i
            Next
        End If

        If sIV.Length > 8 Then
            sIV = sIV.Substring(0, 8)
        ElseIf sIV.Length < 8 Then
            Dim n As Integer = 8 - sIV.Length

            For i As Integer = 0 To n - 1
                sIV &= i
            Next
        End If

        ' Conversion des deux chaînes en tableaux d'octets
        Key = ue.GetBytes(sKey)
        IV = ue.GetBytes(sIV)
    End Sub

    ' Accesseur pour la clé de cryptage
    Public Property Key() As Byte()
        Get
            Return m_Key
        End Get
        Set(ByVal Value As Byte())
            m_Key = Value
        End Set
    End Property

    ' Accesseur pour le vecteur d'interruption
    Public Property iv() As Byte()
        Get
            Return m_IV
        End Get
        Set(ByVal Value As Byte())
            m_IV = Value
        End Set
    End Property

    ' Crypte le texte reçu du flux input et dirige le resultat vers le flux output
    Public Sub Encrypt(ByVal input As Stream, ByVal output As Stream)
        ' L'écriture dans le flux cs produira une sortie cryptée selon l'algorithme de Rijndael
        ' avec m_Key comme clé et m_IV comme vecteur d'interruption
        Dim cs As New CryptoStream(output, Rijndael.CreateEncryptor(m_Key, m_IV), CryptoStreamMode.Write)
        Dim buffer(100) As Byte ' tampon utilisé pour la lecture et l'écriture
        Dim n As Integer        ' nombre d'octets réellement lus ou écrits

        ' La suite ressemble à un transfert de données entre deux flux quelconques
        Do
            n = input.Read(buffer, 0, 100)
            cs.Write(buffer, 0, n)
        Loop Until n < 100

        cs.Close()
    End Sub

    ' Décrypte le texte reçu du flux input et dirige le resultat vers le flux output
    Public Sub Decrypt(ByVal input As Stream, ByVal output As Stream)
        ' La lecture à partir du flux cs est précédée du décryptage des données d'origine.
        ' La clé de cryptage et le vecteur d'interruption doivent être les mêmes que
        ' ceux utilisés lors du cryptage des données.
        Try
            Dim cs As New CryptoStream(input, Rijndael.CreateDecryptor(m_Key, m_IV), CryptoStreamMode.Read)
            Dim buffer(200) As Byte
            Dim n As Integer

            Do
                n = cs.Read(buffer, 0, 100)
                output.Write(buffer, 0, n)
            Loop Until n < 100

            cs.Close()
        Catch Ex As System.Security.Cryptography.CryptographicException

            Dim infosLog As String = ""

            infosLog = "Sub Decrypt() :Not OK " & vbCrLf & Ex.Message

            Form_Demarrage.Log_Idealtake(My.Application.Info.DirectoryPath, infosLog)

        End Try

    End Sub

    ' Crypte un texte en mémoire et retourne le texte crypté
    Public Function EncryptText(ByVal original As String) As String
        ' On crée un MemoryStream avec le texte original
        Dim UE As New UnicodeEncoding
        Dim msIn As New MemoryStream(UE.GetBytes(original))
        ' et on en crée un autre pour la sortie
        Dim msOut As New MemoryStream

        ' Les deux flux sont transmis à la méthode encrypt
        Encrypt(msIn, msOut)
        msIn.Close()

        ' Puis, on recupère le contenu du flux de sortie
        Dim encrypted As String = UE.GetString(msOut.ToArray)
        msOut.Close()

        ' Qu'on retourne en guise de resultat
        Return encrypted
    End Function

    ' Décrypte un texte en mémoire et retourne le texte décrypté
    ' Le principe est analogue à celui de EncryptText
    Public Function DecryptText(ByVal encrypted As String) As String
        Dim UE As New UnicodeEncoding
        Dim msIn As New MemoryStream(UE.GetBytes(encrypted))
        Dim msOut As New MemoryStream

        Decrypt(msIn, msOut)
        msIn.Close()

        Dim decrypted As String = UE.GetString(msOut.ToArray)
        msOut.Close()

        Return decrypted
    End Function

    ' Crypte un fichier texte et produit un autre fichier texte complètement illisible
    Public Sub EncryptFile(ByVal original As String, ByVal encrypted As String)
        ' On obtient un flux qui représente le fichier original
        Dim fsIn As New FileStream(original, FileMode.Open)
        ' On en obtient un autre qui représente le fichier crypté
        Dim fsOut As New FileStream(encrypted, FileMode.Create)

        ' On trasmet les deux flux à la méthode Encrypt et le travail est fait
        Encrypt(fsIn, fsOut)
        fsIn.Close()
        fsOut.Close()
    End Sub

    ' Décrypte un fichier texte crypté pour reconstituer le fichier original
    ' Le principe est analogue à celui de EncryptFile
    Public Sub DecryptFile(ByVal encrypted As String, ByVal decrypted As String)
        Dim fsIn As New FileStream(encrypted, FileMode.Open)
        Dim fsOut As New FileStream(decrypted, FileMode.Create)

        Decrypt(fsIn, fsOut)
        fsIn.Close()
        fsOut.Close()
    End Sub

    ' Crypte un texte en mémoire à destination d'un fichier sur disque
    Public Sub EncryptageTexteduFichier(ByVal original As String, ByVal fileName As String)
        Try
            ' On crée un MemoryStream avec le texte original
            Dim UE As New UnicodeEncoding
            Dim msIn As New MemoryStream(UE.GetBytes(original))
            ' On obtient un flux de sortie qui représente le fichier crypté
            Dim fsOut As New FileStream(fileName, FileMode.Create)

            ' On trasmet les deux flux à la méthode Encrypt et le travail est fait
            Encrypt(msIn, fsOut)
            msIn.Close()
            fsOut.Close()
        Catch ex As System.IO.IOException
            MessageBox.Show(ex.Message, "Erreur EncryptageTexteduFichier()", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Dim infosLog As String = ""
            infosLog = "Sub EncryptageTexteduFichier() :Not OK " & vbCrLf & ex.Message
            Form_Demarrage.Log_Idealtake(My.Application.Info.DirectoryPath, infosLog)

        End Try

    End Sub

    ' Décrypte un fichier texte crypté et retourne le texte original
    ' sous la forme d'une très longue chaîne de caractères.
    ' Le principe est analogue à celui de EncryptTextToFile
    Public Function DecryptageTexteduFichier(ByVal fileName As String) As String
        Dim decrypted As String = ""
        Try
            Dim UE As New UnicodeEncoding
            Dim fsIn As New FileStream(fileName, FileMode.Open)
            Dim msOut As New MemoryStream

            Decrypt(fsIn, msOut)
            fsIn.Close()

            decrypted = UE.GetString(msOut.ToArray)
            msOut.Close()

        Catch ex As System.IO.IOException
            MessageBox.Show(ex.Message, "Erreur DecryptageTexteduFichier()", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Dim infosLog As String = ""
            infosLog = "Sub DecryptageTexteduFichier() :Not OK " & vbCrLf & ex.Message
            Form_Demarrage.Log_Idealtake(My.Application.Info.DirectoryPath, infosLog)

        End Try

        Return decrypted

    End Function

End Class

