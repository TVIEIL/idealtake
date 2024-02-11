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
'Copyright (C) 2007-2011 VIEIL Thierry Ce programme est un logiciel libre ;
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


NotInheritable Class Message_Box_Question


    Private m_Titre_Boite As String = ""
    Private m_Question_Boite As String = ""
    Private m_Oui_Boite As String = ""
    Private m_Non_Boite As String = ""
    Private m_Réponse_question As Boolean = False

    Public Property Titre_Boite() As String

        Get
            Return m_Titre_Boite
        End Get

        Set(ByVal value As String)
            m_Titre_Boite = value
        End Set

    End Property

    Public Property Question_Boite() As String

        Get
            Return m_Question_Boite
        End Get

        Set(ByVal value As String)
            m_Question_Boite = value
        End Set

    End Property

    Public Property Oui_Boite() As String

        Get
            Return m_Oui_Boite
        End Get

        Set(ByVal value As String)
            m_Oui_Boite = value
        End Set

    End Property

    Public Property Non_Boite() As String

        Get
            Return m_Non_Boite
        End Get

        Set(ByVal value As String)
            m_Non_Boite = value
        End Set

    End Property

    Public Property Réponse_question() As Boolean

        Get
            Return m_Réponse_question
        End Get

        Set(ByVal value As Boolean)
            m_Réponse_question = value
        End Set

    End Property


    Function OuiButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OuiButton.Click
        Me.Hide()
        Réponse_question = True

        Return True
    End Function

    Function NonButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NonButton.Click
        Me.Hide()
        Réponse_question = False

        Return True
    End Function


    Public Sub informations(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Titre_Boite <> "" Then
            Me.Text = Titre_Boite
        End If

        Question_Boite = "   " & Question_Boite
        Me.ListBox1.Items(2) = Question_Boite
        Dim Longeur_Question_Boite As Integer = 0
        Longeur_Question_Boite = Question_Boite.Length

        If (Longeur_Question_Boite > 40) Then
            Me.Width = Me.Width + ((Longeur_Question_Boite - 40) * 6)
        End If

        If Oui_Boite <> "" Then
            Me.OuiButton.Text = Oui_Boite
        End If

        If Non_Boite <> "" Then
            Me.NonButton.Text = Non_Boite
        End If

        Me.Show()
        Me.Enabled = True

        'sender.Enabled = False

        Me.SetDesktopLocation((My.Computer.Screen.WorkingArea.Width / 2) - 130, _
                               (My.Computer.Screen.WorkingArea.Height / 2) - 70)
        'Me.StartPosition  = FormStartPosition.CenterScreen

        Me.BringToFront()
        'MsgBox(sender.ToString)

        Me.Titre_Boite = ""
        Me.Question_Boite = ""
        Me.Question_Boite = ""
        Me.Question_Boite = ""
        Me.Oui_Boite = ""
        Me.Non_Boite = ""

    End Sub

End Class