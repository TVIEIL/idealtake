Public Class UserControl1

    Private Sub Effacer_les_Références_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Effacer_les_Références.Click
        Liste_des_références_TextBox.Text = ""
    End Sub


    Private Sub Liste_des_références_Menu_Strip_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles Liste_des_références_Menu_Strip.ItemClicked
        'MsgBox(sender)
    End Sub



End Class
