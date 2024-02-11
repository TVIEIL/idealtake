Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions

<PermissionSet(SecurityAction.Demand, Name:="FullTrust")> _
<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
Public Class Form1
    Inherits Form


    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        WebBrowser1.ObjectForScripting = Me

        WebBrowser1.DocumentText = _
    "<html><head><script>" & _
    "function test(message) { alert(message); }" & _
    "</script></head><body><button " & _
    "onclick=""window.external.Test('called from script code')"" > " & _
    "call client code from script code</button>" & _
    "</body></html>"

    End Sub

    Public Sub Test(ByVal message As String)
        MessageBox.Show(message, "Infos provenant du javascript")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        WebBrowser1.Document.InvokeScript("test", _
    New String() {"Salut les filles!"})
    End Sub
End Class
