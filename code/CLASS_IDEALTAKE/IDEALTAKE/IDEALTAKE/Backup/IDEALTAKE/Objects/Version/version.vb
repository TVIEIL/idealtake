Namespace CLASS_IDEALTAKE.Objects



    Public Class version

        Private m_Version As String

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim String_Version_CLASS_IDEALTAKE As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.version
        '''Mssgbox(String_Version_CLASS_IDEALTAKE.VERSION_CLASS_IDEALTAKE))
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La fonction <c>VERSION_CLASS_IDEALTAKE()</c> retourne une chaîne identifiant le nom interne de la classe et sa version.
        ''' </summary>
        ''' 
        Public Function VERSION_CLASS_IDEALTAKE() As String

            'on récupère la version de CLASS_IDEALTAKE

            Dim Assembly_CLASS_IDEALTAKE As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
            Dim FileVersionInfo_CLASS_IDEALTAKE As System.Diagnostics.FileVersionInfo = FileVersionInfo.GetVersionInfo(Assembly_CLASS_IDEALTAKE.Location)

            m_Version = FileVersionInfo_CLASS_IDEALTAKE.InternalName.ToString & " - VERSION : " & _
                        FileVersionInfo_CLASS_IDEALTAKE.ProductVersion.ToString

            'FileVersionInfo_CLASS_IDEALTAKE.ProductMajorPart.ToString & "." & _
            'FileVersionInfo_CLASS_IDEALTAKE.ProductMinorPart.ToString & "." & _
            'FileVersionInfo_CLASS_IDEALTAKE.ProductBuildPart.ToString & "." & _
            'FileVersionInfo_CLASS_IDEALTAKE.PrivateBuild.ToString & "*-* " & _
            'FileVersionInfo_CLASS_IDEALTAKE.InternalName.ToString & "*-*" & _

            Return m_Version

        End Function

    End Class

End Namespace