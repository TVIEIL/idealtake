Namespace CLASS_IDEALTAKE.Objects

    ''' <summary>
    ''' La classe Chrono_IDRef met en relation une référence 'Ref' avec un numéro d'enregistrement unique 'Chrono'.
    ''' </summary>
    Public Class Chrono_IDRef
        Private m_Chrono As Integer = -1
        Private m_IDRef As Integer = -1
        Private m_Quantite_Ref As String = ""

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_Chrono As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Chrono_IDRef
        '''num_Chrono.Chrono = 4
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' Le champ <c>Chrono</c> indique le numéro d'enregistrement.
        ''' </summary>
        Public Property Chrono() As Integer
            Get
                Return m_Chrono

            End Get
            Set(ByVal value As Integer)
                m_Chrono = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_IDRef As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Chrono_IDRef
        '''num_IDRef.IDRef = 23
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' Le champ <c>IDRef</c> indique le numéro d'une référence.
        ''' </summary>
        Public Property IDRef() As Integer
            Get
                Return m_IDRef

            End Get
            Set(ByVal value As Integer)
                m_IDRef = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Quantité_Ref As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Chrono_IDRef
        '''Quantité_Ref.Quantite_Ref = "2 ou 3"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Quantite_Ref</c> indique le nombre (sous la forme d'une chaîne) de référence ou une description des besoins.
        ''' </summary>
        Public Property Quantite_Ref() As String
            Get
                Return m_Quantite_Ref

            End Get
            Set(ByVal value As String)
                m_Quantite_Ref = value
            End Set
        End Property

    End Class

End Namespace