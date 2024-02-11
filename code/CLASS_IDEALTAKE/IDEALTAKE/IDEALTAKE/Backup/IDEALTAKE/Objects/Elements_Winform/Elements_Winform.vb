

Namespace CLASS_IDEALTAKE.Objects


    ''' <summary>
    ''' La classe Elements_Winform stocke les paramètres des différents winform de l'application.
    ''' </summary>
    Public Class Elements_Winform

        Private m_IDElements_Windows_Forms As Integer = -1
        Private m_IDConstructeur As Integer = -1
        Private m_AutoCompleteMode As String = ""
        Private m_AutoCompleteSource As String = ""
        Private m_BackColor As String = ""
        Private m_Background_image As String = ""
        Private m_BackGroungImageLayout As String = ""
        Private m_BorderStyle As String = ""

        'Fonction affecté à cet événement
        Private m_CheckStateChanged As String = ""

        'Fonction affecté à cet événement
        Private m_Click As String = ""

        Private m_Cursor As String = ""
        Private m_Dock As String = ""
        Private m_DropDownHeight As String = ""
        Private m_DropDownStyle As String = ""
        Private m_DropDownWidth As String = ""
        Private m_Enable As String = ""
        Private m_Font As String = ""
        Private m_ForeColor As String = ""
        Private m_Form_Parent As String = ""
        Private m_FormatString As String = ""
        Private m_hauteur As Integer = -1
        Private m_Image As String = ""
        Private m_ImageLocation As String = ""
        Private m_ItemHeight As Integer = -1
        Private m_Items_Allemand As String = ""
        Private m_Items_Anglais As String = ""
        Private m_Items_Espagnole As String = ""
        Private m_Items_Francais As String = ""
        Private m_Items_Italien As String = ""
        Private m_largeur As Integer = -1
        Private m_MaxLength As Integer = -1
        Private m_Multiline As String = ""
        Private m_Nom_Elements_Windows_Forms As String = ""
        Private m_Opacity As Integer = -1
        Private m_Position_X As Integer = -1
        Private m_Position_Y As Integer = -1
        Private m_ReadOnly As String = ""
        Private m_ScrollBars As String = ""

        'Fonction affecté à cet événement
        Private m_SelectedValueChanged As String = ""

        Private m_Sorted As Boolean = False
        Private m_Tab_Index As Integer = -1
        Private m_Tab_Stop As String = ""
        Private m_Text_Allemand As String = ""
        Private m_Text_Anglais As String = ""
        Private m_Text_Espagnole As String = ""
        Private m_Text_Francais As String = ""
        Private m_Text_Italien As String = ""
        Private m_TextChanged As String = ""
        Private m_Text_Français As String = ""
        Private m_ToolTip_Allemand As String = ""
        Private m_ToolTip_Anglais As String = ""
        Private m_ToolTip_Espagnole As String = ""
        Private m_ToolTip_Français As String = ""
        Private m_ToolTip_Italien As String = ""
        Private m_Type_Elements_Windows_Forms As String = ""
        Private m_Visibilité As Boolean = False
        Private m_WordWrap As String = ""



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_IDElements_Windows_Forms As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''num_IDElements_Windows_Forms.IDElements_Windows_Forms = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDElements_Windows_Forms</c> est un entier identifiant de façon unique un élément winform.
        ''' </summary>
        Public Property IDElements_Windows_Forms() As Integer
            Get
                Return m_IDElements_Windows_Forms

            End Get
            Set(ByVal value As Integer)
                m_IDElements_Windows_Forms = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim num_IDConstructeur As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''num_IDConstructeur.IDConstructeur = 2
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>IDConstructeur</c> est un entier identifiant de façon unique un constructeur.
        ''' </summary>
        Public Property IDConstructeur() As Integer
            Get
                Return m_IDConstructeur

            End Get
            Set(ByVal value As Integer)
                m_IDConstructeur = value
            End Set
        End Property


        ''' <remarks>
        ''' AutoCompleteMode.None = 0,
        ''' AutoCompleteMode.Suggest = 1,                                                          
        ''' AutoCompleteMode.Append = 2,                                                             
        ''' AutoCompleteMode.SuggestAppend = 3;                                                    
        ''' 
        ''' <example>
        ''' <code>
        '''Dim num_IDConstructeur As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''num_IDConstructeur.AutoCompleteMode = ...
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>AutoCompleteMode</c>
        ''' </summary>
        Public Property AutoCompleteMode() As String
            Get
                Return m_AutoCompleteMode

            End Get
            Set(ByVal value As String)
                m_AutoCompleteMode = value
            End Set
        End Property


        ''' <remarks>
        ''' AutoCompleteSource.FileSystem = 1,
        ''' AutoCompleteSource.HistoryList = 2,
        ''' AutoCompleteSource.RecentlyUsedList = 4,
        ''' AutoCompleteSource.AllUrl = 6,
        ''' AutoCompleteSource.AllSystemSources = 7,
        ''' AutoCompleteSource.FileSystemDirectories = 32,
        ''' AutoCompleteSource.CustomSource = 64,
        ''' AutoCompleteSource.None = 128,
        ''' AutoCompleteSource.ListItems = 256;
        ''' 
        ''' <example>
        ''' <code>
        '''Dim num_IDConstructeur As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''num_IDConstructeur.AutoCompleteSource = ...
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' propriété <c>AutoCompleteSource</c>
        ''' </summary>
        Public Property AutoCompleteSource() As String
            Get
                Return m_AutoCompleteSource

            End Get
            Set(ByVal value As String)
                m_AutoCompleteSource = value
            End Set
        End Property


        ''' <remarks>
        ''' 
        ''' aliceblue #F0F8FF,
        ''' antiquewhite #FAEBD7,
        ''' aqua #00FFFF,
        ''' aquamarine #7FFFD4,
        ''' azure #F0FFFF,
        ''' beige #F5F5DC,
        ''' bisque #FFE4C4,
        ''' black #000000,
        ''' blanchedalmond #FFEBCD,
        ''' blue #0000FF,
        ''' blueviolet #8A2BE2,
        ''' brown #A52A2A,
        ''' burlywood #DEB887,
        ''' cadetblue #5F9EA0,
        ''' chartreuse #7FFF00,
        ''' chocolate #D2691E,
        ''' coral #FF7F50,
        ''' cornflowerblue #6495ED,
        ''' cornsilk #FFF8DC,
        ''' crimson #DC143C,
        ''' cyan #00FFFF,
        ''' darkblue #00008B,
        ''' darkcyan #008B8B,
        ''' darkgoldenrod #B8860B,
        ''' darkgray #A9A9A9,
        ''' darkgreen #006400,
        ''' darkkhaki #BDB76B,
        ''' darkmagenta #8B008B,
        ''' darkolivegreen #556B2F,
        ''' darkorange #FF8C00,
        ''' darkorchid #9932CC,
        ''' darkred #8B0000,
        ''' darksalmon #E9967A,
        ''' darkseagreen #8FBC8F,
        ''' darkslateblue #483D8B,
        ''' darkslategray #2F4F4F,
        ''' darkturquoise #00CED1,
        ''' darkviolet #9400D3,
        ''' deeppink #FF1493,
        ''' deepskyblue #00BFFF,
        ''' dimgray #696969,
        ''' dodgerblue #1E90FF,
        ''' firebrick #B22222,
        ''' floralwhite #FFFAF0,
        ''' forestgreen #228B22,
        ''' fuchsia #FF00FF,
        ''' gainsboro #DCDCDC,
        ''' ghostwhite #F8F8FF,
        ''' gold #FFD700,
        ''' goldenrod #DAA520,
        ''' gray #808080,
        ''' green #008000,
        ''' greenyellow #ADFF2F,
        ''' honeydew #F0FFF0,
        ''' hotpink #FF69B4,
        ''' indianred #CD5C5C,
        ''' indigo #4B0082,
        ''' ivory #FFFFF0,
        ''' khaki #F0E68C,
        ''' lavender #E6E6FA,
        ''' lavenderblush #FFF0F5,
        ''' lawngreen #7CFC00,
        ''' lemonchiffon #FFFACD,
        ''' lightblue #ADD8E6,
        ''' lightcoral #F08080,
        ''' lightcyan #E0FFFF,
        ''' lightgoldenrodyellow #FAFAD2,
        ''' lightgreen #90EE90,
        ''' lightgrey #D3D3D3,
        ''' lightpink #FFB6C1,
        ''' lightsalmon #FFA07A,
        ''' lightseagreen #20B2AA,
        ''' lightskyblue #87CEFA,
        ''' lightslategray #778899,
        ''' lightsteelblue #B0C4DE,
        ''' lightyellow #FFFFE0,
        ''' lime #00FF00,
        ''' limegreen #32CD32,
        ''' linen #FAF0E6,
        ''' magenta #FF00FF,
        ''' maroon #800000,
        ''' mediumaquamarine #66CDAA,
        ''' mediumblue #0000CD,
        ''' mediumorchid #BA55D3,
        ''' mediumpurple #9370DB,
        ''' mediumseagreen #3CB371,
        ''' mediumslateblue #7B68EE,
        ''' mediumspringgreen #00FA9A,
        ''' mediumturquoise #48D1CC,
        ''' mediumvioletred #C71585,
        ''' midnightblue #191970,
        ''' mintcream #F5FFFA,
        ''' mistyrose #FFE4E1,
        ''' moccasin #FFE4B5,
        ''' navajowhite #FFDEAD,
        ''' navy #000080,
        ''' oldlace #FDF5E6,
        ''' olive #808000,
        ''' olivedrab #6B8E23,
        ''' orange #FFA500,
        ''' orangered #FF4500,
        ''' orchid #DA70D6,
        ''' palegoldenrod #EEE8AA,
        ''' palegreen #98FB98,
        ''' paleturquoise #AFEEEE,
        ''' palevioletred #DB7093,
        ''' papayawhip #FFEFD5,
        ''' peachpuff #FFDAB9,
        ''' peru #CD853F,
        ''' pink #FFC0CB,
        ''' plum #DDA0DD,
        ''' powderblue #B0E0E6,
        ''' purple #800080,
        ''' red #FF0000,
        ''' rosybrown #BC8F8F,
        ''' royalblue #4169E1,
        ''' saddlebrown #8B4513,
        ''' salmon #FA8072,
        ''' sandybrown #F4A460,
        ''' seagreen #2E8B57,
        ''' seashell #FFF5EE,
        ''' sienna #A0522D,
        ''' silver #C0C0C0,
        ''' skyblue #87CEEB,
        ''' slateblue #6A5ACD,
        ''' slategray #708090,
        ''' snow #FFFAFA,
        ''' springgreen #00FF7F,
        ''' steelblue #4682B4,
        ''' tan #D2B48C,
        ''' teal #008080,
        ''' thistle #D8BFD8,
        ''' tomato #FF6347,
        ''' turquoise #40E0D0,
        ''' violet #EE82EE,
        ''' wheat #F5DEB3,
        ''' white #FFFFFF,
        ''' whitesmoke #F5F5F5,
        ''' yellow #FFFF00,
        ''' yellowgreen #9ACD32
        ''' ;
        ''' 
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.BackColor = "white"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>BackColor</c> est une chaîne représentant la couleur de fond d'un élément winform.
        ''' </summary>
        Public Property BackColor() As String
            Get
                Return m_BackColor

            End Get
            Set(ByVal value As String)
                m_BackColor = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Background_image = "image.jpg"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Background_image</c> est une chaîne représentant le nom de fichier d'une image d'un élément winform.
        ''' </summary>
        Public Property Background_image() As String
            Get
                Return m_Background_image

            End Get
            Set(ByVal value As String)
                m_Background_image = value
            End Set
        End Property


        ''' <remarks>
        ''' ImageLayout.None = 0,
        ''' ImageLayout.Tile = 1,
        ''' ImageLayout.Center = 2,
        ''' ImageLayout.Stretch =3,
        ''' ImageLayout.Zoom = 4;
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.BackGroungImageLayout = "(AUCUN)"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>BackGroungImageLayout</c> est une chaîne qui définit l'image d'arrière
        ''' plan pour le contrôle.
        ''' 
        ''' </summary>
        Public Property BackGroungImageLayout() As String
            Get
                Return m_BackGroungImageLayout
            End Get
            Set(ByVal value As String)
                m_BackGroungImageLayout = value
            End Set
        End Property


        ''' <remarks>
        ''' 0 = BorderStyle.None,
        ''' 1 = BorderStyle.FixedSingle,
        ''' 2 = BorderStyle.Fixed3D;
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.BorderStyle = ...
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' Propriété <c>BorderStyle</c>
        ''' 
        ''' 
        ''' </summary>
        Public Property BorderStyle() As String
            Get
                Return m_BorderStyle
            End Get
            Set(ByVal value As String)
                m_BorderStyle = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.CheckStateChanged = "Function01"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>CheckStateChanged</c> indique la fonction répondant à cet événement
        ''' </summary>
        Public Property CheckStateChanged() As String
            Get
                Return m_CheckStateChanged

            End Get
            Set(ByVal value As String)
                m_CheckStateChanged = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Click = "Function01"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Click</c> indique la fonction répondant à cet événement
        ''' </summary>
        Public Property Click() As String
            Get
                Return m_Click

            End Get
            Set(ByVal value As String)
                m_Click = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Cursor = "Hand"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Cursor</c> est une chaîne indiquant l'apparence du curseur
        ''' du curseur au passage du pointeur sur le contrôle.
        ''' 
        ''' </summary>
        Public Property Cursor() As String
            Get
                Return m_Cursor
            End Get
            Set(ByVal value As String)
                m_Cursor = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Dock = "None"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Dock</c> est une chaîne qui définit les bordures du
        ''' contrôle liées au conteneur.
        ''' 
        ''' </summary>
        Public Property Dock() As String
            Get
                Return m_Dock
            End Get
            Set(ByVal value As String)
                m_Dock = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.DropDownHeight = 8
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>DropDownHeight</c> est un entier indiquant la hauteur d'un DropDown.
        ''' </summary>
        Public Property DropDownHeight() As Integer
            Get
                Return m_DropDownHeight
            End Get
            Set(ByVal value As Integer)
                m_DropDownHeight = value
            End Set
        End Property


        ''' <remarks>
        ''' 0 = ComboBoxStyle.Simple,
        ''' 1 = ComboBoxStyle.DropDown,
        ''' 2 = ComboBoxStyle.DropDownList;
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.DropDownStyle = "DropDown"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>DropDownStyle</c> est une chaîne indiquant l'apparence et la fonctionnalité
        ''' d'une zonne de liste déroulante.
        ''' 
        ''' </summary>
        Public Property DropDownStyle() As String
            Get
                Return m_DropDownStyle
            End Get
            Set(ByVal value As String)
                m_DropDownStyle = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.DropDownWidth = 250
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>DropDownWidth</c> est un entier indiquant la largeur d'un DropDown.
        ''' </summary>
        Public Property DropDownWidth() As Integer
            Get
                Return m_DropDownWidth
            End Get
            Set(ByVal value As Integer)
                m_DropDownWidth = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Enable = true
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Enable</c> est un booléen qui indique l'accessibilité d'un élément winform.
        ''' </summary>
        Public Property Enable() As Boolean
            Get
                Return m_Enable

            End Get
            Set(ByVal value As Boolean)
                m_Enable = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Font = "Microsoft Sans Serif; 8,25pt"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Font</c> est une chaîne qui définit la police
        ''' de caractères, taille.
        ''' 
        ''' </summary>
        Public Property Font() As String
            Get
                Return m_Font
            End Get
            Set(ByVal value As String)
                m_Font = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ForeColor = "black"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ForeColor</c> est une chaîne représentant la couleur forecolor d'un élément winform.
        ''' </summary>
        Public Property ForeColor() As String
            Get
                Return m_ForeColor

            End Get
            Set(ByVal value As String)
                m_ForeColor = value
            End Set
        End Property



        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Form_Parent = "Form_Demarrage"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Form_Parent</c> est une chaîne qui définit le formulaire
        ''' dans lequel le contrôle apparait.
        ''' 
        ''' </summary>
        Public Property Form_Parent() As String
            Get
                Return m_Form_Parent
            End Get
            Set(ByVal value As String)
                m_Form_Parent = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.FormatString = "N3"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>FormatString</c> est une chaîne qui définit comment une valeur
        ''' doit être affichée.
        ''' 
        ''' </summary>
        Public Property FormatString() As String
            Get
                Return m_FormatString
            End Get
            Set(ByVal value As String)
                m_FormatString = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim pos1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''pos1.hauteur = 17
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>hauteur</c> est un entier qui représente la hauteur d'un élément winform.
        ''' </summary>
        Public Property hauteur() As Integer
            Get
                Return m_hauteur

            End Get
            Set(ByVal value As Integer)
                m_hauteur = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Image = "nom_de_l'image"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Image</c> est une chaîne qui définit le nom
        ''' de l'image dans un contrôle Button.
        ''' 
        ''' </summary>
        Public Property Image() As String
            Get
                Return m_Image
            End Get
            Set(ByVal value As String)
                m_Image = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ImageLocation = "c:\path\image01.jpg"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ImageLocation</c> est une chaîne.
        ''' Obtient ou définit le chemin d'accès ou l'URL de l'image à afficher dans un PictureBox.
        ''' </summary>
        Public Property ImageLocation() As String
            Get
                Return m_ImageLocation
            End Get
            Set(ByVal value As String)
                m_Image = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ItemHeight = 8
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ItemHeight</c> est un entier indiquant la hauteur d'un Item.
        ''' </summary>
        Public Property ItemHeight() As Integer
            Get
                Return m_ItemHeight
            End Get
            Set(ByVal value As Integer)
                m_ItemHeight = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Items = "Element1;Element2;Element3"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Items</c> est une chaîne qui définit une liste
        ''' d'éléments d'une liste déroulante Allemand.
        ''' 
        ''' </summary>
        Public Property Items_Allemand() As String
            Get
                Return m_Items_Allemand
            End Get
            Set(ByVal value As String)
                m_Items_Allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Items = "Element1;Element2;Element3"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Items</c> est une chaîne qui définit une liste
        ''' d'éléments d'une liste déroulante Anglais.
        ''' 
        ''' </summary>
        Public Property Items_Anglais() As String
            Get
                Return m_Items_Anglais
            End Get
            Set(ByVal value As String)
                m_Items_Anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Items = "Element1;Element2;Element3"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Items</c> est une chaîne qui définit une liste
        ''' d'éléments d'une liste déroulante Espagnole.
        ''' 
        ''' </summary>
        Public Property Items_Espagnole() As String
            Get
                Return m_Items_Espagnole
            End Get
            Set(ByVal value As String)
                m_Items_Espagnole = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Items = "Element1;Element2;Element3"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Items</c> est une chaîne qui définit une liste
        ''' d'éléments d'une liste déroulante Francais.
        ''' 
        ''' </summary>
        Public Property Items_Francais() As String
            Get
                Return m_Items_Francais
            End Get
            Set(ByVal value As String)
                m_Items_Francais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Items = "Element1;Element2;Element3"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Items</c> est une chaîne qui définit une liste
        ''' d'éléments d'une liste déroulante Italien.
        ''' 
        ''' </summary>
        Public Property Items_Italien() As String
            Get
                Return m_Items_Italien
            End Get
            Set(ByVal value As String)
                m_Items_Italien = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim pos1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''pos1.largeur = 150
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>largeur</c> est un entier qui représente la largeur d'un élément winform.
        ''' </summary>
        Public Property largeur() As Integer
            Get
                Return m_largeur

            End Get
            Set(ByVal value As Integer)
                m_largeur = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.MaxLength = 10
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>MaxLength</c> est un entier indiquant le nombre maximal de caractères pouvant être insérer dans la zone de liste.
        ''' </summary>
        Public Property MaxLength() As Integer
            Get
                Return m_MaxLength
            End Get
            Set(ByVal value As Integer)
                m_MaxLength = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Multiline = "True"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Multiline</c> est une chaîne qui définit l'aspect
        ''' d'un TextBox par exemple.
        ''' 
        ''' </summary>
        Public Property Multiline() As String
            Get
                Return m_Multiline
            End Get
            Set(ByVal value As String)
                m_Multiline = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim nom1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''nom1.Nom_Elements_Windows_Forms = "label1"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Nom_Elements_Windows_Forms</c> est une chaîne contenant le nom d'un élément winform.
        ''' </summary>
        Public Property Nom_Elements_Windows_Forms() As String
            Get
                Return m_Nom_Elements_Windows_Forms

            End Get
            Set(ByVal value As String)
                m_Nom_Elements_Windows_Forms = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Opacity = 25
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Opacity</c> est un entier qui indique la transparence en % d'un élément winform.
        ''' </summary>
        Public Property Opacity() As Integer
            Get
                Return m_Opacity

            End Get
            Set(ByVal value As Integer)
                m_Opacity = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim pos1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''pos1.Position_X = 25
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Position_X</c> est un entier qui représente la position en abscisse d'un élément winform.
        ''' </summary>
        Public Property Position_X() As Integer
            Get
                Return m_Position_X

            End Get
            Set(ByVal value As Integer)
                m_Position_X = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim pos1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''pos1.Position_Y = 17
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Position_Y</c> est un entier qui représente la position en ordonnée d'un élément winform.
        ''' </summary>
        Public Property Position_Y() As Integer
            Get
                Return m_Position_Y

            End Get
            Set(ByVal value As Integer)
                m_Position_Y = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Element_ReadOnly = "True"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Element_ReadOnly</c> est une chaîne qui spécifie 
        ''' Qu'une variable ou une propriété peut être lue, mais pas écrite.
        ''' 
        ''' </summary>
        Public Property Element_ReadOnly() As String
            Get
                Return m_ReadOnly
            End Get
            Set(ByVal value As String)
                m_ReadOnly = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' ScrollBars.None = 0, 
        ''' ScrollBars.Horizontal = 1,
        ''' ScrollBars.Vertical = 2, 
        ''' ScrollBars.Both = 3 
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ScrollBars = "..."
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ScrollBars</c> est une chaîne qui spécifie 
        ''' Obtient ou définit le type de barres de défilement à afficher
        ''' Dans le contrôle RichTextBox.
        ''' </summary>
        Public Property ScrollBars() As String
            Get
                Return m_ScrollBars
            End Get
            Set(ByVal value As String)
                m_ScrollBars = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.SelectedValueChanged = "Function01"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>SelectedValueChanged</c> indique la fonction répondant à cet événement
        ''' </summary>
        Public Property SelectedValueChanged() As String
            Get
                Return m_SelectedValueChanged

            End Get
            Set(ByVal value As String)
                m_SelectedValueChanged = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Sorted = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Sorted</c> est un booléen indiquant si une liste déroulante est trié.
        ''' </summary>
        Public Property Sorted() As Boolean
            Get
                Return m_Sorted
            End Get
            Set(ByVal value As Boolean)
                m_Sorted = value
            End Set
        End Property
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Tab_Index = 1
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Tab_Index</c> est un entier représentant l'ordre de tabulation d'un élément winform.
        ''' </summary>
        Public Property Tab_Index() As Integer
            Get
                Return m_Tab_Index

            End Get
            Set(ByVal value As Integer)
                m_Tab_Index = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Tab_Stop = True
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Tab_Stop</c> est un booléen qui définit si le focus
        ''' peut être donné au contrôle par la touche TABULATION.
        ''' 
        ''' </summary>
        Public Property Tab_Stop() As Boolean
            Get
                Return m_Tab_Stop
            End Get
            Set(ByVal value As Boolean)
                m_Tab_Stop = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Text_Allemand = "Menu"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Text_Allemand</c> est une chaîne représentant le texte d'un élément winform en Allemand.
        ''' </summary>
        Public Property Text_Allemand() As String
            Get
                Return m_Text_Allemand

            End Get
            Set(ByVal value As String)
                m_Text_Allemand = value
            End Set
        End Property
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Text_Anglais = "Menu"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Text_Anglais</c> est une chaîne représentant le texte d'un élément winform en Anglais.
        ''' </summary>
        Public Property Text_Anglais() As String
            Get
                Return m_Text_Anglais

            End Get
            Set(ByVal value As String)
                m_Text_Anglais = value
            End Set
        End Property
        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Text_Espagnole = "Menu"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Text_Espagnole</c> est une chaîne représentant le texte d'un élément winform en Espagnole.
        ''' </summary>
        Public Property Text_Espagnole() As String
            Get
                Return m_Text_Espagnole

            End Get
            Set(ByVal value As String)
                m_Text_Espagnole = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Text_Français = "Menu"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Text_Français</c> est une chaîne représentant le texte d'un élément winform en Français.
        ''' </summary>
        Public Property Text_Français() As String
            Get
                Return m_Text_Français

            End Get
            Set(ByVal value As String)
                m_Text_Français = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Text_Italien = "Menu"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Text_Italien</c> est une chaîne représentant le texte d'un élément winform en Italien.
        ''' </summary>
        Public Property Text_Italien() As String
            Get
                Return m_Text_Italien

            End Get
            Set(ByVal value As String)
                m_Text_Italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.TextChanged = "Function01"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>TextChanged</c> indique la fonction répondant à cet événement
        ''' </summary>
        Public Property TextChanged() As String
            Get
                Return m_TextChanged

            End Get
            Set(ByVal value As String)
                m_TextChanged = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ToolTip_Allemand = "Information sur le contrôle"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ToolTip_Allemand</c> est une chaîne représentant le tooltip d'un élément winform en Allemand.
        ''' </summary>
        Public Property ToolTip_Allemand() As String
            Get
                Return m_ToolTip_Allemand

            End Get
            Set(ByVal value As String)
                m_ToolTip_Allemand = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ToolTip_Anglais = "Information sur le contrôle"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ToolTip_Anglais</c> est une chaîne représentant le tooltip d'un élément winform en Anglais.
        ''' </summary>
        Public Property ToolTip_Anglais() As String
            Get
                Return m_ToolTip_Anglais

            End Get
            Set(ByVal value As String)
                m_ToolTip_Anglais = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ToolTip_Français = "Information sur le contrôle"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ToolTip_Français</c> est une chaîne représentant le tooltip d'un élément winform en Français.
        ''' </summary>
        Public Property ToolTip_Français() As String
            Get
                Return m_ToolTip_Français

            End Get
            Set(ByVal value As String)
                m_ToolTip_Français = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ToolTip_Espagnole = "Information sur le contrôle"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ToolTip_Espagnole</c> est une chaîne représentant le tooltip d'un élément winform en Espagnole.
        ''' </summary>
        Public Property ToolTip_Espagnole() As String
            Get
                Return m_ToolTip_Espagnole

            End Get
            Set(ByVal value As String)
                m_ToolTip_Espagnole = value
            End Set
        End Property


        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.ToolTip_Italien = "Information sur le contrôle"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>ToolTip_Italien</c> est une chaîne représentant le tooltip d'un élément winform en Italien.
        ''' </summary>
        Public Property ToolTip_Italien() As String
            Get
                Return m_ToolTip_Italien

            End Get
            Set(ByVal value As String)
                m_ToolTip_Italien = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''Type1.Type_Elements_Windows_Forms = "System.Windows.Forms.TextBox"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Type_Elements_Windows_Forms</c> est une chaîne contenant le type de contrôle winform.
        ''' </summary>
        Public Property Type_Elements_Windows_Forms() As String
            Get
                Return m_Type_Elements_Windows_Forms

            End Get
            Set(ByVal value As String)
                m_Type_Elements_Windows_Forms = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim obj1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''obj1.Visibilité = true
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>Visibilité</c> est un booléen qui indique la visibilité d'un élément winform.
        ''' </summary>
        Public Property Visibilité() As Boolean
            Get
                Return m_Visibilité

            End Get
            Set(ByVal value As Boolean)
                m_Visibilité = value
            End Set
        End Property

        ''' <remarks>
        ''' <example>
        ''' <code>
        '''Dim Type1 As New CLASS_IDEALTAKE.CLASS_IDEALTAKE.Objects.Elements_Winform
        '''Type1.WordWrap = "True"
        ''' </code>
        ''' </example>
        ''' </remarks>
        ''' <summary>
        ''' La propriété <c>WordWrap</c> est une chaîne qui  Indique si un contrôle zone de texte 
        ''' Multiligne renvoie les mots au début de la ligne suivante lorsque cela est nécessaire. 
        ''' </summary>
        Public Property WordWrap() As String
            Get
                Return m_WordWrap

            End Get
            Set(ByVal value As String)
                m_WordWrap = value
            End Set
        End Property


        'Méthode de chargement adandonnée
        'Création d'une fonction en décembre 2009
        '
        'Public Function Modification_éléments_formulaires(ByVal sender As System.Object, ByVal Langue As Integer, ByVal Table_des_éléments As Array, ByVal Nombre_ligne As Integer, Optional ByVal ID_CONSTRUCTEUR As Integer = 0) As String
        'Dim index_Table_des_éléments As Integer


        'Valeur de Langue
        ' 0 = Allemand
        ' 1 = Anglais
        ' 2 = Espagnole
        ' 3 = Français
        ' 4 = Italien



        'Test colonne 49 = Type_Elements_Windows_Forms?
        'MsgBox(Table_des_éléments(1, 49).ToString & " - elements_totaux : " & Table_des_éléments.Length.ToString & " -nombre éléments dimension 0 : " & Table_des_éléments.GetLength(0))
        'MsgBox(Table_des_éléments(1, 49).ToString & " -nombre éléments: " & Nombre_ligne.ToString)



        'For index_Table_des_éléments = 1 To Nombre_ligne

        'Les labels
        'If Table_des_éléments(index_Table_des_éléments, 49) = "System.Windows.Forms.Label" Then

        'Les objets Label
        'Dim Label_obj As Object


        'Form_Parent -> position 16
        'Nom_Elements_Windows_Forms -> position 29
        ' -> Position 49

        'on charge le texte du label
        'Select Case Langue
        '                Case Langue = 0 'Allemand

        '                Case Langue = 1 'Anglais

        '                Case Langue = 2 'Espagnole

        '                Case Langue = 3 'Français

        '                Case Langue = 4 'Italien

        '               Case Else


        '            End Select

        '        End If

        '    Next


        '    Return True
        ' End Function

    End Class

End Namespace