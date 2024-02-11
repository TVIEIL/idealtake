Public Class logiciel_dessin
    Dim G As Graphics
    Dim AncX, AncY As Single
    Dim point1_x, point1_y As Single
    Dim point2_x, point2_y As Single
    Dim old_x, old_y As Single
    Dim OutilEnCours As Integer = 0
    Dim pen As New Pen(Color.Black)
    Dim pen_rouge As New Pen(Color.Red)
    Dim clear_pen As New Pen(Color.White)
    Dim br = New SolidBrush(Color.Red)
    Dim Tracé As Boolean
    Dim compteur_point As Integer
    Dim Couleur_Bloc As Color = Color.LightYellow
    Dim Indice_de_ligne_tracée_sur_écran As Integer = 0
    Dim postion_pointeur_maintenant_x, postion_pointeur_maintenant_y As Single
    Dim Indice_de_ligne_tracée_sur_écran_à_supprimer As Integer = 0

    Dim Ligne(1000, 4)


    '' Ce Tableau va stocker les coordonnées des lignes
    '' tracées
    '' Ligne(x,1) --> Point1_x
    '' Ligne(x,2) --> Point1_y
    '' Ligne(x,3) --> Point2_x
    '' Ligne(x,4) --> Point2_y



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Initialisation du Tableau Ligne
        Dim tempo01 As Integer
        For tempo01 = 0 To 999

            '' Ligne(x,1) --> Point1_x
            '' Ligne(x,2) --> Point1_y
            '' Ligne(x,3) --> Point2_x
            '' Ligne(x,4) --> Point2_y

            Ligne(tempo01, 1) = 0
            Ligne(tempo01, 2) = 0
            Ligne(tempo01, 3) = 0
            Ligne(tempo01, 4) = 0
        Next

        'Dim newBitmap As Bitmap = New Bitmap(800, 600)
        'Dim G As Graphics = Graphics.FromImage(newBitmap)
        G = Me.CreateGraphics()
        G.Clear(Color.White)
        Tracé = False


    End Sub

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked
        Select Case e.ClickedItem.Text

            Case "Ligne"
                ToolStripStatusLabel1.Text = "Outil de création de Ligne actif"
                OutilEnCours = 1

            Case "Rectangle"
                ToolStripStatusLabel1.Text = "Outil de création de Rectangle actif"
                OutilEnCours = 2

            Case "Supprimer une Ligne"
                ToolStripStatusLabel1.Text = "Outil de Suppression de ligne actif"
                OutilEnCours = 3


            Case "Supprimer un Rectangle"
                ToolStripStatusLabel1.Text = "Outil de Suppression de Rectangle actif"
                OutilEnCours = 4

        End Select
    End Sub

    Private Sub Form1_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        actualise_affichage_des_lignes()

    End Sub

    Private Function actualise_affichage_des_lignes()

        'Dim Ligne(100, 4)
        '' Ce Tableau va stocker les coordonnées des lignes
        '' tracées
        '' Ligne(x,1) --> Point1_x
        '' Ligne(x,2) --> Point1_y
        '' Ligne(x,3) --> Point2_x
        '' Ligne(x,4) --> Point2_y
        '' Ligne(Indice_de_ligne_tracée_sur_écran, 1) = point1_x
        '' Ligne(Indice_de_ligne_tracée_sur_écran, 2) = point1_y
        '' Ligne(Indice_de_ligne_tracée_sur_écran, 3) = point2_x
        '' Ligne(Indice_de_ligne_tracée_sur_écran, 4) = point2_y

        Dim T As Integer
        For T = 1 To Indice_de_ligne_tracée_sur_écran

            'G.DrawLine(pen, point1_x, point1_y, point2_x, point2_y)
            G.DrawLine(pen, Ligne(T, 1), Ligne(T, 2), Ligne(T, 3), Ligne(T, 4))

        Next


        Return True

    End Function


    'Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e _
    'As System.Windows.Forms.MouseEventArgs) Handles _
    'MyBase.MouseDown
    'Tracé = True
    'AncX = e.X
    'AncY = e.Y
    'End Sub


    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.MouseEventArgs) Handles _
    MyBase.MouseDown

        Tracé = True
        'AncX = e.X
        'AncY = e.Y

        '' Création de Ligne
        If (OutilEnCours = 1) Then

            compteur_point += 1
            'MsgBox(compteur_point)
            If (compteur_point = 1) Then
                point1_x = e.X
                point1_y = e.Y

            End If

            If (compteur_point = 2) Then
                point2_x = e.X
                point2_y = e.Y
                Indice_de_ligne_tracée_sur_écran += 1
                'Dim Ligne(100, 4)
                '' Ce Tableau va stocker les coordonnées des lignes
                '' tracées
                '' Ligne(x,1) --> Point1_x
                '' Ligne(x,2) --> Point1_y
                '' Ligne(x,3) --> Point2_x
                '' Ligne(x,4) --> Point2_y
                Ligne(Indice_de_ligne_tracée_sur_écran, 1) = point1_x
                Ligne(Indice_de_ligne_tracée_sur_écran, 2) = point1_y
                Ligne(Indice_de_ligne_tracée_sur_écran, 3) = point2_x
                Ligne(Indice_de_ligne_tracée_sur_écran, 4) = point2_y

                G.DrawLine(pen, point1_x, point1_y, point2_x, point2_y)

                compteur_point = 0
                point1_x = 0
                point1_y = 0
                point2_x = 0
                point2_y = 0
                Actualise_coordonnées_dans_barre_etat()

            End If
        End If

        'Suppression de Ligne
        If (OutilEnCours = 3) Then

            Dim Indice_temporaire As Integer = 0
            postion_pointeur_maintenant_x = e.X
            postion_pointeur_maintenant_y = e.Y



            Indice_temporaire = Test_Coordonnées_pointeurs_avec_points_depart_arrivé_des_segements()
           
            If (Indice_temporaire > Indice_de_ligne_tracée_sur_écran And Indice_temporaire > 0) Then
                Indice_de_ligne_tracée_sur_écran_à_supprimer = 0
                Indice_temporaire = 0
            Else
                Indice_de_ligne_tracée_sur_écran_à_supprimer = Indice_temporaire
                Indice_temporaire = 0
                GoTo label1
            End If

            'MsgBox(Indice_de_ligne_tracée_sur_écran_à_supprimer & "/" & Indice_de_ligne_tracée_sur_écran)
            'MsgBox(Indice_de_ligne_tracée_sur_écran_à_supprimer)


            Indice_temporaire = Test_Coordonnées_pointeurs_avec_un_segement_Verticale()

            If (Indice_temporaire > Indice_de_ligne_tracée_sur_écran And Indice_temporaire > 0) Then
                Indice_de_ligne_tracée_sur_écran_à_supprimer = 0
                Indice_temporaire = 0
            Else
                Indice_de_ligne_tracée_sur_écran_à_supprimer = Indice_temporaire
                Indice_temporaire = 0
                GoTo label1
            End If

            'MsgBox(Indice_de_ligne_tracée_sur_écran_à_supprimer & "/" & Indice_de_ligne_tracée_sur_écran)
            'MsgBox(Indice_de_ligne_tracée_sur_écran_à_supprimer)

            Indice_temporaire = TestCoordonnéespointeursAvecUnSegementHorizontale()

            If (Indice_temporaire > Indice_de_ligne_tracée_sur_écran And Indice_temporaire > 0) Then
                Indice_de_ligne_tracée_sur_écran_à_supprimer = 0
                Indice_temporaire = 0
            Else
                Indice_de_ligne_tracée_sur_écran_à_supprimer = Indice_temporaire
                Indice_temporaire = 0
                GoTo label1
            End If

            Indice_temporaire = TestCoordonnéespointeursAvecUnSegementOblique()
            If (Indice_temporaire > Indice_de_ligne_tracée_sur_écran And Indice_temporaire > 0) Then
                Indice_de_ligne_tracée_sur_écran_à_supprimer = 0
                Indice_temporaire = 0
            Else
                Indice_de_ligne_tracée_sur_écran_à_supprimer = Indice_temporaire
                Indice_temporaire = 0
                GoTo label1
            End If
label1:
            If (Indice_de_ligne_tracée_sur_écran_à_supprimer = 0) Then
                Exit Sub
            End If

            'MsgBox(Indice_de_ligne_tracée_sur_écran_à_supprimer & "/" & Indice_de_ligne_tracée_sur_écran)
            'MsgBox(Indice_de_ligne_tracée_sur_écran_à_supprimer)

            G.DrawLine(pen_rouge, Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 1), Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 2), Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 3), Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 4))
            G.DrawLine(pen_rouge, Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 1) + 1, Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 2), Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 3) + 1, Ligne(Indice_de_ligne_tracée_sur_écran_à_supprimer, 4))

            Dim Question_Supprimer_Ligne As Windows.Forms.DialogResult
            'MsgBox(Indice_de_ligne_tracée_sur_écran_à_supprimer)
            Question_Supprimer_Ligne = MessageBox.Show("Voulez vous supprimer cette Ligne?", "Supprimer une Ligne", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)

            If (Question_Supprimer_Ligne = Windows.Forms.DialogResult.Yes) Then
                SupprimerUNeLIgneDeLaTableLigne(Indice_de_ligne_tracée_sur_écran_à_supprimer)
            End If

            G.Clear(Color.White)
            actualise_affichage_des_lignes()
            'Réinitialisation de l'indice
            Indice_de_ligne_tracée_sur_écran_à_supprimer = 0


        End If


    End Sub

    Public Function SupprimerUNeLIgneDeLaTableLigne(ByVal Position As Integer)
        Dim Indice_Balayage As Integer

        For Indice_Balayage = Indice_de_ligne_tracée_sur_écran_à_supprimer To (Indice_de_ligne_tracée_sur_écran + 1)
            Ligne(Indice_Balayage, 1) = Ligne(Indice_Balayage + 1, 1)
            Ligne(Indice_Balayage, 2) = Ligne(Indice_Balayage + 1, 2)
            Ligne(Indice_Balayage, 3) = Ligne(Indice_Balayage + 1, 3)
            Ligne(Indice_Balayage, 4) = Ligne(Indice_Balayage + 1, 4)
        Next

        Return True
    End Function


    Public Function Test_Coordonnées_pointeurs_avec_points_depart_arrivé_des_segements()
        Dim T As Integer
        'MsgBox(Indice_de_ligne_tracée_sur_écran)
        For T = 1 To Indice_de_ligne_tracée_sur_écran + 1

            '3ème version
            ''Test sur les abscisses des 2 points départ et fin des segments
            If (((postion_pointeur_maintenant_x <= Ligne(T, 1) + 2) And _
                 (postion_pointeur_maintenant_x >= Ligne(T, 1) - 2)) Or _
                ((postion_pointeur_maintenant_x <= Ligne(T, 3) + 2) And _
                 (postion_pointeur_maintenant_x >= Ligne(T, 3) - 2))) Then

                ''Test sur les ordonnées des 2 points départ et fin des segments
                If (((postion_pointeur_maintenant_y <= Ligne(T, 2) + 2) And _
                     (postion_pointeur_maintenant_y >= Ligne(T, 2) - 2)) Or _
                    ((postion_pointeur_maintenant_y <= Ligne(T, 4) + 2) And _
                     (postion_pointeur_maintenant_y >= Ligne(T, 4) - 2))) Then

                    'Indice_de_ligne_tracée_sur_écran_à_supprimer = T
                    'G.DrawLine(pen, point1_x, point1_y, point2_x, point2_y)
                    'G.DrawLine(pen_rouge, Ligne(T, 1), Ligne(T, 2), Ligne(T, 3), Ligne(T, 4))
                    Exit For

                End If
            End If

        Next

        Return T
    End Function

    Public Function TestCoordonnéespointeursAvecUnSegementOblique()
        Dim T As Integer
        Dim U As Integer
        Dim V As Integer = 0
        Dim Xpt1, Ypt1 As Single
        Dim Xpt2, Ypt2 As Single
        Dim Calcul_de_y As Single
        Dim a As Single
        Dim b As Single
        Dim Coordonnées_points_1_ligne(10000, 2) As Single

        For T = 1 To Indice_de_ligne_tracée_sur_écran + 1

            If ((Ligne(T, 1) <> Ligne(T, 3)) And (Ligne(T, 2) <> Ligne(T, 4))) Then

                Select Case Ligne(T, 1) - Ligne(T, 3)

                    Case Is > 0
                        Xpt1 = Ligne(T, 3)
                        Ypt1 = Ligne(T, 4)
                        Xpt2 = Ligne(T, 1)
                        Ypt2 = Ligne(T, 2)

                    Case Is < 0
                        Xpt1 = Ligne(T, 1)
                        Ypt1 = Ligne(T, 2)
                        Xpt2 = Ligne(T, 3)
                        Ypt2 = Ligne(T, 4)

                End Select

                'la ligne est une oblique
                'y = ax + b
                '1/ calcul de a
                a = (Ypt2 - Ypt1) / (Xpt2 - Xpt1)

                '2/ calcul de b
                ' b = y - ax
                'Application numérique avec PT2
                b = Ypt2 - (a * Xpt2)

                '3/ calcul dans la table 'Coordonnées_points_1_ligne'
                'des points du segment de la valeur de Xpt1 à
                'la valeur de Xpt2

                For U = Xpt1 To Xpt2 Step 1
                    V += 1
                    Calcul_de_y = a * U + b
                    Coordonnées_points_1_ligne(V, 1) = U
                    Coordonnées_points_1_ligne(V, 2) = Fix(Calcul_de_y)
                    'G.DrawLine(Pens.Blue, U, Calcul_de_y, U + 1, Calcul_de_y + 1)
                Next

                For U = 1 To V
                    'If (a < -1.8 Or a > 1.8) Then

                    ''Test sur les abscisses des points des segments
                    'If ((postion_pointeur_maintenant_x <= Coordonnées_points_1_ligne(U, 1) + 2) Or _
                    '(postion_pointeur_maintenant_x <= Coordonnées_points_1_ligne(U, 1) - 2)) Then
                    ''Test sur les ordonnées des points des segments
                    'If ((postion_pointeur_maintenant_y <= Coordonnées_points_1_ligne(U, 2) + 2) Or _
                    '(postion_pointeur_maintenant_y <= Coordonnées_points_1_ligne(U, 2) - 2)) Then
                    'GoTo LABEL2
                    'End If
                    'End If

                    'Else

                    ''Test sur les abscisses des points des segments
                    If (postion_pointeur_maintenant_x = Coordonnées_points_1_ligne(U, 1)) Then
                        ''Test sur les ordonnées des points des segments
                        If (postion_pointeur_maintenant_y = Coordonnées_points_1_ligne(U, 2)) Then
                            GoTo LABEL2
                        End If
                    End If

                    'End If


                Next  'Next U

            End If ' les ordonnées de PT1 et PT2 sont <>, les absisses de PT1 et PT2 sont <>,
        Next ' Next T

LABEL2:
        Return T

    End Function

    Public Function Test_Coordonnées_pointeurs_avec_un_segement_Verticale()
        Dim T As Integer
        'MsgBox(Indice_de_ligne_tracée_sur_écran)
        For T = 1 To Indice_de_ligne_tracée_sur_écran + 1

            If (Ligne(T, 1) = Ligne(T, 3)) Then 'forcement xPt1 = Xpt2

                'If (((postion_pointeur_maintenant_x <= Ligne(T, 1) + 2) And _
                '(postion_pointeur_maintenant_x >= Ligne(T, 1) - 2)) Or _
                '((postion_pointeur_maintenant_x <= Ligne(T, 3) + 2) And _
                '(postion_pointeur_maintenant_x >= Ligne(T, 3) - 2))) Then

                ''Test sur les abscisses des 2 points départ et fin des segments
                If ((postion_pointeur_maintenant_x <= Ligne(T, 1) + 2) And _
                     (postion_pointeur_maintenant_x >= Ligne(T, 1) - 2)) Then

                    Dim yPT1 As Integer = 0
                    Dim yPT2 As Integer = 0

                    If (Ligne(T, 2) > Ligne(T, 4)) Then
                        yPT1 = Ligne(T, 4)
                        yPT2 = Ligne(T, 2)
                    Else
                        yPT1 = Ligne(T, 2)
                        yPT2 = Ligne(T, 4)
                    End If

                    ''Test sur les ordonnées comprisent entre yPT1 < yCurseur < yPT2
                    If ((postion_pointeur_maintenant_y > yPT1) And _
                         (postion_pointeur_maintenant_y < yPT2)) Then
                        Exit For
                    End If

                    ''Ou Test sur les ordonnées comprisent entre yPT1 < yCurseur < yPT2
                    'If ((postion_pointeur_maintenant_y < Ligne(T, 2)) And _
                    '    (postion_pointeur_maintenant_y > Ligne(T, 4))) Then
                    ' Exit For
                    ' End If

                End If
            End If

        Next

        Return T
    End Function

    Public Function TestCoordonnéespointeursAvecUnSegementHorizontale()
        Dim T As Integer
        'MsgBox(Indice_de_ligne_tracée_sur_écran)
        For T = 1 To Indice_de_ligne_tracée_sur_écran + 1

            '                If (((postion_pointeur_maintenant_y <= Ligne(T, 2) + 2) And _
            '(postion_pointeur_maintenant_y >= Ligne(T, 2) - 2)) Or _
            '((postion_pointeur_maintenant_y <= Ligne(T, 4) + 2) And _
            '(postion_pointeur_maintenant_y >= Ligne(T, 4) - 2))) Then

            If (Ligne(T, 2) = Ligne(T, 4)) Then 'forcement yPt1 = ypt2
                ''Test sur les ordonnées des 2 points départ et fin des segments
                If ((postion_pointeur_maintenant_y <= Ligne(T, 2) + 2) And _
                     (postion_pointeur_maintenant_y >= Ligne(T, 2) - 2)) Then


                    Dim xPT1 As Integer = 0
                    Dim xPT2 As Integer = 0

                    If (Ligne(T, 1) > Ligne(T, 3)) Then
                        xPT1 = Ligne(T, 3)
                        xPT2 = Ligne(T, 1)
                    Else
                        xPT1 = Ligne(T, 1)
                        xPT2 = Ligne(T, 3)
                    End If

                    ''Test sur les ordonnées comprisent entre xPT1 < yCurseur < xPT2
                    If ((postion_pointeur_maintenant_x > xPT1) And _
                         (postion_pointeur_maintenant_x < xPT2)) Then
                        Exit For
                    End If



                End If
            End If

        Next

        Return T
    End Function

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.MouseEventArgs) Handles _
    MyBase.MouseMove

        'G.DrawLine(pen, AncX, AncY, e.X, e.Y)
        AncX = e.X
        AncY = e.Y
        If (OutilEnCours = 1) Then
            If (point1_x <> 0 And point1_y <> 0) Then
                G.DrawLine(clear_pen, point1_x, point1_y, old_x, old_y)
                G.DrawLine(pen, point1_x, point1_y, e.X, e.Y)
                old_x = e.X
                old_y = e.Y
            End If

            Actualise_coordonnées_dans_barre_etat()
        End If

    End Sub


    Private Function Actualise_coordonnées_dans_barre_etat()

        If (OutilEnCours <> 0) Then
            If (compteur_point = 0) Then
                ToolStripStatusLabel1.Text = "Coordonnées du point de Départ de la ligne X : " & AncX.ToString & " ; Y : " & AncY.ToString
            End If
            If (compteur_point = 1) Then
                ToolStripStatusLabel1.Text = "Coordonnées du point de Fin de la ligne X : " & AncX.ToString & " ; Y : " & AncY.ToString & _
                " - Coordonnées du point de Départ de la ligne(" & point1_x.ToString & "," & point1_y.ToString & ")"
            End If

            'ToolStripStatusLabel1.Text = " X : " & AncX.ToString & " ; Y :" & AncY.ToString & _
            '" P1(" & point1_x.ToString & "," & point1_y.ToString & ")" & _
            '" P2(" & point2_x.ToString & "," & point2_y.ToString & ")"
        End If

        Return True
    End Function

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e _
    As System.Windows.Forms.MouseEventArgs) Handles _
    MyBase.MouseUp
        Tracé = False
        Select Case OutilEnCours
            Case 1
                'G.DrawLine(pen, AncX, AncY, e.X, e.Y)

                'If (compteur_point = 2) Then
                'G.DrawLine(pen, point1_x, point1_y, point2_x, point2_y)
                'If compteur_point > 1 Then
                'compteur_point = 0
                'End If
                'End If

            Case 2
                G.DrawRectangle(pen, AncX, AncY, e.X - AncX, e.Y - AncY)
                G.FillRectangle(br, AncX, AncY, e.X - AncX, e.Y - AncY)
            Case 3
                G.DrawEllipse(pen, AncX, AncY, e.X - AncX, e.Y - AncY)
                G.FillEllipse(br, AncX, AncY, e.X - AncX, e.Y - AncY)
        End Select

    End Sub



End Class
