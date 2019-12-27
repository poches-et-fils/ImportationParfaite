Option Compare Text

Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element     As Variant
    On Error GoTo IsInArrayError:        'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = TRUE
            Exit Function
        End If
    Next element
    Exit Function
    IsInArrayError:
    On Error GoTo 0
    IsInArray = FALSE
End Function

Sub ToutesPochesToutesBases()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = FALSE
    
    Set ImportSource = ActiveWorkbook
    NomUser = InputBox("Quel est ton nom de user")
    NomFichier = InputBox("Comment veux-tu appeler le fichier?")
    Workbooks.Add
    
    ActiveWorkbook.SaveAs Filename:="C:\Users\" & NomUser & "\pochesetfils.com\PUBLIC - Documents\004 Web\01Outils\01ImportationProduit\" & Format(Date, "yyyy-mm-dd") & "_" & NomFichier & ".csv"
    Dim IMPORT      As Workbook
    Set IMPORT = ActiveWorkbook
    Dim couleuradecliner As New Collection
    Dim couleursinterdites As New Collection
    
    For Each m In Array(8)        'numéro des pages de chandails poches
        'Vneck = 1
        'Bébé = 2
        'tshirt femme boyfriend = 3
        'tshirt homme = 4
        'tshirt ado = 5
        'longsleeves homme = 6
        'longsleeves femme = 7
        'cami femme = 8
        'Cami homme = 9
        'robes = 10
        
        ImportSource.Activate
        If m = 1 Then
            
            'design à décliner
            Page = "Femmes Vnecks"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "femme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "femme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "femme"
            tagsCollections = ", collection:Femme, collection:Femme - Tops, collection:Femme - V-neck à poche, collection:All"
            googlegender = "femme"        'voir onglet google merchant fr
            prix = "36.99"
            codetee = "203"
            typeprod = "V-neck à poche"        ' modle seoo
            typebarre = "vneck_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 7
            couleursoriginales = Array(3, 4, 5, 6, 7)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 13        ' dernire taillee
            seo = "T-shirt à poche à col en v pour femme 50% COTON 50% POLYESTER offert en noir, blanc, charbon, bleu marin et gris. On peut se tutoyer? Avec une belle poche de même, tu vas être funky, confortable et juste bien dans ta peau. Porte fièrement ton v-neck Poches & Fils, il est conu et empoché par des humains à Montréal, Québec, Canada."
            
        ElseIf m = 2 Then
            Page = "Bébé TShirt"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "bébé"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "bebe"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "bébé"
            tagsCollections = ", collection:Enfant, collection:Bébé - T-shirt à poche, collection:All, collection:Enfant et bébé"
            googlegender = "unisexe"        'voir onglet google merchant fr
            prix = "26.99"
            codetee = "101"
            typeprod = "T-shirt à poche"        ' modle seoo
            typebarre = "t-shirt_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 5
            couleursoriginales = Array(3, 4, 5)
            googleage = "bébés"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 2        'premire taille à déclinerr
            b = 8        ' dernire taillee
            seo = "T-shirt à poche à col rond pour bébé 50% COTON 50% POLYESTER offert en noir, blanc et charbon. On a vérifié et il        'y a rien de plus cute qu'un bébé avec un t-shirt à poche Poches & Fils funky et confortable. Ce t-shirt conu et empoché par des humains à Montréal, Québec, Canada vous rendra tout simplement gaga #aww."
            
        ElseIf m = 3 Then
            Page = "Femme TShirt"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "femme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "femme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "femme"
            tagsCollections = ", collection:Boxing Day - Femme, collection:Boxing Day"
            googlegender = "femme"        'voir onglet google merchant fr
            prix = "20.00"
            codetee = "217"
            typeprod = "T-shirt à poche"        ' modle seoo
            typebarre = "t-shirt_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 6
            couleursoriginales = Array(4)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 13        ' dernire taillee
            seo = "T-shirt à poche à col rond pour femme coupe Boyfriend 50% COTON 50% POLYESTER offert en noir, charbon et gris. On peut se tutoyer? Avec une belle poche de même, tu vas être funky, assumée et confortable, surtout. Porte fièrement ton chandail Poches & Fils, il est conçu et empoché par des humains à Montréal, Québec, Canada."
            
        ElseIf m = 4 Then
            Page = "Homme TShirt"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "homme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "homme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "homme"
            tagsCollections = ", collection:Homme, collection:Homme - Tops, collection:Homme - T-shirt à poche, collection:All"
            googlegender = "homme"        'voir onglet google merchant fr
            prix = "20.00"
            codetee = "301"
            typeprod = "T-shirt à poche"        ' modle seoo
            typebarre = "t-shirt_a_poche_cer"        'modele web
            couleurdebut = 3
            couleurfin = 7
            couleursoriginales = Array(10)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 14        ' dernire taillee
            seo = "T-shirt à poche à col rond pour homme 50% COTON 50% POLYESTER offert en noir, blanc, charbon et gris. Porter sa poche sur son coeur, a met du coeur sur sa poche? Porte fièrement ton chandail Poches & Fils, il est baptisé        'un nom drôle que seuls les vrais connaissent et est conçu et empoché par des humains à Montréal."
            
        ElseIf m = 5 Then
            Page = "Enfant TShirt"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "enfant"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "enfant"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "enfant"
            tagsCollections = ", collection:Enfant, collection:Enfant - T-shirt à poche, collection:All, collection:Enfant et bébé"
            googlegender = "unisexe"        'voir onglet google merchant fr
            prix = "30.99"
            codetee = "501"
            typeprod = "T-shirt à poche"        ' modle seoo
            typebarre = "t-shirt_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 5
            couleursoriginales = Array(3, 4, 5)
            googleage = "enfants"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 10        'premire taille à déclinerr
            b = 12        ' dernire taillee
            seo = "T-shirt à poche à col rond pour enfant 50% COTON 50% POLYESTER offert en noir, blanc et charbon. Votre enfant cherche à exprimer sa personnalité ou à matcher funky avec Maman et Papa (tout en étant cool)? #yolesjeunes!        'est facile avec un chandail Poches & Fils conçu et empoché par des humains à Montréal, Québec, Canada"
            
        ElseIf m = 6 Then
            Page = "Homme Longsleeves"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "homme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "homme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "homme"
            tagsCollections = ", collection:Homme, collection:Homme - Tops, collection:Homme - T-shirt à manches longues à poche, collection:All"
            googlegender = "homme"        'voir onglet google merchant fr
            prix = "51.99"
            codetee = "312"
            typeprod = "T-shirt à manches longues à poche"        ' modle seoo
            typebarre = "t-shirt_a_manches_longues_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 9
            couleursoriginales = Array(3, 5, 6, 8, 9)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 14        ' dernire taillee
            seo = "T-shirt à manches longues à poche unisexe 93% COTON 7% ÄLASTHANE offert en noir, gris, charbon, bleu royal et sarcelle. On peut se tutoyer? Avec cet unique t-shirt à manches longues conçu et empoché par des humains à Montréal, Québec, Canada, tu vas pouvoir montrer ton côté givré-saucé, tout en étant aussi confortable que dans un Snuggy."
            
        ElseIf m = 7 Then
            Page = "Femme Longsleeves"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "femme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "femme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "femme"
            tagsCollections = ", collection:Femme, collection:Femme - Tops, collection:Femme - T-shirt à manches longues à poche, collection:All"
            googlegender = "femme"        'voir onglet google merchant fr
            prix = "51.99"
            codetee = "212"
            typeprod = "T-shirt à manches longues à poche"        ' modle seoo
            typebarre = "t-shirt_a_manches_longues_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 9
            couleursoriginales = Array(3, 5, 6, 8, 9)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 14        ' dernire taillee
            seo = "T-shirt décontracté à manches longues à poche unisexe 93% COTON 7% ÄLASTHANE offert en noir, gris, charbon, bleu royal et sarcelle. On peut se tutoyer? Avec cet unique t-shirt à manches longues conçu et empoché par des humains à Montréal, Québec, Canada, tu vas pouvoir montrer ton côté givré-saucé, tout en étant aussi confortable que dans un Snuggy."
            
        ElseIf m = 8 Then
            Page = "Femmes Camisoles"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "femme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "femme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "femme"
            tagsCollections = ", collection:Boxing Day - Femme, collection:Boxing Day"
            googlegender = "femme"        'voir onglet google merchant fr
            prix = "15.00"
            codetee = "202"
            typeprod = "Camisole à poche"        ' modle seoo
            typebarre = "camisole_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 5
            couleursoriginales = Array(3, 4, 5)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 13        ' dernire taillee
            seo = "Camisole à poche pour femme 50% COTON 50% POLYESTER offert en noir, blanc et charbon. Que tu aimes avoir les aisselles au vent et/ou montrer ta poche au gym, avec cette camisole, tu vas être funky et confortable. Porte fièrement ton chandail Poches & Fils, il est conçu et empoché par des humains à Montréal, Québec, Canada."
            
        ElseIf m = 9 Then
            Page = "Hommes Camisoles"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "homme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "homme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "homme"
            tagsCollections = ", collection:Boxing Day - Homme, collection:Boxing Day"
            googlegender = "homme"        'voir onglet google merchant fr
            prix = "15.00"
            codetee = "302"
            typeprod = "Camisole à poche"        ' modle seoo
            typebarre = "camisole_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 5
            couleursoriginales = Array(3, 4)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 14        ' dernire taillee
            seo = "Camisole à poche pour homme 50% COTON 50% POLYESTER offert en noir, blanc et charbon. Que tu aimes avoir les aisselles au vent et/ou montrer ta poche au gym, avec cette camisole, tu vas être funky et confortable. Porte firement ton chandail Poches & Fils, il est conçu et empoché par des humains à Montréal, Québec, Canada."
            
        ElseIf m = 10 Then
            Page = "Femmes Robes"        ' ATTENTION LONGSLEEVES COULEURS
            gender = "femme"        ' bébé, enfant, homme, femme, unisexe
            genderbarre = "femme"        '  bebe, enfant, homme, femme, unisexe
            tagsGenre = "femme"
            tagsCollections = ", collection:Femme, collection:Femme - Tops, collection:Femme - Robe à poche, collection:All"
            googlegender = "femme"        'voir onglet google merchant fr
            prix = "56.99"
            codetee = "210"
            typeprod = "Robe à poche"        ' modle seoo
            typebarre = "robe_a_poche"        'modele web
            couleurdebut = 3
            couleurfin = 9
            couleursoriginales = Array(3, 5, 6, 8, 9)
            googleage = "adultes"        'voir onglet google merchant fr / ATTENTION BÄBÄS ET PETIT KID
            a = 9        'premire taille à déclinerr
            b = 13        ' dernire taillee
            seo = "Robe à poche semi-ajustée (qui n’est pas un t-shirt) 93% COTON 7% ÉLASTHANE offert en noir, gris et charbon. Son tissu soyeux - touche-le - est assez épais pour cacher ce qu’il faut, sans être trop chaud. Cette robe est comme toi: polyvalente, belle, confortable dans toutes les situations et prête pour l’été!"
            
        End If
        
        ' Toutes les poches, toutes les bases
        Sheets(Page).Activate
        Range(Sheets(Page).Cells(2, 1), Sheets(Page).Cells(10000, 1000)).Select
        Selection.Clear
        n = 2
        'Integration des tabs Unepoche toutes les bases => For r = 1 To Sheets("UnePocheToutesLesBases").Cells(1, 1).CurrentRegion.Rows.Count
        For I = 1 To Sheets("poches à décliner").Cells(1, 1).CurrentRegion.Rows.Count
            Set couleursadecliner = New Collection
            Set couleursinterdites = New Collection
            
            If m = 2 Or m = 5 Then
                
                'poches à ne pas décliner pour enfants
                
                If IsInArray(Sheets("poches à décliner").Cells(I, 1), Array("0006", "0008", "0093", "0141", "C271", "C264", "C282", "C268")) Then I = I + 1
            End If
            
            'à modifier
            For k = 3 To 20
                If CStr(Sheets("couleurs prioritaires").Cells(1, k)) = CStr(codetee) Then col = k
                If CStr(Sheets("couleurs prioritaires").Cells(1, k)) = "Do Not do" Then col_interdite = k
                
            Next k
            
            For k = 2 To Sheets("couleurs prioritaires").Cells(1, 1).CurrentRegion.Rows.Count
                
                If CStr(Sheets("couleurs prioritaires").Cells(k, 1)) = CStr(Sheets("poches à décliner").Cells(I, 1)) Then
                    couleurfav = Sheets("couleurs prioritaires").Cells(k, col)
                    lignepoche = k
                    
                    Exit For
                End If
                
            Next k
            
            k = col_interdite
            
            While Sheets("couleurs prioritaires").Cells(lignepoche, k) <> ""
                couleursinterdites.Add (Sheets("couleurs prioritaires").Cells(lignepoche, k))
                k = k + 1
            Wend
            
            'MsgBox (couleursinterdites(1) & couleursinterdites(2))
            
            couleurstemp = couleursoriginales
            
            If IsInArray(couleurfav, couleurstemp) And couleurfav <> 3 Then
                pos = Application.Match(couleurfav, couleurstemp, False) - 1
                temp = couleurstemp(0)
                couleurstemp(0) = couleurfav
                couleurstemp(pos) = temp
            End If
            
            For Each k In couleurstemp
                couleursadecliner.Add k
                For Each Z In couleursinterdites
                    If couleursadecliner(couleursadecliner.Count) = Z Then
                        couleursadecliner.Remove (couleursadecliner.Count)
                    End If
                Next Z
            Next k
            
            nbcouleur = couleursadecliner.Count
            'MsgBox (couleursadecliner(1) & couleursadecliner(2))
            'Dim tempo(10) As Integer
            '
            'For k = 1 To nbcouleur
            'If k = 1 Then
            '    tempo(1) = couleurfav
            '    ElseIf k > pos Then
            '        tempo(k) = couleursadecliner(k)
            '    ElseIf k <= pos Then
            '    tempo(k) = couleursadecliner(k - 1)
            '
            'End If
            '
            'Next k
            
            'MsgBox (couleursadecliner(1))
            
            'End If
            
            debut = 1
            For Each p In couleursadecliner
                'c = couleurdebut + p - 1
                'fav = 0
                '
                '  If (codetee = "312" Or codetee = "210") And (couleurfav = 4 Or couleurfav = 5 Or couleurfav = 6) Then
                '
                'ElseIf couleurfav <= couleurfin And couleurfav > couleurdebut Then
                '
                '
                '
                'fav = 1
                '    If p = 1 Then
                '    c = couleurfav
                '    End If
                '
                '    If p > 1 Then
                '    c = couleurdebut + p - 2
                '    If (codetee = "312" Or codetee = "210") And (couleurfav = 4 Or couleurfav = 5 Or couleurfav = 6) Then c = c + 1
                '
                '    If c >= couleurfav Then c = c + 1
                '    End If
                '
                'End If
                '
                '
                '
                'If codetee = "312" And c = 4 Or codetee = "312" And c = 7 Then
                'p = p + 1
                'c = c + 1 ' pas de longsleeves blanc et marine
                'End If
                'If codetee = "210" And c = 4 Or codetee = "210" And c = 7 Then
                'p = p + 1
                'c = c + 1 ' pas de robes blanc et marine
                'End If
                'If codetee = "217" And c = 4 Then
                'p = p + 1
                'c = c + 1 ' pas de boyfriend tee blanc
                'End If
                
                couleur = Sheets("FR").Cells(p, 9)
                couleurnom = Sheets("FR").Cells(p, 10)
                codecouleur = Sheets("FR").Cells(p, 8)
                
                For j = a To b        'tailles à décliner
                    codetaille = Sheets("FR").Cells(j, 19)
                    
                    If j = a And debut = 1 Then
                        Sheets(Page).Cells(n, 1) = typebarre & "_" & genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2)
                        Sheets(Page).Cells(n, 2) = Sheets("poches à décliner").Cells(I, 3)
                        Sheets(Page).Cells(n, 3) = Sheets("poches à décliner").Cells(I, 4)
                        Sheets(Page).Cells(n, 4) = Replace(Sheets("poches à décliner").Cells(I, 5),        '", "_")
                        
                        Sheets(Page).Cells(n, 5) = typeprod        ' type de produit
                        
                        Sheets(Page).Cells(n, 6) = "gender:" & tagsGenre & tagsCollections & ", collection:" & Sheets("poches à décliner").Cells(I, 6) & ", collection:" & Sheets("poches à décliner").Cells(I, 6)
                        If gender = "homme" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & "-Homme"
                        If gender = "femme" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & "-Femme"
                        If gender = "enfant" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & "-Enfant"
                        If gender = "bébé" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & "-Bebe"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & ", B2S19"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & ", collection:Nouveautés FW19"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & ", collection:Nouveauté"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And gender = "homme" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & ", collection:Homme - Nouvelles poches FW19"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And gender = "femme" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & ", collection:Femme - Nouvelles poches FW19"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And gender = "enfant" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & ", collection:Enfant - Nouvelles poches FW19"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And gender = "bébé" Then Sheets(Page).Cells(n, 6) = Sheets(Page).Cells(n, 6) & ", collection:Bébé - Nouvelles poches FW19"
                        
                        Sheets(Page).Cells(n, 7) =        'true"
                        Sheets(Page).Cells(n, 8) = "Size"
                        Sheets(Page).Cells(n, 9) = Sheets("FR").Cells(j, 18)        'taille
                        Sheets(Page).Cells(n, 10) = "Color"
                        Sheets(Page).Cells(n, 11) = couleurnom
                        Sheets(Page).Cells(n, 14) = codetee & codecouleur & Sheets("poches à décliner").Cells(I, 1) & "-" & codetaille        ' changer le début seulement
                        Sheets(Page).Cells(n, 16) = "shopify"
                        Sheets(Page).Cells(n, 17) = 100000
                        Sheets(Page).Cells(n, 18) = "deny"
                        Sheets(Page).Cells(n, 19) = "manual"
                        Sheets(Page).Cells(n, 20) = prix
                        
                        Sheets(Page).Cells(n, 22) =        'true"
                        Sheets(Page).Cells(n, 23) =        'true"
                        Sheets(Page).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & typebarre & "_" & couleur & "_" & genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                        Sheets(Page).Cells(n, 26) = 1
                        Sheets(Page).Cells(n, 27) = typeprod & " pour " & gender & " avec poche " & Sheets("poches à décliner").Cells(I, 3) & " à motif de " & Sheets("poches à décliner").Cells(I, 7)
                        Sheets(Page).Cells(n, 28) =        'false"
                        Sheets(Page).Cells(n, 29) = Sheets("poches à décliner").Cells(I, 3) & " - " & typeprod & " " & couleurnom & " " & gender & " | Poches & Fils"
                        Sheets(Page).Cells(n, 30) = seo
                        Sheets(Page).Cells(n, 31) =        'Google merchant FR'!R2C1:R50C2,2,0)"
                        Sheets(Page).Cells(n, 32) = googlegender
                        Sheets(Page).Cells(n, 33) = googleage
                        Sheets(Page).Cells(n, 34) = Left(Sheets(Page).Cells(n, 14), 9)
                        Sheets(Page).Cells(n, 35) = Sheets(Page).Cells(n, 5)
                        Sheets(Page).Cells(n, 37) = "neuf"
                        'à modifier
                        Sheets(Page).Cells(n, 44) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & typebarre & "_" & couleur & "_" & genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                        
                        n = n + 1
                        
                    Else
                        'à modifier
                        Sheets(Page).Cells(n, 1) = typebarre & "_" & genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2)
                        Sheets(Page).Cells(n, 9) = Sheets("FR").Cells(j, 18)        'taille
                        Sheets(Page).Cells(n, 11) = couleurnom
                        'à modifier
                        Sheets(Page).Cells(n, 14) = codetee & codecouleur & Sheets("poches à décliner").Cells(I, 1) & "-" & codetaille        ' changer le début seulement
                        Sheets(Page).Cells(n, 16) = "shopify"
                        Sheets(Page).Cells(n, 17) = 100000
                        Sheets(Page).Cells(n, 18) = "deny"
                        Sheets(Page).Cells(n, 19) = "manual"
                        Sheets(Page).Cells(n, 20) = prix
                        
                        Sheets(Page).Cells(n, 22) =        'true"
                        Sheets(Page).Cells(n, 23) =        'true"
                        
                        '---Enlever les poches non-convenable pour les enfants
                        
                        '---Inventaire des bases---
                        'Camisoles Homme
                        
                        'Longsleeve Homme et Femme
                        
                        '***********************************************************************
                        If Left(Sheets(Page).Cells(n, 1), 17) = "t-shirt_a_manches" Then
                            If Sheets(Page).Cells(n, 11) = "bleu-royal" Then
                                If Sheets(Page).Cells(n, 9) = "S" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                        End If
                        '*************************************************************************
                        
                        If Left(Sheets(Page).Cells(n, 1), 5) = "vneck" Then
                            
                            If Sheets(Page).Cells(n, 11) = "blanc" Then
                                If Sheets(Page).Cells(n, 9) = "XL" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "charbon" Then
                                If Sheets(Page).Cells(n, 9) = "XL" Or Sheets(Page).Cells(n, 9) = "L" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "gris" Then
                                If Sheets(Page).Cells(n, 9) = "L" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "bleu-marine" Then
                                If Sheets(Page).Cells(n, 9) = "XL" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                        End If
                        
                        '*********************************************************************************
                        If Left(Sheets(Page).Cells(n, 1), 21) = "t-shirt_a_poche_ceris" Then
                            If Sheets(Page).Cells(n, 11) = "blanc" Then
                                If Sheets(Page).Cells(n, 9) = "XXL" Or Sheets(Page).Cells(n, 9) = "XS" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "charbon" Then
                                If Sheets(Page).Cells(n, 9) = "XXL" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "yogourt-aux-cerises" Then
                                If Sheets(Page).Cells(n, 9) = "S" Or Sheets(Page).Cells(n, 9) = "L" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                        End If
                        
                        '************************************************************************************
                        
                        If Left(Sheets(Page).Cells(n, 1), 4) = "robe" Then
                            If Sheets(Page).Cells(n, 11) = "noir" Then
                                If Sheets(Page).Cells(n, 9) = "XXL" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "charbon" Then
                                If Sheets(Page).Cells(n, 9) = "XXL" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "bleu-royal" Then
                                If Sheets(Page).Cells(n, 9) = "XS" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "sarcelle" Then
                                If Sheets(Page).Cells(n, 9) = "XS" Or Sheets(Page).Cells(n, 9) = "S" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                        End If
                        '*********************************************************************************
                        '
                        
                        If Left(Sheets(Page).Cells(n, 1), 21) = "t-shirt_a_poche_femme" Then
                            If Sheets(Page).Cells(n, 11) = "noir" Then
                                If Sheets(Page).Cells(n, 9) <> "XS" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                                
                            End If
                            
                            If Sheets(Page).Cells(n, 11) = "gris" Then
                                If Sheets(Page).Cells(n, 9) = "S" Or Sheets(Page).Cells(n, 9) = "M" Or Sheets(Page).Cells(n, 9) = "XL" Then
                                    Sheets(Page).Cells(n, 17) = 0
                                End If
                            End If
                            
                        End If
                        
                        '*********************************************************************************
                        
                        'On ajoute les bases qui ont une image de dos disponible
                        If Left(Sheets(Page).Cells(n, 14), 3) = "312" Or Left(Sheets(Page).Cells(n, 14), 3) = "212" Or Left(Sheets(Page).Cells(n, 14), 3) = "203" Or Left(Sheets(Page).Cells(n, 14), 3) = "217" Or Left(Sheets(Page).Cells(n, 14), 3) = "301" Or Left(Sheets(Page).Cells(n, 14), 3) = "210" Then
                            If Sheets(Page).Cells(n - 1, 26) = 1 Then
                                Sheets(Page).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & typebarre & "_" & couleur & "_" & genderbarre & "-back.jpg"
                                Sheets(Page).Cells(n, 26) = 2
                                
                                Sheets(Page).Cells(n, 27) = "color:" & Sheets("FR").Cells(p, 10)
                                
                            ElseIf Sheets(Page).Cells(n - 1, 26) > 1 And Sheets(Page).Cells(n - 1, 26) <= nbcouleur Then
                                Sheets(Page).Cells(n, 26) = Sheets(Page).Cells(n - 1, 26) + 1
                                Sheets(Page).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & typebarre & "_" & Sheets("FR").Cells(couleursadecliner(Sheets(Page).Cells(n - 1, 26)), 9) & "_" & genderbarre & "-back.jpg"
                                Sheets(Page).Cells(n, 27) = "color:" & Sheets("FR").Cells(couleursadecliner(Sheets(Page).Cells(n - 1, 26)), 10)
                            Else
                            End If
                        End If
                        
                        Sheets(Page).Cells(n, 30) = seo
                        Sheets(Page).Cells(n, 31) =        'Google merchant FR'!R2C1:R50C2,2,0)"
                        Sheets(Page).Cells(n, 32) = googlegender
                        Sheets(Page).Cells(n, 33) = googleage
                        If Sheets(Page).Cells(n, 9) = "18m" Or Sheets(Page).Cells(n, 9) = "2t" Or Sheets(Page).Cells(n, 9) = "3t" Or Sheets(Page).Cells(n, 9) = "4t" Or Sheets(Page).Cells(n, 9) = "5t6t" Then Sheets(Page).Cells(n, 33) = "tout-petits"
                        
                        Sheets(Page).Cells(n, 34) = Left(Sheets(Page).Cells(n, 14), 9)
                        'à modifier
                        Sheets(Page).Cells(n, 44) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & typebarre & "_" & couleur & "_" & genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                        
                        n = n + 1
                    End If
                    
                Next j
                debut = 0
            Next p
            
        Next I
        
        If m = 6 Or m = 7 Then
            p = 1
            While Sheets(Page).Cells(p, 1) <> ""
                If Sheets(Page).Cells(p, 26) = 3 Then Sheets(Page).Cells(p, 26) = ""
                If Sheets(Page).Cells(p, 26) = 4 Then Sheets(Page).Cells(p, 26) = 3
                If Sheets(Page).Cells(p, 26) = 5 Then Sheets(Page).Cells(p, 26) = 4
                p = p + 1
            Wend
            
        End If
        
        Range(Sheets(Page).Cells(1, 1), Sheets(Page).Cells(1, 46)).Select
        Selection.Copy
        
        IMPORT.Activate
        IMPORT.Sheets(1).Cells(1, 1).Select
        ActiveSheet.Paste
        
        ImportSource.Activate
        
        Range(Sheets(Page).Cells(2, 1), Sheets(Page).Cells(Sheets(Page).Cells(1, 1).CurrentRegion.Rows.Count, 46)).Select
        Selection.Copy
        
        IMPORT.Activate
        IMPORT.Sheets(1).Cells(IMPORT.Sheets(1).Cells(1, 1).CurrentRegion.Rows.Count + 1, 1).Select
        ActiveSheet.Paste
        
    Next m
    
    Application.ScreenUpdating = TRUE
    Application.Calculation = xlCalculationAutomatic
End Sub