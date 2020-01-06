Option Compare Text
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
	Dim element     As Variant
	On Error GoTo IsInArrayError: 'array is empty
	For Each element In arr
		If element = valToBeFound Then
			IsInArray = True
			Exit Function
		End If
		Next element
		Exit Function
		IsInArrayError:
		On Error GoTo 0
		IsInArray = False
	End Function
Sub ToutesPochesToutesBases()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Set FichierImportProduit = ActiveWorkbook
    NomUser = InputBox("Quel est ton nom de user") 'INFO=Pour le filepath de sauvegarde
    NomFichier = Format(Date, "yyyy-mm-dd") & "_" & InputBox("Comment veux-tu appeler le fichier?") 'Pour le filename
    Workbooks.Add 'TODO=Intégrer l'argument Fileformat = xlCSVUTF8
    ActiveWorkbook.SaveAs Filename:="C:\Users\" & NomUser & "\pochesetfils.com\PUBLIC - Documents\004 Web\01Outils\01ImportationProduit\" & NomFichier
    Dim ImportFile As Workbook
    Set ImportFile = ActiveWorkbook
    Dim couleuradecliner As New Collection
    Dim couleursinterdites As New Collection
    For Each m In Array(101) 'INFO=3 premiers chiffres des VPN à décliner
        'TODO=Intégrer db.csv au fichier d'import ??
        'FUTUREUSE=FichierImportProduit.Activate
        'FUTUREUSE=Worksheets("db").Activate
        Dim db As Workbook
        Set db = Workbooks.Open(Filename:="C:\Users\" & NomUser & "\pochesetfils.com\PUBLIC - Documents\008 Opérations\09 macros\ImportationParfaite\db.csv")
        Dim oStyle As New clsStyle 'INFO=oStyle est un objet custom qui sert à contenir tous les attributs de chaque produit et variante
        n = Application.WorksheetFunction.Match(m, Range("A1:A20"), 0)
            oStyle.Page = Cells(n, 2)
            oStyle.gender = Cells(n, 3)
            oStyle.genderbarre = Cells(n, 4)
            oStyle.tagsGenre = Cells(n, 5)
            oStyle.tagsCollections = Cells(n, 6)
            oStyle.googlegender = Cells(n, 7)
            oStyle.prix = Cells(n, 8)
            oStyle.codetee = Cells(n, 9)
            oStyle.typeprod = Cells(n, 10)
            oStyle.typebarre = Cells(n, 11)
            oStyle.couleurdebut = Cells(n, 12)
            oStyle.couleurfin = Cells(n, 13)
            oStyle.couleursoriginales = Split(Cells(n, 14), ",")
            oStyle.googleage = Cells(n, 15)
            oStyle.a = Cells(n, 16)
            oStyle.b = Cells(n, 17)
            oStyle.seo = Cells(n, 18)
        FichierImportProduit.Sheets(oStyle.Page).Activatel
        Range(Sheets(oStyle.Page).Cells(2, 1), Sheets(oStyle.Page).Cells(10000, 1000)).Clear 'TODO=Rendre dynamique la taille du .Clear au lieu de 10k, 1k
        n = 2
        For I = 1 To Sheets("poches à décliner").Cells(1, 1).CurrentRegion.Rows.Count
            Set couleursadecliner = New Collection
            Set couleursinterdites = New Collection
            If m = 101 Or m = 501 Then 'INFO=Cette condition skip les lignes des poches qu'on ne décline pas pour enfants
                PochesPourAdultes = Array("0006", "0008", "0093", "0141", "C271", "C264", "C282", "C268")
                If IsInArray(Sheets("poches à décliner").Cells(I, 1), PochesPourAdultes) Then I = I + 1
            End If
            'INFO=k un compteur qu'on réutilise dans plusieurs boucles hermétiques
            For k = 3 To 20 'INFO=Nb de colonnes dans page "couleurs prioritaires"
                'TODO=Étudier le fonctionnement du loop ici
                If CStr(Sheets("couleurs prioritaires").Cells(1, k)) = CStr(oStyle.codetee) Then col = k
                If CStr(Sheets("couleurs prioritaires").Cells(1, k)) = "Do not do" Then col_interdite = k
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
        couleurstemp = oStyle.couleursoriginales
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
        debut = 1
        For Each p In couleursadecliner                                
                        couleur = Sheets("FR").Cells(p, 9)
                        couleurnom = Sheets("FR").Cells(p, 10)
                        codecouleur = Sheets("FR").Cells(p, 8)
            For j = oStyle.a To oStyle.b 'tailles à décliner
                codetaille = Sheets("FR").Cells(j, 19)
                If j = oStyle.a And debut = 1 Then
                    Sheets(oStyle.Page).Cells(n, 1) = oStyle.typebarre & "_" & oStyle.genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2)
                    Sheets(oStyle.Page).Cells(n, 2) = Sheets("poches à décliner").Cells(I, 3)
                    Sheets(oStyle.Page).Cells(n, 3) = Sheets("poches à décliner").Cells(I, 4)
                    Sheets(oStyle.Page).Cells(n, 4) = Replace(Sheets("poches à décliner").Cells(I, 5), "'", "_")
                    Sheets(oStyle.Page).Cells(n, 5) = oStyle.typeprod ' type de produit
                    Sheets(oStyle.Page).Cells(n, 6) = "gender:" & oStyle.tagsGenre & oStyle.tagsCollections & ", collection:" & Sheets("poches à décliner").Cells(I, 6) & ", collection:" & Sheets("poches à décliner").Cells(I, 6)
                    If oStyle.gender = "homme" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & "-Homme"
                    If oStyle.gender = "femme" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & "-Femme"
                    If oStyle.gender = "enfant" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & "-Enfant"
                    If oStyle.gender = "bébé" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & "-Bebe"
                    If Sheets("poches à décliner").Cells(I, 8) = "new" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & ", B2S19"
                    If Sheets("poches à décliner").Cells(I, 8) = "new" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & ", collection:Nouveautés FW19"
                    If Sheets("poches à décliner").Cells(I, 8) = "new" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & ", collection:Nouveauté"
                    If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "homme" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & ", collection:Homme - Nouvelles poches FW19"
                    If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "femme" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & ", collection:Femme - Nouvelles poches FW19"
                    If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "enfant" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & ", collection:Enfant - Nouvelles poches FW19"
                    If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "bébé" Then Sheets(oStyle.Page).Cells(n, 6) = Sheets(oStyle.Page).Cells(n, 6) & ", collection:Bébé - Nouvelles poches FW19"
                    Sheets(oStyle.Page).Cells(n, 7) = "'true"
                    Sheets(oStyle.Page).Cells(n, 8) = "Size"
                    Sheets(oStyle.Page).Cells(n, 9) = Sheets("FR").Cells(j, 18) 'taille
                    Sheets(oStyle.Page).Cells(n, 10) = "Color"
                    Sheets(oStyle.Page).Cells(n, 11) = couleurnom
                    Sheets(oStyle.Page).Cells(n, 14) = oStyle.codetee & codecouleur & Sheets("poches à décliner").Cells(I, 1) & "-" & codetaille ' changer le début seulement
                    Sheets(oStyle.Page).Cells(n, 16) = "shopify"
                    Sheets(oStyle.Page).Cells(n, 17) = 100000
                    Sheets(oStyle.Page).Cells(n, 18) = "deny"
                    Sheets(oStyle.Page).Cells(n, 19) = "manual"
                    Sheets(oStyle.Page).Cells(n, 20) = oStyle.prix
                    Sheets(oStyle.Page).Cells(n, 22) = "'true"
                    Sheets(oStyle.Page).Cells(n, 23) = "'true"
                    Sheets(oStyle.Page).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.typebarre & "_" & couleur & "_" & oStyle.genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                    Sheets(oStyle.Page).Cells(n, 26) = 1
                    Sheets(oStyle.Page).Cells(n, 27) = oStyle.typeprod & " pour " & oStyle.gender & " avec poche " & Sheets("poches à décliner").Cells(I, 3) & " à motif de " & Sheets("poches à décliner").Cells(I, 7)
                    Sheets(oStyle.Page).Cells(n, 28) = "'false"
                    Sheets(oStyle.Page).Cells(n, 29) = Sheets("poches à décliner").Cells(I, 3) & " - " & oStyle.typeprod & " " & couleurnom & " " & oStyle.gender & " | Poches & Fils"
                    Sheets(oStyle.Page).Cells(n, 30) = oStyle.seo
                    Sheets(oStyle.Page).Cells(n, 31) = "=VLOOKUP(LEFT(RC[3],3),'Google merchant FR'!R2C1:R50C2,2,0)"
                    Sheets(oStyle.Page).Cells(n, 32) = oStyle.googlegender
                    Sheets(oStyle.Page).Cells(n, 33) = oStyle.googleage
                    Sheets(oStyle.Page).Cells(n, 34) = Left(Sheets(oStyle.Page).Cells(n, 14), 9)
                    Sheets(oStyle.Page).Cells(n, 35) = Sheets(oStyle.Page).Cells(n, 5)
                    Sheets(oStyle.Page).Cells(n, 37) = "neuf"
                    'à modifier
                    Sheets(oStyle.Page).Cells(n, 44) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.typebarre & "_" & couleur & "_" & oStyle.genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                    n = n + 1
                Else
                    'à modifier
                    Sheets(oStyle.Page).Cells(n, 1) = oStyle.typebarre & "_" & oStyle.genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2)
                    Sheets(oStyle.Page).Cells(n, 9) = Sheets("FR").Cells(j, 18) 'taille
                    Sheets(oStyle.Page).Cells(n, 11) = couleurnom
                    'à modifier
                    Sheets(oStyle.Page).Cells(n, 14) = oStyle.codetee & codecouleur & Sheets("poches à décliner").Cells(I, 1) & "-" & codetaille ' changer le début seulement
                    Sheets(oStyle.Page).Cells(n, 16) = "shopify"
                    Sheets(oStyle.Page).Cells(n, 17) = 100000
                    Sheets(oStyle.Page).Cells(n, 18) = "deny"
                    Sheets(oStyle.Page).Cells(n, 19) = "manual"
                    Sheets(oStyle.Page).Cells(n, 20) = oStyle.prix
                                        
                    Sheets(oStyle.Page).Cells(n, 22) = "'true"
                    Sheets(oStyle.Page).Cells(n, 23) = "'true"
                    '---Inventaire des bases---
                    'On ajoute les bases qui ont une image de dos disponible
                    If Left(Sheets(oStyle.Page).Cells(n, 14), 3) = "312" Or Left(Sheets(oStyle.Page).Cells(n, 14), 3) = "212" Or Left(Sheets(oStyle.Page).Cells(n, 14), 3) = "203" Or Left(Sheets(oStyle.Page).Cells(n, 14), 3) = "217" Or Left(Sheets(oStyle.Page).Cells(n, 14), 3) = "301" Or Left(Sheets(oStyle.Page).Cells(n, 14), 3) = "210" Then
                        If Sheets(oStyle.Page).Cells(n - 1, 26) = 1 Then
                            Sheets(oStyle.Page).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.typebarre & "_" & couleur & "_" & oStyle.genderbarre & "-back.jpg"
                            Sheets(oStyle.Page).Cells(n, 26) = 2
                            Sheets(oStyle.Page).Cells(n, 27) = "color:" & Sheets("FR").Cells(p, 10)
                        ElseIf Sheets(oStyle.Page).Cells(n - 1, 26) > 1 And Sheets(oStyle.Page).Cells(n - 1, 26) <= nbcouleur Then
                            Sheets(oStyle.Page).Cells(n, 26) = Sheets(oStyle.Page).Cells(n - 1, 26) + 1
                            Sheets(oStyle.Page).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.typebarre & "_" & Sheets("FR").Cells(couleursadecliner(Sheets(oStyle.Page).Cells(n - 1, 26)), 9) & "_" & oStyle.genderbarre & "-back.jpg"
                            Sheets(oStyle.Page).Cells(n, 27) = "color:" & Sheets("FR").Cells(couleursadecliner(Sheets(oStyle.Page).Cells(n - 1, 26)), 10)
                        Else
                        End If
                    End If            
                    Sheets(oStyle.Page).Cells(n, 30) = oStyle.seo
                    Sheets(oStyle.Page).Cells(n, 31) = "=VLOOKUP(LEFT(RC[3],3),'Google merchant FR'!R2C1:R50C2,2,0)"
                    Sheets(oStyle.Page).Cells(n, 32) = oStyle.googlegender
                    Sheets(oStyle.Page).Cells(n, 33) = oStyle.googleage
                    If Sheets(oStyle.Page).Cells(n, 9) = "18m" Or Sheets(oStyle.Page).Cells(n, 9) = "2t" Or Sheets(oStyle.Page).Cells(n, 9) = "3t" Or Sheets(oStyle.Page).Cells(n, 9) = "4t" Or Sheets(oStyle.Page).Cells(n, 9) = "5t6t" Then Sheets(oStyle.Page).Cells(n, 33) = "tout-petits"
                        Sheets(oStyle.Page).Cells(n, 34) = Left(Sheets(oStyle.Page).Cells(n, 14), 9)
                                        'à modifier
                        Sheets(oStyle.Page).Cells(n, 44) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.typebarre & "_" & couleur & "_" & oStyle.genderbarre & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                        n = n + 1
                    End If
                Next j
                debut = 0
            Next p
        Next I
        If m = 6 Or m = 7 Then
            p = 1
            While Sheets(oStyle.Page).Cells(p, 1) <> ""
                If Sheets(oStyle.Page).Cells(p, 26) = 3 Then Sheets(oStyle.Page).Cells(p, 26) = ""
                If Sheets(oStyle.Page).Cells(p, 26) = 4 Then Sheets(oStyle.Page).Cells(p, 26) = 3
                If Sheets(oStyle.Page).Cells(p, 26) = 5 Then Sheets(oStyle.Page).Cells(p, 26) = 4
                p = p + 1
            Wend
        End If
        Range(Sheets(oStyle.Page).Cells(1, 1), Sheets(oStyle.Page).Cells(1, 46)).Select
        Selection.Copy
        ImportFile.Activate
        ImportFile.Sheets(1).Cells(1, 1).Select
        ActiveSheet.Paste 
        FichierImportProduit.Activate
        Range(Sheets(oStyle.Page).Cells(2, 1), Sheets(oStyle.Page).Cells(Sheets(oStyle.Page).Cells(1, 1).CurrentRegion.Rows.Count, 46)).Select
        Selection.Copy
        ImportFile.Activate
        ImportFile.Sheets(1).Cells(ImportFile.Sheets(1).Cells(1, 1).CurrentRegion.Rows.Count + 1, 1).Select
        ActiveSheet.Paste                            
    Next m
    db.Close                             
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub