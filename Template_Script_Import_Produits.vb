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
'Loop over les bases m
    'Loop over les poches à décliner I
        'Skip poches pour adulte
        'Cible colonne poches prioritaires
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Set FichierImportProduit = ActiveWorkbook
    NomUser = InputBox("Quel est ton nom de user") 'INFO=Pour le filepath de sauvegarde
    NomFichier = Format(Date, "yyyy-mm-dd") & "_" & InputBox("Comment veux-tu appeler le fichier?") 'Pour le filename
    Workbooks.Add 'TODO=Intégrer l'argument Fileformat = xlCSVUTF8
    ActiveWorkbook.SaveAs Filename:="C:\Users\" & NomUser & "\pochesetfils.com\PUBLIC - Documents\004 Web\01Outils\01ImportationProduit\" & NomFichier
    Dim ImportFile As Workbook
    Set ImportFile = ActiveWorkbook
    Dim couleursadecliner As New Collection
    Dim couleursinterdites As New Collection
    For Each m In Array(217) 'INFO=3 premiers chiffres des VPN à décliner 101, 203, 210, 212, 217, 301, 312, 501
        'TODO=Intégrer db.csv au fichier d'import ??
        'FUTUREUSE=FichierImportProduit.Activate
        'FUTUREUSE=Worksheets("db").Activate
        Dim db As Workbook
        Set db = Workbooks.Open(Filename:="C:\Users\" & NomUser & "\pochesetfils.com\PUBLIC - Documents\008 Opérations\09 macros\ImportationParfaite\db.csv")
        Dim oStyle As New clsStyle 'INFO=oStyle est un objet custom qui sert à contenir tous les attributs de chaque produit et variante
        Dim oVariant As New clsVariant 'INFO=Doit-on déclarer cet objet à cet endroit ou aller plus loin dans le nesting ?
        n = Application.WorksheetFunction.Match(m, Range("A1:A20"), 0)
            oStyle.SheetName = Cells(n, 2)
            oStyle.gender = Cells(n, 3)
            oStyle.GenderASCII = Cells(n, 4)
            oStyle.tagsGenre = Cells(n, 5)
            oStyle.tagsCollections = Cells(n, 6)
            oStyle.googlegender = Cells(n, 7)
            oStyle.prix = Cells(n, 8)
            oStyle.VPN123 = Cells(n, 9)
            oStyle.Style = Cells(n, 10)
            oStyle.style_snake_case = Cells(n, 11)
            oStyle.couleurdebut = Cells(n, 12)
            oStyle.couleurfin = Cells(n, 13)
            oStyle.couleursoriginales = Split(Cells(n, 14), ",")
            oStyle.googleage = Cells(n, 15)
            oStyle.a = Cells(n, 16)
            oStyle.b = Cells(n, 17)
            oStyle.seo = Cells(n, 18)
        FichierImportProduit.Sheets(oStyle.SheetName).Activate
        Range(Sheets(oStyle.SheetName).Cells(2, 1), Sheets(oStyle.SheetName).Cells(10000, 1000)).Clear 'TODO=Rendre dynamique la taille du .Clear au lieu de 10k, 1k
        n = 2
        For I = 1 To Sheets("poches à décliner").Cells(1, 1).CurrentRegion.Rows.Count 'INFO=Itère sur chaque poche à décliner
            Set couleursadecliner = New Collection
            Set couleursinterdites = New Collection
            If m = 101 Or m = 501 Then 'INFO=Skip poches pour adulte
                PochesPourAdultes = Array("0006", "0008", "0093", "0141", "C271", "C264", "C282", "C268")
                If IsInArray(Sheets("poches à décliner").Cells(I, 1), PochesPourAdultes) Then I = I + 1
            End If
            'TODO?=On pourrait déplacer les déclarations de k pour couleurs prioritaires et couleurs interdites à la racine du loop
            'INFO=k un compteur de col qu'on réutilise dans plusieurs boucles hermétiques
            For k = 3 To 20
                If Sheets("couleurs prioritaires").Cells(1, k) = oStyle.VPN123 Then col = k
            Next k
            For k = 2 To Sheets("couleurs prioritaires").Cells(1, 1).CurrentRegion.Rows.Count
                If CStr(Sheets("couleurs prioritaires").Cells(k, 1)) = CStr(Sheets("poches à décliner").Cells(I, 1)) Then
                    couleurfav = CStr(Sheets("couleurs prioritaires").Cells(k, col))
                    lignepoche = k
            Exit For
                End If
            Next k
            col_interdite = Application.Match("Do not do", Sheets("couleurs prioritaires").Range("A1:T1"), 0)
            k = col_interdite
            While Sheets("couleurs prioritaires").Cells(lignepoche, k) <> "" 'INFO=tant qu'il existe une couleur prioritaire
                couleursinterdites.Add CStr((Sheets("couleurs prioritaires").Cells(lignepoche, k)))
                k = k + 1
            Wend
            'MsgBox (couleursinterdites(1) & couleursinterdites(2))
            couleurstemp = oStyle.couleursoriginales
            If IsInArray(couleurfav, couleurstemp) And couleurfav <> 3 Then
                pos = Application.Match(couleurfav, couleurstemp, False) - 1 'SEARCH=-1 car uno-indexé?
                temp = couleurstemp(0)
                couleurstemp(0) = couleurfav 'SEARCH=0 car zéro-indexé?
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
                couleur_snake_case = Sheets("FR").Cells(p, 9)
                couleur_kebab_case = Sheets("FR").Cells(p, 10)
                codecouleur = Sheets("FR").Cells(p, 8)
            For j = oStyle.a To oStyle.b 'tailles à décliner
                codetaille = Sheets("FR").Cells(j, 19)
                If j = oStyle.a And debut = 1 Then
                    oVariant.Handle = oStyle.style_snake_case & "_" & oStyle.GenderASCII & "_" & Sheets("poches à décliner").Cells(I, 2)
                    oVariant.Title = Sheets("poches à décliner").Cells(I, 3)
                    oVariant.Body_HTML = Sheets("poches à décliner").Cells(I, 4)
                    oVariant.Vendor = Replace(Sheets("poches à décliner").Cells(I, 5), "'", "_")
                    x = "gender:" _
                        & oStyle.tagsGenre _
                        & oStyle.tagsCollections _
                        & ", collection:" _
                        & Sheets("poches à décliner").Cells(I, 6) _
                        & ", collection:" _
                        & Sheets("poches à décliner").Cells(I, 6)
                        If oStyle.gender = "homme" Then x = x & "-Homme"
                        If oStyle.gender = "femme" Then x = x & "-Femme"
                        If oStyle.gender = "enfant" Then x = x & "-Enfant"
                        If oStyle.gender = "bébé" Then x = x & "-Bebe"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" Then x = x & ", B2S19"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" Then x = x & ", collection:Nouveauté"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "homme" Then x = x & ", collection:Homme - Nouvelles Poches"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "femme" Then x = x & ", collection:Femme - Nouvelles Poches"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "enfant" Then x = x & ", collection:Enfant - Nouvelles Poches"
                        If Sheets("poches à décliner").Cells(I, 8) = "new" And oStyle.gender = "bébé" Then x = x & ", collection:Bébé - Nouvelles Poches"
                    oVariant.Tags = x
                    oVariant.Published = "'true"
                    oVariant.Option1_Name = "Size"
                    oVariant.Option1_Value = Sheets("FR").Cells(j, 18)
                    oVariant.Option2_Name = "Color"
                    oVariant.Option2_Value = couleur_kebab_case
                    oVariant.Variant_SKU = oStyle.VPN123 & codecouleur & Sheets("poches à décliner").Cells(I, 1) & "-" & codetaille ' changer le début seulement
                    oVariant.Variant_Inventory_Tracker = "shopify"
                    oVariant.Variant_Inventory_Qty = 10000
                    oVariant.Variant_Inventory_Policy = "deny"
                    oVariant.Variant_Fulfillment_Service = "manual"
                    oVariant.Variant_Requires_Shipping = "'true"
                    oVariant.Variant_Taxable = "'true"
                    oVariant.Image_Src = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" _
                        & oStyle.style_snake_case & "_" _
                        & couleur_snake_case & "_" _
                        & oStyle.GenderASCII & "_" _
                        & Sheets("poches à décliner").Cells(I, 2) & ".jpg"

                    Sheets(oStyle.SheetName).Cells(n, 1) = oVariant.Handle
                    Sheets(oStyle.SheetName).Cells(n, 2) = oVariant.Title
                    Sheets(oStyle.SheetName).Cells(n, 3) = oVariant.Body_HTML
                    Sheets(oStyle.SheetName).Cells(n, 4) = oVariant.Vendor
                    Sheets(oStyle.SheetName).Cells(n, 5) = oStyle.Style
                    Sheets(oStyle.SheetName).Cells(n, 6) = oVariant.Tags
                    Sheets(oStyle.SheetName).Cells(n, 7) = oVariant.Published
                    Sheets(oStyle.SheetName).Cells(n, 8) = oVariant.Option1_Name
                    Sheets(oStyle.SheetName).Cells(n, 9) = oVariant.Option1_Value
                    Sheets(oStyle.SheetName).Cells(n, 10) = oVariant.Option2_Name
                    Sheets(oStyle.SheetName).Cells(n, 11) = oVariant.Option2_Value
                    Sheets(oStyle.SheetName).Cells(n, 14) = oVariant.Variant_SKU
                    Sheets(oStyle.SheetName).Cells(n, 16) = oVariant.Variant_Inventory_Tracker
                    Sheets(oStyle.SheetName).Cells(n, 17) = oVariant.Variant_Inventory_Qty
                    Sheets(oStyle.SheetName).Cells(n, 18) = oVariant.Variant_Inventory_Policy
                    Sheets(oStyle.SheetName).Cells(n, 19) = oVariant.Variant_Fulfillment_Service
                    Sheets(oStyle.SheetName).Cells(n, 20) = oStyle.prix
                    Sheets(oStyle.SheetName).Cells(n, 22) = oVariant.Variant_Requires_Shipping
                    Sheets(oStyle.SheetName).Cells(n, 23) = oVariant.Variant_Taxable
                    Sheets(oStyle.SheetName).Cells(n, 25) = oVariant.Image_Src
                    Sheets(oStyle.SheetName).Cells(n, 26) = 1
                    Sheets(oStyle.SheetName).Cells(n, 27) = oStyle.Style & " pour " _
                                                            & oStyle.gender & " avec poche " _
                                                            & Sheets("poches à décliner").Cells(I, 3) _
                                                            & " à motif de " _
                                                            & Sheets("poches à décliner").Cells(I, 7)

                    'TODO=Continuer le pont entre les deux sections (reste 11 variables)

                    Sheets(oStyle.SheetName).Cells(n, 28) = "'false"
                    Sheets(oStyle.SheetName).Cells(n, 29) = Sheets("poches à décliner").Cells(I, 3) & " - " & oStyle.Style & " " & couleur_kebab_case & " " & oStyle.gender & " | Poches & Fils"
                    Sheets(oStyle.SheetName).Cells(n, 30) = oStyle.seo
                    Sheets(oStyle.SheetName).Cells(n, 31) = "=VLOOKUP(LEFT(RC[3],3),'Google merchant FR'!R2C1:R50C2,2,0)"
                    Sheets(oStyle.SheetName).Cells(n, 32) = oStyle.googlegender
                    Sheets(oStyle.SheetName).Cells(n, 33) = oStyle.googleage
                    Sheets(oStyle.SheetName).Cells(n, 34) = Left(Sheets(oStyle.SheetName).Cells(n, 14), 9)
                    Sheets(oStyle.SheetName).Cells(n, 35) = Sheets(oStyle.SheetName).Cells(n, 5)
                    Sheets(oStyle.SheetName).Cells(n, 37) = "neuf"
                    'à modifier
                    Sheets(oStyle.SheetName).Cells(n, 44) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.style_snake_case & "_" & couleur_snake_case & "_" & oStyle.GenderASCII & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                    n = n + 1
                Else
                    'à modifier
                    Sheets(oStyle.SheetName).Cells(n, 1) = oStyle.style_snake_case & "_" & oStyle.GenderASCII & "_" & Sheets("poches à décliner").Cells(I, 2)
                    Sheets(oStyle.SheetName).Cells(n, 9) = Sheets("FR").Cells(j, 18) 'taille
                    Sheets(oStyle.SheetName).Cells(n, 11) = couleur_kebab_case
                    'à modifier
                    Sheets(oStyle.SheetName).Cells(n, 14) = oStyle.VPN123 & codecouleur & Sheets("poches à décliner").Cells(I, 1) & "-" & codetaille ' changer le début seulement
                    Sheets(oStyle.SheetName).Cells(n, 16) = "shopify"
                    Sheets(oStyle.SheetName).Cells(n, 17) = 10000
                    Sheets(oStyle.SheetName).Cells(n, 18) = "deny"
                    Sheets(oStyle.SheetName).Cells(n, 19) = "manual"
                    Sheets(oStyle.SheetName).Cells(n, 20) = oStyle.prix
                                        
                    Sheets(oStyle.SheetName).Cells(n, 22) = "'true"
                    Sheets(oStyle.SheetName).Cells(n, 23) = "'true"
                    '---Inventaire des bases---
                    'On ajoute les bases qui ont une image de dos disponible
                    If Left(Sheets(oStyle.SheetName).Cells(n, 14), 3) = "312" Or Left(Sheets(oStyle.SheetName).Cells(n, 14), 3) = "212" Or Left(Sheets(oStyle.SheetName).Cells(n, 14), 3) = "203" Or Left(Sheets(oStyle.SheetName).Cells(n, 14), 3) = "217" Or Left(Sheets(oStyle.SheetName).Cells(n, 14), 3) = "301" Or Left(Sheets(oStyle.SheetName).Cells(n, 14), 3) = "210" Then
                        If Sheets(oStyle.SheetName).Cells(n - 1, 26) = 1 Then
                            Sheets(oStyle.SheetName).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.style_snake_case & "_" & couleur_snake_case & "_" & oStyle.GenderASCII & "-back.jpg"
                            Sheets(oStyle.SheetName).Cells(n, 26) = 2
                            Sheets(oStyle.SheetName).Cells(n, 27) = "color:" & Sheets("FR").Cells(p, 10)
                        ElseIf Sheets(oStyle.SheetName).Cells(n - 1, 26) > 1 And Sheets(oStyle.SheetName).Cells(n - 1, 26) <= nbcouleur Then
                            Sheets(oStyle.SheetName).Cells(n, 26) = Sheets(oStyle.SheetName).Cells(n - 1, 26) + 1
                            Sheets(oStyle.SheetName).Cells(n, 25) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.style_snake_case & "_" & Sheets("FR").Cells(couleursadecliner(Sheets(oStyle.SheetName).Cells(n - 1, 26)), 9) & "_" & oStyle.GenderASCII & "-back.jpg"
                            Sheets(oStyle.SheetName).Cells(n, 27) = "color:" & Sheets("FR").Cells(couleursadecliner(Sheets(oStyle.SheetName).Cells(n - 1, 26)), 10)
                        Else
                        End If
                    End If
                    Sheets(oStyle.SheetName).Cells(n, 30) = oStyle.seo
                    Sheets(oStyle.SheetName).Cells(n, 31) = "=VLOOKUP(LEFT(RC[3],3),'Google merchant FR'!R2C1:R50C2,2,0)"
                    Sheets(oStyle.SheetName).Cells(n, 32) = oStyle.googlegender
                    Sheets(oStyle.SheetName).Cells(n, 33) = oStyle.googleage
                    If Sheets(oStyle.SheetName).Cells(n, 9) = "18m" Or Sheets(oStyle.SheetName).Cells(n, 9) = "2t" Or Sheets(oStyle.SheetName).Cells(n, 9) = "3t" Or Sheets(oStyle.SheetName).Cells(n, 9) = "4t" Or Sheets(oStyle.SheetName).Cells(n, 9) = "5t6t" Then Sheets(oStyle.SheetName).Cells(n, 33) = "tout-petits"
                        Sheets(oStyle.SheetName).Cells(n, 34) = Left(Sheets(oStyle.SheetName).Cells(n, 14), 9)
                                        'à modifier
                        Sheets(oStyle.SheetName).Cells(n, 44) = "https://raw.githubusercontent.com/poches-et-fils/volume8-images/master/" & oStyle.style_snake_case & "_" & couleur_snake_case & "_" & oStyle.GenderASCII & "_" & Sheets("poches à décliner").Cells(I, 2) & ".jpg"
                        n = n + 1
                    End If
            Next j
            debut = 0
            Next p
        Next I
        If m = 6 Or m = 7 Then
            p = 1
            While Sheets(oStyle.SheetName).Cells(p, 1) <> ""
                If Sheets(oStyle.SheetName).Cells(p, 26) = 3 Then Sheets(oStyle.SheetName).Cells(p, 26) = ""
                If Sheets(oStyle.SheetName).Cells(p, 26) = 4 Then Sheets(oStyle.SheetName).Cells(p, 26) = 3
                If Sheets(oStyle.SheetName).Cells(p, 26) = 5 Then Sheets(oStyle.SheetName).Cells(p, 26) = 4
                p = p + 1
            Wend
        End If
        Range(Sheets(oStyle.SheetName).Cells(1, 1), Sheets(oStyle.SheetName).Cells(1, 46)).Select
        Selection.Copy
        ImportFile.Activate
        ImportFile.Sheets(1).Cells(1, 1).Select
        ActiveSheet.Paste
        FichierImportProduit.Activate
        Range(Sheets(oStyle.SheetName).Cells(2, 1), Sheets(oStyle.SheetName).Cells(Sheets(oStyle.SheetName).Cells(1, 1).CurrentRegion.Rows.Count, 46)).Select
        Selection.Copy
        ImportFile.Activate
        ImportFile.Sheets(1).Cells(ImportFile.Sheets(1).Cells(1, 1).CurrentRegion.Rows.Count + 1, 1).Select
        ActiveSheet.Paste
    Next m
    db.Close SaveChanges:=False
    'On ouvre l'inventaire
    Dim Inventaire As Workbook
    Set Inventaire = Workbooks.Open(Filename:="C:\Users\" & NomUser & _
    "\pochesetfils.com\PUBLIC - Documents\008 Opérations\02GESTION DE COMMANDES\Inventaire P&F.xlsm")

    'Déclaration des variables
    Dim Sku As String, LeftSku As String, RightSku As String, GenericSku As String

    'Trouve la colonne "AVAILABLE TO SELL" dans Inventaire
    Inventaire.Worksheets("Produits à poches").Activate
    ColAvailToSell = Application.WorksheetFunction.Match("AVAILABLE TO SELL", Range("A4:Z4"), 0)

    'Pour chaque chandail à poche dans ImportFile, mettre à -100 si la base est en rupture d'inventaire
    For ligne = 2 To ImportFile.Worksheets(1).Cells(1, 1).CurrentRegion.Rows.Count
        Sku = ImportFile.Worksheets(1).Cells(ligne, 14)
        LeftSku = Left(Sku, 5)
        If Left(LeftSku, 3) = "212" Then LeftSku = Replace(LeftSku, "212", "312") 'Convertir VPN 212 à VPN 312
        RightSku = Right(Sku, 2)
        GenericSku = LeftSku & "XXXX-" & RightSku
        qt = Application.VLookup(GenericSku, Inventaire.Worksheets("Produits à poches").Range("A5:Z500"), ColAvailToSell, False)
        If IsError(qt) = True Then qt = 0
        If qt <= 0 Then ImportFile.Worksheets(1).Cells(ligne, 17) = -100
    Next ligne

    'On ferme l'inventaire
    Inventaire.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub