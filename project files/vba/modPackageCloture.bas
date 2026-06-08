Attribute VB_Name = "modPackageCloture"
'=====================================================================
' MODULE  : modPackageCloture
' PROJET  : Atlas Agro Holding - Automatisation du reporting de cloture
' OBJET   : Genere le package de cloture mensuel en un clic.
'           1) Actualise les donnees (Power Query / connexions)
'           2) Recalcule le modele
'           3) Horodate la cloture
'           4) Exporte Dashboard + Package + Synthese en un PDF date
'           5) Archive le PDF dans un sous-dossier dedie
' AUTEUR  : Ismail Abgar
'=====================================================================
Option Explicit

' --- Parametres du module ---
Private Const FEUILLES_PACKAGE As String = "DASHBOARD,PACKAGE_PnL,SYNTHESE_BU"
Private Const DOSSIER_ARCHIVE  As String = "Archives_Cloture"

'---------------------------------------------------------------------
' PROCEDURE PRINCIPALE : a lancer via le bouton du dashboard
'---------------------------------------------------------------------
Public Sub GenererPackage()

    Dim tDebut As Double
    Dim cheminPDF As String
    tDebut = Timer

    On Error GoTo GestionErreur

    ' 1) Optimisation de l'affichage pendant le traitement
    OptimiserPerformance True

    ' 2) Actualisation des donnees
    Application.StatusBar = "Etape 1/4 : actualisation des donnees..."
    ActualiserDonnees

    ' 3) Recalcul complet
    Application.StatusBar = "Etape 2/4 : recalcul du modele..."
    Application.CalculateFull

    ' 4) Horodatage
    Application.StatusBar = "Etape 3/4 : horodatage de la cloture..."
    HorodaterCloture

    ' 5) Export PDF + archivage
    Application.StatusBar = "Etape 4/4 : generation du PDF..."
    cheminPDF = ExporterPackagePDF()

    ' 6) Restauration + message de confirmation
    OptimiserPerformance False
    Application.StatusBar = False

    MsgBox "Package de cloture genere avec succes." & vbCrLf & vbCrLf & _
           "Fichier : " & cheminPDF & vbCrLf & _
           "Duree   : " & Format(Timer - tDebut, "0.0") & " s", _
           vbInformation, "Atlas Agro - Cloture"
    Exit Sub

GestionErreur:
    OptimiserPerformance False
    Application.StatusBar = False
    MsgBox "Une erreur est survenue lors de la generation :" & vbCrLf & vbCrLf & _
           "N. " & Err.Number & " - " & Err.Description, _
           vbCritical, "Atlas Agro - Erreur"
End Sub

'---------------------------------------------------------------------
' Actualise toutes les requetes / connexions, en mode synchrone
'---------------------------------------------------------------------
Private Sub ActualiserDonnees()
    Dim cn As WorkbookConnection

    ' Forcer l'attente de la fin de chaque requete (pas d'arriere-plan)
    On Error Resume Next
    For Each cn In ThisWorkbook.Connections
        cn.OLEDBConnection.BackgroundQuery = False
        cn.ODBCConnection.BackgroundQuery = False
    Next cn
    On Error GoTo 0

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
End Sub

'---------------------------------------------------------------------
' Inscrit la date/heure de generation dans l'onglet PARAM
'---------------------------------------------------------------------
Private Sub HorodaterCloture()
    With ThisWorkbook.Worksheets("PARAM")
        .Range("B18").Value = "Derniere generation"
        .Range("C18").Value = Now
        .Range("C18").NumberFormat = "dd/mm/yyyy hh:mm"
    End With
End Sub

'---------------------------------------------------------------------
' Exporte les feuilles du package en un PDF unique et l'archive.
' Retourne le chemin complet du PDF.
'---------------------------------------------------------------------
Private Function ExporterPackagePDF() As String
    Dim cheminDossier As String
    Dim cheminPDF As String
    Dim feuilles() As String

    ' Le classeur doit etre enregistre (pour connaitre son dossier)
    If ThisWorkbook.Path = "" Then
        Err.Raise vbObjectError + 513, , _
            "Enregistre d'abord le classeur (.xlsm) avant de generer le package."
    End If

    ' Creer le dossier d'archive s'il n'existe pas
    cheminDossier = ThisWorkbook.Path & Application.PathSeparator & DOSSIER_ARCHIVE
    If Dir(cheminDossier, vbDirectory) = "" Then MkDir cheminDossier

    ' Nom de fichier date
    cheminPDF = cheminDossier & Application.PathSeparator & NomFichierPDF()

    ' Mise en page propre de chaque feuille
    feuilles = Split(FEUILLES_PACKAGE, ",")
    PreparerMiseEnPage feuilles

    ' Selection des feuilles puis export PDF
    ThisWorkbook.Sheets(feuilles).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=cheminPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Revenir au dashboard
    ThisWorkbook.Worksheets("DASHBOARD").Select

    ExporterPackagePDF = cheminPDF
End Function

'---------------------------------------------------------------------
' Construit le nom du PDF : Cloture_Atlas_Agro_Mars_2026.pdf
'---------------------------------------------------------------------
Private Function NomFichierPDF() As String
    Dim moisNum As Integer, annee As Integer
    Dim moisNoms As Variant
    moisNoms = Array("Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", _
                     "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre")
    With ThisWorkbook.Worksheets("PARAM")
        moisNum = CInt(.Range("C4").Value)
        annee = CInt(.Range("C5").Value)
    End With
    NomFichierPDF = "Cloture_Atlas_Agro_" & moisNoms(moisNum - 1) & "_" & annee & ".pdf"
End Function

'---------------------------------------------------------------------
' Applique une mise en page homogene (paysage, ajuste largeur, footer)
'---------------------------------------------------------------------
Private Sub PreparerMiseEnPage(feuilles() As String)
    Dim i As Integer
    Dim ws As Worksheet
    For i = LBound(feuilles) To UBound(feuilles)
        Set ws = ThisWorkbook.Worksheets(Trim(feuilles(i)))
        With ws.PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .LeftMargin = Application.InchesToPoints(0.3)
            .RightMargin = Application.InchesToPoints(0.3)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.4)
            .CenterHeader = "&""Arial,Bold""&12 Atlas Agro Holding - Package de cloture"
            .LeftFooter = "&""Arial""&8 " & ws.Name
            .CenterFooter = "&""Arial""&8 Page &P / &N"
            .RightFooter = "&""Arial""&8 &D"
        End With
    Next i
End Sub

'---------------------------------------------------------------------
' Active (True) / desactive (False) les optimisations d'affichage
'---------------------------------------------------------------------
Private Sub OptimiserPerformance(actif As Boolean)
    Application.ScreenUpdating = Not actif
    Application.EnableEvents = Not actif
    If actif Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub
