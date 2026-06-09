Attribute VB_Name = "modOutils"

Option Explicit

' --- Noms des fichiers CSV attendus ---
Private Const CSV_REEL_26  As String = "GL_Reel_2026.csv"
Private Const CSV_BUDGET   As String = "GL_Budget_2026.csv"
Private Const CSV_REEL_25  As String = "GL_Reel_2025.csv"


Public Sub AfficherPilotage()

    Dim moisNoms As Variant
    Dim moisActuel As Integer, seuilActuel As Double
    Dim choixMois As String, choixSeuil As String
    Dim doRefresh As VbMsgBoxResult, doExport As VbMsgBoxResult
    
    moisNoms = Array("Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", _
                     "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre")
    
    ' Lire les parametres actuels
    On Error Resume Next
    moisActuel = ThisWorkbook.Worksheets("PARAM").Range("C4").Value
    seuilActuel = ThisWorkbook.Worksheets("PARAM").Range("C15").Value
    If seuilActuel = 0 Then seuilActuel = 50
    On Error GoTo 0
    
    ' --- ETAPE 1 : Choix du mois ---
    choixMois = InputBox( _
        "ATLAS AGRO HOLDING — Pilotage de la cloture" & vbCrLf & vbCrLf & _
        "Mois de cloture actuel : " & moisActuel & " (" & moisNoms(moisActuel - 1) & ")" & vbCrLf & vbCrLf & _
        "Saisissez le numero du mois de cloture (1 a 12) :" & vbCrLf & _
        "  1 = Janvier     7 = Juillet" & vbCrLf & _
        "  2 = Fevrier     8 = Aout" & vbCrLf & _
        "  3 = Mars        9 = Septembre" & vbCrLf & _
        "  4 = Avril      10 = Octobre" & vbCrLf & _
        "  5 = Mai        11 = Novembre" & vbCrLf & _
        "  6 = Juin       12 = Decembre", _
        "Mois de cloture", CStr(moisActuel))
    
    If choixMois = "" Then Exit Sub ' annule
    If Not IsNumeric(choixMois) Or CInt(choixMois) < 1 Or CInt(choixMois) > 12 Then
        MsgBox "Mois invalide. Saisissez un nombre entre 1 et 12.", vbExclamation
        Exit Sub
    End If
    
    ' --- ETAPE 2 : Seuil de materialite ---
    choixSeuil = InputBox( _
        "Seuil de materialite actuel : " & seuilActuel & " k" & Chr(8364) & vbCrLf & vbCrLf & _
        "Saisissez le nouveau seuil (en k" & Chr(8364) & ") :" & vbCrLf & _
        "(les ecarts inferieurs a ce seuil seront ignores)", _
        "Seuil de materialite", CStr(seuilActuel))
    
    If choixSeuil = "" Then Exit Sub
    If Not IsNumeric(choixSeuil) Then
        MsgBox "Seuil invalide. Saisissez un nombre.", vbExclamation
        Exit Sub
    End If
    
    ' --- ETAPE 3 : Options ---
    doRefresh = MsgBox( _
        "Voulez-vous actualiser les donnees (Power Query) avant de generer ?", _
        vbYesNo + vbQuestion, "Actualisation")
    
    doExport = MsgBox( _
        "Voulez-vous exporter le package en PDF ?", _
        vbYesNo + vbQuestion, "Export PDF")
    
    ' --- ETAPE 4 : Confirmation ---
    Dim recap As String
    recap = "=== RECAPITULATIF ===" & vbCrLf & vbCrLf & _
            "Mois de cloture : " & choixMois & " (" & moisNoms(CInt(choixMois) - 1) & ")" & vbCrLf & _
            "Seuil materialite : " & choixSeuil & " k" & Chr(8364) & vbCrLf & _
            "Actualiser les donnees : " & IIf(doRefresh = vbYes, "Oui", "Non") & vbCrLf & _
            "Exporter en PDF : " & IIf(doExport = vbYes, "Oui", "Non") & vbCrLf & vbCrLf & _
            "Confirmer et lancer ?"
    
    If MsgBox(recap, vbOKCancel + vbInformation, "Atlas Agro - Confirmation") = vbCancel Then
        MsgBox "Operation annulee.", vbInformation
        Exit Sub
    End If
    
    ' --- ETAPE 5 : Execution ---
    Application.ScreenUpdating = False
    Application.StatusBar = "Mise a jour des parametres..."
    
    ' Mettre a jour PARAM
    With ThisWorkbook.Worksheets("PARAM")
        .Range("C4").Value = CInt(choixMois)
        .Range("C15").Value = CDbl(choixSeuil)
    End With
    
    ' Actualiser si demande
    If doRefresh = vbYes Then
        Application.StatusBar = "Actualisation des donnees (Power Query)..."
        Dim cn As WorkbookConnection
        On Error Resume Next
        For Each cn In ThisWorkbook.Connections
            cn.OLEDBConnection.BackgroundQuery = False
            cn.ODBCConnection.BackgroundQuery = False
        Next cn
        On Error GoTo 0
        ThisWorkbook.RefreshAll
        Application.CalculateUntilAsyncQueriesDone
    End If
    
    ' Recalculer
    Application.StatusBar = "Recalcul du modele..."
    Application.CalculateFull
    
    ' Exporter si demande
    If doExport = vbYes Then
        Application.StatusBar = "Export du PDF..."
        GenererPackage  ' appelle la macro existante du module modPackageCloture
    Else
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Parametres mis a jour et modele recalcule." & vbCrLf & _
               "Mois : " & moisNoms(CInt(choixMois) - 1) & " | Seuil : " & choixSeuil & " k" & Chr(8364), _
               vbInformation, "Atlas Agro - Termine"
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub


'=====================================================================
' 2. IMPORT DYNAMIQUE DES FICHIERS CSV


Public Sub ImporterDonnees()

    Dim dossierSource As String
    Dim dossierCible As String
    Dim fichiers(1 To 3) As String
    Dim trouves(1 To 3) As Boolean
    Dim i As Integer
    Dim msg As String
    Dim nbTrouves As Integer
    
    fichiers(1) = CSV_REEL_26
    fichiers(2) = CSV_BUDGET
    fichiers(3) = CSV_REEL_25
    
    ' --- Selecteur de dossier ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selectionnez le dossier contenant les CSV ERP"
        .ButtonName = "Selectionner"
        If .Show = -1 Then
            dossierSource = .SelectedItems(1)
        Else
            MsgBox "Import annule.", vbInformation
            Exit Sub
        End If
    End With
    
    ' --- Verification des fichiers attendus ---
    msg = "=== ANALYSE DU DOSSIER ===" & vbCrLf
    msg = msg & dossierSource & vbCrLf & vbCrLf
    
    nbTrouves = 0
    For i = 1 To 3
        If Dir(dossierSource & Application.PathSeparator & fichiers(i)) <> "" Then
            trouves(i) = True
            nbTrouves = nbTrouves + 1
            ' Compter les lignes pour info
            Dim nbLignes As Long
            nbLignes = CompterLignesCSV(dossierSource & Application.PathSeparator & fichiers(i))
            msg = msg & "  OK  " & fichiers(i) & " (" & Format(nbLignes, "#,##0") & " lignes)" & vbCrLf
        Else
            trouves(i) = False
            msg = msg & "  --  " & fichiers(i) & " (non trouve)" & vbCrLf
        End If
    Next i
    
    msg = msg & vbCrLf & nbTrouves & " fichier(s) sur 3 detecte(s)."
    
    If nbTrouves = 0 Then
        MsgBox msg & vbCrLf & vbCrLf & "Aucun fichier CSV trouve dans ce dossier.", _
               vbExclamation, "Atlas Agro - Import"
        Exit Sub
    End If
    
    ' --- Confirmation avant copie ---
    msg = msg & vbCrLf & vbCrLf & "Copier les fichiers trouves vers le repertoire de travail ?"
    If MsgBox(msg, vbYesNo + vbQuestion, "Atlas Agro - Import") = vbNo Then
        Exit Sub
    End If
    
    ' --- Determiner le dossier cible (le dossier du classeur ou un sous-dossier "data") ---
    dossierCible = ThisWorkbook.Path & Application.PathSeparator & "data"
    If Dir(dossierCible, vbDirectory) = "" Then
        ' Essayer directement le dossier du classeur
        dossierCible = ThisWorkbook.Path
    End If
    
    ' --- Copie des fichiers ---
    Dim nbCopies As Integer: nbCopies = 0
    On Error GoTo ErreurCopie
    For i = 1 To 3
        If trouves(i) Then
            FileCopy dossierSource & Application.PathSeparator & fichiers(i), _
                     dossierCible & Application.PathSeparator & fichiers(i)
            nbCopies = nbCopies + 1
        End If
    Next i
    On Error GoTo 0
    
    ' --- Actualisation automatique ---
    Dim doRefresh As VbMsgBoxResult
    doRefresh = MsgBox( _
        nbCopies & " fichier(s) copie(s) dans :" & vbCrLf & _
        dossierCible & vbCrLf & vbCrLf & _
        "Voulez-vous actualiser les donnees maintenant ?", _
        vbYesNo + vbQuestion, "Atlas Agro - Import termine")
    
    If doRefresh = vbYes Then
        Application.ScreenUpdating = False
        Application.StatusBar = "Actualisation des donnees..."
        Dim cn As WorkbookConnection
        On Error Resume Next
        For Each cn In ThisWorkbook.Connections
            cn.OLEDBConnection.BackgroundQuery = False
            cn.ODBCConnection.BackgroundQuery = False
        Next cn
        On Error GoTo 0
        ThisWorkbook.RefreshAll
        Application.CalculateUntilAsyncQueriesDone
        Application.CalculateFull
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Import et actualisation termines avec succes.", _
               vbInformation, "Atlas Agro"
    End If
    Exit Sub

ErreurCopie:
    MsgBox "Erreur lors de la copie : " & Err.Description, vbCritical
End Sub


'=====================================================================
' UTILITAIRE : Compte les lignes d'un fichier CSV

Private Function CompterLignesCSV(chemin As String) As Long
    Dim f As Integer
    Dim ligne As String
    Dim compteur As Long
    
    f = FreeFile
    Open chemin For Input As #f
    compteur = 0
    Do While Not EOF(f)
        Line Input #f, ligne
        compteur = compteur + 1
    Loop
    Close #f
    CompterLignesCSV = compteur - 1  ' moins l'en-tete
End Function
