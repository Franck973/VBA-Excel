Attribute VB_Name = "F_Copier_Dossier"
Public Function Copier_Dossier(D_source As String, D_Copie As String)
'+-------------------------------------------------------------------------------------------------------------------+
'|                                                                                                                   |
'| Copier d'un dossier                                                                                               |
'|                                                                                                                   |
'+-------------------------------------------------------------------------------------------------------------------+
'| Crée par Franck DUBOURTHOUMIEU le 06/08/2019                                                                      |
'+-------------------------------------------------------------------------------------------------------------------+
'|                                                                                                                   |
'| D_source  = Fichier source ex : C:\Test\MonFichier.txt                                                            |
'| D_copie = Fichier copien ex : X:\MesSauvegardes\MonFichier.txt                                                    |
'|                                                                                                                   |
'+-------------------------------------------------------------------------------------------------------------------+
'
'+-------------------------------------------------------------------------------------------------------------------+
'|  Déclaration des variables du module                                                                              |
'+-------------------------------------------------------------------------------------------------------------------+
'
'+-------------------------------------------------------------------------------------------------------------------+
'|  Set des Variables                                                                                                |
'+-------------------------------------------------------------------------------------------------------------------+
'
CopieDossier = False
'
'+-------------------------------------------------------------------------------------------------------------------+
'|  Code du Module                                                                                                   |
'+-------------------------------------------------------------------------------------------------------------------+
'
'vérifie si les arguments ne sont pas vides
If D_source = "" Or D_Copie = "" Then Exit Function

'vérifie s'il n'y a pas des (back)slash à la fin des deux chemins
If Right(D_source, 1) = "\" Or Right(D_source, 1) = "/" Then D_source = Left(D_source, Len(D_source) - 1)
If Right(D_Copie, 1) = "\" Or Right(D_Copie, 1) = "/" Then D_Copie = Left(D_Copie, Len(D_Copie) - 1)

'teste si le Dossier original existe, si oui, copie le dossier et retourne valeur True, si non, retourne valeur False
If Len(Dir(D_source, vbDirectory)) = 0 Then
    Exit Function
Else
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFolder D_source, D_Copie, True
    CopieDossier = True
End If
'
'+-------------------------------------------------------------------------------------------------------------------+
'|  Fin du Module                                                                                                    |
'+-------------------------------------------------------------------------------------------------------------------+
'
End Function
