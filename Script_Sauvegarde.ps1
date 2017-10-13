# --------------------------------------------- #
#
# Script de sauvegarde de configuration windows
# Auteur : Kopry
# Version: 1.0.0
# 
# --------------------------------------------- #
# Version: 1.0.0
# - Ecriture du script
# Version: 1.1.0
# - Corriger les erreurs possibles (Try and Catch)
# - Correction menu
# - Corriger multiple instance IE
# --------------------------------------------- #

# Chemin du dossier temporaire, remplacez SSP par ce que vous voulez
$Chemin = "C:\Users\$env:UserName\SSP"
# Verification de l'existence de ce dossier
$VerifierDossier = Test-Path $Chemin
#Definition des noms pour le rapport complet
$MAJ = "liste_des_mises_a_jour_de_securite.html"
$LOGI = "liste_des_logiciels.html"
$RPF = "liste_regles_du_parefeu.html"
$GPO = "liste_des_gpo.html"
$titre_document = "Rapport"

If ($VerifierDossier -eq $False){
    # Si pas de dossier temporaire creation d'un dossier temporaire
    mkdir C:\Users\$env:UserName\SSP\
    }# Sinon continuer

$Menu = 0
while ($Menu -lt '7'){ 
# Menu du script
cls
# Declaration d'un objet "Navigateur" pour afficher nos resultats HTML
$Navigateur=new-object -com internetexplorer.application
"(Menu) Export des parametres windows
 (1)   Ou les fichiers sont-ils sauvegardés ?
 (2)   Liste des patchs de sécurité
 (3)   Programmes installés
 (4)   Regles du Pare-Feu
 (5)   GPO actives
 (6)   Construction d'un rapport complet
 (7)   Quitter
     "
$menu = Read-Host -Prompt '>'
switch ($menu) 
    { 
        1 {"Par defaut le script enregistre ses fichiers dans le dossier personnel de l'utilisateur
            > $Chemin
           "
                    # Chemin d'enregistrement des fichiers stockant les resultats
                    pause
          } 
        2 {"Liste des patchs de sécurité"
                    # Liste complete des mises à jour (Update / Security Update / Hotfixes)
                    # Verification de l'existence du fichier
                    if ((Test-Path -path "$Chemin\$MAJ") -eq $False){
                    wmic qfe list full /format:htable > "$Chemin\$MAJ"
                    } # Sinon afficher le fichier
                    # Appel de l'objet IE pour charger le fichier dontenant les resultats
                    $Navigateur.navigate2("$Chemin\$MAJ")
                    $Navigateur.visible=$True
          } 
        3 {"Programmes installés"
                    # Programmes installés
                    if (!(Test-Path -path "$Chemin\$LOGI")){ 
                    Get-WmiObject -class Win32_Product | Select-Object -Property Caption,Vendor,Version | ConvertTo-Html | Set-Content $Chemin\$LOGI
                    } # Sinon afficher le fichier
                    $Navigateur.navigate2("$Chemin\$LOGI")
                    $Navigateur.visible=$True
          } 
        4 {"Regles du Pare-Feu"
                    # Liste des regles du parefeu
                    # Verification de l'existence du fichier
                    if (!(Test-Path -path "$Chemin\$RPF")){
                    Get-NetFirewallRule | sort direction,applicationName | Format-Table -wrap -property DisplayName,DisplayGroup,Profile,Enabled,direction | ConvertTo-Html | Set-Content $Chemin\$RPF
                    }# Sinon afficher le fichier
                    $Navigateur.navigate2("$Chemin\$RPF")
                    $Navigateur.visible=$True
          } 
        5 {"GPO actives" 
                    # Liste des GPO
                    # Verification de l'existence du fichier
                    If (!(Test-Path -path "$Chemin\$GPO")){
                    gpresult /h "$Chemin\$GPO"
                    }# Sinon afficher le fichier
                    $Navigateur.navigate2("$Chemin\$GPO")
                    $Navigateur.visible=$True
                    while ($Navigateur.busy) {sleep -milliseconds 50}
          }
        6 {"Rapport complet"
                    # Construction d'un rapport complet
                    if ((Test-Path -path "$Chemin\$MAJ") -eq $False){
                    wmic qfe list full /format:htable > "$Chemin\$MAJ"
                    }
                    if (!(Test-Path -path "$Chemin\$LOGI")){ 
                    Get-WmiObject -class Win32_Product | Select-Object -Property Caption,Vendor,Version | ConvertTo-Html | Set-Content $Chemin\$LOGI
                    }
                    if (!(Test-Path -path "$Chemin\$RPF")){
                    Get-NetFirewallRule | sort direction,applicationName | Format-Table -wrap -property DisplayName,DisplayGroup,Profile,Enabled,direction | ConvertTo-Html | Set-Content $Chemin\$RPF
                    }
                    If (!(Test-Path -path "$Chemin\$GPO")){
                    gpresult /h "$Chemin\$GPO"
                    }

                    # Generation du sommaire
                    echo "<html><body>
                    <h1>$titre_document</h1><br/>
                    <a href='./$MAJ' target='conteneur'>$MAJ</a><br/>
                    <a href='./$LOGI' target='conteneur'>$LOGI</a><br/>
                    <a href='./$RPF' target='conteneur'>$RPF</a><br/>
                    <a href='./$GPO' target='conteneur'>$GPO</a><br/>
                    </body></html>" > $Chemin\sommaire.html

                    # Generation de l'index
                    echo "<html><head><title>$titre_document</title></head>
                    <frameset cols='300,*'>
                    <frame src='sommaire.html' name='sommaire'>
                    <frame src='blank' name='conteneur'>
                    </frameset>
                    </html>
                    " > $Chemin\Rapport_complet.html
                    $Navigateur.navigate("$Chemin\Rapport_complet.html")
                    $Navigateur.visible=$True
          }
        7 {"Quitter"}
        # Si l'utilisateur entre n'importe quoi
        default {"Erreur."}
    }
}