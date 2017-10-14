# --------------------------------------------- #
#
# Script de sauvegarde de configuration windows
# Et autre config
# Auteur : Kopry
# Version: 1.2.0
# 
# --------------------------------------------- #
# Version: 1.2.0
# - Configuration reseau (Menu 7)
# - Ajout des dates
# - Correction des commandes
# - Amelioration du menu
# - Correction bug menu 2, menu 8
#
# Version 1.3.0
# - Correction de l'heure
# - Ajout des services windows
#
# Version 1.4.0
# - Corriger les erreurs possibles (Try and Catch)
# - Ajout d'un menu diff
# - Ameliorer le template HTML
# - Ajout d'un menu config
# --------------------------------------------- #

# ----------------------------------------- -----------------------------------------
# Parametre global du script
# ----------------------------------------- -----------------------------------------

# Chemin du dossier temporaire, remplacez SSP par ce que vous voulez
$Chemin = "C:\Users\$env:UserName\SSP"
# Verification de l'existence de ce dossier
$VerifierDossier = Test-Path $Chemin

# Definition des noms pour les rapports complets
$MAJ = "liste_des_mises_a_jour_de_securite.html"
$LOGI = "liste_des_logiciels.html"
$RPF = "liste_regles_du_parefeu.html"
$GPO = "liste_des_gpo.html"
$CR = "liste_configuration_reseau.html"
$SW = "liste_des_services_windows.html"

# Datation
$Date = Get-Date -UFormat "%Y / %m / %d / %A à %H:%M"
# Titre du document HTML
$titre_document = "Rapport"

If (!($VerifierDossier)){
    # Si pas de dossier temporaire creation d'un dossier temporaire
    mkdir C:\Users\$env:UserName\SSP\
    }# Sinon continuer

# ----------------------------------------- -----------------------------------------
# Fonction de recuperations des rapports
# ----------------------------------------- -----------------------------------------

# Fonction pour generer un rapport HTML sur les patchs de sécurités installés
function Get-USUH {
        param($Chemin,
              $MAJ,
              $Etat_navigateur)
        # Verification de l'existence du fichier
        if (!(Test-Path -path "$Chemin\$MAJ")){

        # Creation d'un objet windows update pour rechercher les maj installées
        $Session = New-Object -ComObject Microsoft.Update.Session
        # Recherche des mise à jour
        $Index = $Session.CreateUpdateSearcher()
        # Puis enregistrement dans un rapport
        $Index.Search("IsInstalled=1").Updates | ConvertTo-Html -Property Date,Title,LastDeploymentChangeTime | Set-Content $Chemin\$MAJ

        } # Sinon afficher le fichier
        # Appel de l'objet IE pour charger le fichier dontenant les resultats
        # Desactivé si $Etat_navigateur = False
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$MAJ")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport HTML sur les programmes installés
function Get-PI {
        param($Chemin,
              $PI,
              $Etat_navigateur)
        if (!(Test-Path -path "$Chemin\$LOGI")){ 
        Get-WmiObject -class Win32_Product | ConvertTo-Html -Property Caption,Vendor,Version | Set-Content $Chemin\$LOGI
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$LOGI")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport HTML sur les regles du parefeu windows
function Get-RPF {
        param($Chemin,
              $RPF,
              $Etat_navigateur)
        if (!(Test-Path -path "$Chemin\$RPF")){
        Get-NetFirewallRule | sort direction,applicationName | ConvertTo-Html -property DisplayName,Profile,Enabled,direction | Set-Content $Chemin\$RPF
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$RPF")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport HTML sur les GPO actives
function Get-GPO {
        param($Chemin,
              $GPO,
              $Etat_navigateur)
        If (!(Test-Path -path "$Chemin\$GPO")){
        gpresult /h "$Chemin\$GPO"
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$GPO")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport HTML sur la configuration reseau de l'ordinateur
function Get-CR {
        param($Chemin,
              $CR,
              $Etat_navigateur)
        If (!(Test-Path -path "$Chemin\$CR")){
        Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName . | Select-Object -Property [a-z]* -ExcludeProperty IPX*,WINS* | ConvertTo-Html -property Description,MACAddress,@{l="IPAddress";e={$_.IPAddress -join " "}},@{l="DefaultIPGateway";e={$_.DefaultIPGateway -join " "}}  | Set-Content $Chemin\$CR
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$CR")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport HTML sur les services windows de l'ordinateur
function Get-SW {
        param($Chemin,
              $SW,
              $Etat_navigateur)
        If (!(Test-Path -path "$Chemin\$SW")){
        Get-Service | Sort-Object Status | ConvertTo-Html | Set-Content $Chemin\$SW
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$SW")
        $Navigateur.visible=$True
        }
}

# ----------------------------------------- -----------------------------------------
# Menu de configuration
# ----------------------------------------- -----------------------------------------

# A rediger

# ----------------------------------------- -----------------------------------------
# Menu principal
# ----------------------------------------- -----------------------------------------

$Menu = 0
while ($Menu -lt '8'){ 
# Menu du script
cls
# Declaration d'un objet "Navigateur" pour afficher nos resultats HTML 
# (celui-ci est redeclaré a chaque boucle, a cause d'un bug)
$Navigateur=new-object -com internetexplorer.application
"
 ----------------------------------------------
|(Menu) Export des parametres windows          |
 ----------------------------------------------
| (1)   Ou les fichiers sont-ils sauvegardés ? |
| (2)   Liste des patchs de sécurité           |
| (3)   Programmes installés                   |
| (4)   Regles du Pare-Feu                     |
| (5)   GPO actives                            |
| (6)   Configuration Réseau                   |
| (7)   Services windows                       |
| (8)   Construction d'un rapport complet      |
| (9)   Quitter                                |
 ----------------------------------------------
"
$menu = Read-Host -Prompt '>'
switch ($menu) 
    { 
        1 {"Par defaut le script enregistre ses fichiers dans le dossier personnel de l'utilisateur
            > $Chemin
           "
                    # Chemin d'enregistrement des fichiers stockant les resultats
                    # Parametre pour chaques elemens (sous menu)
                    # Emplacement des fichier et suppression des rapports déjà construit
                    pause
          } 
        2 {"Liste des patchs de sécurité"
                    # Liste complete des mises à jour (Update / Security Update / Hotfixes)
                    Get-USUH $Chemin $MAJ
          } 
        3 {"Programmes installés"
                    # Programmes installés
                    Get-PI $Chemin $LOGI
          } 
        4 {"Regles du Pare-Feu"
                    # Liste des regles du parefeu
                    Get-RPF $Chemin $RPF
          } 
        5 {"GPO actives" 
                    # Liste des GPO
                    Get-GPO $Chemin $GPO
          }
        6 {"Configuration Reseau" 
                    # Liste de la configuration des interfaces
                    Get-CR $Chemin $CR
          }
        6 {"Configuration Reseau" 
                    # Liste de la configuration des interfaces
                    Get-SW $Chemin $SW
          }
        8 {"Rapport complet"

                    $Etat_navigateur = "False"
                    # Construction de chaque rapport
                    Get-USUH $Chemin $MAJ $Etat_navigateur
                    Get-PI $Chemin $LOGI $Etat_navigateur
                    Get-RPF $Chemin $RPF $Etat_navigateur
                    Get-GPO $Chemin $GPO $Etat_navigateur
                    Get-CR $Chemin $CR $Etat_navigateur
                    Get-SW $Chemin $SW $Etat_navigateur
                    $Etat_navigateur = "True"

                    # Generation du sommaire
                    echo "<html><body>
                    <h1>$titre_document</h1><br/>
                    <a href='./$MAJ' target='conteneur'>$MAJ</a><br/>
                    <a href='./$LOGI' target='conteneur'>$LOGI</a><br/>
                    <a href='./$RPF' target='conteneur'>$RPF</a><br/>
                    <a href='./$GPO' target='conteneur'>$GPO</a><br/>
                    <a href='./$CR' target='conteneur'>$CR</a><br/>
                    <a href='./$SW' target='conteneur'>$SW</a><br/>
                    <br/><br/><br/><br/><h2>$Date</h2>
                    </body></html>" > $Chemin\sommaire.html

                    # Generation de l'index
                    echo "<html><head><title>$titre_document</title></head>
                    <frameset cols='300,*'>
                    <frame src='sommaire.html' name='sommaire'>
                    <frame src='$MAJ' name='conteneur'>
                    </frameset>
                    </html>
                    " > $Chemin\Rapport_complet.html

                    # Affichage de l'index du rapport
                    $Navigateur.navigate("$Chemin\Rapport_complet.html")
                    $Navigateur.visible=$True
          }
        9 {"Quitter"
                    cls
                    break
          }
        # Si l'utilisateur entre n'importe quoi
        default {"Erreur."}
    }
}