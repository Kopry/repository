# --------------------------------------------- #
#
# Script de sauvegarde de configuration windows
# Et autre config
# Auteur : Kopry
# Version: 1.1.0
# 
# --------------------------------------------- #
# Version: 1.0.0
# - Ecriture du script
#
# Version: 1.1.0
# - Correction menu (Couleur et Forme)
# - Correction multiple instance IE
# - Amelioration du code (fonction)
# - Affichage d'un rapport par defaut dans le fichier html
#
# Version: 1.2.0
# - Corriger les erreurs possibles (Try and Catch)
# - Ameliorer le template HTML
# - Ajout des dates
# - Ajout d'un menu diff
# - Ajout d'un menu config
# --------------------------------------------- #

# Chemin du dossier temporaire, remplacez SSP par ce que vous voulez
$Chemin = "C:\Users\$env:UserName\SSP"
# Verification de l'existence de ce dossier
$VerifierDossier = Test-Path $Chemin
# Definition des noms pour les rapports complets
$MAJ = "liste_des_mises_a_jour_de_securite.html"
$LOGI = "liste_des_logiciels.html"
$RPF = "liste_regles_du_parefeu.html"
$GPO = "liste_des_gpo.html"
# Titre du document HTML
$titre_document = "Rapport"

If ($VerifierDossier -eq $False){
    # Si pas de dossier temporaire creation d'un dossier temporaire
    mkdir C:\Users\$env:UserName\SSP\
    }# Sinon continuer

# Fonction pour generer un rapport HTML sur les patchs de sécurités installés
function Get-USUH {
        param($Chemin,
              $MAJ,
              $Etat_navigateur)
        # Verification de l'existence du fichier
        if (!(Test-Path -path "$Chemin\$MAJ")){

        # Creation d'un objet windows update pour rechercher les maj installées
        $Index = New-Object -ComObject Microsoft.Update.Session
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
function Get-PI ($Chemin, $LOGI, $Etat_navigateur) {
        if (!(Test-Path -path "$Chemin\$LOGI")){ 
        Get-WmiObject -class Win32_Product | ConvertTo-Html -Property Caption,Vendor,Version | Set-Content $Chemin\$LOGI
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$LOGI")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport HTML sur les regles du parefeu windows
function Get-RPF ($Chemin, $RPF, $Etat_navigateur){
        if (!(Test-Path -path "$Chemin\$RPF")){
        Get-NetFirewallRule | sort direction,applicationName | ConvertTo-Html -property DisplayName,Profile,Enabled,direction| Set-Content $Chemin\$RPF
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$RPF")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport HTML sur les GPO actives
function Get-GPO ($Chemin, $GPO, $Etat_navigateur){
        If (!(Test-Path -path "$Chemin\$GPO")){
        gpresult /h "$Chemin\$GPO"
        }
        If (!$Etat_navigateur){
        $Navigateur.navigate2("$Chemin\$GPO")
        $Navigateur.visible=$True
        }
}

# Fonction pour generer un rapport complet

$Menu = 0
while ($Menu -lt '7'){ 
# Menu du script
#cls
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
        6 {"Rapport complet"

                    $Etat_navigateur = "False"
                    # Construction d'un rapport complet
                    Get-USUH $Chemin $MAJ $Etat_navigateur
                    Get-PI $Chemin $LOGI $Etat_navigateur
                    Get-RPF $Chemin $RPF $Etat_navigateur
                    Get-GPO $Chemin $GPO $Etat_navigateur

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
                    <frame src='$MAJ' name='conteneur'>
                    </frameset>
                    </html>
                    " > $Chemin\Rapport_complet.html

                    # Affichage du rapport
                    $Navigateur.navigate("$Chemin\Rapport_complet.html")
                    $Navigateur.visible=$True
                    $Etat_navigateur = "True"
          }
        7 {"Quitter"}
        # Si l'utilisateur entre n'importe quoi
        default {"Erreur."}
    }
}

