<#
	.SYNOPSIS
		Ce script permet d'activer/désactiver la possibilité d'installer des mises à jour avec DISM (Deployment Image Servicing and Management) facilement et rapidement
        directement via le menu contextuel de windows
	.DESCRIPTION
    
        installState et uninstallState permettent de savoir si les clefs dans le registres
        sont inscrite ou non :

        installState = "Désactivé" Option contextuelle d'installation des fichier cab non affichée
        uninstallState = "Désactivé" Option contextuelle de désinstallation des fichier cab non affichée
        
        installState = "Activé" Option contextuelle d'installation des fichier cab affichée
        uninstallState = "Activé" Option contextuelle de désinstallation des fichier cab affichée
        
        /!\ Il ne semble pas possible d'avoir deux commandes simultanément active dans le menu contextuel de windows permettant d'obtenir un accès administrateur
        /!\ RunAS n'accepte pas plus d'une commande... C'est la raison de l'existence de ce script, il permet de permuter entre plusieurs options
	.EXAMPLE
		Le script demandera a l'utilisateur de choisir quel fonction il veut ajouter et l'ajoutera dans le menu contextuel	
	.OUTPUTS
		Présence/Absence d'une option dans le menu contextuel
	.NOTES
		AUTHOR	: D.Patrick
		DATE	: 28/11/2017
		VERSION HISTORY	:
			☑ 1.0 | 
				Version initiale
#>

<# ----------------------------------------- -----------------------------------------
    Définition des parametres
----------------------------------------- ----------------------------------------- #>

$Keys="RunAs"
$version = '1.0'
# Par défaut les options dans le menu sont désactivées
$installState = "Désactivé"
$uninstallState = "Désactivé"

<# ----------------------------------------- -----------------------------------------
    Auto élévation des privilèges
    Utile dans le cas ou le script serait lancé sans droits particuliers alors que celui ci
    necessite des droits étendus
----------------------------------------- ----------------------------------------- #>

# Obtient l'ID de sécurité de l'utilisateur courrant
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)

# Obtient l'ID de sécurité de l'administrateur
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

# Vérifie d'abord si l'on est soi même un administrateur
if ($myWindowsPrincipal.IsInRole($adminRole))

   {
   # La commande console courrante est administrateur changement de la couleur de la console pour bien indiquer cet etat
   $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Administrateur)"
   $Host.UI.RawUI.BackgroundColor = "DarkBlue"
   clear-host
   }
else
   {
   # Toujours pas administrateur, redémarrage du script en administrateur
   # Création d'un nouveau processus Powershell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   # Specification du chemin du script courrant et son nom en parametre
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;
   # Demande a ce que le processus soit élévé (en utilisant runas)
   $newProcess.Verb = "runas";
   # Démarrage du nouveau processus
   [System.Diagnostics.Process]::Start($newProcess);
# Quitte depuis le processus courrant (qui n'est pas administrateur)
exit
}

Write-Host -NoNewLine "Appuyez sur n'importe quelle touche pour continuer..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

<# ----------------------------------------- -----------------------------------------
     Définition des fonctions du script
----------------------------------------- ----------------------------------------- #>

Function Set-Cab_Install_Option {
    #Creation d'un repertoire RunAs
    New-Item -Path Registry::HKCR\CABFolder\shell\ -Name $Keys
    #Clef par defaut (installer ce fichier cab)
    Set-ItemProperty -Path Registry::HKCR\CABFolder\shell\$Keys -Name "(Default)" -Value "Installer ce fichier (.cab)"
    #HasluaShield
    New-ItemProperty -Path Registry::HKCR\CABFolder\shell\$Keys -Name "HasLUAShield" -PropertyType "String" -Value ""
    #Command
    New-Item -Path Registry::HKCR\CABFolder\shell\$Keys -Name "Command"
    Set-ItemProperty -Path Registry::HKCR\CABFolder\shell\$Keys\Command -Name "(Default)" -Value "cmd /k dism /online /add-package /packagepath:%1"
}

Function Set-Cab_Uninstall_Option {
    #Creation d'un repertoire RunAs
    New-Item -Path Registry::HKCR\CABFolder\shell\ -Name $Keys
    #Clef par defaut (installer ce fichier cab)
    Set-ItemProperty -Path Registry::HKCR\CABFolder\shell\$Keys -Name "(Default)" -Value "Désinstaller ce fichier (.cab)"
    #HasluaShield
    New-ItemProperty -Path Registry::HKCR\CABFolder\shell\$Keys -Name "HasLUAShield" -PropertyType "String" -Value ""
    #Command
    New-Item -Path Registry::HKCR\CABFolder\shell\$Keys -Name "Command"
    Set-ItemProperty -Path Registry::HKCR\CABFolder\shell\$Keys\Command -Name "(Default)" -Value "cmd /k dism /online /remove-package /packagepath:%1"
}

Function UnSet-Cab_Install_Uninstall_Option {
    #Suppression du menu
    Remove-Item -Path Registry::HKCR\CABFolder\shell\$Keys -Recurse
}

<# ----------------------------------------- -----------------------------------------
     Menu principal
----------------------------------------- ----------------------------------------- #>

Set-Cab_Install_Option
$menu = 0
while ($menu -lt '3'){ 
cls
# Menu du script
"
-------------------------------------------------------------------------
(Menu) Activer/Desactiver l'installation facile de fichier .cab          
-------------------------------------------------------------------------
(Version)              | $version                                      
-------------------------------------------------------------------------
 (1)   Installation    | etat :  "+ $installState +"                           
 (2)   Desinstallation | etat :  "+ $uninstallState +"                         
 (3)   Quitter         | note :  tout sera désactivé
-------------------------------------------------------------------------
                                                                       TA
"
$menu = Read-Host -Prompt '>'
switch ($menu) 
    { 
        1 {"Installation"
            If ( $installState -like 'Désactivé'){
                #Si installation actif alors désinstallation inactive (cf note dans la description)
                UnSet-Cab_Install_Uninstall_Option
                $uninstallState = "Désactivé"
                #Changement de l'état a Activé
                $installState = "Activé"
                Set-Cab_Install_Option
                }
            else {
                #Changement de l'état a Désactivé
                $installState = "Désactivé"
                UnSet-Cab_Install_Uninstall_Option
                }
        }

        2 {"Désinstallation"
                If ( $uninstallState -like 'Désactivé'){
                UnSet-Cab_Install_Uninstall_Option
                $installState = "Désactivé"
                #Changement de l'état a Activé
                $uninstallState = "Activé"
                Set-Cab_Uninstall_Option
                }
                else {
                $uninstallState = "Activé"
                UnSet-Cab_Install_Uninstall_Option
                }
        } 
    }
}
# Purge du registre avant de quitter
If (( $uninstallState -like 'Activé') -or ($installState -like 'Activé')){UnSet-Cab_Install_Uninstall_Option}
cls           