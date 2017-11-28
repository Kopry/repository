<#
	.SYNOPSIS
        **MBSA_Update.ps1**
		Ce script récupere un fichier de rapport de MBSA (Microsoft Baseline Security Analyzer) et télécharge tous les 
        packages manquant listés dans ce rapport, il générera ensuite un script automatique d'installation de ces paquets
        a utiliser sur un poste distant ne pouvant pas se connecter sur internet et necessitant une mise à jour.
	.DESCRIPTION (DETAIL)
		Le présent script utilise pour fonctionner les utilitaires pkgmgr.exe / dism.exe / wget.exe
        l'utilitaire wget n'est pas nativement inclu dans une installation windows, il vous faudra
        lorsque vous utiliserez l'option adéquate du script télécharger wget ou utiliser celui inclu dans l'archive.
	.EXAMPLE
		Téléchargement du fichier KB 3035132
        Démarrage du téléchargement  Security Update for Windows 7 (KB3035132) .	
	.OUTPUTS
		Fichiers KB######
        Script _Install.bat
	.NOTES
		AUTHOR	: *
		DATE	: 2017/11/17
		VERSION HISTORY	:
			☑ 1.0 | 
				Version initiale
			☑ 1.1 | Bugs
				Correction des bugs initiaux au script comme le mauvais encodage UTF-8 - BOM (incorrect) pour le script d'installation
                L'utilisation de l'utilitaire pkgmgr.exe est déprécié, mise à jour en conséquence pour l'utilitaire dism.exe
                Ajout d'une méthode de téléchargement annexe (WGET) dans le cas ou les modules pour télécharger nativement avec
                powershell windows ne soient pas disponible
                Ajout User Agent personalisé (Firefox) dans le cas ou un filtrage s'effectuerait sur l'UA
            ☑ 1.2 | CLI - menu
                Construction d'un menu pour personaliser les parametres du script

#>

#Erreur plus détaillée
#$error[0]|format-list -force

#WGET version
#Choix de la méthode de téléchargement
#Mettre sur true pour utiliser uniquement le téléchargement windows sinon false pour utiliser wget (necessite wget.exe)
$downloadMethod = $false

#Definition de quelques variables
$proxy = 'proxy'
Write-Host 'Définition du proxy', $proxy
Write-Host 'Définition des localisation'
$UpdateXML = “updates.xml”
Write-Host 'fichier a utiliser > ', $UpdateXML
$toFolder = “updates”
Write-Host 'dossier de stockage des mise à jour > ',$toFolder
$installFile = $toFolder +”\_Install.bat”
Write-Host 'lien vers le script > ', $installFile
$userAgent = 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:45.0) Gecko/20100101 Firefox/45.0'
Write-Host 'User Agent défini > ', $userAgent



#Inutile si WGET
#Initialisation de l'objet Webclient (pour télécharger les données)
#$webclient = New-Object Net.Webclient
#$webClient.UseDefaultCredentials = $true
#$webclient.Headers.Add($userAgent) 
#$webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
#Write-Host 'Initialisation de', $webClient
#Fin Inutile


#Trie le contenu du fichier xml (récupere la liste des mises à jour a télécharger
#[XML]$xmlList = Get-Content –Path $xml
#$Updates = $xmlList.SelectSingleNode('//UpdateData[@IsInstalled="false"]')
$Updates = [xml](Get-Content $UpdateXML)
Write-Host 'Lecture du fichier XML'
#TO DO - verifier recuperation uniquement des fichier manquant


“@Echo Off” | Out-File $installFile -encoding utf8
“REM Ce script va installer les patchs” | Out-File $installFile -Append -encoding utf8

foreach ($Check in $Updates.XMLOut.Check)
{
    Write-Host “Verification pour”, $Check.Name
    Write-Host $Check.Advice.ToString()

    #Verification des fichiers a télécharger
    foreach ($UpdateData in $Check.Detail.UpdateData)
    {
        if ($UpdateData.IsInstalled -eq $false)
    {
    Write-Host “Téléchargement du fichier KB”, $UpdateData.KBID
    Write-Host “Démarrage du téléchargement “, $UpdateData.Title, “.”
    $url = [URI]$UpdateData.References.DownloadURL
    $fileName = $url.Segments[$url.Segments.Count – 1]
    $toFile = $toFolder +”\”+ $fileName

    if ($downloadMethod -eq $true)
    {
        #Méthode native
        $webClient.DownloadFile($url, $toFile)
    }
    elseif ($downloadMethod -eq $false)
    {
        #Méthode annexe
        #Utilisation du module WGET dans le cas ou ca ne fonctionnerait pas (ps necessite d'avoir le binaire wget a côté)
        #Wget étant très bavard le cmdlt '2>&1 | Out-Null' permet de le rendre muet et d'avoir un affichage propre
        #cependant on ne voit pas les erreurs qu'il retourne, supprimez cette commande si jamais un probleme semble s'etre produit
        .\wget.exe -c -U $userAgent -e use_proxy=yes -e http_proxy=$proxy $url -O $toFile
        Write-Host “Téléchargement terminé”
    }

    #GENERATION DU SCRIPT D'INSTALLATION

    #Affiche le package courant
    “@echo "+ $fileName | Out-File $installFile -Append -encoding utf8
    #Compteur pour afficher la progression totale
    "set /A count+=1" | Out-File $installFile -Append -encoding utf8
    #Indique que l'on traite ce package
    "@echo Démarrage de l'installation %count% : “+ $fileName | Out-File $installFile -Append -encoding utf8
    #Si l'extension du fichier est .msu on utilise l'utilitaire wusa.exe (windows update service agent)
    if ($fileName.EndsWith(“.msu”))
    {
        “wusa.exe “+ $fileName + ” /quiet /norestart /log:%SystemRoot%\Temp\KB”+$UpdateData.KBID+”.log” | Out-File $installFile -Append -encoding utf8
    }
    #Sinon si l'extension du fichier est .cab on utilise l'utilitaire dism.exe
    elseif ($fileName.EndsWith(“.cab”))
    {
        “dism.exe /norestart /online /add-package /packagepath:”+ $fileName | Out-File $installFile -Append -encoding utf8
        # ??? + ” /quiet /nostart /l:%SystemRoot%\Temp\KB”+$UpdateData.KBID+”.log”
    }
    #Autrement on tente de le lancer directement
    else
    {
        $fileName + ” /passive /norestart /log %SystemRoot%\Temp\KB”+$UpdateData.KBID+”.log” | Out-File $installFile -Append -encoding utf8
    }
    #Indique si une erreur a pu se produire (note error 0 signifie pas d'erreur)
    “@echo L'installation a retourné le code suivant %ERRORLEVEL%” | Out-File $installFile -Append -encoding utf8
    #Espace pour faire joli
    “@echo.” | Out-File $installFile -Append -encoding utf8
    Write-Host
    }
}
Write-Host
}