# Script de génération de la liste des équipes Teams et export dans un fichier CSV puis envoie de ce fichier par FTP pour être exploité par un fichier Excel
# Réalisé par Sigalou 06/12/2020 sigalou@@@sigalou-domotique.fr
# Utilisez le mails pour toutes questions, demandes d'amélioration et n'hésitez pas à proposer des émalioration
# Merci de laisser ces lignes

# Identification FPT
	param([string]$UserName, [string]$Password, [switch]$MFA,[int]$Action) 
	$passwordFTP = ConvertTo-SecureString -AsPlainText "xxxxx" -Force
	$credentialsFTP = New-Object System.Management.Automation.PSCredential "xxxxx_ftp", $passwordFTP
	Import-Module PSFTP

# Identification M365
	$UserName = "xxxxx@xxxxx.com"
	$Password = "xxxx"

#Connect to Microsoft Teams
	$Module=Get-Module -Name MicrosoftTeams -ListAvailable 
	if($Module.count -eq 0)
	{
	 Write-Host MicrosoftTeams module is not available  -ForegroundColor yellow 
	 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
	 if($Confirm -match "[yY]")
	 {
	  Install-Module MicrosoftTeams
	 }
	 else
	 {
	  Write-Host MicrosoftTeams module is required.Please install module using Install-Module MicrosoftTeams cmdlet.
	  Exit
	 }
	}
	Write-Host `n`n`n`n`n`nLancement ... -ForegroundColor Yellow

	#Autentication using MFA
	if($mfa.IsPresent)
	{
	 $Team=Connect-MicrosoftTeams
	}
	#Authentication using non-MFA
	else
	{
	 #Storing credential in script for scheduling purpose/ Passing credential as parameter
	 if(($UserName -ne "") -and ($Password -ne ""))
	 {
	  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
	  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
	  $Team=Connect-MicrosoftTeams -Credential $Credential
	 }
	 else
	 {  
	  $Team=Connect-MicrosoftTeams
	 }
	}

	#Check for Teams connectivity
	If($Team -ne $null)
	{
	 Write-host `nConnecté avec succès à Microsoft Teams -ForegroundColor Green
	}
	else
	{
	 Write-Host Error occurred while creating Teams session. Please try again -ForegroundColor Red
	 exit
	}

	# Mettre $True pour avoir un debug à l'ecran pour suivre la génération
	$modeDEBUG=$False
	$Result=""  
	$Results=@() 
	$Path="./TeamsData.csv"
	If (Test-Path $Path){
	Remove-Item $Path
	}
	Write-Host `nRécupération des données...
	$Count=0
	$Teams=Get-Team -Visibility Private 
	$nbTeamsPrivate=$Teams.Count
	#$Teams=Get-Team -Visibility Public -User xxxxx@xxxxxxxxxxxxxxxx.fr
	$Teams=Get-Team -Visibility Public 
	$nbTeams=$Teams.Count
	$ColA=$nbTeams.ToString()+"-"+$nbTeamsPrivate.ToString()
	$Result=@{'ColA'=$ColA;'ColB'="Actualisation du $((Get-Date -format 'dd-MM-yyyy` HH:mm').ToString())";}
	$Results= New-Object psobject -Property $Result
	$Results | select 'ColA','ColB' | Export-Csv $Path -NoTypeInformation -Append
	Write-Host `nTraitement des données...
	$Teams | Sort-Object DisplayName | foreach {
		$compteurLignes=0
		$TeamName=$_.DisplayName
		$Description=$_.Description
		$GroupId=$_.GroupId
		$TeamUser=Get-TeamUser -GroupId $GroupId
		$TeamMemberCount=$TeamUser.Count
		$TeamPrefix=""
		$mettreVirgule=0
		$NomProprio=""
		$TouslesProprios=$TeamUser| Where-Object { $_.Role -eq 'Owner' }
		$TouslesProprios | foreach {
			if ($mettreVirgule -eq 1) {
				$mettreVirgule=0 
				$NomProprio=$NomProprio+", "
			}
			$NomProprio=$NomProprio+$_.Name
			$mettreVirgule=1
		}
		$NomProprio=$NomProprio+" et "+$TeamMemberCount+" membres"
		if ($TeamMemberCount -gt 80) {
		   $NomProprio=""
		   $TeamPrefix="🌐 "
		   } 
		if ($_.Archived -eq "True") {
		   $TeamPrefix="🗄️ "
		   $NomProprio="[ Archivée ]"
		   } 
		If($Description -eq $TeamName) {$Description=""}
		Write-Progress -Activity "   $Count/$nbTeams $TeamName "
		if ($modeDEBUG) { Write-Host `n╔═══════[$TeamName] $Description -ForegroundColor green}
		$Count++
		Get-TeamChannel -MembershipType Standard -GroupId $GroupId | foreach {
			$ChannelName=$_.DisplayName
			If($ChannelName -ne "general") {
				Write-Progress -Activity "   $Count/$nbTeams $TeamName // $ChannelName"
				#$Description=$_.Description
				if ($modeDEBUG) { Write-Host ╠═══> [$ChannelName] $ChannelMemberCount $Description}
				$Result=@{'ColA'="";'ColB'=$ChannelName;}
			} else {
				$Result=@{'ColA'=$TeamPrefix+$TeamName;'ColB'=$NomProprio;}
			}
			$Results= New-Object psobject -Property $Result
			$Results | select 'ColA','ColB' | Export-Csv $Path -NoTypeInformation -Append
			$compteurLignes++
		} 
		$canauxprives=Get-TeamChannel -MembershipType Private -GroupId $GroupId
		$canauxprivescount=$canauxprives.Count
		If($canauxprivescount -eq "1") {
			if ($modeDEBUG) { Write-Host ╠═══> +$canauxprivescount canal privé }
			$Result=@{'ColA'="";'ColB'="+1 canal privé";'ColC'="";'ColD'="";}
			$Results= New-Object psobject -Property $Result
			$Results | select 'ColA','ColB' | Export-Csv $Path -NoTypeInformation -Append
			$compteurLignes++
		}
		elseIf($canauxprivescount -gt 1) {
			if ($modeDEBUG) { Write-Host ╠═══> +$canauxprivescount canaux privés }
			$Result=@{'ColA'="";'ColB'="+$canauxprivescount canaux privés";}
			$Results= New-Object psobject -Property $Result
			$Results | select 'ColA','ColB' | Export-Csv $Path -NoTypeInformation -Append
			$compteurLignes++
		}
		if ($compteurLignes -eq '1') {
			if ($modeDEBUG) { Write-Host ╠═══> On ajoute une ligne vide }
			$Result=@{'ColA'="";'ColB'=" ";}
			$Results= New-Object psobject -Property $Result
			$Results | select 'ColA','ColB' | Export-Csv $Path -NoTypeInformation -Append
			}
		if ($modeDEBUG) { Write-Host ╚══════════════════════════════════════════════════════ -ForegroundColor green }
	}
	Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName  `n Currently Processing Channel: $ChannelName"  -Completed
	if((Test-Path -Path $Path) -eq "True") 
	{
	Write-Host `nReport available in $Path -ForegroundColor Green
	}
	  
	Disconnect-MicrosoftTeams
	
# Envoyer le fichier TeamsData.csv dans ftp 
	Set-FTPConnection -Credentials $credentialsFTP -Server ftp.xxxx.org -Session FloFTP -ignoreCert -UseBinary -KeepAlive
	$Session = Get-FTPConnection -Session FloFTP
	Get-ChildItem ".\TeamsData.csv" | Add-FTPItem -Session $Session -Path /Teams/ -Overwrite	  
