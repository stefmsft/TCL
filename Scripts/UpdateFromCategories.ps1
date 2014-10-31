# Injection de la dll .Net dans l'espace de travail Powershell
##
[reflection.assembly]::LoadFrom("C:\Scratch\Projets\Tracking\Scripts\TrackingObject.dll")

#Instanciation de l'object TrackingCol
##
$TC = new-object TrackingObject.TrackingCol

#Positionnement des Property de debut et fin de recherche des rendez vous
##
$Today=[DateTime]::now
$TC.Start=$Today.AddMonths(-1).ToShortDateString()
$TC.End=$Today.ToShortDateString()

#Appel de la method d'extraction des Rendez vous depuis Outlook
##
$TC.Extract()

#Affichage du nombre de Rendez vous retourné
##
$TC.count

#Definition des catégorie correspondant à votre recherche parmi le "set" de rendez-vous retourné
## CHANGEZ CETTE LISTE AVEC VOTRE/VOS CATEGORIE
##
$Categories = @("CAL:Avant Vente","CAL:Avant Vente (OOF)","CAL:Avant Vente (ConfCall)")

#Pour chaque rendez vous
##
foreach ($rdv in $TC)
{
#   Pour chaque Catégorie définies dans un rendez-vous
##
	foreach ($category in $rdv.Categories)
	{
		$CatMatch = $false
		#Pour chaque catégorie qui m'interresse, je verifie si elle presente dans le Rendez-vous
		##
		foreach ($c in $Categories)
		{
			if ($category -eq $c) { $CatMatch = $true}
		}
		if ($CatMatch)
		{
		#Si le Rendez-vous est du type de ceux qui m'interresse
		##
			if (!($rdv.Subject.StartsWith("STU")))
			{
			#Si un code "STU" n'est pas deja présent dans le Rendez-vous
			##
				Write-Host("Update trigered : " + $rdv.Subject)

				#La logique qui suis implemente la recherche d'un nom de client qui s'il existe dans le sujet devrait etre ecrit en majuscule
				#vous n'etes pas obligé de suivre cette convention ou meme de voir renseigner un nom de client
				##
				$res = $rdv.Subject | Select-String -Casesensitive -AllMatches '([A-Z])+' | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
				if ($res -ne $null)
				{
					$longuest=0
					$ClientName = "CUSTOMER"
					foreach ($s in $res)
					{
						if ($s.Length -gt $longuest)
						{
							$ClientName = $s
							$longuest = $s.length
						}
					}
				}

				#Creation de la nouvelle chaine correspondant au sujet pour mise à jour
				##
				$ns = "STUB-" + $ClientName + " " + $rdv.Subject
				Write-Host("Nouveau Sujet : $ns")
				Write-Host("")
				$rdv.Subject = $ns

				#Cette commande va mettre à jour l'objet dans Outlook
				## Decommentez la ligne lorsque vous etes satisfait de la detection et de la modification
				##
#				$rdv.Update()
			}
		}	
	}
}