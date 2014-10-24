[reflection.assembly]::LoadFrom("C:\Scratch\Projets\Tracking\Scripts\TrackingObject.dll")
$TC = new-object TrackingObject.TrackingCol
$Today =  [DateTime]::now
$TC.Start=$Today.AddMonths(-1).ToShortDateString()
$TC.End=$Today.ToShortDateString()
$TC.Extract()
$TC.count

$Categories = @("CAL:Avant Vente","CAL:Avant Vente (OOF)","CAL:Avant Vente (ConfCall)")

foreach ($rdv in $TC)
{
#	Write-Host("Found " + $rdv.Subject)
	foreach ($category in $rdv.Categories)
	{
		$CatMatch = $false
		foreach ($c in $Categories)
		{
			if ($category -eq $c) { $CatMatch = $true}
		}
		if ($CatMatch)
		{
#			Write-Host(" Match 1 Catégorie Business")
			if (!($rdv.Subject.StartsWith("STU")))
			{
				Write-Host("Update trigered : " + $rdv.Subject)
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
				$ns = "STUB-" + $ClientName + " " + $rdv.Subject
				Write-Host("Nouveau Sujet : $ns")
				Write-Host("")
				$rdv.Subject = $ns
				$rdv.Update()
			}
		}	
	}
}