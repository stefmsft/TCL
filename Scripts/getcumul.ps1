[reflection.assembly]::LoadWithPartialName("TrackingObject") > $null
$TC = new-object TrackingObject.TrackingCol
$Today =  [DateTime]::now
$TC.End=$Today.ToShortDateString()
$TC.Start=$Today.AddMonths(-1).ToShortDateString()
$TC.Extract() > $null
$TC.Categories["mssa-Activity"] | ForEach-Object {$_.Cumul} | sort
