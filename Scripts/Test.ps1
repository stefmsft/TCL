[reflection.assembly]::LoadWithPartialName("TrackingObject")
[appdomain]::currentdomain.GetAssemblies()
$TC = new-object TrackingObject.TrackingCol
$Today =  [DateTime]::now
$TC.End=$Today.ToShortDateString()
$TC.Start="27/03/2007"
$TC.Start=$Today.AddMonths(-1).ToShortDateString()
$TC.Extract()
$TC.count
$TC | foreach { $_ | Format-Table Start,End,Subject -hideTableHeaders}
$TC.Categories.Keys
$TC.Categories.keys | foreach { $_}
$TC.Categories.values | ForEach-Object  { $_.Cumul }
$TC.Categories.values | ForEach-Object  { $_.Occurence }
$TC.Categories.values | ForEach-Object  { $_.Labels }