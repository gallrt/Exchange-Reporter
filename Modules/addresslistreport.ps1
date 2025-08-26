#L.Heinz 2020-10-20 v1.0 based on Module dyndgreport

$addresslistreport = Generate-ReportHeader "addresslistreport.png" "$l_addresslist_header"

$cells = @("$l_addresslist_name", "$l_addresslist_hasmember", "$l_addresslist_membercount")
$addresslistreport += Generate-HTMLTable "$l_addresslist_t1header" $cells

$addresslists = Get-Addresslist | Sort-Object Name
foreach ($addresslist in $addresslists) {
	$addresslistname = $addresslist.Name
	$addresslistmembers = Get-Recipient -RecipientPreviewFilter $addresslist.RecipientFilter -OrganizationalUnit $addresslist.RecipientContainer -ResultSize Unlimited | Sort-Object Name
	if ($addresslistmembers) {
		$addresslisthasmember = "$l_addresslist_memberyes"
		$addresslistmemcount = $addresslistmembers.Count

		$memcells = @("Name", "Typ")
		$addresslistmemberreport += Generate-HTMLTable "$l_addresslist_t2header $addresslistname" $memcells

		foreach ($addresslistmember in $addresslistmembers) {
			$memname = $addresslistmember.Name
			$memtyp = $addresslistmember.RecipientType

			$memcells = @("$memname", "$memtyp")
			$addresslistmemberreport += New-HTMLTableLine $memcells
		}
		$addresslistmemberreport += End-HTMLTable
	} else {
		$addresslisthasmember = "$l_addresslist_memberno"
		$addresslistmemcount = "0"
	}
	$cells = @("$addresslistname", "$addresslisthasmember", "$addresslistmemcount")
	$addresslistreport += New-HTMLTableLine $cells
}

$addresslistreport += End-HTMLTable

$addresslistreport += $addresslistmemberreport

$addresslistreport | Set-Content "$tmpdir\addresslistreport.html"
$addresslistreport | Add-Content "$tmpdir\report.html"