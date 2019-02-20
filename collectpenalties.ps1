$start = Get-Date;
$strPath = $PSScriptRoot;

$numberColumn = 1;
$nameColumn = 2;
$penaltyColumn = 21;
$jamsColumn = 23;

$skaterPenalties = @{}
$skaterJams = @{}
$statbooks = gci $strPath | Where-Object {$_.FullName -like "*xlsx*"};

$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false

foreach ($statbook in $statbooks)
{
	$workbook = $objExcel.Workbooks.Open($statbook.FullName)
	$worksheet = $workbook.sheets.item("Penalty Summary")
	
	if ($worksheet.cells.item(2,1).value2.ToLower().Contains("steel"))
	{
		$startRow = 4;
	}
	else
	{
		$startRow = 33;
	}
	
	for ($row = $startRow; $row -le $startRow + 19; $row++)
	{
		$skaterNumber = $worksheet.cells.item($row, $numberColumn).value2.ToString().Trim();
        $skaterName = $worksheet.cells.item($row, $nameColumn).value2
		$penalties = $worksheet.cells.item($row, $penaltyColumn).value2
        $jams = $worksheet.cells.item($row, $jamsColumn).value2

        # Gotta clean up the skater number
        $skaterNumber = $skaterNumber -replace '\*','';
        
        #correct for Red
        if ($skaterNumber.ToLower().Contains("v3"))
        {
            $skaterNumber = [string]3;
        }

        if (($skaterNumber -eq $null) -or ($skaterNumber -eq ""))
        {
            continue;
        }

        Write-Host "In bout $statbook, skater: $skaterName, #$skaterNumber got $penalties major penalties across $jams jams";
		
		if ($skaterPenalties.ContainsKey($skaterNumber))
		{
			$curPenaltyInfo = $skaterPenalties.Get_Item($skaterNumber);
			$curPenaltyInfo.Penalties = $curPenaltyInfo.Penalties + $penalties;
            $curPenaltyInfo.Jams = $curPenaltyInfo.Jams + $jams;
		}
		else
		{
            $penaltyInfoProps = @{'Penalties'=$penalties; 'Jams'=$jams};
            $penaltyInfo = New-Object -TypeName PSObject -Prop $penaltyInfoProps;
			$skaterPenalties.Add($skaterNumber, $penaltyInfo)
		}
	}

    $workbook.close($False);
    
}

$objExcel.quit();
$skaterPenalties.GetEnumerator() | Select-Object -Property Name, @{Name="Penalties";Expression={$_.Value.Penalties}}, @{Name="Jams";Expression={$_.Value.Jams}}, @{Name="Jam / Pen";Expression={ "{0:N3}" -f ([int] $_.Value.Jams / $_.Value.Penalties)}} | Sort-Object -Property "Jam / Pen" | Format-Table
Write-Host "Time to run: $(((Get-Date) - $start).totalseconds)s"