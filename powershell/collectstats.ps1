$start = Get-Date;
$strPath = $PSScriptRoot;

$numberColumn = 1;
$nameColumn = 2;
$ptsForColumn = 17;
$ptsAgainstColumn = 18;
$jamsColumn = 6;

$skaterPoints = @{}
$skaterJams = @{}
$statbooks = gci $strPath | Where-Object {$_.FullName -like "*xlsx*"};

$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false

foreach ($statbook in $statbooks)
{
	$workbook = $objExcel.Workbooks.Open($statbook.FullName)
	$worksheet = $workbook.sheets.item("Game Summary")
	
	if ($worksheet.cells.item(5,2).value2.ToLower().Contains("steel"))
	{
		$startRow = 6;
	}
	else
	{
		$startRow = 28;
	}
	
	for ($row = $startRow; $row -le $startRow + 19; $row++)
	{
		$skaterNumber = $worksheet.cells.item($row, $numberColumn).value2.ToString().Trim();
        $skaterName = $worksheet.cells.item($row, $nameColumn).value2
		$ptsFor = $worksheet.cells.item($row, $ptsForColumn).value2
        $ptsAgainst = $worksheet.cells.item($row, $ptsAgainstColumn).value2
        $jams = $worksheet.cells.item($row, $jamsColumn).value2

        # Gotta clean up the skater number
        $skaterNumber = $skaterNumber -replace '\*','';
        
        if (($skaterNumber -eq $null) -or ($skaterNumber -eq ""))
        {
            continue;
        }

        if (($ptsFor -le 1) -and ($ptsAgainst -le 1))
        {
            continue;
        }

        Write-Host "In bout $statbook, skater: $skaterName, #$skaterNumber got $ptsFor points for and $ptsAgainst points against across $jams jams";
		
		if ($skaterPoints.ContainsKey($skaterNumber))
		{
			$curPointsInfo = $skaterPoints.Get_Item($skaterNumber);
			$curPointsInfo.PtsFor = $curPointsInfo.PtsFor + $ptsFor;
			$curPointsInfo.PtsAgainst = $curPointsInfo.PtsAgainst + $ptsFor;
            $curPointsInfo.Jams = $curPointsInfo.Jams + $jams;
		}
		else
		{
            $pointsInfoProps = @{'PtsFor'=$ptsFor; 'PtsAgainst'=$ptsAgainst; 'Jams'=$jams};
            $pointsInfo = New-Object -TypeName PSObject -Prop $pointsInfoProps;
			$skaterPoints.Add($skaterNumber, $pointsInfo)
		}
	}

    $workbook.close($False);    
}

$objExcel.quit();
$skaterPoints.GetEnumerator() | Select-Object -Property Name, @{Name="Points For";Expression={$_.Value.PtsFor}}, @{Name="Points Against";Expression={$_.Value.PtsAgainst}}, @{Name="Jams";Expression={$_.Value.Jams}}, @{Name="PMPJ";Expression={ "{0:N3}" -f ([int] ($_.Value.PtsFor - $_.Value.PtsAgainst) / $_.Value.Jams)}} | Sort-Object -Property "PMPJ" | Format-Table
Write-Host "Time to run: $(((Get-Date) - $start).totalseconds)s"