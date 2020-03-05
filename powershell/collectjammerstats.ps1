$start = Get-Date;
$strPath = $PSScriptRoot;

$jammerPoints = @{}
$jammerBouts = @{}
$startRow = 4;
$statbooks = gci $strPath | Where-Object {$_.FullName -like "*xlsx*"};

$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false

foreach ($statbook in $statbooks)
{
	$workbook = $objExcel.Workbooks.Open($statbook.FullName)
	$worksheet = $workbook.sheets.item("Score")
	
	if ($worksheet.cells.item(1,1).value2.ToLower().Contains("steel"))
	{
		$numberColumn = 2;
        $scoreColumn = 17;
	}
	else
	{
		$numberColumn = 21;
        $scoreColumn = 36;
	}

    $priorJammer = 0;
	
	for ($row = $startRow; $row -le $startRow + 79; $row++)
	{
		$jammerNumber = $worksheet.cells.item($row, $numberColumn).value2;
		$ptsFor = $worksheet.cells.item($row, $scoreColumn).value2
        
        if (($jammerNumber -eq $null) -or ($jammerNumber -eq "") -or ($jammerNumber.ToString().Trim().ToLower() -match '\D'))
        {
            #figure out if we're star passing
            $spValue = $worksheet.cells.item($row, $numberColumn - 1).value2;
            if ($spValue -ne $null)
            {
                if ($spValue.ToString().Trim().ToLower().Contains("sp*"))
                {
                    $jammerNumber = $priorJammer;
                }
                else
                {
                    continue;
                }
            }
            else
            {
                continue;
            }
        }

        # Gotta clean up the jammer number
        $jammerNumber = $jammerNumber.ToString().Trim() -replace '\*','';        
        $priorJammer = $jammerNumber;

        Write-Host "In bout $statbook, jam # $row, jammer: #$jammerNumber got $ptsFor points";
		
		if ($jammerPoints.ContainsKey($jammerNumber))
		{
			$curPointsInfo = $jammerPoints.Get_Item($jammerNumber);
            $curPointsInfo.PtsFor = $curPointsInfo.PtsFor + $ptsFor;
            $curPointsInfo.Jams++;
            if ($ptsFor -gt $curPointsInfo.GreatestJamPoints)
            {
                $curPointsInfo.GreatestJamPoints = $ptsFor;
            }
		}
		else
		{
            $pointsInfoProps = @{'PtsFor'=$ptsFor; 'GreatestJamPoints'=$ptsFor; 'Jams'=1};
            $pointsInfo = New-Object -TypeName PSObject -Prop $pointsInfoProps;
			$jammerPoints.Add($jammerNumber, $pointsInfo)
		}
	}

    ForEach ($jammer in $jammerPoints.GetEnumerator())
    {
        if ($jammerBouts.ContainsKey($jammer.Name))
        {
            $curBoutInfo = $jammerBouts.Get_Item($jammer.Name);
            $curBoutInfo.TotalJams = $curBoutInfo.TotalJams + $jammer.Value.Jams;
            $curBoutInfo.TotalPoints = $curBoutInfo.TotalPoints + $jammer.Value.PtsFor;

            if ($jammer.Value.PtsFor -gt $curBoutInfo.GreatestBoutPoints)
            {
                $curBoutInfo.GreatestBoutPoints = $jammer.Value.PtsFor;
                $curBoutInfo.GBBout = $statbook;
            }

            if ($jammer.Value.GreatestJamPoints -gt $curBoutInfo.GreatestJamPoints)
            {
                $curBoutInfo.GreatestJamPoints = $jammer.Value.GreatestJamPoints;
                $curBoutInfo.GJBout = $statbook;
            }
        }
        else
        {
            $boutInfoProps = @{'GreatestJamPoints'=$jammer.Value.GreatestJamPoints; 'GreatestBoutPoints'=$jammer.Value.PtsFor; 'GBBout'=$statbook; 'GJBout'=$statbook; 'TotalJams'=$jammer.Value.Jams; 'TotalPoints'=$jammer.Value.PtsFor};
            $boutInfo = New-Object -TypeName PSObject -Prop $boutInfoProps;
            $jammerBouts.Add($jammer.Name, $boutInfo);
        }

        $jammer.Value.PtsFor = 0;
        $jammer.Value.GreatestJamPoints = 0;
        $jammer.Value.Jams = 0;
    }

    $workbook.close($False);    
}

$objExcel.quit();
$jammerBouts.GetEnumerator() | Select-Object -Property Name, @{Name="Total Points"; Expression={$_.Value.TotalPoints}}, @{Name="Total Jams Jammed"; Expression={$_.Value.TotalJams}}, @{Name="Greatest Jam Points";Expression={$_.Value.GreatestJamPoints}}, @{Name="GJ Bout";Expression={$_.Value.GJBout}}, @{Name="Greatest Bout Points";Expression={$_.Value.GreatestBoutPoints}}, @{Name="GB Bout";Expression={$_.Value.GBBout}} | Sort-Object -Property "Greatest Jam Points" | Format-Table
Write-Host "Time to run: $(((Get-Date) - $start).totalseconds)s"