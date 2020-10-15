cls
$today = (Get-Date)

for ($y=2017; $y -le 2020; $y++)
{
    for ($i=1; $i -le 12; $i++)
    {
        $periodStart = (Get-Date -Year $y -Month $i -Day 01)

        if ($periodStart -lt $today)
        {
            $periodEnd = $periodStart.AddMonths(1).AddDays(-1)
            if ($periodEnd -ge $today)
            {
                $periodEnd = $today.AddDays(-1);
            }

            $startString = $periodStart.ToString("dd/MM/yyyy")
            $endString = $periodEnd.ToString("dd/MM/yyyy")
            $outputFile = "$($periodStart.ToString("yyyy-MM")).csv"

            Write-Host "Fetching readings for period $($startString) to $($endString)"
            Start-Process -Wait -NoNewWindow -FilePath ".\OvoData.exe" -ArgumentList "-u your.login@yourdomain.co.uk -p yourPassword -f $($startString) -t $($endString) -o .\OvoData-$($outputFile)"
            Sleep -Seconds 15
        }
        else
        {
            Break Script
        }
    }
}
