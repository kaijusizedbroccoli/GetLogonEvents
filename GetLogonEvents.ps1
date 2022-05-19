function Get-LogonEvents {
    param (
        [string[]]$ComputerName,
        [int]$PastNumberOfDays,
        [string]$OutputPath,
        [switch]$OutputExcel
    )

    Import-Module ImportExcel

    foreach ($Computer in $ComputerName) {

        $cleanedResults = [System.Collections.ArrayList]@()
        $results        = [System.Collections.ArrayList]@()
        
        $filterTable = @{
            LogName = 'System'
            StartTime = $((Get-Date).AddDays(-$PastNumberOfDays))
            ID = '7001','7002'
        }
    
        $winEventLogs = Get-WinEvent -ComputerName $Computer -FilterHashtable $filterTable | Where-Object {$_.LevelDisplayName -like "Information"}
    
        # ------- REGION START: Get Logs ------- #
        foreach ($Log in $winEventLogs) {
            
            $type = switch ($Log.Id) {
                7001 { "Logon" }
                7002 { "Logoff" }
                Default { Continue }
            }
    
            $userName = [System.Security.Principal.SecurityIdentifier]::new($Log.Properties[1].Value).Translate([System.Security.Principal.NTAccount])
    
            if($userName -match "^.*\\.*$"){
    
                $userString = $userName -replace "^.*\\", ""
                $adUser = (Get-ADUser $userString).Name
    
                $string = "{0},{1},{2},{3}" -f $adUser,$type,$log.TimeCreated.ToLongDateString(),$log.TimeCreated.ToShortTimeString()
                $results.Add($string) | Out-Null
            } 
            else {
                Continue
            }
    
        }
        # ------- REGION END: Robocopy Cmd ------- #

        # ------- REGION START: Process Data ------- #
    
        # $results INDEX KEY: 0 = User, 1 = Type, 2 = Date, 3 = Time
    
        $uniqueUsers = $results | ForEach-Object { $_.Split(",")[0] } | Sort-Object -Unique
    
        foreach ($user in $uniqueUsers) {
            
            [System.Collections.ArrayList]$uniqueUserEvents = $results -match $user
            [System.Collections.ArrayList]$logonEvents      = $uniqueUserEvents -match "Logon"
            [System.Collections.ArrayList]$logoffEvents     = $uniqueUserEvents -match "Logoff"
    
            if($logonEvents.Count -ne $logoffEvents.Count){
                $logonEvents.Remove($logonEvents[0])
            }
    
            for ($i = 0; $i -lt $logonEvents.Count; $i++){
              
                if ($logonEvents[$i].Date -eq $logoffEvents[$i].Date) {
    
                    $timeSpan  = New-TimeSpan -Start $($logonEvents[$i].Split(",")[3]) -End $($logoffEvents[$i].Split(",")[3])
                    $hourSpan   = [string]$([math]::Round($timeSpan.Hours))
                    $minuteSpan = [string]$([math]::Round($timeSpan.Minutes))

                    $obj = [PSCustomObject]@{
                        User = $user
                        Date = [DateTime]($logonEvents[$i].Split(",")[2])
                        TimePeriod = "{0} - {1}" -f $logonEvents[$i].Split(",")[3],$logoffEvents[$i].Split(",")[3]
                        LengthOfTime = "{0} Hours {1} Minutes" -f $hourSpan,$minuteSpan
                    }
                    $cleanedResults.Add($obj) | Out-Null
                }
            }
        }
        # ------- REGION END: Process Data ------- #

        # ------- REGION START: Output Excel ------- #
        if($OutputExcel){
            $cleanedResults | 
            Sort-Object -Property Date -Descending | 
            Export-Excel -Path "$OutputPath\$Computer.xlsx" -WorksheetName "$Computer" -TableName DetailedUsage -AutoSize

            $AdditionalStatsObj |
            Export-Excel -Path "$OutputPath\$Computer.xlsx" -WorksheetName "$Computer Additional Stats" -TableName AdditionalStats -AutoSize

        } else {
            Write-Output ($cleanedResults | Sort-Object -Property Date -Descending)
            Write-Output $AdditionalStatsObj
        }
        # ------- REGION END: Output Excel ------- #
        
    }
    
}