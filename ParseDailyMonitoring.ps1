<#
    NAME: ParseDailyMonitoring.ps1
    DESCRIPTION: Parse Daily Monitoring txt files for SQL input
    AUTHOR: Austin Vargason
    DATE MODIFIED: 10/14/18
#>


function Get-ParsedDailyReport() {
    param (
        [Parameter(Mandatory=$true)]
        [PsObject[]]$fileInput
    )

    <#

        Col1: Client
        Col2: Data Written GB
        Col3: # Files
        Col4: # Objects
        Col5: # Completed DA
        Col6: # Failed DA
        Col7: # Running DA
        Col8: # Pending DA
        Col9: Success

        HEADER:
        Report Selection Parameters    Value                                                                                           
        ________________________________________________________________________________________________________________________________
        Backup specification(s)        IDB IDB,DECOM_Final_Backup,COSD_Plano_Tulsa_Cutover,COSD_OS_Full,COSD_Data,COSD_Config_X,ADHOC_U
        Timeframe                      24 24                                                                                           

        Client                  Data Written [GB]    # Files      # Objects # Completed DA # Failed DA # Running DA # Pending DA Success
        _________________________________________________________________________________________________________________________________

     #>
    
    #save the header
    $header = $fileInput[0..11]

    #get the date from the header
    $date = $header[3]

    #remove the beginning of Date row
    $date = $date.Remove(0, 15)

    #create a result array
    $resultArray = @()

    #skip the header and read the rows of the text
    for ($i = 12; $i -lt $fileInput.Count; $i++) {

        #get the row
        $row = $fileInput[$i]

        #create an object to represent the row
        $obj = New-Object -TypeName PsObject

        #values array to store the read values
        $values = @()

        #parse the items in the column
        foreach ($item in $row.Split()) {
            
            if ( $item.Length -ne 0 ) {
                
                if ($item -like "*.cosd.*") {
                    $item = $item.Remove($item.IndexOf("."))
                }

                #add to the values array
                $values += $item
            }
        }

        #add the values to the object
        $obj | Add-Member -Name Client -Value $values[0] -MemberType NoteProperty
        $obj | Add-Member -Name DataWrittenGB -Value $values[1] -MemberType NoteProperty
        $obj | Add-Member -Name FileCount -Value $values[2] -MemberType NoteProperty
        $obj | Add-Member -Name ObjectCount -Value $values[3] -MemberType NoteProperty
        $obj | Add-Member -Name CompletedDA -Value $values[4] -MemberType NoteProperty
        $obj | Add-Member -Name FailedDA -Value $values[5] -MemberType NoteProperty
        $obj | Add-Member -Name RunningDA -Value $values[6] -MemberType NoteProperty
        $obj | Add-Member -Name PendingDA -Value $values[7] -MemberType NoteProperty
        $obj | Add-Member -Name Success -Value $values[8] -MemberType NoteProperty
        $obj | Add-Member -Name Run_Date -Value $date -MemberType NoteProperty

        #add to the result array
        $resultArray += $obj
    }

    #return the resultArray
    return $resultArray
}


$emailTexts = Get-ChildItem -Path .\EmailTexts | Where {$_.Name -like "*.txt"}
$count = $emailTexts.Count
$i = 0

foreach ($file in $emailTexts) {
    $output = Get-ParsedDailyReport -fileInput (Get-Content -Path $file.FullName)

    $timeRan = $output.Run_Date | Select -First 1 | Out-String

    try {
        $timeOut = [DateTime]::Parse($timeRan)

        $timeOut = $timeOut.ToString("MMddyy-HHmm")

        if ($timeOut.Contains("0800") -or $timeOut.Contains("1700") -or $timeOut.Contains("2300")) {
        
            $fileNameSplit = $file.Name.Split("_")

            $serverInName = $fileNameSplit[0]

            $fileOutputName = ".\output\BUR_Summary_$serverInName" + "_$timeOut.csv"

            $output | Export-Csv -Path $fileOutputName -NoTypeInformation
        }
    }
    catch {
        Write-Host "Issue with fileName: " $file.Name -ForegroundColor Yellow
    }

    $i++
    Write-Progress -Activity "Getting BUR output" -Status "Completed output for $fileOutputName" -PercentComplete (($i / $count) * 100)
}



$files = Get-ChildItem -Path .\output

$hash = @{}


foreach ($file in $files) {
    $date = $file.Name.Substring($file.Name.Length - 15).Replace(".csv","")
    $data = Import-Csv -Path $file.FullName
    $numFailed = $data | Where {$_.FailedDA -gt 0}
    $totalClients = $data.Count
    $numFailed = $numFailed.Count

    if ($numFailed -eq $null ) {
        $numFailed = 0
    }

    if ($hash.ContainsKey($date)) {


        $hash[$date].TotalClients = $totalClients + $hash[$date].TotalClients
        $hash[$date].TotalClientsSuccessful = ($totalClients - $numFailed) + $hash[$date].TotalClientsSuccessful
        $hash[$date].TotalClientsFailed = $numFailed + $hash[$date].TotalClientsFailed
        $hash[$date].FilesParsed += 1
    }
    else {
        $obj = New-Object -TypeName PsObject

        $obj | Add-Member -Name "TotalClients" -Value $totalClients -MemberType NoteProperty
        $obj | Add-Member -Name "TotalClientsSuccessful" -Value ($totalClients - $numFailed) -MemberType NoteProperty
        $obj | Add-Member -Name "TotalClientsFailed" -Value $numFailed -MemberType NoteProperty
        $obj | Add-Member -Name "FilesParsed" -Value 1 -MemberType NoteProperty

        $hash.Add($date, $obj)

        Write-Host $numFailed,$hash[$date].TotalClientsFailed,$file.Name
    }

}

$resArray = @()

foreach ($key in $hash.Keys) {
    $obj = New-Object -TypeName PsObject

    $obj | Add-Member -Name "DateTime" -Value $key -MemberType NoteProperty
    $obj | Add-Member -Name "TotalClients" -Value ($hash[$key].TotalClients) -MemberType NoteProperty
    $obj | Add-Member -Name "TotalClientsSuccessful" -Value $hash[$key].TotalClientsSuccessful -MemberType NoteProperty
    $obj | Add-Member -Name "TotalClientsFailed" -Value $hash[$key].TotalClientsFailed -MemberType NoteProperty
    $obj | Add-Member -Name "FilesParsed" -Value $hash[$key].FilesParsed -MemberType NoteProperty
    $obj | Add-Member -Name "DailySuccessAverage" -Value ([math]::Round(((($hash[$key].TotalClientsSuccessful) / $hash[$key].TotalClients) * 100), 2)) -MemberType NoteProperty

    $resArray += $obj
}


foreach ($row in $resArray) {
    $row.DateTime = "$($row.DateTime.Substring(0,2))/$($row.DateTime.Substring(2,2))/$($row.DateTime.Substring(4,2))"
}

$resGroups = $resArray | Group-Object -Property DateTime

$newOutput = @()

foreach ($group in $resGroups) {

    $newOutput += $group.Group | Sort-Object -Property DailySuccessAverage -Descending | Select -First 1
}


$newOutput | Export-Excel -Path .\outputSummary.xlsx -WorkSheetname "DailySuccessAVG" -TableName "DateSuccessAVG" -Show