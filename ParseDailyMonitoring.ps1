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
    
    #create variables for header and the date
    $header = ""
    $date = ""
    $startRow = 0

    if ($fileInput[0][0] -eq "#") {
        $header = $fileInput[0..8]
        $date = $header[2]
        $startRow = 9
    }
    else {
        #save the header
        $header = $fileInput[0..11]

        #get the date from the header
        $date = $header[3]
        $startRow = 12
    }

    #remove the beginning of Date row
    $date = $date.Remove(0, 15)

    $date = $date.Trim()

    #create a result array
    $resultArray = @()

    #skip the header and read the rows of the text
    for ($i = $startRow; $i -lt $fileInput.Count; $i++) {

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

# function to move the archive files to archive directory and remove old report data
function Set-ArchiveOutputFiles {

    # get the output files in the output directory
    $outputFiles = Get-ChildItem -Path .\output | Where-Object {$_.Name -like "*.csv"}

    # get the month we are archiving for
    $archiveMonth = (Get-Date (Get-Date).AddMonths(-2) -format "MM")

    # remove the old report data
    if (Test-Path -Path .\outputSummary.xlsx) {
        Remove-Item -Path .\outputSummary.xlsx
    }

    # create the archive month folder under archive
    if (!(Test-Path -Path .\archive\$archiveMonth)) {
        $null = New-Item -ItemType Directory -Path .\archive\$archiveMonth 
    }

    # move the output files to the created archive folder
    foreach ($file in $outputFiles) {

        Copy-Item -Path $file.FullName -Destination .\archive\$archiveMonth
        Remove-Item $file.FullName
    }
}

# function to produce output files from parsed text files
function Get-OutputFiles {

    # set the email texts to the emailed text files downloaded to the specific folder
    $emailTexts = Get-ChildItem -Path .\EmailTexts | Where-Object {$_.Name -like "*.txt"}

    # save the count to help with iterating performance
    $count = $emailTexts.Count

    # initialize a counter at 0
    $i = 0

    # loop through the emailed files
    foreach ($file in $emailTexts) {

        # get the output based on the imported text file content
        $output = Get-ParsedDailyReport -fileInput (Get-Content -Path $file.FullName)

        # get the run date from the output result
        $timeRan = $output.Run_Date | Select-Object -First 1 | Out-String

        # try catch for summary parsing
        try {
            # parse timeout into a datetime
            $timeOut = [DateTime]::Parse($timeRan)

            # specify a datetime string format
            $timeOut = $timeOut.ToString("MMddyy-HHmm")

            # filter to make sure we are getting correctly timed files
            if (($timeOut.Contains("0800") -or $timeOut.Contains("1700") -or $timeOut.Contains("2300")) -and $timeOut.StartsWith((Get-Date (Get-Date).AddMonths(-1) -format "MM"))) {
                
                # get the array split based on splitting the file name
                $fileNameSplit = $file.Name.Split("_")

                # get the server in the name from the filename split
                #TODO: This could be condensed into a single line
                $serverInName = $fileNameSplit[0]

                # specify the new file output name exported as a csv
                $fileOutputName = ".\output\BUR_Summary_$serverInName" + "_$timeOut.csv"

                # export the csv without type information
                $output | Export-Csv -Path $fileOutputName -NoTypeInformation
            }
        }
        catch {
            Write-Host "Issue with fileName: " $file.Name -ForegroundColor Yellow
        }

        # increase the counter and write to the progress bar
        $i++
        Write-Progress -Activity "Getting BUR output" -Status "Completed output for $fileOutputName" -PercentComplete (($i / $count) * 100)
    }
}

# function to get the monthly bur report from calling needed functions
function Get-MonthlyBURReport {

    # set the archived output files
    Set-ArchiveOutputFiles

    # get the output files
    Get-OutputFiles
    
    # get the output files in the output folder
    $files = Get-ChildItem -Path .\output

    # initialize a new hash
    $hash = @{}

    # loop through the files and summarize
    foreach ($file in $files) {
        # date can be gathered from the filename substring 
        $date = $file.Name.Substring($file.Name.Length - 15).Replace(".csv","")

        # import the spreadsheet
        $data = Import-Csv -Path $file.FullName

        # get the num failed objects
        $numFailed = $data | Where-Object {$_.FailedDA -gt 0}

        # count total variables for average calculations
        $totalClients = $data.Count
        $numFailed = $numFailed.Count

        # null check
        if ($null -eq $numFailed) {
            $numFailed = 0
        }

        # if the hash contains the date key then fill out the hash data accordingly
        # else create a new object and add to the hash representation
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

    # intialize an empty array
    $resArray = @()

    # loop through the hash to fill out the summary array with custom members
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

    # update datetime of the row data
    foreach ($row in $resArray) {
        $row.DateTime = "$($row.DateTime.Substring(0,2))/$($row.DateTime.Substring(2,2))/$($row.DateTime.Substring(4,2))"
    }

    # group the data
    $resGroups = $resArray | Group-Object -Property DateTime

    # initialize a new empty array
    $newOutput = @()

    # add each group to the output
    foreach ($group in $resGroups) {

        $newOutput += $group.Group | Sort-Object -Property DailySuccessAverage -Descending | Select-Object -First 1
    }

    # output the new output
    Write-Output $newOutput
}

# final funciton calls
Get-MonthlyBURReport | Export-Excel -Path .\outputSummary.xlsx -WorkSheetname "DailySuccessAVG" -TableName "DateSuccessAVG" -Show