<#
.SYNOPSIS

    Use to pull leaderboard data off mwomercs.com

.DESCRIPTION
	Parses data from mwomercs.com leaderboards. This script
    should be expected to take a long time as it has to go parse multiple
    pages.

.PARAMETER global
Only pulls global data instead of all classes. 
#>


[cmdletbinding()]
param (
    [switch]$global,

    [Parameter(Mandatory=$True)]
    [string]$season
)


function ParseTable {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.HtmlWebResponseObject] $WebRequest
    )

    #Extract the tables out of the web request
    try {
        $tables = @($WebRequest.ParsedHtml.getElementsByTagName("table"))
    }
    catch{
        Write-Error "An Error was encountered while trying to pull element by tag `
        name table."
    }
    $table = $tables[0]
    $titles = @()
    $rows = @($table.Rows)

    #Go through all of the rows in the table
    foreach($row in $rows){
        $cells = @($row.Cells)
        if($cells[0].tagName -eq "th"){
            $titles = @($cells | % { ("" + $_.InnerText).Trim() })
            continue
        }

        #create titles if not found
        if(-not $titles){
            $titles = @(1..($cells.Count + 2) | % { "P$_" })
        }

        #Now go through the cells in the the row. For each, try to find the
        #title that represents that column and create a hashtable mapping those
        #titles to content
        $resultObject = [Ordered] @{}
        for($counter = 0; $counter -lt $cells.Count; $counter++){
            $title = $titles[$counter]
            if(-not $title) { continue }
            $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
        }

        #And finally cast that hashtable to a PSCustomObject
        [PSCustomObject] $resultObject
    }
}



#Set SSL to updated version
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Configuration varriables
$username = Read-Host "MWO Username (email)?"
$passwordString = read-host -AsSecureString "Password?"
#subtract 1 due to MWO leaderboard format for season
$season = $season - 1
$Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordString))
$loginUrl = "https://mwomercs.com/profile/leaderboards"
$savepath = [Environment]::GetFolderPath("MyDocuments")
$ErrorCount = 0
$leaderboards =@{
    "Global" = 0
    "Light"  = 1
    "Medium" = 2
    "Heavy"  = 3
    "Assualt"= 4
}

if ($global){
    $leaderboards =@{
        "Global" = 0
    }
}


#Pull web object

$r = Invoke-WebRequest -Uri('https://mwomercs.com/do/login') -SessionVariable mwo
$form = $r.Forms
$form.fields['email'] = $username
$form.fields['password'] = $password

#set cookies
$sortcookie = New-Object System.Net.Cookie   
$sortcookie.Name = "leaderboard__rank_by"
$sortcookie.Value = "0"
$sortcookie.Domain = ".mwomercs.com"
$mwo.Cookies.Add($sortcookie);

$seasoncookie = New-Object System.Net.Cookie   
$seasoncookie.Name = "leaderboard_season"
$seasoncookie.Value = "$season"
$seasoncookie.Domain = ".mwomercs.com"
$mwo.Cookies.Add($seasoncookie);


#Submit loginform
$r=Invoke-WebRequest -Uri ('https://mwomercs.com/do/login') -WebSession $mwo -Method POST -Body $form.Fields
sleep 5

#pull leaderboards 
foreach ($leaderboard in $leaderboards.GetEnumerator()){
    $page= 0
    $rawtables = @()
    $leaderboardpage = $null
    while ($leaderboardpage.Content -notlike "*No results found.*"){
        do{
            $ParseFail = $null
            $progressPreference = 'silentlyContinue'
            $leaderboardpage=Invoke-WebRequest -Uri ("https://mwomercs.com/profile/leaderboards?page=$page&type=$($leaderboard.value)") -WebSession $mwo
            $progressPreference = 'Continue'
            if ($leaderboardpage.Content -notlike "*No results found.*"){
                Try {
                        ParseTable $leaderboardpage -ErrorAction Stop | `
                        Export-Csv "$savepath\$($leaderboard.name +"_"+ $seasonquestion).csv" -NoTypeInformation -Append
                    }
                catch {
                    Write-Warning "Error encountered during parse. Retrying..."
                    $ParseFail = $true
                    $ErrorCount++
                    if ($ErrorCount -ge 1){
                        Write-Warning "Retry number $ErrorCount."
                    }
                    Start-Sleep -Seconds 30
                    }
            }
        }until ((!$ParseFail) -or ($ErrorCount -ge 30))
        if (($ParseFail) -and ($ErrorCount -ge 30)){
                Write-Error "Parse error retry exceeded"
                exit
            }
        $page++
        Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Status "Page: $page"
    }
    Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Completed
    Write-Output "$($leaderboard.name) saved to $savepath\$($leaderboard.name +"_"+ $seasonquestion).csv"  
}
