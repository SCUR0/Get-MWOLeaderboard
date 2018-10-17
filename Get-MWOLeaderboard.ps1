<#
.SYNOPSIS

    Use to pull leaderboard data off mwomercs.com

.DESCRIPTION
	Parses data from mwomercs.com leaderboards. This script
    should be expected to take a long time as it has to go parse multiple
    pages.

.PARAMETER global
Only pulls global data instead of all classes.

.PARAMETER season
The season that you would like to parse.

.PARAMETER savepath
The location you want to save. Script will dynamically default to documents folder if parameter is not used.
#>


[cmdletbinding()]
param (
    [switch]$global,
    [Parameter(Mandatory=$True)]
    [string]$season,
    $savepath = [Environment]::GetFolderPath("MyDocuments")
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
            $titles = @(
                foreach($cell in $cells){
                    ("" + $cell.InnerText).Trim() 
                }
            )
            continue
        }

        #create titles if not found
        if(-not $titles){
            $titles = @(
                foreach ($cell in (1..($cells.Count + 2))){
                     "P$cell" 
                }
            )
        }

        <#Now go through the cells in the the row. For each, try to find the
        title that represents that column and create a hash#>
        $resultObject = [Ordered] @{}
        for($counter = 0; $counter -lt $cells.Count; $counter++){
            $title = $titles[$counter]
            if(-not $title) { continue }
            $resultObject[$title] = ("" + $cells[$counter].InnerHTML
        }

        #hashtable to PSCustomObject
        [PSCustomObject] $resultObject
    }
}

#Load Parallel Script/Module
$IParallelLocation="C:\Program Files\WindowsPowerShell\Modules\Invoke-Parallel\Invoke-Parallel.ps1"
try{
    . $IParallelLocation
}catch{
    Write-Error "Invoke-Parallel script is required. Download at https://github.com/RamblingCookieMonster/Invoke-Parallel."
    Write-Output "Place Invoke-Parallel script at $IParallelLocation"
    pause
    exit
}

#Set SSL to updated version
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Configuration varriables
$username = Read-Host "MWO Username (email)?"
$passwordString = read-host -AsSecureString "Password?"
#Respect Original Progress Preference
$OriginalProgressPreference=$ProgressPreference
#subtract 1 due to MWO leaderboard format for season
$seasonquery = $season - 1
$Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordString))
$loginUrl = "https://mwomercs.com/profile/leaderboards"
$ErrorCount = 0
$leaderboards =@{
    "Global" = 0
    "Light"  = 1
    "Medium" = 2
    "Heavy"  = 3
    "Assault"= 4
}

if ($global){
    $leaderboards =@{
        "Global" = 0
    }
}

#checks to make sure documents path is clear for parse
$ClassArray=@(
    "Global_",
    "Light_",
    "Medium_",
    "Heavy_"
    "Assault_"
)
Foreach ($Class in $ClassArray){
    $ClassDocumentPath ="$($savepath)\$($Class)$season.csv"
    If (Test-Path $ClassDocumentPath){
        Write-Error "$class$season file already exists!"
        Write-Output "Remove file at $ClassDocumentPath and restart script."
        pause
        exit
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
$seasoncookie.Value = "$seasonquery"
$seasoncookie.Domain = ".mwomercs.com"
$mwo.Cookies.Add($seasoncookie);


#Submit loginform
$r=Invoke-WebRequest -Uri ('https://mwomercs.com/do/login') -WebSession $mwo -Method POST -Body $form.Fields
Write-Host "Parse request initialized." 
start-sleep 3

#pull leaderboards 

$Leaderboards.GetEnumerator() | Invoke-Parallel -ImportVariables -ImportFunctions -ScriptBlock {
    $page= 0
    $rawtables = @()
    $leaderboardpage = $null
    write-host "Parsing $($_.name)..."
    while ($leaderboardpage.Content -notlike "*No results found.*"){
        do{
            $ParseFail = $null
            $ProgressPreference = 'silentlyContinue'
            $leaderboardpage=Invoke-WebRequest -Uri ("https://mwomercs.com/profile/leaderboards?page=$page&type=$($_.value)") -WebSession $mwo
            $ProgressPreference = $OriginalProgressPreference
            if ($leaderboardpage.Content -notlike "*No results found.*"){
                Try {
                        ParseTable $leaderboardpage -ErrorAction Stop | `
                        Export-Csv "$savepath\$($_.name +"_"+ $season).csv" -Delimiter "`t" -NoTypeInformation -Append
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
        }until ((!$ParseFail) -or ($ErrorCount -ge 5))
        if (($ParseFail) -and ($ErrorCount -ge 5)){
                Write-Error "Parse error retry exceeded"
                pause
                exit
            }
        $page++
        #Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Status "Page: $page"
    }
    #Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Completed
    Write-Output "$($_.name) saved to $savepath\$($_.name +"_"+ $season).csv"  
}
