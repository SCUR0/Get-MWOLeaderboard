<#
.SYNOPSIS

    Use to pull leaderboard data off mwomercs.com

.DESCRIPTION
    Parses data from mwomercs.com leaderboards. Due to the nature of having to parse many pages from a website's output,
it should be expected to take a long time. Usually 20+ minutes for all tables running parallel.

.PARAMETER Credential
    MWO email and password to login to mwomercs.com

.PARAMETER Global
    Grab global leaderboards. If no leaderboard parameter specified all will be pulled

.PARAMETER Light
    Grab light leaderboards. If no leaderboard parameter specified all will be pulled

.PARAMETER Medium
    Grab medium leaderboards. If no leaderboard parameter specified all will be pulled

.PARAMETER Heavy
    Grab heavy leaderboards. If no leaderboard parameter specified all will be pulled

.PARAMETER Assault
    Grab assault leaderboards. If no leaderboard parameter specified all will be pulled

.PARAMETER Season
    The season that you would like to parse.

.PARAMETER Threads
    How many parallel parses to run at once. Defaults to 5.

.PARAMETER Savepath
    Output path for parsed CSVs. If not provided the script will use current directory
#>


[cmdletbinding()]
param (
    [switch]$Global,
    [switch]$Light,
    [switch]$Medium,
    [switch]$Heavy,
    [switch]$Assault,
    [System.Management.Automation.PSCredential]$Credential,
    [Parameter(Mandatory=$True)]
    [string]$Season,
    [int]$Threads = 5,
    $SavePath
)


function ParseTable {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.HtmlWebResponseObject] $WebRequest,
        [switch]$FirstParse,
        [Parameter(Mandatory = $true)]
        [System.IO.StreamWriter]$TextStream

    )
    
    $table = @($WebRequest.ParsedHtml.getElementsByTagName("table"))
    if (!$table){
        Write-Error "An Error was encountered while trying to pull element by tag `
        name table."
    }
    $rownum = 0
    #Get all the rows in the tables
    ForEach($row in $table.rows){
        #Treat the first row as a header
        if($rownum -eq 0){
            if ($FirstParse){
                $RowArray = @()
                ForEach($cell in $row.cells){
                    $RowArray += "`"$(($cell.innerHTML))`""
                }
                $TextStream.WriteLine($RowArray -join "`t")
            }           
        }else{
            $RowArray = @()
            ForEach($cell in $row.cells){
                $RowArray += "`"$(($cell.innerHTML))`""
            }
            $TextStream.WriteLine($RowArray -join "`t")
        }
        $rownum++
    }
}

#Load Parallel Script/Module
$IParallelLocation="C:\Program Files\WindowsPowerShell\Modules\Invoke-Parallel\Invoke-Parallel.ps1"
try{
    . $IParallelLocation
}catch{
    Write-Error "Invoke-Parallel script is required. Download at https://github.com/RamblingCookieMonster/Invoke-Parallel."
    Write-Output "Place Invoke-Parallel script at $IParallelLocation"
    exit
}

#Set SSL to updated version
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


if (!$Credential){
    #request password if left out of parameter
    $Credential = Get-Credential -Username "email@domain" -Message "MWOMercs.com login"
}

if ((!$Credential.UserName) -or (!$Credential.GetNetworkCredential().Password)){
    Write-Error "Username and password is required"
    exit
}

#Configuration varriables
#Respect Original Progress Preference
$OriginalProgressPreference=$ProgressPreference
$loginUrl = "https://mwomercs.com/profile/leaderboards"
$ErrorCount = 0
if (!$SavePath){
    $SavePath = (Get-Location).Path
}


if ($Global -or $Light -or $Medium -or $Heavy -or $Assault){
    $Leaderboards = @{}
    if ($Global){
        $Leaderboards['Global']  = 0
    }
    if ($Light){
        $Leaderboards['Light']   = 1
    }
    if ($Medium){
        $Leaderboards['Medium']  = 2
    }
    if ($Heavy){
        $Leaderboards['Heavy']   = 3
    }
    if ($Assault){
        $Leaderboards['Assault'] = 4
    }
}else{
    $Leaderboards =@{
        "Global" = 0
        "Light"  = 1
        "Medium" = 2
        "Heavy"  = 3
        "Assault"= 4
    }
}

#checks to make sure documents path is clear for parse
Foreach ($Class in $Leaderboards.GetEnumerator()){
    $ClassDocumentPath ="$($SavePath)\$($Class.Name)_$Season.csv"
    If (Test-Path $ClassDocumentPath){
        Write-Error "$Class$Season file already exists!"
        Write-Output "Remove file at $ClassDocumentPath and restart script."
        exit
    }
}


#Pull web object

$r = Invoke-WebRequest -Uri('https://mwomercs.com/do/login') -SessionVariable mwo
$form = $r.Forms
$form.fields['email'] = $Credential.UserName
$form.fields['password'] = $Credential.GetNetworkCredential().Password

#set cookies
$sortcookie = New-Object System.Net.Cookie   
$sortcookie.Name = "leaderboard__rank_by"
$sortcookie.Value = "0"
$sortcookie.Domain = ".mwomercs.com"
$mwo.Cookies.Add($sortcookie);

$seasoncookie = New-Object System.Net.Cookie   
$seasoncookie.Name = "leaderboard_season"
$seasoncookie.Value = "$Season"
$seasoncookie.Domain = ".mwomercs.com"
$mwo.Cookies.Add($seasoncookie);


#Submit loginform
Invoke-WebRequest -Uri ('https://mwomercs.com/do/login') -WebSession $mwo -Method POST -Body $form.Fields | Out-Null
Write-Verbose "Parse request initialized." -Verbose
start-sleep 2

#pull leaderboards 

$Leaderboards.GetEnumerator() | Invoke-Parallel -ImportVariables -Throttle $Threads -ImportFunctions -ScriptBlock {
    $page= 0
    $rawtables = @()
    $leaderboardpage = $null
    $FirstParse = $true
    $TextStream=[System.IO.StreamWriter]"$SavePath\$($_.name +"_"+ $Season).csv"
    Write-Verbose "Parsing $($_.name)..." -Verbose
    while ($leaderboardpage.Content -notlike "*No results found.*"){
        do{
            $ParseFail = $null
            $ProgressPreference = 'silentlyContinue'
            $leaderboardpage=Invoke-WebRequest -Uri ("https://mwomercs.com/profile/leaderboards?page=$page&type=$($_.value)") -WebSession $mwo
            $ProgressPreference = $OriginalProgressPreference
            if ($leaderboardpage.Content -notlike "*No results found.*"){
                Try {
                        if ($FirstParse){
                            ParseTable -WebRequest $leaderboardpage -FirstParse -TextStream $TextStream -ErrorAction Stop
                        }else{
                            ParseTable -WebRequest $leaderboardpage -TextStream $TextStream -ErrorAction Stop
                        }
                    }
                catch {
                    Write-Warning "Error encountered during parse of $($_.name). Retrying..."
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
            $TextStream.close()    
            Write-Error "Parse error or retry limit exceeded"
            exit
        }
        $FirstParse=$false
        $page++
    }
    $TextStream.close()
    Write-Verbose "$($_.name) saved to $SavePath\$($_.name +"_"+ $Season).csv" -Verbose
}
