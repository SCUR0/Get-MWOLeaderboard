<#
.SYNOPSIS

    Use to pull leaderboard data off mwomercs.com

.DESCRIPTION
	Parses data from mwomercs.com leaderboards. This script
    should be expected to take a long time as it has to go parse multiple
    pages.

.PARAMETER mwoEmail
MWO email used to login mwo.

.PARAMETER password
Used if you want to automate login. Can be left blank if you prefer manual input.

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
    [string]$mwoEmail,
    $password,
    [Parameter(Mandatory=$True)]
    [string]$season,
    $savepath = [Environment]::GetFolderPath("MyDocuments")
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
    pause
    exit
}

#Set SSL to updated version
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


if (!$password){
    #request password if left out of parameter
    $passwordString = read-host -AsSecureString "Password?"
    $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordString))
}

#Configuration varriables
#Respect Original Progress Preference
$OriginalProgressPreference=$ProgressPreference
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
$form.fields['email'] = $mwoEmail
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
Invoke-WebRequest -Uri ('https://mwomercs.com/do/login') -WebSession $mwo -Method POST -Body $form.Fields | Out-Null
Write-Verbose "Parse request initialized." -Verbose
start-sleep 2

#pull leaderboards 

$Leaderboards.GetEnumerator() | Invoke-Parallel -ImportVariables -ImportFunctions -ScriptBlock {
    $page= 0
    $rawtables = @()
    $leaderboardpage = $null
    $FirstParse = $true
    $TextStream=[System.IO.StreamWriter]"$savepath\$($_.name +"_"+ $season).csv"
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
            Write-Error "Parse error retry exceeded"
            pause
            exit
        }
        $FirstParse=$false
        $page++
    }
    $TextStream.close()
    Write-Verbose "$($_.name) saved to $savepath\$($_.name +"_"+ $season).csv" -Verbose
}
