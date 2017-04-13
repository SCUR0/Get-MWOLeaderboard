<#
.SYNOPSIS

    Use to pull leaderboard data off mwomercs.com

.DESCRIPTION
	Parses data from mwomercs.com leaderboards. This script
    should expected to take a long time as it has to go through
    multiple pages.

.PARAMETER global
Only pulls global data instead of all classes.
#>


[cmdletbinding()]
param (
    [switch]$global
)


function ParseTable {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.HtmlWebResponseObject] $WebRequest
    )

    #Extract the tables out of the web request
    try {
        $tables = @($WebRequest.ParsedHtml.getElementsByTagName("TABLE"))
    }
    catch{
        Write-Error "An Error was encountered while trying to pull element by tag `
        name table."
        exit
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
$seasonquestion = Read-Host "Season?"
$season = $seasonquestion - 1
$Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordString))
$loginUrl = "https://mwomercs.com/profile/leaderboards"
$savepath = [Environment]::GetFolderPath("MyDocuments")
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
sleep 1

#pull leaderboards 
foreach ($leaderboard in $leaderboards.GetEnumerator()){
    $page= 0
    $rawtables = @()
    $leaderboardpage = $null
    while ($leaderboardpage.Content -notlike "*No results found.*"){
        $leaderboardpage=Invoke-WebRequest -Uri ("https://mwomercs.com/profile/leaderboards?page=$page&type=$($leaderboard.value)") -WebSession $mwo
        if ($leaderboardpage.Content -notlike "*No results found.*"){
            $rawtables += ParseTable $leaderboardpage
        }
        $page++
        Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Status "Page: $page"
    }
    Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Completed
    $rawtables | Export-Csv "$savepath\$($leaderboard.name +"_"+ $seasonquestion).csv" -NoTypeInformation
    Write-Output "$($leaderboard.name) saved to $savepath\$($leaderboard.name +"_"+ $seasonquestion).csv"  
}
