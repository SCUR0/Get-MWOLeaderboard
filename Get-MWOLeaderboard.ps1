function ParseTable {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.HtmlWebResponseObject] $WebRequest
    )

    #Extract the tables out of the web request

    $tables = @($WebRequest.ParsedHtml.body.getElementsByTagName("TABLE"))

    $table = $tables[0]

    $titles = @()

    $rows = @($table.Rows)

    #Go through all of the rows in the table

    foreach($row in $rows){

        $cells = @($row.Cells)

        # If we've found a table header, remember its titles

        if($cells[0].tagName -eq "th"){

            $titles = @($cells | % { ("" + $_.InnerText).Trim() })

            continue

        }

        #If we haven't found any table headers, make up names "P1", "P2", etc.

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
        $rawtables += ParseTable $leaderboardpage
        $page++
        Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Status "Page: $page"
    }
    Write-Progress -Activity "Scanning $($leaderboard.name) Pages..." -Completed
    $rawtables | Export-Csv "$savepath\$($leaderboard.name).csv"
    Write-Output "$($leaderboard.name) saved to $savepath\$($leaderboard.name).csv"  
}
