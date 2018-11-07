# Get-MWOLeaderboard
Parses MWO leaderboard data from PGI's website

Requires Invoke-Parallel (https://github.com/RamblingCookieMonster/Invoke-Parallel)

Ouputs CSV files to your Documents folder if parameter left empty (usually C:\users\username\documents).

Script requires login credentials. This is needed because the leaderboard requires login.
You can check source code. Optional password parameter if you would like to pass password parameter.


### Note
CSV files will show a data size of zero until completed. The script usually takes around 45 minutes
depending on how many players participated that season.
