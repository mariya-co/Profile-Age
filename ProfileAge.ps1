<# Profile age script - Writen to enumrate profiles and whether they're being actively used from a remote machine.
This was re-developed and refactored from this microsoft docs page https://learn.microsoft.com/en-us/troubleshoot/windows-server/support-tools/scripts-to-retrieve-profile-age
With the aim to make a more simple, less complicated solution. It sort of works - it doesn't neccessairly produce super accurate results due to the nature of the script.
For example - someone that never logs out and just uses sleep would show as having a super old profile that never logs in, even though they use it constantly.
#>
# Covert the low/high regkeys, into human readable (or slightly more so) format for use further down.
Function ConvertToDate {
    param(
        [uint32]$lowpart, 
        [uint32]$highpart
    )
    $ft64 = ([UInt64]$highpart -shl 32) -bor $lowpart
    [datetime]::FromFileTime($ft64)
}

# Translate SID to NT Account name - makes it easier to work with the functions further down in this script
Function TranslateSID2Name {
    param(
        [string]$sid
        )
    try {
        $osid = New-Object System.Security.Principal.SecurityIdentifier($sid)
        $osid.Translate([System.Security.Principal.NTAccount])
    } catch {
        Write-Output "Error translating SID: $sid"
        Write-Output $_.Exception.Message
        return $sid
    }
}

# Enumerate OST files from the machine - used to retrieve the full names of the profile, without using AD or ADSI.
Function Get-OSTFilesRemote {
    param(
        [string]$computerName, 
        [string]$profileImagePath
    )
    $username = Split-Path -Path $profileImagePath -Leaf
    Invoke-Command -ComputerName $computerName -ScriptBlock {
        param($username)
        $outlookPath = "C:\Users\$username\AppData\Local\Microsoft\Outlook" # appdata path that contains the OST files, which we use to figure out the name of the profile.
        Get-ChildItem -Path $outlookPath -Filter '*.ost' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
    } -ArgumentList $username
}

# Get profile age and display name
Function Get-ProfileAge {
    param(
        [string]$profileSid,
        [string[]]$defaultSids,
        [string]$computerName
    )

    if ($profileSid -in $defaultSids) { # Use this list of default sids to return null, once again, not relevant to us.
        return $null
    }

    $profileKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$profileSid"
    try { # We're pulling the individual keys here from the profile key - and getting the value within. If any errors, throw to null.
        $profileImagePath = Get-ItemPropertyValue -Path $profileKeyPath -Name ProfileImagePath -ErrorAction Stop
        $loadTimeLow = Get-ItemPropertyValue -Path $profileKeyPath -Name LocalProfileLoadTimeLow -ErrorAction Stop
        $loadTimeHigh = Get-ItemPropertyValue -Path $profileKeyPath -Name LocalProfileLoadTimeHigh -ErrorAction Stop
    } catch {
        return $null
    }

    $profileLoadDate = ConvertToDate -lowpart $loadTimeLow -highpart $loadTimeHigh
    $profileAge = (Get-Date) - $profileLoadDate # Calculate the age of the profile by subtracting the load date from the current date.
    $profileName = TranslateSID2Name -sid $profileSid

    # Run the OST function - this checks the OST path to see whats there, and further down, brings up a menu that lets the user chose which one looks right.
    $ostFiles = Get-OSTFilesRemote -computerName $computerName -profileImagePath $profileImagePath
    # Prompt user locally if multiple OST files found
    if ($ostFiles.Count -gt 1) { # Make a little nice menu for us to select which .OST file looks correct. This is mainly used when people have shared mailboxes and the like.
        Write-Host "Multiple OST files found for $profileName"
        for ($i = 0; $i -lt $ostFiles.Count; $i++) {
            Write-Host "$($i + 1): $($ostFiles[$i])"
        }
        $selection = Read-Host "Select an OST file by number"
        $selectedOstFile = $ostFiles[$selection - 1] # Register the selections
    } elseif ($ostFiles.Count -eq 1) { # If there's one, theres none (no reason to do anything - continue as normal and assume the email there is their name.)
        $selectedOstFile = $ostFiles[0]
    } else {
        $selectedOstFile = $null # Some people don't use the outlook client, this will return null.
    }

    if ($selectedOstFile) {
        $email = $selectedOstFile -replace '\.ost$', '' # Regex that chops and changes x.x@domain to x x so that it looks like a real name, and not an email address.
        $displayName = $email.Split('@')[0] -replace '\.', ' ' 
    } else {
        $displayName = $profileName # Condition that occurs when there's no ost files - we'll just show their username instead. Works better than nothing.
    }

    [PSCustomObject]@{
        "Profile Name"        = $profileName
        "Display Name"        = $displayName
        "Profile Load Date"   = $profileLoadDate.ToString("dd/MM/yyyy")
        "Last login (days)"   = $profileAge.Days
    }
}

# Main - filter out default sids, get host name, get the reg keys, get the profile age, draw a table with the results.
$defaultSids = @("S-1-5-18", "S-1-5-19", "S-1-5-20") # Default SIDS to exclude - not relevant to what we're looking for
$remoteComputer = Read-Host 'Enter hostname'

# Get profile SIDs remotely
$profileSids = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
    Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
        Where-Object { $_.PSChildName -notlike "*.old" } | # Filter out .old profiles, select just the specific profile keys required.
        Select-Object -ExpandProperty PSChildName
}

$profileData = @()
foreach ($sid in $profileSids) {
    $profile = Get-ProfileAge -profileSid $sid -defaultSids $defaultSids -computerName $remoteComputer
    if ($profile) {
        $profileData += $profile
    }
}

$profileData | Select-Object "Profile Name", "Display Name", "Profile Load Date", "Last login (days)" | Format-Table -AutoSize
 