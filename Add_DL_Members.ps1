# Version 1.0

# functions
function Show-Introduction
{
    Write-Host "This script adds a list of owners or members to a Distribution List." -ForegroundColor "DarkCyan"
    Read-Host "Press Enter to continue"
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor DarkCyan
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again"
        }
    }
}

function TryConnect-AzureAD
{
    $connected = Test-ConnectedToAzureAD

    while (-not($connected))
    {
        Write-Host "Connecting to Azure AD..." -ForegroundColor "DarkCyan"
        Connect-AzureAD -ErrorAction SilentlyContinue | Out-Null

        $connected = Test-ConnectedToAzureAD
        if (-not($connected))
        {
            Write-Warning "Failed to connect to Azure AD."
            Read-Host "Press Enter to try again"
        }
    }
}

function Test-ConnectedToAzureAD
{
    try
    {
        Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue | Out-Null
    }
    catch
    {
        return $false
    }
    return $true
}

function PromptFor-DL
{
    $dlEmail = Read-Host "Enter the email address of the DL"
    $dlEmail = $dlEmail.Trim()
    $dl = Get-DistributionGroup -Identity $dlEmail -ErrorAction "Stop"
    return $dl
}

function PromptFor-UserCsvInputs
{
    Write-Host "Script requires CSV list of users and must include a hearer named `"UserPrincipalName`"." -ForegroundColor "DarkCyan"
    $csvPath = Read-Host "Enter path to user CSV (must be .csv)"
    $csvPath = $csvPath.Trim('"')
    return Import-Csv -Path $csvPath
}

function Confirm-CSVHasCorrectHeaders($importedCSV)
{
    $firstRecord = $importedCSV | Select-Object -First 1
    $validCSV = $true

    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name "UserPrincipalName"))
    {
        Write-Warning "This CSV file is missing a header called 'UserPrincipalName'."
        $validCSV = $false
    }

    if (-not($validCSV))
    {
        Write-Host "Please make corrections to the CSV."
        Read-Host "Press Enter to exit"
        Exit
    }
}

function Prompt-YesOrNo($question)
{
    Write-Host "$question`n[Y] Yes  [N] No"

    do
    {
        $response = Read-Host
        $validResponse = $response -imatch '^\s*[yn]\s*$' # regex matches y or n but allows spaces
        if (-not($validResponse)) 
        {
            Write-Warning "Please enter y or n."
        }
    }
    while (-not($validResponse))

    if ($response -imatch '^\s*y\s*$') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false
}

function Grant-Members($dl, $userCsv, $excludeDisabledUsers)
{
    foreach ($user in $userCsv)
    {
        Write-Progress -Activity "Granting members..." -Status $user.UserPrincipalName
        if ($excludeDisabledUsers)
        {
            $userEnabled = Confirm-UserEnabled $user.UserPrincipalName
            if ($null -eq $userEnabled)
            { 
                Log-Warning "The user $($user.UserPrincipalName) was not found. Skipping user."
                continue 
            }

            if (-not($userEnabled)) 
            { 
                Log-Warning "The user $($user.UserPrincipalName) is disabled. Skipping user."
                continue 
            }
        }

        Grant-DLMember -DL $dl -UPN $user.UserPrincipalName
    }
}

function Confirm-UserEnabled($upn)
{
    $upn = $upn.Trim()
    $user = Get-AzureADUser -ObjectId $upn -ErrorAction "SilentlyContinue"
    if ($null -eq $user) { return }
    return $user.AccountEnabled
}

function Grant-DLMember($dl, $upn)
{
    if ($null -eq $upn) { return }

    $dlEmail = $dl.PrimarySmtpAddress.Trim()
    $upn = $upn.Trim()

    try
    {
        Add-DistributionGroupMember -Identity $dlEmail -Member $upn -BypassSecurityGroupManagerCheck -ErrorAction "Stop"
    }    
    catch
    {
        $errorRecord = $_
        Log-Warning "An error occurred when adding the member $upn : `n$errorRecord"
    } 
}

function Log-Warning($message, $logPath = "$PSScriptRoot\logs.txt")
{
    $message = "[$(Get-Date -Format 'yyyy-MM-dd hh:mm tt') W] $message"
    Write-Output $message | Tee-Object -FilePath $logPath -Append | Write-Host -ForegroundColor "Yellow"
}

# main
Show-Introduction
TryConnect-ExchangeOnline
$dl = PromptFor-DL
$userCsv = PromptFor-UserCsvInputs
$excludeDisabledUsers = Prompt-YesOrNo "Exclude disabled users from script inputs? (Takes longer)"
if ($excludeDisabledUsers) { TryConnect-AzureAD }
Grant-Members -DL $dl -UserCsv $userCsv -ExcludeDisabledUsers $excludeDisabledUsers
Write-Host "All done!" -ForegroundColor "Green"
Read-Host -Prompt "Press Enter to exit"


<#
Testing
Input file with 1 user
Input file with 2 users

User is not already a member
User is already a member
User is already an owner
User email was not found
User is disabled
User is enabled
DL is not found
Logging is correct
#>