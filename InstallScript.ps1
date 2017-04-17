if (!$PSScriptRoot) {
    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
}

Write-Output "Retrieving signature card templates..."

$CommonCardTemplates = @{
    "A-1" = Get-Content $PSScriptRoot"\A-1.txt";
    "A-2" = Get-Content $PSScriptRoot"\A-2.txt";
    "A-3" = Get-Content $PSScriptRoot"\A-3.txt";
    "B-1" = Get-Content $PSScriptRoot"\B-1.txt";
    "B-2" = Get-Content $PSScriptRoot"\B-2.txt";
    "B-3" = Get-Content $PSScriptRoot"\B-3.txt";
    "C-1" = Get-Content $PSScriptRoot"\C-1.txt";
    "C-2" = Get-Content $PSScriptRoot"\C-2.txt";
    "C-3" = Get-Content $PSScriptRoot"\C-3.txt";
    "D-1" = Get-Content $PSScriptRoot"\D-1.txt";
    "D-2" = Get-Content $PSScriptRoot"\D-2.txt";
    "D-3" = Get-Content $PSScriptRoot"\D-3.txt";
}

$CommonCardTemplatesNK = ($CommonCardTemplates.Keys | Measure-Object).Count
$CommonCardTemplatesNV = ($CommonCardTemplates.Values | Measure-Object).Count
if (-not $CommonCardTemplatesNK -eq $CommonCardTemplatesNV ) {
    Write-Error "At least one Common Card template file could not be found or accessed in this folder. Please make sure it exists and then try again. Aborting."
    exit
}

$CommonCardTemplates = @{
    "A-1" = $CommonCardTemplates["A-1"] | Out-String;
    "A-2" = $CommonCardTemplates["A-2"] | Out-String;
    "A-3" = $CommonCardTemplates["A-3"] | Out-String;
    "B-1" = $CommonCardTemplates["B-1"] | Out-String;
    "B-2" = $CommonCardTemplates["B-2"] | Out-String;
    "B-3" = $CommonCardTemplates["B-3"] | Out-String;
    "C-1" = $CommonCardTemplates["C-1"] | Out-String;
    "C-2" = $CommonCardTemplates["C-2"] | Out-String;
    "C-3" = $CommonCardTemplates["C-3"] | Out-String;
    "D-1" = $CommonCardTemplates["D-1"] | Out-String;
    "D-2" = $CommonCardTemplates["D-2"] | Out-String;
    "D-3" = $CommonCardTemplates["D-3"] | Out-String;
}

$CustomCardTemplates = @{}
ForEach ($CustomCardFile in Get-ChildItem $PSScriptRoot -Filter "*@*.txt") {
    $CustomCardTemplates[$CustomCardFile.BaseName] = Get-Content $CustomCardFile.VersionInfo.FileName | Out-String
}

$CustomCardTemplatesNK = ($CustomCardTemplates.Keys | Measure-Object).Count
$CustomCardTemplatesNV = ($CustomCardTemplates.Values | Measure-Object).Count
if (-not $CustomCardTemplatesNK -eq $CustomCardTemplatesNV ) {
    Write-Error "At least one Custom Card template file could not be read in this folder. Please check permissions and then try again. Aborting."
    exit
}

Write-Output "Uninstalling old signature card templates..."

$OldTransportRules = Get-TransportRule "Signature | *"

$OldTransportRules | %{ Remove-TransportRule -Identity $_.Guid.Guid -Confirm:$false }

Write-Output "Installing new signature card templates..."

New-TransportRule -Name "Signature | Reset" -RemoveHeader "X-SignatureCards"

New-TransportRule -Name "Signature | Start" -Enabled $false `
 -SentToScope "NotInOrganization" -ExceptIfHeaderMatchesMessageHeader "In-Reply-To" -ExceptIfHeaderMatchesPatterns "\w" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Ready"

ForEach ($EmailAddress in $CustomCardTemplates.Keys) {
    $Disclaimer = $CustomCardTemplates[$EmailAddress]
    New-TransportRule -Name "Signature | Custom Card | $($EmailAddress)" `
     -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "Ready" `
     -From $EmailAddress `
     -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $Disclaimer -ApplyHtmlDisclaimerFallbackAction Wrap `
     -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"
}

New-TransportRule -Name "Signature | CC-1 | > IF phonenumber" `
 -SenderADAttributeMatchesPatterns "phonenumber:\w" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "Ready" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1"

New-TransportRule -Name "Signature | CC-2 | >> IF mobilenumber" `
 -SenderADAttributeMatchesPatterns "mobilenumber:\w" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-1" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2"

New-TransportRule -Name "Signature | CC-3 | >>> IF hasvcard" `
 -FromMemberOf "hasvcard" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2 CC-Scope-3"

New-TransportRule -Name "Signature | CC-4 | >>>> IF initials THEN A-2" `
 -SenderADAttributeMatchesPatterns "initials:\w" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["A-2"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-5 | >>>> ELSE A-1" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-4" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["A-1"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-6 | >>> ELSE A-3" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-3" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["A-3"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-7 | >> ELSE" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-2" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-1" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2"

New-TransportRule -Name "Signature | CC-8 | >>> IF hasvcard" `
 -FromMemberOf "hasvcard" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2 CC-Scope-3"

New-TransportRule -Name "Signature | CC-9 | >>>> IF initials THEN B-2" `
 -SenderADAttributeMatchesPatterns "initials:\w" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["B-2"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-10 | >>>> ELSE B-1" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-4" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["B-1"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-11 | >>> ELSE B-3" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-3" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["B-3"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-12 | > ELSE" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-1" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "Ready" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1"

New-TransportRule -Name "Signature | CC-13 | >> IF mobilenumber" `
 -SenderADAttributeMatchesPatterns "mobilenumber:\w" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-1" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2"

New-TransportRule -Name "Signature | CC-14 | >>> IF hasvcard" `
 -FromMemberOf "hasvcard" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2 CC-Scope-3"

New-TransportRule -Name "Signature | CC-15 | >>>> IF initials THEN C-2" `
 -SenderADAttributeMatchesPatterns "initials:\w" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["C-2"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-16 | >>>> ELSE C-1" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-4" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["C-1"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-17 | >>> ELSE C-3" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-3" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["C-3"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-18 | >> ELSE" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-2" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-1" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2"

New-TransportRule -Name "Signature | CC-19 | >>> IF hasvcard" `
 -FromMemberOf "hasvcard" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "CC-Scope-1 CC-Scope-2 CC-Scope-3"

New-TransportRule -Name "Signature | CC-20 | >>>> IF initials THEN D-2" `
 -SenderADAttributeMatchesPatterns "initials:\w" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["D-2"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-21 | >>>> ELSE D-1" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-4" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-3" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["D-1"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | CC-22 | >>> ELSE D-3" `
 -ExceptIfHeaderContainsMessageHeader "X-SignatureCards" -ExceptIfHeaderContainsWords "CC-Scope-3" `
 -HeaderContainsMessageHeader "X-SignatureCards" -HeaderContainsWords "CC-Scope-2" `
 -ApplyHtmlDisclaimerLocation "Append" -ApplyHtmlDisclaimerText $CommonCardTemplates["D-3"] -ApplyHtmlDisclaimerFallbackAction Wrap `
 -SetHeaderName "X-SignatureCards" -SetHeaderValue "Finished"

New-TransportRule -Name "Signature | End" -RemoveHeader "X-SignatureCards"

Write-Output "Enabling new signature card templates..."

Enable-TransportRule -Identity "Signature | Start"

Write-Output "Done."
