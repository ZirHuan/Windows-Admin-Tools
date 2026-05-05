# Import Active Directory module
Import-Module ActiveDirectory

# Main loop
do {
    Clear-Host
    
    # Get all AD groups and sort them by name
    Write-Host "`nRetrieving Active Directory groups..." -ForegroundColor Cyan
    $groups = Get-ADGroup -Filter * | Sort-Object Name

    # Display numbered list of groups
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Active Directory Groups" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green

    for ($i = 0; $i -lt $groups.Count; $i++) {
        Write-Host "$($i + 1). $($groups[$i].Name)"
    }

    # Prompt user to select a group
    Write-Host "`n========================================" -ForegroundColor Green
    $selection = Read-Host "Enter the number of the group to view members (or 'Q' to quit)"

    # Check if user wants to quit
    if ($selection -eq 'Q' -or $selection -eq 'q') {
        Write-Host "Exiting..." -ForegroundColor Yellow
        exit
    }

    # Validate selection
    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $groups.Count) {
        $selectedGroup = $groups[[int]$selection - 1]
        
        Write-Host "`nRetrieving members of group: $($selectedGroup.Name)" -ForegroundColor Cyan
        
        # Get group members
        $members = Get-ADGroupMember -Identity $selectedGroup.DistinguishedName | Sort-Object Name
        
        # Display results
        Write-Host "`n========================================" -ForegroundColor Green
        Write-Host "Members of: $($selectedGroup.Name)" -ForegroundColor Green
        Write-Host "Total Members: $($members.Count)" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        
        if ($members.Count -gt 0) {
            foreach ($member in $members) {
                Write-Host "- $($member.Name) ($($member.objectClass))"
            }
        } else {
            Write-Host "No members found in this group." -ForegroundColor Yellow
        }
        
    } else {
        Write-Host "Invalid selection. Please enter a valid number." -ForegroundColor Red
    }

    # Prompt to continue or exit
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Press ENTER to run again or ESC to exit..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    $continue = $true
    while ($true) {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        if ($key.VirtualKeyCode -eq 13) {  # Enter key
            break
        } elseif ($key.VirtualKeyCode -eq 27) {  # Esc key
            $continue = $false
            break
        }
    }

} while ($continue)

Write-Host "`nGoodbye!" -ForegroundColor Green