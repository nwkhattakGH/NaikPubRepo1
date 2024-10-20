﻿# Define List A and List B
$listA = @("item1", "item2", "item3", "item4")
$listB = @("item3", "item2")

# Function to check if all items in List B exist in List A
function Check-AllItemsExist {
    param (
        [array]$listA,
        [array]$listB
    )
    foreach ($item in $listB) {
        if (-not ($listA -contains $item)) {
            return $false
        }
    }
    return $true
}

# Call the function and store the result
$result = Check-AllItemsExist -listA $listA -listB $listB

# Output the result
if ($result) {
    Write-Host "All items from List B exist in List A." -ForegroundColor Green
} else {
    Write-Host "Not all items from List B exist in List A." -ForegroundColor Yellow
}