function Get-MatchingItems {
    param (
        [array]$List1,
        [array]$List2
    )

    $matchingItems = @()

    foreach ($item in $List1) {
        if ($List2 -contains $item) {
            $matchingItems += $item
        }
    }

    return $matchingItems
}

# Example usage
$list1 = @("user1@domain.onmicrosoft.com", "user2@domain.onmicrosoft.com", "user3@domain.onmicrosoft.com","user4@domain.onmicrosoft.com")
$list2 = @("user2@domain.onmicrosoft.com", "user4@domain.onmicrosoft.com")

$matchingResults = Get-MatchingItems -List1 $list1 -List2 $list2
Write-Output ($matchingResults -join('; '))
