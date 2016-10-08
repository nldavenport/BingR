<#
.Synopsis
   BingR - Open Bing and search
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>

$Vowels = "aeiouy "
$Consenants = "bcdfghjklmnpqrstvwxyz "
$UpperCase = "$($Consenants.ToUpper())$($Vowels.ToUpper())"
$LowerCase = "$($Consenants.ToLower())$($Vowels.ToLower())"
$MixedCase = "$($UpperCase)$($LowerCase)"
$Numbers = "1234567890 "
$SpecialChars = "!@#$%^&*()-=_+[]{}\|;:',./<>? "
$AllowedChars = "$($MixedCase)$($Numbers)$($SpecialChars)"

# Function to return a random character
Function get-rndchar([char]$chartype) { 
    Switch ($chartype) {
        V { $char = $Vowels.Substring((Get-Random -Minimum 1 -Maximum $Vowels.length),1) }
        c { $char = $Consenants.Substring((Get-Random -Minimum 1 -Maximum $Consenants.length),1) }
        u { $char = $UpperCase.Substring((Get-Random -Minimum 1 -Maximum $UpperCase.length),1) }
        l { $char = $LowerCase.Substring((Get-Random -Minimum 1 -Maximum $LowerCase.length),1) }
        m { $char = $MixedCase.Substring((Get-Random -Minimum 1 -Maximum $MixedCase.length),1) }
        n { $char = $Numbers.Substring((Get-Random -Minimum 1 -Maximum $Numbers.length),1) }
        s { $char = $SpecialChars.Substring((Get-Random -Minimum 1 -Maximum $SpecialChars.length),1) }
        a { $char = $AllowedChars.Substring((Get-Random -Minimum 1 -Maximum $AllowedChars.length),1) }
    }
    return $char
}

# Function to return a random character
Function get-rndword([int]$qty) { 
    for ($i=1; $i -le $qty; $i++) {
        $Word = "$($word)$(get-rndchar(Get-Random -input "v","c","l","u","m",""n","s","a"))"
    }
    return $word
}

# Function to Open Edge browser and search Bing multiple times
Function invoke-BingSearch([int]$qty) {

    start microsoft-edge:https://login.live.com/login.srf
    start microsoft-edge:https://account.microsoft.com/rewards/dashboard


    for ($i = 1; $i -le $qty; $i++){
        Switch (Get-Random -input "web","images","videos","news") {
            web		{ start "microsoft-edge:http://bing.com/Search?q=$(get-rndword (Get-Random -Minimum 1 -Maximum 45))" }
            images	{ start "microsoft-edge:http://bing.com/images/Search?q=$(get-rndword (Get-Random -Minimum 1 -Maximum 45))" }
            videos	{ start "microsoft-edge:http://bing.com/videos/Search?q=$(get-rndword (Get-Random -Minimum 1 -Maximum 45))" }
            news	{ start "microsoft-edge:http://bing.com/news/Search?q=$(get-rndword (Get-Random -Minimum 1 -Maximum 45))" }
        }
    }
}

clear-host
invoke-BingSearch 30

<# 
# Set-Variable SpecialChars "!@#$%^&*()-=_+[]{}\|;:',./<>? " –option ReadOnly
# Remove-Variable SpecialChars -force
#>