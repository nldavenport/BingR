<#
.NOTES
	Author: Nathan Davenport
    Date  : 2009/03/18

    Copyright © 1984-  Nathan Davenport. All rights reserved.
    Free for all users, but leave in this header

.SYNOPSIS
   BingR - Open Microsoft Edge browser and randomly search Bing

.DESCRIPTION
   Open Microsoft Edge browser and randomly search Bing
      web - regular web search
      images - search images
      videos - search videos
      news - search news
#>

Function Search-Bing {
    <#
	.SYNOPSIS
        Search-Bing - search bing multiple times

	.DESCRIPTION
        Get-Rndword - Performs random Bing searches a specified qty of times

	.PARAMETER qty
        number of Bing searches to perform

    .EXAMPLE

        PS> Search-Bing 3

	.LINK
        
	#>
    Param(
        [int]$qty = 1,
        [int]$min = 3,
        [int]$max = 30,
        [ValidateSet('web','images','videos','news','shop')] $choice
    )

    Begin{
    }
    Process{
        for ($i = 1; $i -le $qty; $i++){
            if ($choice -eq $null) {
                $type = Get-Random -input "web","images","videos","news","shop"
            }
            else {
                $type = $choice
            }
            Switch ($type) {
                web		{ start "microsoft-edge:http://bing.com/Search?q=$(Get-Rndword (Get-Random -Minimum $min -Maximum $max))" }
                images	{ start "microsoft-edge:http://bing.com/images/Search?q=$(Get-Rndword (Get-Random -Minimum $min -Maximum $max))" }
                videos	{ start "microsoft-edge:http://bing.com/videos/Search?q=$(Get-Rndword (Get-Random -Minimum $min -Maximum $max))" }
                news	{ start "microsoft-edge:http://bing.com/news/Search?q=$(Get-Rndword (Get-Random -Minimum $min -Maximum $max))" }
                shop	{ start "microsoft-edge:http://bing.com/Shop?q=$(Get-Rndword (Get-Random -Minimum $min -Maximum $max))" }
            }
        }
    }
    End {
    }
}