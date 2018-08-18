#requires -version 5
#Author:       Maxim A. Klyunnikov
#Author:       Anatoly S. Demidko

class MailChimp {
    # Connector Reference: https://docs.microsoft.com/en-us/connectors/mailchimp/
    
           [string]         $Version = '2018.08.18'
           [string]         $UserAgent = "PowerShell/MailChimp-API/3.0 (github.com/MaksimKlyu/MailChimp-API/$( $this.Version ))"
           [string]         $ApiEndPoint = 'https://<dc>.api.mailchimp.com/3.0/'
           [int16]          $ApiRequestCount = 10
           [int16]          $ApiRequestTimeoutSec = 15
           [PSCustomObject] $ApiRoot
    hidden [string]         $ApiKey
    hidden [hashtable]      $Headers

    # Create a new instance
    # @param string $api_key      Your MailChimp API key
    # @param string $api_endpoint Optional custom API endpoint
    MailChimp( [string]$ApiKey, [string]$ApiEndPoint ) {
        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'
        
        if ( $ApiEndPoint.Trim() -eq '' ) {
            $this.ApiEndPoint = $this.ApiEndPoint -replace '<dc>', $this.GetDataCenter($ApiKey)
        } else {
            $this.ApiEndPoint = $ApiEndPoint;
        }

        $this.SetHeaders($ApiKey)
        
        #This resource is nothing more than links to other resources available through the API.
        $this.ApiRoot = $this.InvokeWebRequest( $this.ApiEndPoint, 'GET' )
	}

    #
    [string]GetDataCenter( [string]$ApiKey ){
        if ( $ApiKey -notmatch '^(\w{32})-(us\d{2})$' ) { #TODO: DataCenter [2]
            throw "Invalid MailChimp API key supplied"
        }
        return $Matches[2]
    }

	# Return string The url to the API endpoint
    [string]GetApiEndpoint(){
        return $this.ApiEndPoint
    }

    # Set headers of MailChimp apikey
    hidden[void]SetHeaders( [string]$ApiKey ){
        $EncodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($this.Version):$ApiKey"))
        $this.Headers = @{ Authorization = "Basic $EncodedCreds" }
    }
    
    [void]SetApiRequestCount( [int16]$Count ){
        if ($Count -gt 0) {
            $this.ApiRequestCount = $Count
        }
    }

    [void]SetApiRequestTimeoutSec( [int16]$TimeoutSec ){
        if ($TimeoutSec -gt 0) {
            $this.ApiRequestTimeoutSec = $TimeoutSec
        }
    }

    # Convert an email address into a 'subscriber hash' for identifying the subscriber in a method URL
    # @param  string $email The subscriber's email address
    # @return string        Hashed version of the input
	[string]GetSubscriberHash( [string]$Email ){
		$md5  = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
		$utf8 = New-Object -TypeName System.Text.UTF8Encoding
		return [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes( $Email.ToLower() ))).Replace('-', '').ToLower()
    }
    
    # Get MailChimp objests
    [Object]GetObjects( [string]$Uri, [string]$HttpMethod, [string]$PropertyName, [string]$Fields ){
        $Objects = @()
        $Response = @()
        $RequestOffset = 0;
        do {
            $Uri += "?offset=$RequestOffset"
            $Uri += "&count=$( $this.ApiRequestCount )"
            if ( $Fields.Trim() -ne '') {
                $Uri += '&fields=' + $Fields.Trim() + ',total_items'
            }
            $Response = $this.InvokeWebRequest( $Uri, $HttpMethod )
            $Objects += $Response.$( $PropertyName )
            $RequestOffset += $this.ApiRequestCount
        } while ( $RequestOffset -lt $Response.total_items )
        return $Objects
    }
    
    #
    [Object]GetLinkObject ( [PSCustomObject]$AnyObject, [string]$rel ){
        $link = $AnyObject._links | Where-Object rel -eq $rel
        if ( -not $link ) {
            throw "REST API Object Link not found for '$rel'"
        }
        return $link
    }

    # Get information about all lists
    [Object]GetLists() {
        $link = $this.GetLinkObject( $this.ApiRoot, 'lists' )
        return $this.GetObjects( $link.href, $link.method, 'lists', $null )
    }

    # Get information about a list's segments
    [Object]GetListSegments( [PSCustomObject]$ListObject ) {
        $link = $ListObject._links | Where-Object rel -eq 'segments'
        return $this.GetObjects( $link.href, $link.method, 'segments', $null )
    }

    # Get information about a list's interest categories
    [Object]GetListInterestCategories( [PSCustomObject]$ListObject ) {
        $link = $ListObject._links | Where-Object rel -eq 'interest-categories'
        return $this.GetObjects( $link.href, $link.method, 'categories', $null )
    }

    # Get all interests in a specific category
    [Object]GetListInterests( [PSCustomObject]$InterestCategoryObject ) {
        $link = $InterestCategoryObject._links | Where-Object rel -eq 'interests'
        return $this.GetObjects( $link.href, $link.method, 'interests', $null )
    }

    # Get information about members in a specific MailChimp list
    [Object]GetListMembers( [PSCustomObject]$ListObject, [string]$Fields ) {
        $link = $this.GetLinkObject( $ListObject, 'members' )
        return $this.GetObjects( $link.href, $link.method, 'members', $fields )
    }

    [PSCustomObject]InvokeWebRequest( [string]$Uri, [string]$HttpMethod ){
        try {
            return Invoke-WebRequest -Uri $Uri -Method $HttpMethod -Headers $this.Headers -UserAgent $this.UserAgent -TimeoutSec $this.ApiRequestTimeoutSec | ConvertFrom-Json
        } catch {
            throw "Request '$Uri' failed: $( $_.ErrorDetails.Message )"
        }
    }

    [PSCustomObject]InvokeRestMethod( [string]$Uri, [string]$HttpMethod ){
        try {
            return Invoke-RestMethod -Uri $Uri -Method $HttpMethod -Headers $this.Headers -UserAgent $this.UserAgent -TimeoutSec $this.ApiRequestTimeoutSec
        } catch {
            throw "Request '$Uri' failed: $( $_.ErrorDetails.Message )"
        }
    }
}