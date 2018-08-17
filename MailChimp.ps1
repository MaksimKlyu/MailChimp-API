#requires -version 5
#Author:       Maxim A. Klyunnikov
#Author:       Anatoly S. Demidko

class MailChimp {
    # Connector Reference: https://docs.microsoft.com/en-us/connectors/mailchimp/
    
           [string]         $Version = '0.0.1'
           [string]         $ApiEndPoint = 'https://<dc>.api.mailchimp.com/3.0/'
           [int16]          $ApiRequestCount = 10
           [int16]          $ApiRequestTimeoutSec = 15
           [PSCustomObject] $Self
    hidden [string]         $ApiKey
    hidden [hashtable]      $Headers

    # Create a new instance
    # @param string $api_key      Your MailChimp API key
    # @param string $api_endpoint Optional custom API endpoint
    MailChimp( [string]$ApiKey, [string]$ApiEndPoint ) {
        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'
        if ( $ApiEndPoint.Trim() -eq '' ) {
            $this.ApiEndPoint = $this.ApiEndPoint -replace '<dc>', $this.GetDataCenter($ApiKey) #TODO: [2] -> DataCenter
        } else {
            $this.ApiEndPoint = $ApiEndPoint;
        }
        $this.SetHeaders($ApiKey)
        $this.Self = $this.InvokeWebRequest( $this.ApiEndPoint, 'GET' )
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

    [void]SetApiApiRequestTimeoutSec( [int16]$TimeoutSec ){
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
        $fields = ''
        $link = $this.GetLinkObject( $this.Self, 'lists' )
        try {
            return $this.GetObjects( $link.href, $link.method, 'lists', $fields )
        } catch {
            throw "Failed to Get mailchim all lists: $( $_.Exception.Message )"
        }
    }

    # Get information about a list's interest categories
    [Object]GetInterestCategories( [PSCustomObject]$ListObject ) {
        $fields = ''
        $link = $ListObject._links | Where-Object rel -eq 'interest-categories'
        try {
            return $this.GetObjects( $link.href, $link.method, 'categories', $fields )
        } catch {
            throw "Failed to Get mailchim all lists: $( $_.Exception.Message )"
        }
    }

    # Get all interests in a specific category
    [Object]GetInterests( [PSCustomObject]$InterestCategoryObject ) {
        $fields = ''
        $link = $InterestCategoryObject._links | Where-Object rel -eq 'interests'
        try {
            return $this.GetObjects( $link.href, $link.method, 'interests', $fields )
        } catch {
            throw "Failed to Get mailchim all lists: $( $_.Exception.Message )"
        }
    }

    # Get information about members in a specific MailChimp list
    [Object]GetListMembers( [PSCustomObject]$ListObject ) {
        
        $fields = 'members.id,members.email_address,members.status,members.merge_fields,members.interests'
        $link = $this.GetLinkObject( $ListObject, 'members' )
        try {
            return $this.GetObjects( $link.href, $link.method, 'members', $fields )
        } catch {
            throw "Failed to Get mailchim all lists: $( $_.Exception.Message )"
        }
    }


    [PSCustomObject]InvokeWebRequest( [string]$Uri, [string]$HttpMethod ){
        $UserAgent = 'PowerShell/MailChimp-API/3.0 (github.com/MaksimKlyu/MailChimp-API)'
        try {
            return Invoke-WebRequest -Uri $Uri -Method $HttpMethod -Headers $this.Headers -UserAgent $UserAgent -TimeoutSec $this.ApiRequestTimeoutSec | ConvertFrom-Json
        } catch {
            throw "Web Request '$Uri' failed : $( $_.ErrorDetails.Message )"
        }
    }

    # [string]ToString(){
    #     return "Hello, I'm MailChimp-API/3.0`nClass version: {0}`nApiEndPoint: {1}" -f $this.Version, $this.ApiEndPoint
    # }
}

# Examle
#
# $MailChimp          = New-Object MailChimp ( '<32>-us17', $null )
# $Lists              = $MailChimp.GetLists()
# $List               = $lists | Where-Object Name -eq 'Internal news' #'Test List'
# $InterestCategories = $MailChimp.GetInterestCategories( $List )
# $InterestCategory   = $InterestCategories | Where-Object Title -eq 'Internal Groups'

# $GetInterests       = $MailChimp.GetInterests( $InterestCategory )
# $ListMembers        = $MailChimp.GetListMembers( $List )
