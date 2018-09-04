#requires -version 5
# Link: https://github.com/MaksimKlyu/MailChimp-API.git
# Author: Maxim A. Klyunnikov

class MailChimp {
	[string]           $Version = '2018.08.31'
	[string]           $UserAgent = "PowerShell/MailChimp-API/3.0 (github.com/MaksimKlyu/MailChimp-API/$( $this.Version ))"
	[string]           $ApiEndPoint = 'https://<dc>.api.mailchimp.com/3.0/'
	[int16]            $ApiRequestCount = 10
	[int16]            $ApiRequestTimeoutSec = 30
	[string]           $ApiRequestContentType = 'application/json; charset=utf-8'
	[PSCustomObject]   $ApiRoot #The Experimental Method
	hidden [string]    $ApiKey
	hidden [hashtable] $Headers

	MailChimp( [string]$ApiKey, [string]$ApiEndPoint ) {
		Set-StrictMode -Version Latest
		$ErrorActionPreference = 'Stop'
		if ( $ApiEndPoint.Trim() -eq '' ) {
			$this.ApiEndPoint = $this.ApiEndPoint -replace '<dc>', $this.GetDataCenter($ApiKey)
		} else {
			$this.ApiEndPoint = $ApiEndPoint
		}
		$this.SetHeaders($ApiKey)

		# This resource is nothing more than links to other resources available through the API
		$this.ApiRoot = $this.InvokeGetRequest( $this.ApiEndPoint )
	}

	# Get the Data Center <dc> value from a Mailchimp API key
	[string]GetDataCenter( [string]$ApiKey ) {
		if ( $ApiKey -notmatch '^\w{32}-(?<DataCenter>[a-zA-Z]{2}\d{1,2})$' ) {
			throw "Invalid MailChimp API key supplied"
		}
		return $Matches.DataCenter
	}

	# Set headers of MailChimp apikey
	hidden[void]SetHeaders( [string]$ApiKey ) {
		$EncodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($this.Version):$ApiKey"))
		$this.Headers = @{ Authorization = "Basic $EncodedCreds" }
	}

	[void]SetApiRequestCount( [int16]$Count ) {
		if ($Count -gt 0) {
			$this.ApiRequestCount = $Count
		}
	}

	[void]SetApiRequestTimeoutSec( [int16]$TimeoutSec ) {
		if ($TimeoutSec -gt 0) {
			$this.ApiRequestTimeoutSec = $TimeoutSec
		}
	}

	# Get information about all lists
	[Object]GetLists() {
		$Link = $this.ApiEndPoint + "lists"
		return $this.GetObjects( $Link, 'lists', $null )
	}

	# Get information about all segments in a list
	[Object]GetListSegments( [string]$ListId ) {
		$Link = $this.ApiEndPoint + "lists/$ListId/segments"
		return $this.GetObjects( $Link, 'segments', $null )
	}

	# Get information about a list’s interest categories
	[Object]GetListInterestCategories( [string]$ListId ) {
		$Link = $this.ApiEndPoint + "lists/$ListId/interest-categories"
		return $this.GetObjects( $Link, 'categories', $null )
	}

	# Get information about a list’s interest categories
	[Object]GetListInterests( [string]$ListId, [string]$InterestCategoryId ) {
		$Link = $this.ApiEndPoint + "lists/$ListId/interest-categories/$InterestCategoryId/interests"
		return $this.GetObjects( $Link, 'interests', $null )
	}

	# Get information about members in a list
	[Object]GetListMembers( [string]$ListId, [string]$Fields ) {
		$Link = $this.ApiEndPoint + "lists/$ListId/members"
		return $this.GetObjects( $Link, 'members', $Fields )
	}

	# Add a new list member
	[Object]AddListMember( [string]$ListId, [hashtable]$Parameters ) {
		$Link = $this.ApiEndPoint + "lists/$ListId/members"
		return $this.InvokePostRequest( $Link, $Parameters )
	}

	# Update a list member
	[Object]UpdateListMember( [string]$ListId, [hashtable]$Parameters ) {
		$SubscriberHash = $this.GetSubscriberHash($Parameters.email_address)
		$Link = $this.ApiEndPoint + "lists/$ListId/members/$SubscriberHash"
		return $this.InvokePatchRequest( $Link, $Parameters )
	}

	# Add or update a list member (This is most useful for syncing subscriber data)
	[Object]SyncListMember( [string]$ListId, [hashtable]$Parameters ) {
		$SubscriberHash = $this.GetSubscriberHash($Parameters.email_address)
		$Link = $this.ApiEndPoint + "lists/$ListId/members/$SubscriberHash"
		return $this.InvokePutRequest( $Link, $Parameters )
	}

	# Remove a list member
	[Object]RemoveListMember( [string]$ListId, [hashtable]$Parameters ) {
		$SubscriberHash = $this.GetSubscriberHash($Parameters.email_address)
		$Link = $this.ApiEndPoint + "lists/$ListId/members/$SubscriberHash"
		return $this.InvokeDeleteRequest( $Link, $Parameters )
	}


	# Get MailChimp objests
	[Object]GetObjects( [string]$Uri, [string]$ObjectName, [string]$Fields ) {
		$Objects = @()
		$Response = @()
		$RequestOffset = 0
		do {
			$UriQuery = $Uri
			$UriQuery += "?offset=$RequestOffset"
			$UriQuery += "&count=$( $this.ApiRequestCount )"
			if ( $Fields.Trim() -ne '') { $UriQuery += '&fields=' + $Fields.Trim() + ',total_items' }
			$Response = $this.InvokeGetRequest( $UriQuery )
			$Objects += $Response.$( $ObjectName )
			$RequestOffset += $this.ApiRequestCount
		} while ( $RequestOffset -lt $Response.total_items )
		return $Objects
	}

	# GET request to retrieve data
	[PSCustomObject]InvokeGetRequest( [string]$Uri ) {
		try {
			return Invoke-RestMethod -Uri $Uri -Method Get -Headers $this.Headers -UserAgent $this.UserAgent -TimeoutSec $this.ApiRequestTimeoutSec
		} catch {
			throw "Request '$Uri' failed: $( $_.ErrorDetails.Message )"
		}
	}

	# POST request to create new resources
	[PSCustomObject]InvokePostRequest( [string]$Uri, [hashtable]$Body ) {
		try {
			$BobyJson = $Body | ConvertTo-Json -Depth 100 -Compress
			return Invoke-RestMethod -Uri $Uri -Method Post -Headers $this.Headers -Body $BobyJson -UserAgent $this.UserAgent -TimeoutSec $this.ApiRequestTimeoutSec -ContentType $this.ApiRequestContentType
		} catch {
			throw "Request '$Uri' failed: $( $_.ErrorDetails.Message )"
		}
	}

	# PATCH request to update a resource
	[PSCustomObject]InvokePatchRequest( [string]$Uri, [hashtable]$Body ) {
		try {
			$BobyJson = $Body | ConvertTo-Json -Depth 100 -Compress
			return Invoke-RestMethod -Uri $Uri -Method Patch -Headers $this.Headers -Body $BobyJson -UserAgent $this.UserAgent -TimeoutSec $this.ApiRequestTimeoutSec -ContentType $this.ApiRequestContentType
		} catch {
			throw "Request '$Uri' failed: $( $_.ErrorDetails.Message )"
		}
	}

	# PUT request to create or update a resource. This is most useful for syncing subscriber data
	[PSCustomObject]InvokePutRequest( [string]$Uri, [hashtable]$Body ) {
		try {
			$BobyJson = $Body | ConvertTo-Json -Depth 100 -Compress
			return Invoke-RestMethod -Uri $Uri -Method Put -Headers $this.Headers -Body $BobyJson -UserAgent $this.UserAgent -TimeoutSec $this.ApiRequestTimeoutSec -ContentType $this.ApiRequestContentType
		} catch {
			throw "Request '$Uri' failed: $( $_.ErrorDetails.Message )"
		}
	}


	# DELETE request to remove a resource [!]
	[PSCustomObject]InvokeDeleteRequest( [string]$Uri, [hashtable]$Body ) {
		try {
			$BobyJson = $Body | ConvertTo-Json -Depth 100 -Compress
			return Invoke-RestMethod -Uri $Uri -Method Delete -Headers $this.Headers -Body $BobyJson -UserAgent $this.UserAgent -TimeoutSec $this.ApiRequestTimeoutSec -ContentType $this.ApiRequestContentType
		} catch {
			throw "Request '$Uri' failed: $( $_.ErrorDetails.Message )"
		}
	}

	# Convert an email address into a 'subscriber hash' for identifying the subscriber in a method URL
	[string]GetSubscriberHash( [string]$Email ) {
		$Md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
		$Utf8 = New-Object -TypeName System.Text.UTF8Encoding
		return [System.BitConverter]::ToString($Md5.ComputeHash($Utf8.GetBytes( $Email.ToLower() ))).Replace('-', '').ToLower()
	}

	# The Experimental Method
	[Object]GetLinkObject ( [PSCustomObject]$AnyObject, [string]$Rel ) {
		$Link = $AnyObject._links | Where-Object rel -eq $Rel
		if ( -not $Link ) {
			throw "REST API Object Link not found. ({1})" -f $AnyObject._links.rel -join ', '
		}
		return $Link
	}

	# The Experimental Method
	[Object]GetCustomObjets( [PSCustomObject]$CustomObject, [string]$rel, [string]$ObjectName ) {
		$Link = $this.GetLinkObject( $CustomObject, $Rel )
		return $this.GetObjects( $Link.href, $ObjectName, $null )
	}
}
