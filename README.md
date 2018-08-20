```
. .\MailChimp-API\MailChimp.ps1

$MailChimp              = New-Object MailChimp ( '<32>-us17', $null )

$Lists                  = $MailChimp.GetLists()

$List                   = $lists | Where-Object Name -eq 'Test List'
$ListInterestCategories = $MailChimp.GetListInterestCategories( $List.id )
$ListInterestCategory   = $ListInterestCategories | Where-Object Title -eq 'Internal Groups'
$ListInterests          = $MailChimp.GetListInterests( $ListInterestCategory.list_id, $ListInterestCategory.id )

$ListSegments           = $MailChimp.GetListSegments( $List.id )

$Filds                  = 'members.id,members.email_address,members.status,members.merge_fields,members.interests'
$ListMembers            = $MailChimp.GetListMembers( $List.id, $Filds )

<#
$Parameters = @{
    'email_address' = 'klyunnikov.maksim@gmail.com'
}
$MailChimp.RemoveListMember( $List.id, $Parameters )
#>



$Parameters = @{
    'status'        = 'subscribed'
    'email_address' = 'klyunnikov.maksim@gmail.com'
    'interests'     = $Interests
    'merge_fields'  = $MergeFields
}
$MailChimp.UpdateListMember( $List.id, $Parameters )
```
