# MailChimp API
- MailChimp API v3
- Requires PowerShell 5.1.
## Installation
You should download the MailChimp.ps1 file and include it manually
```powershell
. .\MailChimp-API\MailChimp.ps1
```
## Examples
Creating an instance with your API key
```powershell
$MailChimp = New-Object MailChimp ( '12345678901234567890123456789012-usXX', $null )
```
Get information about all lists
```powershell
$Lists = $MailChimp.GetLists()
```
```powershell
$List = $lists | Where-Object Name -eq 'ListName'
```
Get information about a list’s interest categories
```powershell
$ListInterestCategories = $MailChimp.GetListInterestCategories( $List.id )
```
```powershell
$ListInterestCategory = $ListInterestCategories | Where-Object Title -eq 'GroupsTitle'
```
Get all interests in a specific category
```powershell
$ListInterests = $MailChimp.GetListInterests( $ListInterestCategory.list_id, $ListInterestCategory.id )
```
Get information about all segments in a list
```powershell
$ListSegments = $MailChimp.GetListSegments( $List.id )
```
Get information about members in a list
```powershell
$Filds = 'members.id,members.email_address,members.status,members.merge_fields,members.interests'
$ListMembers = $MailChimp.GetListMembers( $List.id, $Filds )
```
Add a new list member
```powershell
$Parameters = @{
    'status'        = 'subscribed'
    'email_address' = 'klyunnikov.maksim@gmail.com'
    'interests'     = @{
     }
    'merge_fields'  = @{
     }
}
$MailChimp.AddListMember( $List.id, $Parameters )
```
Update a list member
```powershell
$Parameters = @{
    'status'        = 'unsubscribed'
    'email_address' = 'klyunnikov.maksim@gmail.com'
}
$MailChimp.UpdateListMember( $List.id, $Parameters )
```
Remove a list member
```powershell
$Parameters = @{
    'email_address' = 'klyunnikov.maksim@gmail.com'
}
$MailChimp.RemoveListMember( $List.id, $Parameters )
```




