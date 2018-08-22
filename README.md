# MailChimp API
- MailChimp API v3
- Requires PowerShell 5
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
```console
$Lists | ft
id         web_id name          contact
--         ------ ----          -------
1a2b3c4d5f 123456 Test List     @{company=Contoso; address1=1 The Triangle, Enterp...
1234567890 234567 Test List2    @{company=Contoso; address1=1 The Triangle, Enterp...
2345678901 345678 Test List3    @{company=Contoso; address1=1 The Triangle, Enterp...
```
Select List objects
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
Add or update a list member (This is most useful for syncing subscriber data)
```powershell
$Parameters = @{
    'status'        = 'subscribed'
    'email_address' = 'klyunnikov.maksim@gmail.com'
    'interests'     = @{
     }
    'merge_fields'  = @{
     }
}
$MailChimp.SyncListMember( $List.id, $Parameters )
```
Remove a list member
```powershell
$Parameters = @{
    'email_address' = 'klyunnikov.maksim@gmail.com'
}
$MailChimp.RemoveListMember( $List.id, $Parameters )
```
The Experimental Method
```powershell
$MailChimp.GetCustomObjets( $List, 'merge-fields', 'merge_fields' ) | FT
```
```console
merge_id tag        name         type    required default_value public display_order options                help_text
-------- ---        ----         ----    -------- ------------- ------ ------------- -------                ---------
       3 ADDRESS    Address      address    False                False             4 @{default_country=164}
       5 DEPARTMENT Department   text       False                 True             6 @{size=25}
       1 FNAME      First Name   text       False                 True             2 @{size=25}
       2 LNAME      Last Name    text       False                 True             3 @{size=25}
       4 PHONE      Phone Number phone      False                False             5 @{phone_format=none}
```
