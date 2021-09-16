Import-Module ConfluencePS
Import-Module JiraPS

$user = "MyUsername" #Change Username to Jira/Confluence Automation User
$pw = Get-Content C:\encrypted_password | ConvertTo-SecureString #Get Password from SecureString File ($credential.Password | ConvertFrom-SecureString | Set-Content C:\encrypted_password)
$cred = New-Object System.Management.Automation.PSCredential($user, $pw) #Set Credentials
$pw2 = $cred.GetNetworkCredential().Password

Set-ConfluenceInfo -BaseURi "http://myConfluenceServer.domain" -Credential $cred
Set-JiraConfigServer "http://myJiraServer.domain"
New-JiraSession -Credential $cred                   #Start a new Jira Session


$seite = Get-ConfluenceChildPage -PageID 123456
$seite | ForEach-Object{

    $logpath= "C:\Log\"+$_.Title+".txt"             #Set Log Path and Log title (page title)
        if ((test-path $logpath) -eq $true){return} #If Log already exists end script

$taskkey = $null
$field = $null
$body = $null
$index = 1
$start = 1
$body = $_.body

Start-Transcript $logpath

#While index 
while($index -ne "-1"){

    if($index -gt "0"){
        $index = $body.IndexOf("Projectkey", $start) #Change to Projectkey of Jira Issue
        $taskkey = $body.Substring($index, 9)
        $start = $index+1
        #Write-Host $index
    }
                
    if($taskkey -ne $null -and ($taskkey -match '"' -or $taskkey -match "<" -or $taskkey -match ">")){
        $taskkey = $taskkey.TrimEnd('"', '<', '>')
    }
            
    $field = Get-JiraIssue -Key $taskkey -Fields customfield_ID #Organization field ID
            
    if($field -ne $null -and $field.customfield_ID -eq $null){     
        
        $uri = "$(Get-JiraConfigServer)/rest/api/latest/issue/$($taskkey)?notifyUsers=false"

    $jsonString = @'
    {
        "fields": {
            "customfield_10110":[5]
        }
     }
            }
            }
'@  

    $header = @{
        Authorization = "Basic"+[System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("Usename:$pw2"))
    }

    Invoke-RestMethod -uri $uri -Headers $header -Method PUT -Body $jsonString -ContentType "application/json" -Credential $cred

    Write-Host "Organization added to: $($taskkey)"
    Write-Host "--------------------------------------------------"   

}
}
Stop-Transcript
}