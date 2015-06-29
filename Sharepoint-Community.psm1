

#region libraries

function New-SharePointLibrary {

    [CmdletBinding()]

    Param(

       [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]

       [string]$webUrl,

       [Parameter(Mandatory=$true, Position=1)]

       [string]$LibraryName,

       [Parameter(Mandatory=$true, Position=2)]

       [string]$Description,

       [Parameter(Mandatory=$false, Position=3)]

       [string]$LibraryTemplate

    )

   Process

   {

      Start-SPAssignment -Global 

      $spWeb = Get-SPWeb -Identity $webUrl    

      $spListCollection = $spWeb.Lists  

      $spLibrary = $spListCollection.TryGetList($LibraryName)

      if($spLibrary -ne $null) {

          Write-Host -ForegroundColor Yellow "Library $LibraryName already exists in the site"

      } else {       

          Write-Host -NoNewLine -f yellow "Creating  Library - $LibraryName"

          $spListCollection.Add($LibraryName, $Description, $LibraryTemplate)

          #Always Use the GetList to get library, while once they have been renamed they are not returned list
          $spLibrary = $spWeb.GetList($spWeb.ServerRelativeUrl+'/'+"$LibraryName")

          Write-Host -f Green '...Success!'

      }         

      Stop-SPAssignment -Global  

   }

} 

Function Get-SPLibraryFile {
    
    [cmdletBinding()]
    Param(
        

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [string]
        $LibraryName,

        [Parameter(Mandatory=$false)]
        [string]
        $FileName 

    )

    
    

    if (!($web)){
        $Web = Get-SPweb -Site $Url
    }


        
        if ($FileName){
            $Return = $Web.GetFolder($LibraryName).files | where {$_.name -eq $FileName}
        }else{
            $Return = $Web.GetFolder($LibraryName).files
        }

        return $Return

}


#endregion



#region SpLists

Function Get-SPList {

Param(
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        [string]
        $ListName
)
    if (!($web)){
        $Web = Get-SPweb -Site $Url
    }
 

    if ($Listname){
        $return = $Web.lists | ? {$_.title -eq $ListName} 
        write-verbose "Returning list name: $($List.Title)"
    }else{
        write-verbose 'Returning all lists.'
        $return = $Web.lists
        write-verbose "Total of $($return.count) returned."
    }

Return $return
}

function New-SPList{

    [CmdletBinding()]    param(        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [String]
        $Name,        [Parameter(Mandatory=$false)]
        [String]        $Description,        [Parameter(Mandatory=$true)]        [ValidateSet('Document Library','Form Library','Wiki Page Library','Picture Library','Links','Announcements','Contacts','Calendar','Promoted Links','Discussion Board','Tasks (2010)','Project Tasks',
                        'Tasks',
                        'Issue Tracking',
                        'Custom List',
                        'Custom List in Datasheet View',
                        'External List',
                        'Survey',
                        'Data Sources',
                        'Data Connection Library',
                        'Access App',
                        'Converted Forms',
                        'Custom Workflow Process',
                        'No Code Public Workflows',
                        'No Code Workflows',
                        'Report Library',
                        'Workflow History',
                        'Status List',
                        'Asset Library')]        [string]        $Type,        [Parameter(Mandatory=$false)]        [switch]$Force    )
    
    if (!($Web)){
        $Web = Get-SPWeb -Identity $Url
    }
    
    if ($Web.lists | ? {$_.title -eq $Name} ){
        if ($force){
            write-verbose "Force parameter has been specified. Creating list $($Name)"
            $GUID = $web.lists.add($Name,$Description,$web.listTemplates["$Type"])
            Return $GUID
        }else{
            Write-Warning "A list with $($Name) is already present on site $($Url). Use -Force to override existing list."

        }

    }else{
        $GUID = $web.lists.add($Name,$Description,$web.listTemplates["$Type"])
        Return $GUID
    }
    

}

Function Remove-SpList {

    [CmdletBinding()]    param(
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,        [Parameter(Mandatory=$true)]        $ListName,

        [Parameter(Mandatory=$false)]
        [switch]$Force
    )

    if (!($Web)){
        $Web = Get-SPWeb -Identity $Url
    }
    
    $List = $Web.Lists["$ListName"]
    
    if ($force){
        write-verbose "Deleting the list $($listname) from $($url)"
        $List.Delete()
    }
    else{
        write-verbose "REcycling the list $($listname) from $($url) into recycle bin."
        $List.recycle()
    }


}

Function Get-SpListTemplate {

    [CmdletBinding()]    param(
        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,
        [Parameter(Mandatory=$true)]        [String]        $Name
    )

    if (!($Web)){
        $web = Get-SPWeb -Identity $Url
    }

    $Templates = Get-SPList -Web $Web -ListName 'List Template Gallery'
    
    if ($Name){                                                            
        $ListTemplate = $templates.Items | ? {$_.title -eq $Name}
    }
    else{
        $ListTemplate = $templates.Items
    }
    
    return $ListTemplate

}

Function Remove-SpLisTtemplate {
        [CmdletBinding()]    param(

        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,
        [Parameter(Mandatory=$true)]        [String]        $Name
    )

    $Template = Get-SpListTemplate -url $Url -ListTemplateName $Name                                                            
    
    if ($Template){
        write-verbose 'Deleting $($Name) Template'                                                                                                                                  
        $Template.delete()
    }else{
        write-warning "Could not find 'TemporaryForCopy' template."
    }
}

#endregion



Function Add-SpItemToPromotedLinkList {
    
    [CmdletBinding()]    param(
        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [string]        $ListName,

        [Parameter(Mandatory=$true)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [string]
        $Link,

        [Parameter(Mandatory=$false)]
        [string]
        $Description,

        [Parameter(Mandatory=$false)]
        [switch]
        $force

    )
    if (!($web)){
        $web = Get-SPWeb -Identity $Url
    }
    $LinkList = Get-SPList -Url -ListName $ListName
    if (($linkList.Items).title -contains $Name){
        if ($force){
            write-verbose "Link item already existing. Force parameter has been specified. Overwriting $($Name)."
            $NewLink = $LinkList.Items.Add()
            $NewLink['Title'] = $Name
            $NewLink['LinkLocation'] = $Link
            $NewLink['Description'] = $Description
            $Newlink.Update()
        }
        else{
            Write-warning "Item $($name) is already present on list $($ListName)."
            write-warning "use the 'force' parametmeter to overwrite the existing value. The item has not been set."
        }
    }else{
            $NewLink = $LinkList.Items.Add()
            $NewLink['Title'] = $Name
            $NewLink['LinkLocation'] = $Link
            $NewLink['Description'] = $Description
            $Newlink.Update()
    }
    

    $Return = Get-SpListItem -Url $Url -ListName $ListName -ItemName $Name
    if ($Return){
        write-verbose "Item $($Name) on list $($ListName) has been created successfully."
        Return $Return
    }else{
        Write-warning 'The Item has not been created.'
    }
}

#region rights management


Function New-SpSiteGroup{
    Param(
        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        [string]
        $Owner,

        [Parameter(Mandatory=$false)]
        [string]
        $GroupName,

        [Parameter(Mandatory=$false)]
        [ValidateSet('FullControl','Contribute','Read')]$Permissions,

        [Parameter(Mandatory=$false)]
        [string]
        $description
    )
    if (!($web)){
        $web = Get-SPWeb $Url 
    }
                                                                                     
                                                                                                                                                           
     $user = $web.EnsureUser($owner)   
     $Group = $web.SiteGroups.Add($GroupName,$user,$user,$description)

 #Setting correct permissions
    if ($Permissions -eq 'Full Control'){
        $role = $web.RoleDefinitions['Full Control']
    }Else{
        $role = $web.RoleDefinitions[$Permissions]
    }

     #$assignment = New-Object Microsoft.SharePoint.SPRoleAssignment($Group)
      #$Group.RoleDefinitionBindings.Add($role);
      #$spweb.RoleAssignments.Add($assignment)  
 $web.Update()

}

Function New-SpDefaultSiteGroups {
    Param(
        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        [string]
        $Owner1,

        [Parameter(Mandatory=$False)]
        [string]
        $Owner2

    )
    
    if (!($web)){
        $web = Get-SPWeb $Url 
    }

    if ($Owner1){
        $User1 = $web.EnsureUser($Owner1)
    }

    if ($Owner2){
        $User2 = $web.EnsureUser($Owner2)
    }

    if ($User2){
        $web.CreateDefaultAssociatedGroups($User1.LoginName,$User2.loginName, $web.title)
    }else{
        $web.CreateDefaultAssociatedGroups($User1.LoginName,$null, $web.title)
    }
     
    $web.update()
}

Function Get-SpGroupMembers {
    [cmdletBinding()]
     Param(
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,


        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.SPGroup]
        $GroupObject,

        [Parameter(Mandatory=$false)]
        $GroupName,

        [Parameter(Mandatory=$false)]
        [String]
        $GroupID
    )


        if (!($Web)){
        $Web = Get-SPWeb $Url
    }
    if($GroupObject){
        $Return = $GroupObject.Users
    }
    elseif ($GroupName){
        $Group = Get-SpGroup -Web $Web -Name $GroupName
        $return = $Group.users
    }elseif($GroupID){
        $Group = Get-SpGroup -Web $Web -id $GroupID
        $return = $Group.users
    }else{
        write-warning "Please input value"
    }

    return $Return
}

Function Get-SpSiteGroup {
    [cmdletBinding()]
     Param(
        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        $Name,

        [Parameter(Mandatory=$false)]
        [String]
        $id
    )

    if (!($Web)){
        $Web = Get-SPWeb $Url
    }

    if ($Name){
        $return = $Web.SiteGroups[$Name]
    }elseif($ID){
        $return = $Web.SiteGroups | where-object {$_.ID -eq $ID}
    }else{
        $return = $Web.SiteGroups
    }

    return $return
}

Function Get-SpGroup {
    [cmdletBinding()]
     Param(
        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        $Name,

        [Parameter(Mandatory=$false)]
        [String]
        $id
    )

    if (!($Web)){
        $Web = Get-SPWeb $Url
    }

    if ($Name){
        $return = $Web.Groups[$Name]
    }elseif($ID){
        $return = $Web.Groups | where-object {$_.ID -eq $ID}
    }else{
        $return = $Web.Groups
    }

    return $return
}

Function Add-SpUSerToSpGroup{
        [cmdletBinding()]
     Param(
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [Microsoft.SharePoint.SPUser]
        $UserObject,

        [Parameter(Mandatory=$false)]
        [String]
        $GroupName,

        [Parameter(Mandatory=$false)]
        [String]
        $GroupID
    )

    if (!($Web)){
        $Web = Get-SPWeb $Url
    }

    if ($GroupName){
        write-verbose "Adding user $($userObject.Displayname) to groupname $($GroupName)"
        $Group = Get-SpGroup -Web $Web -Name $GroupName
        $Group.Adduser($UserObject)
    }elseif($GroupID){
        write-verbose "Adding user $($userObject.Displayname) to groupname $($GroupID)"
        $Group = Get-SpGroup -Web $Web -id $GroupID
        $Group.Adduser($UserObject)
    }else{
        write-verbose "Not enough params"
    }

}

#endregion


#region propertyBag

Function Get-SPPropertyBagValue{

    [CmdletBinding()]    param(

        [Parameter(ParameterSetName='url',Mandatory=$true)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        $Name

    )
    
    If(!($web)){
        $Web= Get-SPWeb -Identity $Url
    }
    #$SpSite = Get-SPSite -Identity $Url
    #$Root = $SpSite.Rootweb
    if ($Name){
        $Return = $Web.AllProperties[$Name]
    }else{
        $Return = $Web.AllProperties
    }
    write-verbose "Returning property bag information from site $($Url)."
    Return $Return

}

Function New-SPPropertyBagValue{

    [CmdletBinding()]    param(

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [String]
        $Name,

        [Parameter(Mandatory=$true)]
        [String]
        $Value,

        [Parameter(Mandatory=$true)]
        $InputObject,

        [switch]$Force
        
    )
    
    if (!($web)){
       $Web = Get-SPWeb -Identity $Url
    }
        if ($InputObject){
            foreach ($Property in $InputObject){
                write-verbose "Adding propertybag $($Property.Name) with value $($Property.Value)"
                $Web.AllProperties.Add($Property.Name,$Property.Value)
                
            }
        }else{
            if ($Web.AllProperties.ContainsKey($Name)){
                if($Force){
                    write-verbose "Forcing property $($Name) with value $($Value) "
                    $Web.AllProperties[$Name] = $Value
                }else{
                    write-warning "Property $($Name) is already set. Use -Force to override or Set-SpPropertyBag"
                }
                
            }else{
                $Web.AllProperties.Add($Name,$Value)
            }
        }
        
        $Web.Update()

        if ($InputObject){
            return $Web.AllProperties 
        }Else{
            return $Web.AllProperties[$Name] 
        }
    

}

Function Set-SpPropertyBagValue {
    [CmdletBinding()]    param(

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [String]
        $Name,

        [Parameter(Mandatory=$true)]
        [String]
        $Value,

        [Parameter(Mandatory=$true)]
        $InputObject,

        [switch]$Force
        
    )
    
    if (!($web)){
       $Web = Get-SPWeb -Identity $Url
    }
        if ($InputObject){
            foreach ($Property in $InputObject){
                write-verbose "Adding propertybag $($Property.Name) with value $($Property.Value)"
                $Web.AllProperties.Add($Property.Name,$Property.Value)
                
            }
        }else{
            if ($Web.AllProperties.ContainsKey($Name)){
                
                    write-verbose "Setting property $($Name) with value $($Value) "
                    $Web.AllProperties[$Name] = $Value
                
                
            }else{
                Write-warning "Could not find property $($Name). Use New-SpPropertyBag to set a new value,otherwise specify and existing name."
            }
        }
        
        $Web.Update()

        return $Web.AllProperties[$Name]
    

}

#endregion



Function Get-SharepointVersion{
    [CmdletBinding()]    param(
        
    )

    $Farm = Get-SPFarm
    #$Farm.Products
    $Version = $Farm.BuildVersion.Tostring()

    switch ($version)
    {
        '15.0.4517' {'June 2013 CU'}
        '15.0.4505' {'April 2013 CU'}
        '15.0.4481' {'March Update'}
        '15.0.4420' {'RTM'}
        Default {write-warning "Could not identify the current build $($version)"}
    
    
    }
                       
}

#Latest


Function Get-SpContentType {
    [cmdletBinding()]
     Param(
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        [String]
        $Name,

        [Parameter(Mandatory=$false)]
        [String]
        $id
    )
    Begin{
        if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    }
    Process{
        if($Name){
            $Return = $web.ContentTypes | where-object {$_.name -eq $Name}
        }elseif($id){
            $Return = $Web.ContentTypes | where-object {$_.id -eq $id}
            write-verbose "ID"
        }else{
            $Return = $web.ContentTypes
        }
    }
    End{
        return $Return
    }
}

Function Add-SpFieldToContentType {

    [CmdletBinding()]
    param
    (
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [parameter(Mandatory=$false)]
        [String]
        $FieldName,

        [parameter(Mandatory=$false)]
        [String]
        $ContentTypeName,

        [parameter(Mandatory=$false)]
        [String]
        $ContentTypeid

    )
    
    # Get the Site where the Content Type will be created
    if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }

    $Field = $web.Fields[$FieldName]
    if ($ContentTypeName){
        $ContentType = Get-SpContentType -Web $Web -Name $ContentTypeName
    }
    else{
         $ContentType = Get-SpContentType -Web $Web -id $ContentTypeid
         $ContentType = $Web.ContentTypes | where-object {$_.id -eq $ContentTypeid}
    }
    $link = new-object Microsoft.SharePoint.SPFieldLink $Field
    #$contetype
    $ContentType.FieldLinks.Add($link)
    $ContentType.Update($true)
}

function New-SpcontentType{

    [CmdletBinding()]
    param
    (
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [parameter(Mandatory=$false)]
        [String]
        $Name,

        [parameter(Mandatory=$false)]
        [String]
        $GroupName,

        [Parameter(Mandatory=$false)]
        [String]
        $ParentName,

        [Parameter(Mandatory=$false)]
        [String]
        $ParentID,

        [parameter(Mandatory=$false)]
        [String]
        $Description
    )
    
    # Get the Site where the Content Type will be created
    if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    #$List = Get-SPList -Url $Url -ListName $ParentName
     if($Name){
            $Parent = $web.ContentTypes | ? {$_.name -eq $ParentName}
        }elseif($ParentID){
            $Parent = $web.ContentTypes | ? {$_.id -eq $ParentID}
        }else{
           write-warning 'please either use ParentID or ParentName to specefiy a parent type.'
           Break
        }

   
        #$web.AvailableContentTypes["Document"] somehow did not work (always empty...)
    $documentCTId = New-Object Microsoft.SharePoint.SpContentTypeId $ParentID
    $ct =  $Web.AvailableContentTypes | ?{$_.id -eq $documentCTId}
    $contentType =  New-Object Microsoft.SharePoint.SPContentType -ArgumentList @($ct,$Web.ContentTypes,$Name)
    $contentType.Group = $GroupName
    $contentType.Description = $Description
    $Web.ContentTypes.Add($contentType)
    $Web.Update()
    
    # Dispose of the Web and Site objects and close the loop
    $Web.Dispose()
}

Function Remove-SpContentType{
    Param(
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        [String]
        $Name,

        [Parameter(Mandatory=$false)]
        [String]
        $id
    )
    Begin{
        if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    }
    Process{
        if($Name){
            $ContentType = Get-SpContentType -Web $Web -Name $Name
            $ContentType.delete()
        }elseif($id){
            $ContentType = Get-SpContentType -Web $Web -id $id
            $ContentType.delete()
        }else{
            write-warning 'No corresponding content type has been found.'
            
        }

        
    }
    End{
        
        $Web.Dispose()
    }
}

#region fields

Function New-SpField {

    [CmdletBinding()]
    param
    (
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [parameter(Mandatory=$false)]
        [String]
        $FieldName,

        [parameter(Mandatory=$false)]
        [ValidateSet('Invalid',` 
        'Integer',` 
        'Text',` 
        'Note',` 
        'DateTime',` 
        'Counter',` 
        'Choice',` 
        'Lookup',` 
        'Boolean',` 
        'Number',` 
        'Currency',` 
        'URL',` 
        'Computed',` 
        'Threading',` 
        'Guid',` 
        'MultiChoice',` 
        'GridChoice',` 
        'Calculated',` 
        'File',` 
        'Attachments',` 
        'User',` 
        'Recurrence',` 
        'CrossProjectLink',` 
        'ModStat',` 
        'Error',` 
        'ContentTypeId',` 
        'PageSeparator',` 
        'ThreadIndex',` 
        'WorkflowStatus',` 
        'AllDayEvent',` 
        'WorkflowEventType',` 
        'Geolocation',` 
        'OutcomeChoice',` 
        'MaxItems')]
        [String]
        $Type,

        [parameter(Mandatory=$false)]
        [switch]
        $Mandatory,

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Parameter(ParameterSetName='Choice',Mandatory=$true)]
        [array]$Choices,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [Parameter(ParameterSetName='Choice',Mandatory=$true)]
        $DefaultChoice
        
    )

    Begin{
        if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    }
    Process{
        $NewField = $web.Fields.Add($FieldName,$Type,$Mandatory)
        
        switch ($Type){
            'Choice'{
                    foreach ($c in $Choices){
                        $Web.Fields[$FieldName].addChoice($c)
                    }
                    $Web.Fields[$FieldName].DefaultValue = $DefaultChoice
                    $Web.Fields[$FieldName].Update()
                    
            }
            default{}
        }
    }
    End{
        $Web.Update()
        $web.Dispose()
        
    }
    
}

Function Get-SpField {

    [CmdletBinding()]
    param
    (
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [parameter(Mandatory=$false)]
        [String]
        $Name,

        [Parameter(Mandatory=$false)]
        [String]
        $id

    )
    Begin{
        if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    }
    Process{
        if($Name){
            $Return = $web.Fields | where-object {$_.title -eq $Name}
        }elseif($id){
            $Return = $web.Fields | where-object {$_.id -eq $id}
        }else{
            $Return = $web.Fields
        }
    }
    End{
        return $Return
    }

}

Function Remove-SpField{

    [CmdletBinding()]
    param
    (
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [parameter(Mandatory=$false)]
        [String]
        $Name,

        [Parameter(Mandatory=$false)]
        [String]
        $id

    )
    Begin{
        if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    }
    Process{
        if($Name){
            $Field = Get-SpField -Web $Web -Name $Name
        }elseif($id){
            $Field =  Get-SpField -Web $Web -id $id
        }
    }
    end{
        $Field.Delete()
    }
}

#endregion 

Function Add-SpContentTypeToLibrary{
    [CmdletBinding()]
        param
    (
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [parameter(Mandatory=$false)]
        [String]
        $LibraryName,

        [parameter(Mandatory=$false)]
        [String]
        $ContentTypeName,

        [Parameter(Mandatory=$false)]
        [String]
        $id

    )
    Begin{
        if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    }
    Process{
        $docLibrary = Get-SPList -Url $Url -ListName $LibraryName
        $ContentType = Get-SpContentType -Web $Web -Name $ContentTypeName

        $docLibrary.ContentTypesEnabled = $true
        $docLibrary.Update()

        #Add site content types to the list
        
        $Return = $docLibrary.ContentTypes.Add($ContentType)
        $docLibrary.Update()
    }
    End{
        Return $Return
    }


}

Function Get-SpSiteGroup{
    [CmdletBinding()]
    param
    (
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        [string]
        $Name

    )
    Begin{
        if (!($web)){
            $web = Get-SPWeb -Identity $Url
        }
    }
    Process{
        if($Name){
            $Return =  $Web.SiteGroups[$Name]
            
        }Else{
           $Return = $Web.SiteGroups
        }
    }
    End{
        Return $Return
    }
}

Function Add-LinkToQuickLinks{
    [CmdletBinding()]
    Param(
        
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [string]
        $LinkTitle,

        [Parameter(Mandatory=$true)]
        [string]
        $HyperLink,

        [Parameter(Mandatory=$false)]
        [switch]$AddAsFirst
    )

    if (!($web)){
        $Web = Get-SPWeb -Identity $Url
    }

    $linkNode = New-Object Microsoft.SharePoint.Navigation.SPNavigationNode($LinkTitle, $HyperLink)
    
    if ($AddAsFirst){
        $Web.Navigation.QuickLaunch.AddAsFirst($linkNode)
        $web.update()
        
    }else{
        $Web.Navigation.QuickLaunch.AddAsLast($linkNode)
        $web.update()
    }
    
    
}

#region Listitems --> ok

function Get-SpListItem{
<#
.Synopsis
   returns  one or more item from a specefic list of a sharepoint site
.DESCRIPTION
   if used in conjunction with 
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        # Param2 help description
        [Parameter(Mandatory=$true)]
        [String]
        $ListName,

        [Parameter(Mandatory=$false)]
        $ItemName

    )
    Begin {}
    Process{

        if (!($Web)){
            $ListItems = (Get-SPList -web $Web -ListName $ListName).items
        }else{
            $ListItems = (Get-SPList -Url $Url -ListName $ListName).items
        }

        if ($ItemName){
            $Return = $ListItems | ? {$_.name -eq $ItemName}
        }else{
            $Return = $ListItems
        }
    }
    End{
        Return $Return
    }
}

Function Set-SpItemFieldValue {
[cmdletBinding()]
Param(
    [Parameter(ParameterSetName='url',Mandatory=$false)]
    [string]$Url,

    [Parameter(ParameterSetName='object',Mandatory=$false)]
    [Microsoft.SharePoint.SPWeb]
    $Web,

    [Parameter(Mandatory=$false)]
    [String]
    $ListName,

    [Parameter(Mandatory=$false)]
    [String]
    $ItemName,

    [Parameter(Mandatory=$false)]
    [String]
    $FieldName,

    [Parameter(Mandatory=$true)]
    [String]
    $FieldValue,

    [Parameter(ParameterSetName='list',Mandatory=$false)]
    [Microsoft.SharePoint.SPListItem]
    $InputObject
)

    If (!($InputObject)){

        $InputObject = Get-SpListItem -Url $Url -ListName $ListName -ItemName $ItemName
    }
    
        $InputObject[$Fieldname] = $FieldValue
        $InputObject.Update()

    
}

Function Get-SpItemFieldValue {
[cmdletBinding()]
Param(
    
    [Parameter(ParameterSetName='url',Mandatory=$false)]
    [string]$Url,

    [Parameter(ParameterSetName='object',Mandatory=$false)]
    [Microsoft.SharePoint.SPWeb]
    $Web,

    [Parameter(Mandatory=$false)]
    [String]
    $ListName,

    [Parameter(Mandatory=$false)]
    [String]
    $ItemName,

    [Parameter(Mandatory=$true)]
    [String]
    $FieldName,

    [Microsoft.SharePoint.SPListItem]$InputObject
)

    
    If (!($InputObject)){

        if ($web){
            $InputObject = Get-SpListItem -Web $Web -ListName $ListName -ItemName $ItemName
        }else{
            $InputObject = Get-SpListItem -Url $Url -ListName $ListName -ItemName $ItemName
        }
    }
    
  return $InputObject[$FieldName]
        

    
}

Function Remove-SpListItem {
        [CmdletBinding()]    param(

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [string]
        $ListName,

        [Parameter(Mandatory=$true)]
        [string]
        $ItemName
    )

    $Item = Get-SpListItem  -Url $Url -ListName $ListName -ItemName $ItemName
    if ($Item){
        if ($item.count -eq 1){
            $Item.delete()
            write-verbose "Delete item $($item)"
        }else{
            write-warning "Attempt to delete more then one item from list $($Listname) of url $($url)"
        }
    }else{
        Write-warning "Could not find item $($Name) in list $($listname) at URL $($Url). Please verify the information and try again."
    }
}

#endregion

#region webparts --> ok

Function Get-SpWebPart {
    

    [CmdletBinding()]    param(

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]$Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$false)]
        [String]
        $title,

        [Parameter(Mandatory=$false)]
        [String]
        $ID
    )

    if (!($Web)){
        $Site = Get-SPSite -Identity $Url
        $web=$site.Openweb() 
    }

$WebPartManager = $web.GetLimitedWebPartManager($web.url + $web.RootFolder.WelcomePage, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)

if ($Title){
    write-verbose "Returning webPart with title $($Title)"
    $Return = $WebPartManager.WebParts | Where-Object {$_.Title -eq $Title}
}elseif($ID){
    write-verbose "Returning webPart with ID $($ID)"
    $Return = $WebPartManager.WebParts | Where-Object {$_.ID -eq $ID}
}else{
    write-verbose 'Returning all webParts.'
    $Return = $WebPartManager.WebParts
}

<#
    switch ($PSBoundParameters.Keys)
    {
        'Title' {
            write-verbose "Returning webPart with title $($Title)"
            $Return = $WebPartManager.WebParts | ? {$_.Title -eq $Title} ;break
            }
        'ID' {
            write-verbose "Returning webPart with ID $($ID)"
            $Return = $WebPartManager.WebParts | ? {$_.ID -eq $ID} ;break
        }
        Default {
            
        }
    }
#>
write-verbose "returning: $Return"
return $Return
}

Function New-SpListWebPart{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]
        $Title,

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web

    )

    if (!($web)){
        $Web = Get-SPWeb -Identity $Url
    }
    #Creating New GUID
        $lvwpGuid  = [System.Guid]::NewGuid().ToString()
        $WPID = 'g_' + $lvwpGuid.Replace('-','_')
   
    #Creating New WebPart
        $WebPart = New-Object Microsoft.SharePoint.WebPartPages.XsltListViewWebPart
        $webpart.ID = $WPID
        $WebPart.ListUrl = 'Lists/Links'
        $webpart.Title = $Title   
        $webpart.Visible = $true
        $webpart.PageType = 'PAGE_NORMALVIEW'   
        $webpart.ViewContentTypeId = '0x'
        $webpart.ChromeType = 'None'
        $WebPartManager = $web.GetLimitedWebPartManager($web.url + '/' + $web.RootFolder.WelcomePage, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
        $WebPartManager.AddWebPart($webpart,'WPZ',0)
        $web.dispose
        return $lvwpGuid
        
}

Function Remove-SpWebPart {
    [CmdletBinding()]    param(
        

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [string]
        $Name,

        [Parameter(Mandatory=$true)]
        [string]
        $ID,

        [Parameter(Mandatory=$false)]
        [switch]$RemoveAll
    )


        if (!($web)){
            $Web = Get-SPWeb -Identity $Url
        }

        $WebPartManager = $web.GetLimitedWebPartManager($web.url + $web.RootFolder.WelcomePage, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)



        if ($RemoveAll){
            write-verbose "Removing All Web Parts from page $($Url)"
           $WebParts = Get-SpWebPart -Url $Url | select Title,ID
           foreach ($W in $WebParts){
                write-verbose "Deleting  Web Part $($w.title) with ID $($w.ID) from page $($Url)"
                Remove-SpWebPart -Url $Url -ID $W.ID
           }

        }else{

                if ($Name){
                    $WebPart = Get-SpWebPart -Url $Url -title $Name
                }else{
                    $WebPart = Get-SpWebPart -Url $Url -ID $ID
                    }
            
                if ($WebPart){
                    write-verbose 'Deleting webpart with ID'
                    $WebPartManager.DeleteWebPart($WebPartManager.WebParts[$webpart.ID])
   
                }else{
                    Write-warning "Web part $($Name) with ID $($ID) could not be found on the page $($Url)."
                }
        
            }
        }

Function Add-SpWebPart {

    [CmdletBinding()]    param(

        [Parameter(ParameterSetName='url',Mandatory=$false)]
        [string]
        $Url,

        [Parameter(ParameterSetName='object',Mandatory=$false)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory=$true)]
        [string]
        $Name,

        [Parameter(Mandatory=$true)]
        [string]
        $ID
    )

     if (!($web)){
        $Web = Get-SPWeb -Identity $Url
    }

     # Create  New object of type XsltListViewWebPart

        $webpart = New-Object Microsoft.SharePoint.WebPartPages.XsltListViewWebPart
  
        $List = Get-SPList -Url $Url -ListName $Name
        #$List=$web.Lists.TryGetList($Name)
       
        #Assign the Webart Listid with Current List Id i.e.. "list id that needs to be added in the page"
            $lvwpGuid  = [System.Guid]::NewGuid().ToString()
            $WPID = 'g_' + $lvwpGuid.Replace('-','_')
            
            $webpart.ID = $WPID
            $webpart.ListId = $List.ID
            
        #Setting WebPart
            $WebPartManager = $web.GetLimitedWebPartManager($web.url + $web.RootFolder.WelcomePage, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
            $WebPartManager.AddWebPart($webpart,'WPZ',0)


        #get a reference to wiki page item
            $wpPage = $web.GetFile($web.url + $web.RootFolder.WelcomePage)
            $item = $wpPage.Item

$wikiContent = @"  
<div class="ms-rtestate-read ms-rte-wpbox" contenteditable="false" style="float:left;width:30&%;min-width:300px;">  
 <div class="ms-rtestate-notify  ms-rtestate-read $($lvwpGuid)" id="div_$($lvwpGuid)" unselectable="on"></div>  
 <div id="$($lvwpGuid) " unselectable="on" style="display: none"></div>  
</div>  
  
"@  


        $item['WikiField'] += $wikicontent  


        $item.Update()

}

#endregion