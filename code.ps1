##
# Specify parameters
#
# $PathToStoreCsv                = full path to folder to store everything in, remember a \ on the end!
#                                - example: C:\tmp\Office365\
# $SortCsvByDepartment           = $true/$false whether to create different csv for each deparment found
# $GetOnlyLicensedUsers          = $true/$false whether to get licensed users only or all users
# $GetSpecificTenantByDomainName = full name of a verified domain, only retrieves users for that tenant
#                                - example: contoso.com
##

$PathToStoreCsv = "C:\tmp\o365\"
$SortCsvByDepartment = $false
$GetOnlyLicensedUsers = $false
$GetSpecificTenantByDomainName = ""

if(! (Test-Path $PathToStoreCsv) )
{
    Write-Warning -Message "$($PathToStoreCsv) does not exist, please create it before continuing"
    Exit   
}

##
# SKU overview
# List of SKUs, in order to retrieve friendly names of licenses assigned.
# Thanks to https://gist.github.com/jhochwald/c5eed9727e76f08b8481b0ef5767e119
##

$Sku = @{
      'AAD_BASIC'                        = 'Azure Active Directory Basic'
      'AAD_PREMIUM'                      = 'Azure Active Directory Premium'
      'RIGHTSMANAGEMENT'                 = 'Azure Active Directory Rights'
      'RIGHTSMANAGEMENT_FACULTY'         = 'Azure Active Directory Rights for Faculty'
      'RIGHTSMANAGEMENT_GOV'             = 'Azure Active Directory Rights for Government'
      'RIGHTSMANAGEMENT_STUDENT'         = 'Azure Active Directory Rights for Students'
      'MFA_STANDALONE'                   = 'Azure Multi-Factor Authentication Premium Standalone'
      'EMS'                              = 'Microsoft Enterprise Mobility + Security Suite'
      'EXCHANGESTANDARD_FACULTY'         = 'Exchange (Plan 1 for Faculty)'
      'EXCHANGESTANDARD_STUDENT'         = 'Exchange (Plan 1 for Students)'
      'EXCHANGEENTERPRISE_FACULTY'       = 'Exchange (Plan 2 for Faculty)'
      'EXCHANGEENTERPRISE_STUDENT'       = 'Exchange (Plan 2 for Students)'
      'EXCHANGEARCHIVE'                  = 'Exchange Archiving'
      'EXCHANGEARCHIVE_FACULTY'          = 'Exchange Archiving for Faculty'
      'EXCHANGEARCHIVE_GOV'              = 'Exchange Archiving for Government'
      'EXCHANGEARCHIVE_STUDENT'          = 'Exchange Archiving for Students'
      'EXCHANGESTANDARD_GOV'             = 'Exchange for Government (Plan 1G)'
      'EXCHANGEENTERPRISE_GOV'           = 'Exchange for Government (Plan 2G)'
      'EXCHANGEDESKLESS'                 = 'Exchange Kiosk'
      'EXCHANGEDESKLESS_GOV'             = 'Exchange Kiosk for Government'
      'EXCHANGESTANDARD'                 = 'Exchange Plan 1'
      'EXCHANGEENTERPRISE'               = 'Exchange Plan 2'
      'EOP_ENTERPRISE_FACULTY'           = 'Exchange Protection for Faculty'
      'EOP_ENTERPRISE_GOV'               = 'Exchange Protection for Government'
      'EOP_ENTERPRISE_STUDENT'           = 'Exchange Protection for Student'
      'EXCHANGE_ONLINE_WITH_ONEDRIVE_LITE' = 'Exchange with OneDrive for Business'
      'INTUNE_A'                         = 'Intune'
      'MCOIMP_FACULTY'                   = 'Lync (Plan 1 for Faculty)'
      'MCOIMP_STUDENT'                   = 'Lync (Plan 1 for Students)'
      'MCOSTANDARD_FACULTY'              = 'Lync (Plan 2 for Faculty)'
      'MCOSTANDARD_STUDENT'              = 'Lync (Plan 2 for Students)'
      'MCOVOICECONF'                     = 'Lync (Plan 3)'
      'MCOIMP_GOV'                       = 'Lync for Government (Plan 1G)'
      'MCOSTANDARD_GOV'                  = 'Lync for Government (Plan 2G)'
      'MCOVOICECONF_GOV'                 = 'Lync for Government (Plan 3G)'
      'MCOINTERNAL'                      = 'Lync Internal Incubation and Corp to Cloud'
      'MCOIMP'                           = 'Skype Plan 1'
      'MCOSTANDARD'                      = 'Skype Plan 2'
      'MCOVOICECONF_FACULTY'             = 'Lync Plan 3 for Faculty'
      'MCOVOICECONF_STUDENT'             = 'Lync Plan 3 for Students'
      'CRMENTERPRISE'                    = 'Microsoft Dynamics CRM Online Enterprise'
      'CRMSTANDARD_GCC'                  = 'Microsoft Dynamics CRM Online Government Professional'
      'CRMSTANDARD'                      = 'Microsoft Dynamics CRM Online Professional'
      'DMENTERPRISE'                     = 'Microsoft Dynamics Marketing Online Enterprise'
      'INTUNE_O365_STANDALONE'           = 'Mobile Device Management for Office 365'
      'OFFICE_BASIC'                     = 'Office 365 Basic'
      'O365_BUSINESS'                    = 'Office 365 Business'
      'O365_BUSINESS_ESSENTIALS'         = 'Office 365 Business Essentials'
      'O365_BUSINESS_PREMIUM'            = 'Office 365 Business Premium'
      'DEVELOPERPACK'                    = 'Office 365 Developer'
      'DEVELOPERPACK_GOV'                = 'Office 365 Developer for Government'
      'EDUPACK_FACULTY'                  = 'Office 365 Education for Faculty'
      'EDUPACK_STUDENT'                  = 'Office 365 Education for Students'
      'EOP_ENTERPRISE'                   = 'Office 365 Exchange Protection Enterprise'
      'EOP_ENTERPRISE_PREMIUM'           = 'Office 365 Exchange Protection Premium'
      'STANDARDPACK_GOV'                 = 'Office 365 for Government (Plan G1)'
      'STANDARDWOFFPACK_GOV'             = 'Office 365 for Government (Plan G2)'
      'ENTERPRISEPACK_GOV'               = 'Office 365 for Government (Plan G3)'
      'ENTERPRISEWITHSCAL_GOV'           = 'Office 365 for Government (Plan G4)'
      'DESKLESSPACK_GOV'                 = 'Office 365 for Government (Plan F1G)'
      'STANDARDPACK_FACULTY'             = 'Office 365 Plan A1 for Faculty'
      'STANDARDPACK_STUDENT'             = 'Office 365 Plan A1 for Students'
      'STANDARDWOFFPACK_FACULTY'         = 'Office 365 Plan A2 for Faculty'
      'STANDARDWOFFPACK_STUDENT'         = 'Office 365 Plan A2 for Students'
      'ENTERPRISEPACK_FACULTY'           = 'Office 365 Plan A3 for Faculty'
      'ENTERPRISEPACK_STUDENT'           = 'Office 365 Plan A3 for Students'
      'ENTERPRISEWITHSCAL_FACULTY'       = 'Office 365 Plan A4 for Faculty'
      'ENTERPRISEWITHSCAL_STUDENT'       = 'Office 365 Plan A4 for Students'
      'STANDARDPACK'                     = 'Office 365 Plan E1'
      'STANDARDWOFFPACK'                 = 'Office 365 Plan E2'
      'ENTERPRISEPACK'                   = 'Office 365 Plan E3'
      'ENTERPRISEWITHSCAL'               = 'Office 365 Plan E4'
      'DESKLESSPACK'                     = 'Office 365 Plan F1'
      'DESKLESSPACK_YAMMER'              = 'Office 365 Plan F1 with Yammer'
      'OFFICESUBSCRIPTION'               = 'Office Professional Plus'
      'OFFICESUBSCRIPTION_FACULTY'       = 'Office Professional Plus for Faculty'
      'OFFICESUBSCRIPTION_GOV'           = 'Office Professional Plus for Government'
      'OFFICESUBSCRIPTION_STUDENT'       = 'Office Professional Plus for Students'
      'WACSHAREPOINTSTD_FACULTY'         = 'Office Web Apps (Plan 1 For Faculty)'
      'WACSHAREPOINTSTD_STUDENT'         = 'Office Web Apps (Plan 1 For Students)'
      'WACSHAREPOINTSTD_GOV'             = 'Office Web Apps (Plan 1G for Government)'
      'WACSHAREPOINTENT_FACULTY'         = 'Office Web Apps (Plan 2 For Faculty)'
      'WACSHAREPOINTENT_STUDENT'         = 'Office Web Apps (Plan 2 For Students)'
      'WACSHAREPOINTENT_GOV'             = 'Office Web Apps (Plan 2G for Government)'
      'WACSHAREPOINTSTD'                 = 'Office Web Apps with SharePoint Plan 1'
      'WACSHAREPOINTENT'                 = 'Office Web Apps with SharePoint Plan 2'
      'ONEDRIVESTANDARD'                 = 'OneDrive for Business'
      'ONEDRIVESTANDARD_GOV'             = 'OneDrive for Business for Government (Plan 1G)'
      'WACONEDRIVESTANDARD'              = 'OneDrive for Business with Office Web Apps'
      'WACONEDRIVESTANDARD_GOV'          = 'OneDrive for Business with Office Web Apps for Government'
      'PARATURE_ENTERPRISE'              = 'Parature Enterprise'
      'PARATURE_ENTERPRISE_GOV'          = 'Parature Enterprise for Government'
      'POWER_BI_STANDARD'                = 'Power BI'
      'POWER_BI_STANDALONE'              = 'Power BI for Office 365'
      'POWER_BI_STANDALONE_FACULTY'      = 'Power BI for Office 365 for Faculty'
      'POWER_BI_STANDALONE_STUDENT'      = 'Power BI for Office 365 for Students'
      'PROJECTESSENTIALS'                = 'Project Essentials'
      'PROJECTESSENTIALS_GOV'            = 'Project Essentials for Government'
      'PROJECTONLINE_PLAN_1'             = 'Project Plan 1'
      'PROJECTONLINE_PLAN_1_FACULTY'     = 'Project Plan 1 for Faculty'
      'PROJECTONLINE_PLAN_1_GOV'         = 'Project Plan 1for Government'
      'PROJECTONLINE_PLAN_1_STUDENT'     = 'Project Plan 1 for Students'
      'PROJECTONLINE_PLAN_2'             = 'Project Plan 2'
      'PROJECTONLINE_PLAN_2_FACULTY'     = 'Project Plan 2 for Faculty'
      'PROJECTONLINE_PLAN_2_GOV'         = 'Project Plan 2 for Government'
      'PROJECTONLINE_PLAN_2_STUDENT'     = 'Project Plan 2 for Students'
      'PROJECTCLIENT'                    = 'Project Pro for Office 365'
      'PROJECTCLIENT_FACULTY'            = 'Project Pro for Office 365 for Faculty'
      'PROJECTCLIENT_GOV'                = 'Project Pro for Office 365 for Government'
      'PROJECTCLIENT_STUDENT'            = 'Project Pro for Office 365 for Students'
      'SHAREPOINTSTANDARD_FACULTY'       = 'SharePoint (Plan 1 for Faculty)'
      'SHAREPOINTSTANDARD_STUDENT'       = 'SharePoint (Plan 1 for Students)'
      'SHAREPOINTSTANDARD_YAMMER'        = 'SharePoint (Plan 1 with Yammer)'
      'SHAREPOINTENTERPRISE_FACULTY'     = 'SharePoint (Plan 2 for Faculty)'
      'SHAREPOINTENTERPRISE_STUDENT'     = 'SharePoint (Plan 2 for Students)'
      'SHAREPOINTENTERPRISE_YAMMER'      = 'SharePoint (Plan 2 with Yammer)'
      'SHAREPOINTSTANDARD_GOV'           = 'SharePoint for Government (Plan 1G)'
      'SHAREPOINTENTERPRISE_GOV'         = 'SharePoint for Government (Plan 2G)'
      'SHAREPOINTDESKLESS'               = 'SharePoint Kiosk'
      'SHAREPOINTSTANDARD'               = 'SharePoint Plan 1'
      'SHAREPOINTENTERPRISE'             = 'SharePoint Plan 2'
      'SMB_BUSINESS'                     = 'SMB Business'
      'SMB_BUSINESS_ESSENTIALS'          = 'SMB Business Essentials'
      'SMB_BUSINESS_PREMIUM'             = 'SMB Business Premium'
      'VISIOCLIENT'                      = 'Visio Pro for Office 365'
      'VISIOCLIENT_FACULTY'              = 'Visio Pro for Office 365 for Faculty'
      'VISIOCLIENT_GOV'                  = 'Visio Pro for Office 365 for Government'
      'VISIOCLIENT_STUDENT'              = 'Visio Pro for Office 365 for Students'
      'YAMMER_ENTERPRISE_STANDALONE'     = 'Yammer Enterprise Standalone'
      'RIGHTSMANAGEMENT_ADHOC'           = 'Azure Rights Management Service'
      'ENTERPRISEPREMIUM'                = 'Office 365 Enterprise E5'
      'TEAMS_COMMERCIAL_TRIAL'           = 'Microsoft Teams Commercial Cloud (User Initiated)'
    }


##
# Connect to Msol service
# Enter your Microsoft 365 Partner details or use a delegated account through your Microsoft 365 tenant.
###

$Msol = Connect-MsolService

##
# Retrieves tenants
# This command retrieves all tenants for which you have delegated access to.
##

if ( [string]::IsNullOrEmpty($GetSpecificTenantByDomainName) )
{
    $Tenants = Get-MsolPartnerContract -All | Select-Object TenantId
} else {
    $Tenants = Get-MsolPartnerContract -DomainName $GetSpecificTenantByDomainName | Select-Object TenantId
}

$TenantCount = 0
$TenantsCount = $Tenants.Count

Write-Output "Found $($TenantsCount) tenants!"

foreach ( $Tenant in $Tenants )
{
    # Get all domains
    # And set primary to first domain retrieved in list
    $Domains = Get-MsolDomain -TenantId $Tenant.TenantId
    $Domain = $Domains[0].Name

    # Tell user whats happening
    Write-Output "Gathering information for $($Domain).."

    # Check whether we retrieve only licensed users
    # If not, retrieve all of them
    if($GetOnlyLicensedUsers)
    {
       $Users = Get-MsolUser -TenantID $Tenant.TenantId -All | Where-Object -FilterScript {$_.isLicensed -eq $true} | Select-Object UserPrincipalName, DisplayName, FirstName, LastName, Department, Licenses, WhenCreated
    } else {
       $Users = Get-MsolUser -TenantID $Tenant.TenantId -All | Select-Object UserPrincipalName, DisplayName, FirstName, LastName, Department, Licenses, WhenCreated
    }

    # Tell user whats happening
    Write-Output "Found $($Users.Count) users. Working..."

    if($SortCsvByDepartment)
    {
        $PathToStoreCsvTemp = $PathToStoreCsv + $Domain

        If(!(Test-Path $PathToStoreCsvTemp))
        {
            New-Item -ItemType Directory -Force -Path $PathToStoreCsvTemp > $null
        }

        foreach ( $User in $Users )
        {
            if([string]::IsNullOrEmpty($User.Department))
            {
                $CsvName = "no_department.csv"
                $PathToStoreCsvFull = $PathToStoreCsvTemp + "\$($CsvName)"
                Write-Output "User: $($User.DisplayName), Department: no_department, CsvName: $($PathToStoreCsvFull)"
            } else {
                $CsvName = "$($User.Department).csv"
                $PathToStoreCsvFull = $PathToStoreCsvTemp + "\$($CsvName)"
                Write-Output "User: $($User.DisplayName), Department: $($User.Department), CsvName: $($PathToStoreCsvFull)"
            }

           If(!(Test-Path $PathToStoreCsvFull))
           {
               New-Item -Path $PathToStoreCsvTemp -Name $CsvName -Force -ItemType "file" > $null
           }

           # Get all licenses for the user in un-friendly name (SKUs)
           $MsolUserLicenses = ((Get-MsolUser -TenantId $Tenant.TenantId -UserPrincipalName $User.UserPrincipalName).Licenses).AccountSkuID
           $Licenses = "";

           # For each license SKU, convert them into friendly names
           foreach ( $License in $MsolUserLicenses)
           {
               $LicenseItem = $License -split ":" | Select-Object -Last 1
               $LicenseFriendlyName = $Sku.Item("$($LicenseItem)")
               $Licenses += "$($LicenseFriendlyName), "
           }

           # Prepare the csv data
           $Data = [PSCustomObject]@{
                       UserPrincipalName = $User.UserPrincipalName
                       DisplayName       = $User.DisplayName
                       FirstName         = $User.FirstName
                       LastName          = $User.LastName
                       Department        = $User.Department
                       Licenses          = $Licenses
                       WhenCreated       = $User.WhenCreated
                     }

           # Finally... export!
           $Data | Export-Csv -Append -Path $PathToStoreCsvFull -Encoding UTF8 -NoTypeInformation
         }
    } else {
 
        $CsvName = "$($Domain).csv"

        # Create new csv for tenant
        New-Item -Path $PathToStoreCsv -Name $CsvName -ItemType "file" > $null

        $PathToStoreCsvFull = $PathToStoreCsv + $CsvName

        foreach ( $User in $Users )
        {
           # Get all licenses for the user in un-friendly name (SKUs)
           $MsolUserLicenses = ((Get-MsolUser -TenantId $Tenant.TenantId -UserPrincipalName $User.UserPrincipalName).Licenses).AccountSkuID
           $Licenses = "";

           # For each license SKU, convert them into friendly names
           foreach ( $License in $MsolUserLicenses)
           {
               $LicenseItem = $License -split ":" | Select-Object -Last 1
               $LicenseFriendlyName = $Sku.Item("$($LicenseItem)")
               $Licenses += "$($LicenseFriendlyName), "
           }

           # Prepare the csv data
           $Data = [PSCustomObject]@{
                       UserPrincipalName = $User.UserPrincipalName
                       DisplayName       = $User.DisplayName
                       FirstName         = $User.FirstName
                       LastName          = $User.LastName
                       Department        = $User.Department
                       Licenses          = $Licenses
                       WhenCreated       = $User.WhenCreated
                     }

           # Finally... export!
           $Data | Export-Csv -Append -Path $PathToStoreCsvFull -Encoding UTF8 -NoTypeInformation
         }
     }

     $TenantCount++;
}
