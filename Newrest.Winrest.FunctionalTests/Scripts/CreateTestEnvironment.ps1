Param(
    $codePays,
    $pwdAdminDb,
    $dbOrigName
 )

$OAuthClientId = "4868ee8b-5987-4284-a808-d5b861cfebcc";
$storageConnectionString = "";

$sharedRg = "Winrest-DEV-RG";
$rgSharedConsuption = "Winrest-DEV-RG";
$azFuncConsuptionPrint = "winrest-dev-print-func";
$azFuncConsuptionPrintService = "winrest-dev-printservice-func";
$azFuncPrintServiceOnPlan = "winrest-dev-printserviceonplan-func";
$azFuncPrintOnPlan = "winrest-dev-printonplan-func";

#Redis
$cacheRedisName = "winrest-internal-shared-cache";
$cacheRedisKey = "aK027e8cMJTxyJh1pvbrlu7wVJSEdcDkjPh7CWGoJ7A=";
$cacheRedisConnectionString = "$($cacheRedisName).redis.cache.windows.net:6380,password=$($cacheRedisKey),ssl=True,abortConnect=False,connectTimeout=10000";

#Storage TL
$tlStorageAccount = "winrestdevstorage";
$tlConnectionString = "DefaultEndpointsProtocol=https;AccountName=$($tlStorageAccount);AccountKey=$($tlStaKey);EndpointSuffix=core.windows.net";


$sharedInternalRg = "Winrest-Internal-Shared-RG";
$AppServicePlanName = "Winrest-Internal-Shared-Plan";
$sqlServerAdminLogin = "sa_winrest";
#$sqlServerAdminPassword = "$(SqlServerPassword)";

#Origin Database
$dbOrigRG = "Winrest-ES-RG";
$dbOrigServer = "winrest-es-sqlserver";
#$dbOrigName = "winrest-br-db";

#Destination Database
$dbDestRG = "Winrest-Internal-Shared-RG";
$dbDestServer = "winrest-internal-shared-sqlserver";
$dbDestPool = "winrest-internal-shared-sqlpool";
$dbDestName = "winrest-$($codePays.ToLower())-db";

#ResourceGroup
$rgName = "Winrest-$($codePays)-RG";
$rgLocation = "northeurope";

#Storage
$storageName = "winrest$($codePays.ToLower())storage";
$storageLocation = "northeurope";

#AppService
$appName = "Winrest-$($codePays)-App";
$appLocation = "northeurope";

function LogMessage($message)
{
	$timestamp = Get-Date -format "yyyy-MM-dd_HH:mm.ss";
	echo ($timestamp + " " + $message)
}

#Copy database to the pool
if([string]::IsNullOrEmpty($dbDestPool))
{
	LogMessage ("Copying " + $dbOrigName + " to " + $dbDestName + " in ResourceGroup " + $dbDestRG + " Server " + $dbDestServer);
	New-AzSqlDatabaseCopy   -ResourceGroupName $dbOrigRG -ServerName $dbOrigServer -DatabaseName $dbOrigName `
							-CopyResourceGroupName $dbDestRG -CopyServerName $dbDestServer -CopyDatabaseName $dbDestName;
}
else
{
	LogMessage ("Copying " + $dbOrigName + " to " + $dbDestName + " in ResourceGroup " + $dbDestRG + " Server " + $dbDestServer + " in ElasticPool " + $dbDestPool);
	New-AzSqlDatabaseCopy   -ResourceGroupName $dbOrigRG -ServerName $dbOrigServer -DatabaseName $dbOrigName `
							-CopyResourceGroupName $dbDestRG -CopyServerName $dbDestServer -CopyDatabaseName $dbDestName -ElasticPoolName $dbDestPool;
}

$dataBase = Get-AzSqlDatabase -ResourceGroupName $dbDestRG -DatabaseName $dbDestName -ServerName $dbDestServer -ErrorAction SilentlyContinue
if($null -eq $dataBase)
{
    $cnx = "Server=tcp:$($dbDestServer).database.windows.net,1433;Initial Catalog=$($dbDestName);Persist Security Info=False;User ID=$($sqlServerAdminLogin);Password=$($pwdAdminDb);MultipleActiveResultSets=True;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;Max Pool Size=200;"

    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $cnx
    $SqlConnection.Open()

    $sqlcmd = new-object "System.data.sqlclient.sqlcommand"
    $sqlcmd.connection = $SqlConnection
    $sqlcmd.CommandText = "UPDATE [dbo].[ApplicationSettings] SET [Value] = 'AzureFileAPI' WHERE Name = 'FileService'"
    $rowUpdated = $sqlcmd.ExecuteNonQuery()
    $sqlcmd.CommandText = "UPDATE [dbo].[ApplicationSettings] SET [Value] = 'DEV/TL-FILES-API' WHERE Name = 'WinrestTLApiFolder'"
    $rowUpdated = $sqlcmd.ExecuteNonQuery()
    $sqlcmd.CommandText = "UPDATE [dbo].[ApplicationSettings] SET [Value] = 'https://tldev.newrest.eu/' WHERE Name = 'WinrestTLUrl'"
    $rowUpdated = $sqlcmd.ExecuteNonQuery()
    $sqlcmd.CommandText = "UPDATE [dbo].[Settings] SET [Value] = '1' WHERE KeySetting = 'NewVersionPrint'"
    $rowUpdated = $sqlcmd.ExecuteNonQuery()

    $SqlConnection.Close()
}
else
{
    $cnx = "Server=tcp:$($dbDestServer).database.windows.net,1433;Initial Catalog=$($dbDestName);Persist Security Info=False;User ID=$($sqlServerAdminLogin);Password=$($pwdAdminDb);MultipleActiveResultSets=True;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;Max Pool Size=200;"
}

#Create ResourceGroup
$rg = Get-AzResourceGroup -Name $rgName -ErrorAction SilentlyContinue
# Creation du groupe de ressource
if($null -eq $rg)
{
    New-AzResourceGroup -Name $rgName -Location $rgLocation
}

#Create Storage Account
$sa = Get-AzStorageAccount -ResourceGroupName $rgName -Name $storageName -ErrorAction SilentlyContinue
if($null -eq $sa)
{
# Creation du compte de stockage
    New-AzStorageAccount -ResourceGroupName $rgName -Name $storageName -Location $storageLocation -SkuName Standard_RAGRS -Kind StorageV2 -AccessTier Hot -EnableHttpsTrafficOnly $true

    $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $rgName -AccountName $storageName).Value[0]
    $storageConnectionString = "DefaultEndpointsProtocol=https;AccountName=$($storageName);AccountKey=$($storageKey);EndpointSuffix=core.windows.net"
    $storageContext = New-AzStorageContext  -StorageAccountName $storageName -StorageAccountKey $storageKey

    # Creation des partages
    New-AzStorageShare -Name "upload" -Context $storageContext
    New-AzStorageShare -Name "temp" -Context $storageContext
    New-AzStorageShare -Name "report" -Context $storageContext
}
else
{
	$storageKey = (Get-AzStorageAccountKey -ResourceGroupName $rgName -AccountName $storageName).Value[0]
    $storageConnectionString = "DefaultEndpointsProtocol=https;AccountName=$($storageName);AccountKey=$($storageKey);EndpointSuffix=core.windows.net"
    $storageContext = New-AzStorageContext  -StorageAccountName $storageName -StorageAccountKey $storageKey
}

# Verification de l'existance de la webapp dans le groupe de ressource
$webApp = Get-AzWebApp -Name $appName -ResourceGroupName $rgName -ErrorAction SilentlyContinue

if($null -eq $webApp){
#Verification de l'existance de la webApp dans le groupe de ressource partage

    $webApp = Get-AzWebApp -Name $appName -ResourceGroupName $sharedInternalRg -ErrorAction SilentlyContinue
    if($null -eq $webApp){
        
        $webApp = New-AzWebApp -Name $appName -AppServicePlan $AppServicePlanName -ResourceGroupName $sharedInternalRg -Location $appLocation -ErrorAction Stop
      
    }
    Write-Progress -Activity "WebSite" -Status "Ajout des parametres d'application" -PercentComplete $(5/7*100)

    $resourceName = $appName + "/AppSettings"


    $hash=@{}
    #$hash.Add("APPINSIGHTS_INSTRUMENTATIONKEY", $instrumentationKey)
    $hash.Add('APPINSIGHTS_PROFILERFEATURE_VERSION',"disabled")
    $hash.Add('APPINSIGHTS_SNAPSHOTFEATURE_VERSION',"disabled")
    $hash.Add('ApplicationInsightsAgent_EXTENSION_VERSION',"~2")
    $hash.Add('DiagnosticServices_EXTENSION_VERSION',"disabled")
    $hash.Add('InstrumentationEngine_EXTENSION_VERSION',"disabled")
    $hash.Add('XDT_MicrosoftApplicationInsights_BaseExtensions',"disabled")
    $hash.Add('XDT_MicrosoftApplicationInsights_Mode', "default")
    $hash.Add('OAuth:Client:ClientId',$OAuthClientId)
    $hash.Add('OAuth:Client:Uri',"https://$($appName.ToLower()).azurewebsites.net")
    $hash.Add('Data:Azure:FileAPI:AzureStorage', $storageConnectionString)
    $hash.Add('Data:Azure:FileAPI:AzureStorage_TL', $tlConnectionString)
    $hash.Add('Data:ConnectionStrings:AzureStorage', $storageConnectionString)
    $hash.Add('Data:ConnectionStrings:Redis', $cacheRedisConnectionString)
    $hash.Add('Data:DefaultConnection:ConnectionString', $cnx)
    $hash.Add('Data:ServerName', "WR$($codePays.ToUpper())")
    $hash.Add('WEBSITE_RUN_FROM_PACKAGE', "0")

    New-AzResource -PropertyObject $hash -ResourceName $resourceName -ResourceGroupName $sharedInternalRg -ResourceType "Microsoft.Web/Sites/config"  -Force -ErrorAction Stop

    Write-Progress -Activity "WebSite" -Status "Ajout du AlwaysOn et VirtualPath" -PercentComplete $(5/7*100)
    $WebAppPropertiesObject = @{
            "siteConfig" = @{
                    "AlwaysOn" = $true
                    "virtualApplications" = @(
                                @{ virtualPath = "/"; physicalPath = "site\wwwroot" },
                                @{ virtualPath = "/CustomerPortal"; physicalPath = "site\wwwroot\CustomerPortal" }
                    )
        }                   
    }

    Set-AzResource -ResourceName $appName  -ResourceGroupName $sharedInternalRg  -PropertyObject $WebAppPropertiesObject -Force -ResourceType 'microsoft.web/sites'

    Move-AzResource -DestinationResourceGroupName $rgName -ResourceId $webApp.Id -Confirm:$false -Force
}

Write-Progress -Activity "Fonctions Azure" -Status "Ajout de la configuration" -PercentComplete $(6/7*100)


# Plan de consommation
##Print
$PrintConsuptionName = $azFuncConsuptionPrint + "/AppSettings"

$PrintConsuption = Get-AzWebApp -ResourceGroupName $rgSharedConsuption -Name $azFuncConsuptionPrint
$PrintConsuptionSettings = $PrintConsuption.SiteConfig.AppSettings
$hash = @{}
foreach ($setting in $PrintConsuptionSettings) {
	if($setting.Name -NotLike "*$($codePays)*")
	{
		$hash[$setting.Name] = $setting.Value
	}
}
$hash.Add("WR$($codePays)_AzureStorage", $storageConnectionString)
$hash.Add("WR$($codePays)_ConnectionString", $cnx)
New-AzResource -PropertyObject $hash -ResourceName $PrintConsuptionName -ResourceGroupName $rgSharedConsuption -ResourceType "Microsoft.Web/Sites/config"  -Force -ErrorAction Stop

##Print Service

$PrintServiceConsuptionName = $azFuncConsuptionPrintService + "/AppSettings"

$PrintServiceConsuption = Get-AzWebApp -ResourceGroupName $rgSharedConsuption -Name $azFuncConsuptionPrintService
$PrintServiceConsuptionSettings = $PrintServiceConsuption.SiteConfig.AppSettings
$hash = @{}
foreach ($setting in $PrintServiceConsuptionSettings) {
	if($setting.Name -NotLike "*$($codePays)*")
	{
		$hash[$setting.Name] = $setting.Value
	}
}
$hash.Add("WR$($codePays)_AzureStorage", $storageConnectionString)
$hash.Add("WR$($codePays)_ConnectionString", $cnx)
$hash.Add("WR$($codePays)_PrintOutputDirectory", "printserviceoutput")
$hash.Add("WR$($codePays)_ServerName", "WR$($codePays)")
New-AzResource -PropertyObject $hash -ResourceName $PrintServiceConsuptionName -ResourceGroupName $rgSharedConsuption -ResourceType "Microsoft.Web/Sites/config"  -Force -ErrorAction Stop


# Plan Service
##Print
$PrintOnPlanName = $azFuncPrintOnPlan + "/AppSettings"

$PrintOnPlan = Get-AzWebApp -ResourceGroupName $sharedRg -Name $azFuncPrintOnPlan
$PrintOnPlanSettings = $PrintOnPlan.SiteConfig.AppSettings
$hash = @{}
foreach ($setting in $PrintOnPlanSettings) {
	if($setting.Name -NotLike "*$($codePays)*")
	{
		$hash[$setting.Name] = $setting.Value
	}
}
$hash.Add("WR$($codePays)_AzureStorage", $storageConnectionString)
$hash.Add("WR$($codePays)_ConnectionString", $cnx)
New-AzResource -PropertyObject $hash -ResourceName $PrintOnPlanName -ResourceGroupName $sharedRg -ResourceType "Microsoft.Web/Sites/config"  -Force -ErrorAction Stop

##PrintService
$PrintServiceOnPlanName = $azFuncPrintServiceOnPlan + "/AppSettings"

$PrintServiceOnPlan = Get-AzWebApp -ResourceGroupName $sharedRg -Name $azFuncPrintServiceOnPlan
$PrintServiceOnPlanSettings = $PrintServiceOnPlan.SiteConfig.AppSettings
$hash = @{}
foreach ($setting in $PrintServiceOnPlanSettings) {
	if($setting.Name -NotLike "*$($codePays)*")
	{
		$hash[$setting.Name] = $setting.Value
	}
}
$hash.Add("WR$($codePays)_AzureStorage", $storageConnectionString)
$hash.Add("WR$($codePays)_ConnectionString", $cnx)
$hash.Add("WR$($codePays)_PrintOutputDirectory", "printserviceoutput")
$hash.Add("WR$($codePays)_ServerName", "WR$($codePays)")
New-AzResource -PropertyObject $hash -ResourceName $PrintServiceOnPlanName -ResourceGroupName $sharedRg -ResourceType "Microsoft.Web/Sites/config"  -Force -ErrorAction Stop