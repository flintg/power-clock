[string]$useragent = "com.eagleflint.ps-clock/0.5.1 (PowerShell)";
[string]$appsettings = "WebClock,0.5.1,com.eagleflint.ps-clock";
[string]$requestEndpoint = "";
[string]$uri = "";
[string]$requestString = "";
[boolean]$isDefault = $false;
[string]$version = "";
[boolean]$success = $false;
[string]$msg = "";
[int32]$GMTOffset = 0;
[string]$Lat = "";
[string]$Lon = "";
[string]$credPath = ".\ps-clock.cred.xml";
[string]$settingsPath = ".\ps-Clock.json";
[boolean]$settingsChanged = $false;
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession;
$settingsJson = @"
{
	"url": "",
	"user": "",
	"db": "",
	"savepwd": null
}
"@
$settingsObject = $settingsJson | ConvertFrom-Json


function Main {
	LoadSettings;
	GetLocation;
	DoLogin;
	
	Write-Host "Success:$($script:success)";
	Write-Host "msg:$($script:msg)";
	
	if ($script:success -eq $true) {
		Write-Host "Login successful. Version $($script:version).";
		GetOffset;
		Write-Host "Ahoy! Ye be offset by $($script:GMTOffset) minutes.";
		WebClock;
	} elseif ($success -eq $false) {
		Write-Host "There was a problem while logging in. The server reports the following error: $($script:msg)";
		$retry = Read-Host -Prompt "Would you like to try again?  (yes/no)";
		if ($retry -eq "yes") {
			Main;
		}
	} else {
		Write-Host "Arr! There be a problem with ye request!  The cap'ns response be [$($response.StatusCode) - $($response.StatusDescription)]: $($response.Content)";
	}

	if ($script:settingsChanged -eq $true) {
		SaveSettings;
	}
}

function WebClock {
	$requestEndpoint = "EPclkTable";
	$uri = "$($script:settingsObject.url)$($requestEndpoint)";
	$requestString = "Action=Get";

	$response = Invoke-WebRequest -uri $uri -Method POST -Body $requestString -WebSession $script:session -UserAgent $script:useragent;

	$rjson = $response.Content | ConvertFrom-Json;
	  
	$cfg = $rjson.EPclk.Cfg;
	for($c=0;$c -le $cfg.Count-1;$c++) {
		switch ($cfg.Name[$c]){
			"UseDistributionTable"	{[boolean]$useDistributionTable=[boolean]$cfg.Value[$c]}
			"UseAccount"			{[boolean]$useAccount=[boolean]$cfg.Value[$c]}
			"AccountPicture"		{[string]$accountPicture=[string]$cfg.Value[$c]}
			"UseEarning"			{[boolean]$useEarning=[boolean]$cfg.Value[$c]}
		}
	}
	[boolean]$useMemo = [boolean]$rjson.EPclk.EPCLK_USEMEMO;
	[boolean]$usePosition = [boolean]$rjson.EPclk.EPCLK_USEPOSITION;
	[boolean]$useRate = [boolean]$rjson.EPclk.EPCLK_USERATE;
	[boolean]$useAmount = [boolean]$rjson.EPclk.EPCLK_USEAMOUNT;
	[boolean]$useType = [boolean]$rjson.EPclk.EPCLK_USETYPE;
	[boolean]$useHours = [boolean]$rjson.EPclk.EPCLK_USEHOURS;
	[boolean]$useClient = [boolean]$rjson.EPclk.EPCLK_USECLIENT;
	[int32]$tzOffset = $script:GMTOffset; #360; #Hardcoded, because I haven't figured out how to get this value in minutes automatically
	[string]$clockTime = [string]$rjson.EPclk.EPCLK_CLOCKDATE;
	
	[int32]$status = [int32]$json.EPclk.EPCLK_STATUS;
	$isDefault = $false;
	Write-Host "Choose a direction:";
	Write-Host "===================";
	for ($s=0;$s -le 1;$s++) {
		if($s -eq $status) {
			$isDefault = $true;
		}
		if ($s -ge 1) {
			$displayName = "In";
		} else {
			$displayName = "Out";
		}
		if ($isDefault -eq $true) {
			$displayName = "$($displayName) *";
		}
		Write-Host "$($s). $($displayName)";
		$isDefault=$false;
	}
	Write-Host "===================";
	$statusChoice = Read-Host "Direction";
	if (-Not ([string]::IsNullOrWhiteSpace($statusChoice))) {
		$status = [int32]$statusChoice;
	}
	
	[int32]$earnID = [int32]$rjson.EPclk.EPCLK_PRITEMID;
	if ($status -ge 1) {
		$direction = "in";
		[int32]$earning = 0;
		if ($useEarning -eq $true) {
			$isDefault = $false;
			$earnings = $rjson.EPclk.EPclkItems;
			Write-Host "Choose which earning item to use:";
			Write-Host "=================================";
			for ($e=0;$e -le $earnings.Count-1;$e++) {
				if ($earnings.EPERN_PRITEMID[$e] -eq $rjson.EPclk.EPCLK_PRITEMID) {
					$isDefault=$true;
					[int32]$defaultItemID = [int32]$e;
					[string]$defaultItemName = [string]$earnings.EPERN_ITEMNAME[$e];
				}
				$displayName = $earnings.EPERN_ITEMNAME[$e];
				if ($isDefault -eq $true) {
					$displayName = "$($displayName) *";
				}
				Write-Host "$($e). $($displayName)";
				$isDefault=$false;
			}
			Write-Host "=================================";
			$earningChoice = Read-Host -Prompt "Earning item";
			if ([string]::IsNullOrWhiteSpace($earningChoice)) {
				$earning = [int32]$defaultItemID;
			} else {
				$earning = [int32]$earningChoice;
			}
			
			$earnID = [int32]$earnings.EPERN_PRITEMID[$earning];
			
			$useMemo = [boolean]$earnings.EPERN_USEMEMO[$earning];
			$usePosition = [boolean]$earnings.EPERN_USEPOSITION[$earning];
			$useRate = [boolean]$earnings.EPERN_USERATE[$earning];
			$useDistributionTable = [boolean]$earnings.EPERN_USEDISTRIBUTION[$earning];
			$useAccount = [boolean]$earnings.EPERN_USEACCOUNT[$earning];
			$useAmont = [boolean]$earnings.EPERN_USEAMOUNT[$earning];
			$useType = [boolean]$earnings.EPERN_USETYPE[$earning];
			$useHours = [boolean]$earnings.EPERN_USEHOURS[$earning];
			$useClient = [boolean]$earnings.EPERN_USECLIENT[$earning];
		}
		
		[string]$memo = "";
		if ($useMemo -eq $true) {
			$memo = Read-Host -Prompt "Memo";
		}
		
		[int32]$position = 0;
		if ($usePosition -eq $true) {
			$positions = $rjson.EPclk.EPclkPositions;
			Write-Host "Choose which Position Funding Source to use:";
			Write-Host "============================================";
			for ($p=0;$p -le $positions.Count-1;$p++) {
				Write-Host "$($p). $($positions.PCPOS_POSITIONUNDINGSOURCE[$p])";
			}
			Write-Host "============================================";
			$positionChoice = Read-Host -Prompt "Position Funding Source";
			$position = [int32]$positionChoice;
		}
		
		[decimal]$rate = 0.00;
		if($useRate -eq $true) {
			$rateEntry = Read-Host -Prompt "Rate (i.e. 0.00)";
			$rate = [decimal]$rateEntry;
		}
		
		[int32]$dist = 0;
		if ($useDistributionTable -eq $true) {
			$isDefault = $false;
			[int32]$defaultTableID = 0;
			$dists = $rjson.EPclk.EPclkDists;
			Write-Host "Choose which distribution table to use:";
			Write-Host "=======================================";
			for ($d=0;$d -le $dists.Count-1;$d++) {
				if ($dists.GLDTB_DISTRIBUTIONTABLEID[$d] -eq $rjson.EPclk.EPCLK_DISTRIBUTIONTABLEID) {
					$isDefault=$true;
					$defaultTableID=[int32]$d;
				}
				$displayName = $dists.GLDTB_DISTRIBUTIONTABLE[$d];
				if ($isDefault -eq $true) {
					$displayName = "$($displayName) *";
				}
				Write-Host "$($d). $($displayName)"
				$isDefault=$false;
			}
			Write-Host "=======================================";
			$distChoice = Read-Host -Prompt 'Distribution Table';
			if ([string]::IsNullOrWhiteSpace($distChoice)) {
				$dist = [int32]$defaultTableID;
			} else {
				$dist = [int32]$distChoice;
			}
		}
		
		[string]$mask = "";
		if ($useAccount -eq $true) {
			$mask = Read-Host "Enter an account or mask as $($accountPicture) or blank";
		}
		
		[decimal]$amount = 0.00;
		if ($useAmount -eq $true) {
			$amountEntry = Read-Host -Prompt "Amount (i.e. 0.00)";
			$amount = [decimal]$amountEntry;
		}
		
		[int32]$type = 0;
		if ($useType -eq $true) {
			$isDefault = $false;
			[int32]$defaultTypeID = 0;
			$types = $rjson.EPclk.EPclkTypes;
			Write-Host "Choose which type to use:";
			Write-Host "=========================";
			for ($t=0;$t -le $types.Count-1;$t++) {
				if ($types.AFTYP_TYPEID[$t] -eq $rjson.EPclk.EPCLK_TYPEID) {
					$isDefault=$true;
					$defaultTypeID=[int32]$t;
				}
				$displayName = $types.AFTYP_TYPE[$t];
				if ($isDefault -eq $true) {
					$displayName = "$($displayName) *";
				}
				Write-Host "$($t). $($displayName)"
				$isDefault=$false;
			}
			Write-Host "=========================";
			$typeChoice = Read-Host -Prompt 'Type';
			if ([string]::IsNullOrWhiteSpace($typeChoice)) {
				$type = [int32]$defaultTypeID;
			} else {
				$type = [int32]$typeChoice;
			}
		}
		
		[decimal]$hours = 0.00;
		if ($useHours -eq $true) {
			$hoursEntry = Read-Host -Prompt "Hours (i.e. 0.00)";
			$hours = [decimal]$hoursEntry;
		}
		
		[int32]$client = 0;
		if ($useClient -eq $true) {
			$clientEntry = Read-Host -Prompt "Client entry not supported. Try a value";
			$client = [int32]$clientEntry;
		}
	} else {
		$direction = "out";
	}
	
	$requestEndpoint = "EPclkTable";
	$uri = "$($script:settingsObject.url)$($requestEndpoint)";
	$requestString = "ACTION=Add";
	$requestString = "$($requestString)&EPCLK_GMTOFFSET=$($tzOffset)";
	$requestString = "$($requestString)&EPCLK_STATUS=$($status)";
	$requestString = "$($requestString)&EPCLK_PRITEMID=$($earnID)";
	if ($status -ge 1) {
		$requestString = "$($requestString)&EPCLK_CLIENTID=$($client)"; # Not yet supported
		$requestString = "$($requestString)&EPCLK_MEMO=$($memo)";
		$requestString = "$($requestString)&EPCLK_TYPEID=$($rjson.EPclk.EPclkTypes.AFTYP_TYPEID[$type])";
		$requestString = "$($requestString)&EPCLK_FUNDINGSOURCEID=$($rjson.EPclk.EPclkPositions.PCPOS_EMPLOYEEPOSITIONID[$position])";
		$requestString = "$($requestString)&EPCLK_RATE=$($rate)";
		$requestString = "$($requestString)&EPCLK_DISTRIBUTIONTABLEID=$($rjson.EPclk.EPclkDists.GLDTB_DISTRIBUTIONTABLEID[$dist])";
		$requestString = "$($requestString)&EPCLK_POSTMASK=$($mask)";
		$requestString = "$($requestString)&EPCLK_AMOUNT=$($amount)";
	}
	
	$response = Invoke-WebRequest -uri $uri -Method POST -Body $requestString -WebSession $script:session -UserAgent $script:useragent;
	
	$rjson = $response.Content | ConvertFrom-Json;
	
	if ($rjson.success -eq $true) {	
		Write-Host "You were successfully clocked $($direction) at $($clockTime).";
	} elseif ($rjson.success -eq $false) {
		Write-Host "There was a problem while clocking you $($direction). The server reports the following error: $($rjson.msg)";
		$retry = Read-Host -Prompt "Would you like to try again? (yes/no)";
		if ($retry -eq "yes") {
			WebClock;
		}
	} else {
		Write-Host "Arr! There be a problem with ye request!  The cap'ns response be [$($response.StatusCode) - $($response.StatusDescription)]: $($response.Content)";
	}
}

function DoLogin {
	
	if ($script:settingsObject.savepwd -eq $true) {
		$cred = Import-CliXml $script:credPath;
		$loginPassword = $cred.Password;
	} else {
		$loginPassword = Read-Host -Prompt "Password" -AsSecureString;
	}
	$loginPassword = [string][Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($loginPassword));

	$requestEndpoint = "AccuConfig";
	$uri = "$($script:settingsObject.url)$($requestEndpoint)";
	$requestString = "ACTION=Login";
	if (-Not ($script:settingsObject.db -eq "<no database>")) {
		$requestString = "$($requestString)&database=$($script:settingsObject.db)";
	}
	$requestString = "$($requestString)&userabbr=$($script:settingsObject.user)";
	$requestString = "$($requestString)&password=$($loginPassword)";
	$requestString = "$($requestString)&Latitude=$($script:Lat)";
	$requestString = "$($requestString)&Longitude=$($script:Lon)";
	$requestString = "$($requestString)&appsettings=$($script:appsettings)";

	$response = Invoke-WebRequest -uri $uri -Method POST -Body $requestString -WebSession $script:session -UserAgent $script:useragent;

	$rjson = $response.Content | ConvertFrom-Json;
	
	$script:success = [boolean]$rjson.success;
	$script:msg = [string]$rjson.msg;
	$script:version = [string]$rjson.version;
}

function GetOffset {
	[boolean]$NegOffset = $true;
	$tz = "{0:zzz}" -f (Get-Date);
	[int32]$tzhours = $tz.Split(":",2)[0];
	[int32]$tzminutes = $tz.Split(":",2)[1];
	if ($tzhours -lt 0) {
		$NegOffset = $false;
		$tzhours = $tzhours * -1;
	}
	$tzminutes = $tzminutes + ($tzhours * 60);
	if($NegOffset -eq $true) {
		$tzminutes = $tzminutes * -1;
	}
	$script:GMTOffset = $tzminutes;
}

function GetLocation {
	[int32]$retryCount = 0;
	[int32]$maxRetries = 10;
	# From hxxps://stackoverflow.com/questions/46287792/powershell-getting-gps-coordinates-in-windows-10-using-windows-location-api/46287884 (accessed on 2019-11-18)
	Add-Type -AssemblyName System.Device #Required to access System.Device.Location namespace
	$GeoWatcher = New-Object System.Device.Location.GeoCoordinateWatcher #Create the required object
	$GeoWatcher.Start() #Begin resolving current locaton

	while (($GeoWatcher.Status -ne 'Ready') -and ($GeoWatcher.Permission -ne 'Denied') -and ($retryCount -lt $maxRetries)) {
		Start-Sleep -Milliseconds 100 #Wait for discovery.
		$retryCount = $retryCount + 1;
	}

	if ($GeoWatcher.Permission -eq 'Denied'){
		Write-Host "Access Denied for Location Information";
	} else {
		$script:Lat = [string]$GeoWatcher.Position.Location.Latitude;
		$script:Lon = [string]$GeoWatcher.Position.Location.Longitude;
	}
	Write-Host "Ay matey! Yer reportin ye location as $($script:Lat) Latitude, $($script:Lon) Longitude.";
}

function LoadSettings {

	if (Test-Path -Path $script:settingsPath) {
		$script:settingsObject = Get-Content -Raw -Path $script:settingsPath | ConvertFrom-Json;
	}

	if ([string]::IsNullOrWhiteSpace($script:settingsObject.url)) {
		RequestURL;
	}

	if ([string]::IsNullOrWhiteSpace($script:settingsObject.user)) {
		RequestUser;
	}

	if ([string]::IsNullOrWhiteSpace($script:settingsObject.savepwd)) {
		RequestSavePassword;
	}

	if ((-Not (Test-Path -Path $script:credPath)) -and ($script:settingsObject.savepwd -eq $true)) {
		RequestPassword;
	}

	if ([string]::IsNullOrWhiteSpace($script:settingsObject.db)) {
		RequestDB;
	}
	
	if ($script:settingsChanged) {
		SaveSettings;
	}
}

function SaveSettings {
	Set-Content $script:settingsPath ($script:settingsObject | ConvertTo-Json );
	$script:settingsChanged = $false;
}

function RequestURL {
	$entryURL = Read-Host -Prompt "URL (e.g. https://clock.example.com/)";
	if ([string]::IsNullOrWhiteSpace($entryURL)) {
		Write-Host "Arr! Walk yer fingers, or ye be walkin' the plank!"
		RequestURL; #URL is required
	} else {
		if (-Not ($entryURL.substring($entryURL.length - 1, 1) -eq "/")) {
			$entryURL = "$($entryURL)/";
		}
		$script:settingsObject.url = [string]$entryURL;
		$script:settingsChanged = $true;
	}
}

function RequestUser {
	$entryUser = Read-Host -Prompt "Username";
	if ([string]::IsNullOrWhiteSpace($entryUser)) {
		Write-Host "Arr! Walk yer fingers, or ye be walkin' the plank!"
		RequestUser; #User is required
	} else {
		$script:settingsObject.user = [string]$entryUser;
		$script:settingsChanged = $true;
	}
}

function RequestSavePassword {
	$entrySavePwd = Read-Host -Prompt "Save password for easier login (Yes / No)?";
	if ([string]::IsNullOrWhiteSpace($entrySavePwd)) {
		RequestSavePassword;
	} else {
		$entrySavePwd.ToUpper();
		if ($entrySavePwd.Substring(0,1) -eq "Y") {
			$script:settingsObject.savepwd = $true;
			$script:settingsChanged = $true;
		} elseif ($entrySavePwd.Substring(0,1) -eq "N") {
			$script:settingsObject.savepwd = $false;
			$script:settingsChanged = $true;
		} else {
			Write-Host "Arr! That be the wrong answer, matey! [$($entrySavePwd)]";
			RequestSavePassword;
		}
	}
}

function RequestPassword {
	$credUser = $script:settingsObject.user;
	$cred = Get-Credential -UserName $credUser;
	$cred | Export-CliXml $script:credPath;
}

function RequestDB {
	$entryDB = Read-Host -Prompt "Database";
	if ([string]::IsNullOrWhiteSpace($entryDB)) {
		$entryDB = "<no database>"
		Write-Host "Ay! I suppose there be no persuadin' ya."
	}
	$script:settingsObject.db = [string]$entryDB;
	$script:settingsChanged = $true;
}

Main;