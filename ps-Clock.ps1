[string]$useragent = "com.eagleflint.ps-clock/0.5.1 (PowerShell)";
[string]$appsettings = "WebClock,0.5.1,com.eagleflint.ps-clock";
[string]$requestEndpoint = "";
[string]$uri = "";
[string]$version = "";
[boolean]$success = $false;
[string]$msg = "";
[int32]$GMTOffset = 0;
[string]$Lat = "";
[string]$Lon = "";
[string]$UUID = "";
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
$requestHeaders = @{"X-Requested-With"="com.eagleflint.Power-Clock"};

function Main {
	LoadSettings;
	GetLocation;
	GetDeviceID;
	DoLogin;
	
	Write-Host "Success:$($script:success)";
	Write-Host "msg:$($script:msg)";
	
	if ($script:success -eq $true) {
		Write-Host "Login successful. $($script:version)" -ForegroundColor Green;
		GetOffset;
		Write-Host "Ahoy! Ye be offset by $($script:GMTOffset) minutes.";
		WebClock;
	} elseif ($success -eq $false) {
		Write-Host "There was a problem while logging in. The server reports the following error: $($script:msg)";
		$retry = Read-Host -Prompt "Would you like to try again?  (yes/no)";
		if ($retry.ToUpper().Substring(0,1) -eq "Y") {
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
	[Int32]$defaultChoice = 0;
	$requestBody = [ordered]@{};

	$requestEndpoint = "EPclkTable";
	$uri = "$($script:settingsObject.url)$($requestEndpoint)";
	$requestBody.Add("Action","Get");

	$response = Invoke-WebRequest -uri $uri -Method POST -Body $requestBody -WebSession $script:session -UserAgent $script:useragent -Headers $script:requestHeaders;

	$rjson = $response.Content | ConvertFrom-Json;
	  
	$cfg = $rjson.EPclk.Cfg;
	for($c=0;$c -le $cfg.Count-1;$c++) {
		switch ($cfg.Name[$c]){
			"UseDistributionTable"	{[boolean]$useDistributionTable=[boolean]$cfg.Value[$c]}
			"UseAccount"			{[boolean]$useAccount=[Int32]$cfg.Value[$c]}
			"AccountPicture"		{[string]$accountPicture=[string]$cfg.Value[$c]}
			"UseEarning"			{[Int32]$useEarning=[Int32]$cfg.Value[$c]}
		}
	}
	[boolean]$useAmount		= [boolean]$rjson.EPclk.EPCLK_USEAMOUNT;
	[boolean]$useClient		= [boolean]$rjson.EPclk.EPCLK_USECLIENT;
	[boolean]$useHours		= [boolean]$rjson.EPclk.EPCLK_USEHOURS;
	[boolean]$useMemo		= [boolean]$rjson.EPclk.EPCLK_USEMEMO;
	[boolean]$usePosition	= [boolean]$rjson.EPclk.EPCLK_USEPOSITION;
	[boolean]$useRate		= [boolean]$rjson.EPclk.EPCLK_USERATE;
	[boolean]$useType		= [boolean]$rjson.EPclk.EPCLK_USETYPE;
	[int32]$tzOffset		= $script:GMTOffset;
	[string]$clockTime		= [string]$rjson.EPclk.EPCLK_CLOCKDATE;
	
	[Int32]$status = [Int32]$rjson.EPclk.EPCLK_STATUS;
	GetNewDirection -pStatus ([ref]$status);
	
	[Int32]$earnID = [int32]$rjson.EPclk.EPCLK_PRITEMID;
	if ($status -ge 1) {
		$direction = "in";
		[int32]$earning = 0;
		if ($useEarning -eq 1) {
			$defaultChoice = 0;
			$earnings = $rjson.EPclk.EPclkItems;
			Write-Host "";
			Write-Host "Choose which earning item to use:";
			Write-Host "=================================";
			for ($e=0;$e -le $earnings.Count-1;$e++) {
				Write-Host "$($e). $($earnings.EPERN_ITEMNAME[$e])$(if($earnings.EPERN_PRITEMID[$e] -eq $earnID) {" *";$defaultChoice=$e;})";
			}
			Write-Host "=================================";
			$earningChoice = Read-Host -Prompt "Earning item";
			if ([string]::IsNullOrWhiteSpace($earningChoice)) {
				$earningChoice = $defaultChoice;
			}
			$earning = [int32]$earningChoice;
			$earnID = [Int32]$earnings.EPERN_PRITEMID[$earning];
			$useMemo = [boolean]$earnings.EPERN_USEMEMO[$earning];
			$usePosition = [boolean]$earnings.EPERN_USEPOSITION[$earning];
			$useRate = [boolean]$earnings.EPERN_USERATE[$earning];
			$useDistributionTable = [boolean]$earnings.EPERN_USEDISTRIBUTION[$earning];
			$useAccount = [boolean]$earnings.EPERN_USEACCOUNT[$earning];
			$useAmount = [boolean]$earnings.EPERN_USEAMOUNT[$earning];
			$useType = [boolean]$earnings.EPERN_USETYPE[$earning];
			$useHours = [boolean]$earnings.EPERN_USEHOURS[$earning];
			$useClient = [boolean]$earnings.EPERN_USECLIENT[$earning];
		}
		
		[string]$memo = "";
		if ($useMemo -eq $true) {
			$memo = Read-Host -Prompt "Memo";
		}
		
		[Int32]$fundingSourceID = 0;
		if ($usePosition -eq $true) {
			$positions = $rjson.EPclk.EPclkPositions;
			Write-Host "";
			Write-Host "Choose which Position Funding Source to use:";
			Write-Host "============================================";
			for ($p=0;$p -le $positions.Count-1;$p++) {
				Write-Host "$($p). $($positions.PCPOS_POSITIONUNDINGSOURCE[$p])";
			}
			Write-Host "============================================";
			$positionChoice = Read-Host -Prompt "Position Funding Source";
			$fundingSourceID = [Int32]$positions.PCPOS_EMPLOYEEPOSITIONID[[int32]$positionChoice];
		}
		
		[decimal]$rate = 0.00;
		if($useRate -eq $true) {
			$rateEntry = Read-Host -Prompt "Rate (i.e. 0.00)";
			$rate = [decimal]$rateEntry;
		}
		
		[Int32]$distID = 0;
		if ($useDistributionTable -eq $true) {
			$defaultChoice = 0;
			$dists = $rjson.EPclk.EPclkDists;
			Write-Host "";
			Write-Host "Choose which distribution table to use:";
			Write-Host "=======================================";
			for ($d=0;$d -le $dists.Count-1;$d++) {
				Write-Host "$($d). $($dists.GLDTB_DISTRIBUTIONTABLE[$d])$(if($dists.GLDTB_DISTRIBUTIONTABLEID[$d] -eq $rjson.EPclk.EPCLK_DISTRIBUTIONTABLEID) {" *";$defaultChoice=$d})";
			}
			Write-Host "=======================================";
			$distChoice = Read-Host -Prompt 'Distribution Table';
			if ([string]::IsNullOrWhiteSpace($distChoice)) {
				$distChoice = $defaultChoice;
			}
			$distID = [Int32]$dists.GLDTB_DISTRIBUTIONTABLEID[[Int32]$distChoice];
		}
		
		[string]$mask = "";
		if ($useAccount -eq $true) {
			$mask = Read-Host "Enter an account or mask as $($accountPicture) or blank";
		} else {
			$mask = $rjson.EPclk.EPCLK_ACCOUNT;
		}
		
		[decimal]$amount = 0.00;
		if ($useAmount -eq $true) {
			$amountEntry = Read-Host -Prompt "Amount (i.e. 0.00)";
			$amount = [decimal]$amountEntry;
		}
		
		[int32]$typeID = 0;
		if ($useType -eq $true) {
			$defaultChoice = 0;
			$types = $rjson.EPclk.EPclkTypes;
			Write-Host "";
			Write-Host "Choose which type to use:";
			Write-Host "=========================";
			for ($t=0;$t -le $types.Count-1;$t++) {
				Write-Host "$($t). $($types.AFTYP_TYPE[$t])$(if($types.AFTYP_TYPEID[$t] -eq $rjson.EPclk.EPCLK_TYPEID) {" *";$defaultChoice=$t})";
			}
			Write-Host "=========================";
			$typeChoice = Read-Host -Prompt 'Type';
			if ([string]::IsNullOrWhiteSpace($typeChoice)) {
				$typeChoice = $defaultChoice;
			}
			$typeID = $types.AFTYP_TYPEID[[Int32]$typeChoice];
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
	$requestBody = [ordered]@{}
	$requestBody.Add("Action","Add");
	$requestBody.Add("EPCLK_GMTOFFSE","$($tzOffset)");
	$requestBody.Add("EPCLK_STATUS","$($status)");
	$requestBody.Add("EPCLK_PRITEMID","$($earnID)");
	if ($status -ge 1) {
		$requestBody.Add("EPCLK_AMOUNT","$(if ($useAmount -eq $true){$amount})");
		$requestBody.Add("EPCLK_CLIENTID","$(if ($useClient -eq $true){$client})"); # Not yet supported
		$requestBody.Add("EPCLK_DISTRIBUTIONTABLEID","$(if ($useDistributionTable -eq $true){$distID})");
		$requestBody.Add("EPCLK_FUNDINGSOURCEID","$(if($usePosition -eq $true) {$fundingSourceID})");
		$requestBody.Add("EPCLK_MEMO","$($memo)");
		$requestBody.Add("EPCLK_RATE","$(if($useRate -eq $true){$rate})");
		$requestBody.Add("EPCLK_TYPEID","$(if($useType -eq $true){$typeID})");
		$requestBody.Add("EPCLK_ACCOUNT","$($mask)");
	}
	
	$response = Invoke-WebRequest -uri $uri -Method POST -Body $requestBody -WebSession $script:session -UserAgent $script:useragent -Headers $script:requestHeaders;
	
	$rjson = $response.Content | ConvertFrom-Json;
	
	if ($rjson.success -eq $true) {	
		Write-Host "Aye, Aye! Ye be clocked $($direction) at $($clockTime). $(if($($direction.ToUpper()) -eq "IN"){"Now swab the decks, ye scallywag!"}else{"Feel free to take a Caulk."})" -ForegroundColor Green;
	} elseif ($rjson.success -eq $false) {
		Write-Host "Arr! There be a problem clocking you $($direction). The cap'n says: $($rjson.msg)." -ForegroundColor Red;
		$retry = Read-Host -Prompt "Do ye fancy another try?" -ForegroundColor Yellow;
		if ($retry.ToUpper().Substring(0,1) -eq "Y") {
			WebClock;
		}
	} else {
		Write-Host "Arr! There be a problem with ye request!  The cap'n says [$($response.StatusCode) - $($response.StatusDescription)]: $($response.Content)" -ForegroundColor Red;
	}
}

function DoLogin {
	$requestBody = [ordered]@{}
	if ($script:settingsObject.savepwd -eq $true) {
		$cred = Import-CliXml $script:credPath;
		$loginPassword = $cred.Password;
	} else {
		$loginPassword = Read-Host -Prompt "Password" -AsSecureString;
	}
	$loginPassword = [string][Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($loginPassword));

	$requestEndpoint = "AccuConfig";
	$uri = "$($script:settingsObject.url)$($requestEndpoint)";

	$requestBody.Add("Action","Login");
	if (-Not ($script:settingsObject.db -eq "<no database>")) {
		$requestBody.Add("Database","$($script:settingsObject.db)");
	}
	$requestBody.Add("UserAbbr","$($script:settingsObject.user)");
	$requestBody.Add("Password","$($loginPassword)");
	$requestBody.Add("LAT","$($script:Lat)");
	$requestBody.Add("LONG","$($script:Lon)");
	$requestBody.Add("AppSettings","$($script:appsettings)");
	$requestBody.Add("UUID","$($UUID)");

	$response = Invoke-WebRequest -uri $uri -Method POST -Body $requestBody -WebSession $script:session -UserAgent $script:useragent -Headers $script:requestHeaders;

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
	Add-Type -AssemblyName System.Device; #Required to access System.Device.Location namespace
	$GeoWatcher = New-Object System.Device.Location.GeoCoordinateWatcher; #Create the required object
	$GeoWatcher.Start(); #Begin resolving current locaton

	while (($GeoWatcher.Status -ne 'Ready') -and ($GeoWatcher.Permission -ne 'Denied') -and ($retryCount -lt $maxRetries)) {
		Start-Sleep -Milliseconds 100; #Wait for discovery.
		$retryCount = $retryCount + 1;
	}

	if ($GeoWatcher.Permission -eq 'Denied'){
		Write-Host "Aye!! Good pirates don't be sharin' the locations of treasures." -ForegroundColor Yellow;
	} else {
		$script:Lat = [string]$GeoWatcher.Position.Location.Latitude;
		$script:Lon = [string]$GeoWatcher.Position.Location.Longitude;
	}
	Write-Host "Aye, Matey! Yer reportin ye location as $($script:Lat) Latitude, $($script:Lon) Longitude.";
}

function GetDeviceID {
	$script:UUID = (Get-CimInstance -Class Win32_ComputerSystemProduct).UUID;
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
		Write-Host "Arr! Walk yer fingers, or ye be walkin' the plank!";
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
		Write-Host "Arr! Walk yer fingers, or ye be walkin' the plank!";
		RequestUser; #User is required
	} else {
		$script:settingsObject.user = [string]$entryUser;
		$script:settingsChanged = $true;
	}
}

function RequestSavePassword {
	$entrySavePwd = Read-Host -Prompt "Save password for easier login? (Y/N)";
	if ([string]::IsNullOrWhiteSpace($entrySavePwd)) {
		RequestSavePassword;
	} else {
		$entrySavePwd.ToUpper();
		if ($entrySavePwd.Substring(0,1) -eq "Y") {
			$script:settingsObject.savepwd = $true;
			$script:settingsChanged = $true;
		} elseif ($entrySavePwd.Substring(0,1) -eq "N") {
			Write-Host "Ye can't be sayin' I didn't warn ya!" -ForegroundColor Green;
			$script:settingsObject.savepwd = $false;
			$script:settingsChanged = $true;
		} else {
			Write-Host "Arr! That be the wrong answer, matey! [$($entrySavePwd)] Do it right, or ye'll be flogged!" -ForegroundColor Red;
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
		$entryDB = "<no database>";
		Write-Host "Aye!! I suppose there be no persuadin' ya." -ForegroundColor Green;
	}
	$script:settingsObject.db = [string]$entryDB;
	$script:settingsChanged = $true;
}

function GetNewDirection ([ref]$pStatus) {
	[Int32]$defaultDir=$(if ($pStatus.Value -eq 0){1}else{0});
	Write-Host "";
	Write-Host "Choose a direction:";
	Write-Host "===================";
	for ([Int32]$s=0;$s -le 1;$s++) {
		switch ($s){
			0	{Write-Host "$($s). Out$(if ($s -eq $defaultDir){" *"})"}
			1	{Write-Host "$($s). In $(if ($s -eq $defaultDir){" *"})"}
		}
	}
	Write-Host "===================";
	$statusChoice = Read-Host "Direction";
	if (-Not ([string]::IsNullOrWhiteSpace($statusChoice))) {
		if (([Int32]$statusChoice -eq 0) -or ([Int32]$statusChoice -eq 1)) {
			$pStatus.Value = [int32]$statusChoice;
		} else {
			Write-Host "Arr! Ye have only two options! Try again, or walk the plank!" -ForegroundColor Red;
			GetNewDirection [ref]$pStatus;
		}
	} else {
		$pStatus.Value = $defaultDir;
	}
}

Main;