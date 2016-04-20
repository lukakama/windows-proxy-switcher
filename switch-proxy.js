/*
Copyright 2016 Luca De Petrillo

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

	 http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

// Get the JSON object (not available in every C/WSCRIPT runtime).
var htmlfile = WScript.CreateObject('htmlfile');
htmlfile.write('<meta http-equiv="x-ua-compatible" content="IE=9" />');
var JSON = htmlfile.parentWindow.JSON;
htmlfile.close();

var objShell = WScript.CreateObject('WScript.Shell');

var hasStatusError = false;
openStatusDialog();
try {
	addStatusMessage('Loading...');
	
	var profileName = WScript.Arguments(0);
	
	var fso = WScript.CreateObject('Scripting.FileSystemObject'); 
	
	var scriptFolder = fso.GetFile(WScript.ScriptFullName).ParentFolder;
	var configFileStream = fso.OpenTextFile(scriptFolder + '\\proxy-config.json', 1); 
	var configText = '';
	try {
		configText =  configFileStream.ReadAll();
	} finally {
		configFileStream.Close();
	}
	
	var proxyConfig = JSON.parse(configText);
	
	if (proxyConfig.hasOwnProperty(profileName)) {
		addStatusMessage('Updating proxies for "' + profileName + '"...');
		
		var mavenConfigPath = proxyConfig.mavenConfigPath;
		
		var profileConfig = proxyConfig[profileName];
		
		var proxyHost = profileConfig.proxyHost;
		var proxyPort = profileConfig.proxyPort;
		
		var proxyExclusions = profileConfig.proxyExclusions;
		
		if (proxyHost && proxyPort) {
			// Apply new proxy settings.

			var browserProxyExclusions = proxyExclusions.concat([
					'<local>'
				]);

				
			var totProxyExclusions = browserProxyExclusions.join(';')
				+ ';,' + browserProxyExclusions.join(',')
				+ ',|' + browserProxyExclusions.join('|');

			// Setting IE proxy settings.
			addProgressMessage('systemProxy', 'Setting system proxy...');

			objShell.RegWrite('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ProxyEnable', 1, 'REG_DWORD');
			objShell.RegWrite('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ProxyServer', proxyHost + ':' + proxyPort, 'REG_SZ');
			objShell.RegWrite('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ProxyOverride', totProxyExclusions, 'REG_SZ');
			
			completeProgressMessage('systemProxy');
			
			
			// Update UNIX style environment proxy variables (used by tools like Node.JS).
			addProgressMessage('envVars', 'Setting enviroment variables...');
			
			objShell.RegWrite('HKCU\\Environment\\http_proxy', 'http://' + proxyHost + ':' + proxyPort, 'REG_SZ');
			objShell.RegWrite('HKCU\\Environment\\https_proxy', 'http://' + proxyHost + ':' + proxyPort, 'REG_SZ');
			objShell.RegWrite('HKCU\\Environment\\ftp_proxy', 'http://' + proxyHost + ':' + proxyPort, 'REG_SZ');
			objShell.RegWrite('HKCU\\Environment\\no_proxy', browserProxyExclusions.join(',').replace('*', ''), 'REG_SZ');
			
			completeProgressMessage('envVars');
			
		} else {
			// Remove existing proxy settings.
			addProgressMessage('systemProxy', 'Removing system proxy...');

			// Removing IE proxy settings.
			try {
				objShell.RegDelete('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ProxyEnable');
			} catch (e) {}
			try {
				objShell.RegDelete('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ProxyServer');
			} catch (e) {}
			try {
				objShell.RegDelete('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ProxyOverride');
			} catch (e) {}

			completeProgressMessage('systemProxy');
			
			
			// Removing UNIX style environment proxy variables.
			addProgressMessage('envVars', 'Removing proxy enviroment variables...');
			
			try {
				objShell.RegDelete('HKCU\\Environment\\http_proxy');
			} catch (e) {}
			try {
				objShell.RegDelete('HKCU\\Environment\\https_proxy');
			} catch (e) {}
			try {
				objShell.RegDelete('HKCU\\Environment\\ftp_proxy');
			} catch (e) {}
			try {
				objShell.RegDelete('HKCU\\Environment\\no_proxy');
			} catch (e) {}
			
			completeProgressMessage('envVars');
		}
		
		// Notify the system about the environment variables and IE proxy setting updates.
		sendConfigUpdatedNotifications();
		
		if (mavenConfigPath) {
			// Updating maven proxy settings.

			if (proxyHost && proxyPort) {
				addProgressMessage('maven', 'Setting Maven proxies...');
			} else {
				addProgressMessage('maven', 'Deactivating Maven proxies...');
			}

			// Load XML file
			var doc = WScript.CreateObject('msxml2.DOMDocument.6.0');
			doc.async = false;
			doc.resolveExternals = false;
			doc.validateOnParse = false;

			doc.setProperty('SelectionNamespaces', 'xmlns:def="http://maven.apache.org/SETTINGS/1.0.0"');
			doc.setProperty('SelectionLanguage', 'XPath');

			doc.load(mavenConfigPath);

			// Updating proxy configuration entries (should be already existing).
			// TODO: Handle specific config for http, https and ftp protocols, handling also missing entries.
			var colNodes = doc.selectNodes('//def:settings/def:proxies/def:proxy');
			for (var nodeIdx = 0; nodeIdx < colNodes.length; nodeIdx++) {
				if (proxyHost && proxyPort) {
					colNodes[nodeIdx].selectNodes('./def:active')[0].text = 'true';
					colNodes[nodeIdx].selectNodes('./def:host')[0].text = proxyHost;
					colNodes[nodeIdx].selectNodes('./def:port')[0].text = proxyPort;
					colNodes[nodeIdx].selectNodes('./def:nonProxyHosts')[0].text = proxyExclusions.join('|');
				} else {
					colNodes[nodeIdx].selectNodes('./def:active')[0].text = 'false';
				}
			}
			
			doc.save(mavenConfigPath);
			
			completeProgressMessage('maven');
		}
	
		addStatusMessage('Proxy settings successfully updated for "' + profileName + '"');
		addStatusMessage('This dialog will be closed in 2 seconds...');
	
		WScript.sleep(2000);
		
		closeStatusDialog();
	} else {
		addStatusMessage('No proxy set. Missing configuration for "' + profileName + '".');
	}
} catch (e) {
	addStatusMessage('Error while setting proxy: ' + JSON.stringify(e));
	throw e;
}


function base64Encode(value, charSet) {
	// Create the Stream object to 
	var adodbStream = WScript.CreateObject('ADODB.Stream');
	adodbStream.type = 2; //adTypeText
	adodbStream.charSet = charSet;
	adodbStream.open();
	try {
		adodbStream.writeText(value);
		adodbStream.position = 0;
		adodbStream.type = 1; //adTypeBinary

		// Ignore BOM
		adodbStream.position = 2;

		nodeValue = adodbStream.read();

	} finally {
		adodbStream.close();
	}
	
	// Retrieving the Base64 rapresentation of the binary source.
	var oXML = WScript.CreateObject('msxml2.DOMDocument.6.0');
	var oNode = oXML.createElement('root');
	oNode.dataType = 'bin.Base64';

	oNode.nodeTypedValue = nodeValue;
	
	var base64Value = oNode.text.split('\n').join('');
	return base64Value;
}

function sendConfigUpdatedNotifications() {
	addProgressMessage('sendNotification', 'Sending config updated notifications...');

	var psCommand = 
		'function Reload-InternetOptions { ' +
		'	if (-not ("wininet.nativemethods" -as [type])) { ' +
		'		add-type -Namespace wininet -Name NativeMethods -MemberDefinition \'[DllImport("wininet.dll", SetLastError = true, CharSet=CharSet.Auto)] public static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);\'; ' +
		'	}; ' +
		'	$INTERNET_OPTION_SETTINGS_CHANGED = 39; ' +
		'	$INTERNET_OPTION_REFRESH = 37; ' +
		'	$result1 = [wininet.nativemethods]::InternetSetOption(0, $INTERNET_OPTION_SETTINGS_CHANGED, 0, 0); ' +
		'	$result2 = [wininet.nativemethods]::InternetSetOption(0, $INTERNET_OPTION_REFRESH, 0, 0); ' +
		'	$result1 -and $result2; ' +
		'}; ' + 
		'function Broadcast-WMSettingChange { ' +
		'	if (-not ("win32.nativemethods" -as [type])) { ' +
		'		add-type -Namespace Win32 -Name NativeMethods -MemberDefinition \'[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)] public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult); \'; ' +
		'	}; ' +
		'	$HWND_BROADCAST = [intptr]0xffff; ' +
		'	$WM_SETTINGCHANGE = 0x1a; ' +
		'	$result = [uintptr]::zero; ' +
		'	[win32.nativemethods]::SendMessageTimeout($HWND_BROADCAST, $WM_SETTINGCHANGE, [uintptr]::Zero, "Environment", 2, 5000, [ref]$result); ' +
		'}; ' + 
		'Reload-InternetOptions -and Broadcast-WMSettingChange; ';	
	var retVal = objShell.Run('powershell.exe -NoProfile -EncodedCommand ' + base64Encode(psCommand, 'unicode'), 0, true);
			
	if (retVal !== 0) {
		throw new Error('Error executing sendConfigUpdatedNotifications powershell command: ' + retVal);
	}
	
	completeProgressMessage('sendNotification');
}

var objExplorer = null;
function openStatusDialog() {
	objExplorer = WScript.CreateObject('InternetExplorer.Application');
	
	objExplorer.Navigate2('about:blank');
	while (objExplorer.readyState != 4) {
		WScript.Sleep(200);
	}
	
	objExplorer.ToolBar = false;
	objExplorer.StatusBar = false;
	objExplorer.Left = 200;
	objExplorer.Top = 200;
	objExplorer.Width = 350;
	objExplorer.Height = 150; 
	objExplorer.Visible = true;
	objExplorer.Resizable = false;

	objExplorer.document.title = 'Proxy setting - WScript Progress Dialog';
	objExplorer.document.body.innerHTML = '';
	
	var psCommand = 
			'function BringIeToFront { ' +
			'	if (-not ("win32.nativemethods" -as [type])) { ' +
			'		add-type -Namespace Win32 -Name NativeMethods -MemberDefinition \'[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)] public static extern bool SetForegroundWindow(IntPtr hWnd);\'; ' +
			'	}; ' +
			'	$HWND = [intptr]' + objExplorer.Hwnd + '; ' +
			'	$result1 = [win32.nativemethods]::SetForegroundWindow($HWND); ' +
			'}; ' + 
			'BringIeToFront; ';
	var retVal = objShell.Run('powershell.exe -NoProfile -EncodedCommand ' + base64Encode(psCommand, 'unicode'), 0, true);
	if (retVal !== 0) {
		throw new Error('Error executing openStatusDialog powershell command: ' + retVal);
	}

}

function addStatusMessage(message) {
	if (objExplorer) {
		objExplorer.document.body.innerHTML += '<div style="margin-top: 5px; text-align: center; font-family: Tahoma, Verdana, Segoe, sans-serif; font-size: 14px; ">' + message + '</div>';
		var divs = objExplorer.document.getElementsByTagName('div');
		divs[divs.length - 1].scrollIntoView();
	}
}

function addProgressMessage(key, message) {
	if (objExplorer) {
		objExplorer.document.body.innerHTML += '<div style="margin-top: 5px; text-align: center; font-family: Tahoma, Verdana, Segoe, sans-serif; font-size: 11px; "><span id="' + key + '">' + message + '</span></div>';
		var divs = objExplorer.document.getElementsByTagName('div');
		divs[divs.length - 1].scrollIntoView();
	}
}

function completeProgressMessage(key) {
	if (objExplorer) {
		objExplorer.document.getElementById(key).innerHTML += " DONE";
		var divs = objExplorer.document.getElementsByTagName('div');
		divs[divs.length - 1].scrollIntoView();
	}
}

function closeStatusDialog() {
	if (objExplorer) {
		objExplorer.Quit();
	}
}
