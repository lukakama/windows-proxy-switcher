# windows-proxy-switcher

Simple WScript that can be used to change system proxies (with support for Maven and NodeJS tools).

To use it, download both [switch-proxy.js](switch-proxy.js) and [proxy-config.json](proxy-config.json) files and place them in the same directory. Then, edit the `proxy-config.json` to set the `mavenConfigPath` property to the full path of the maven settings.xml file (or leave it `null` to disable maven settings modifications) and add needed proxy configuraion profiles in the following form:
```json
  "profileName": {
		"proxyHost" : "the-proxy-host",
		"proxyPort" : "the-proxy-port",
		"proxyExclusions": [
			"pattern-to-exclude1",
			"pattern-to-exclude2"
		]
	}
```

To disable the proxy, set a `null` `proxyHost` or `proxyPort` for a profile and activate it.


Example `proxy-config.json`:
```json
{
	"mavenConfigPath" : "C:\\apache-maven-3.0.4\\conf\\settings.xml",
	
	"no-proxy": {
		"proxyHost" : null,
		"proxyPort" : null,
		"proxyExclusions": null
	},
	
	 "local-proxy": {
		"proxyHost" : "127.0.0.1",
		"proxyPort" : "8080",
		"proxyExclusions": [
			"192.168.0.*"
		]
	},
	
	"work-proxy": {
		"proxyHost" : "10.0.0.1",
		"proxyPort" : "3128",
		"proxyExclusions": [
			"10.0.0.*",
			"*.companydomain.com"
		]
	}
}
```


To activate a profile, just open a command line and execute the script issueing a `switch-proxy.js profile-name` in the direcotry where you placed the script (ex: `switch-proxy.js local-proxy`).
