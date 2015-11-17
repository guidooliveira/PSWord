PSWord
====================

Powershell module for creation and manipulation of Docx files without needing Microsoft Office Installed. This module will enable you to automate project documentations into word Directly.

[![Build status](https://ci.appveyor.com/api/projects/status/vuacqoejrkw534v8?svg=true)](https://ci.appveyor.com/project/guidooliveira/psword)

Usage
-----
Assuming you've installed the module somewhere in your module path, just import the module in your profile, e.g.:

```powershell
Import-Module PSWord
```

Installing
----------
If you have PowerShell V5, or have installed [PowerShellGet](https://www.microsoft.com/en-us/download/details.aspx?id=49186) for V3, you can install right away with:

```powershell
Install-Module -Name PSWord
```

Alternatively, you can install in your personal modules folder (e.g. ~\Documents\WindowsPowerShell\Modules), iwith:

```powershell
iex (new-object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/guidooliveira/PSWord/master/install.ps1')
```

If you want to install elsewhere, you can download Install.ps1 (using the URL above) and run it, passing in the directory you want to install.

# Issues / Feedback

For any issues or feedback related to this module, please register for GitHub, and post your inquiry to this project's issue tracker.
