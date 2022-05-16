# OpConsole
A console-like program for OpCon.  Allows you to view reports and do other admin tasks from a console interface.

# Prerequisites
* OpCon v21.4 (most things will work on earlier versions)
* Powershell 7.0+ (https://github.com/PowerShell/PowerShell)
* OpConModule.psm1 (https://github.com/SMATechnologies/opcon-rest-api-client-powershell)

# Instructions
To run this program you must have all the files in the same directory.  The "config" file is not required but makes connecting to environments must faster and easier when starting each time.

You are also able to create new connections at the prompts and save them to the configuration file.

To add "OpConsole" to the Windows Terminal app, open the Settings and add the following lines under Profiles -> List:

```
        {
          "name": "OpConsole",
          "commandline": "pwsh.exe -ExecutionPolicy Bypass -File \"C:\\OpConsole.ps1\"",
          "hidden": false
        }
```

# Disclaimer
No Support and No Warranty are provided by SMA Technologies for this project and related material. The use of this project's files is on your own risk.

SMA Technologies assumes no liability for damage caused by the usage of any of the files offered here via this Github repository.

# License
Copyright 2020 SMA Technologies

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at [apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

# Contributing
We love contributions, please read our [Contribution Guide](CONTRIBUTING.md) to get started!

# Code of Conduct
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-v2.0%20adopted-ff69b4.svg)](code-of-conduct.md)
SMA Technologies has adopted the [Contributor Covenant](CODE_OF_CONDUCT.md) as its Code of Conduct, and we expect project participants to adhere to it. Please read the [full text](CODE_OF_CONDUCT.md) so that you can understand what actions will and will not be tolerated.
