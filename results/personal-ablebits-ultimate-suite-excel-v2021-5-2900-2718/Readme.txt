Ablebits.com Ultimate Suite for Microsoft Excel
--------------------------------------------------------------------------------

This document contains:
  - Short description of the add-in
  - List of supported Microsoft Excel versions
  - List of supported Microsoft Windows versions
  - How to install - Personal edition
  - How to install - Business edition
  - How to install - for Terminal Server Users
  - Trial version limitations
  - Support and web-resources


SHORT DESCRIPTION
--------------------------------------------------------------------------------
Ultimate Suite for Microsoft Excel includes all Ablebits add-ins to automate
many complex Excel tasks. It will be your trusted helper when you need to merge,
split, find, remove, modify, randomize or manage spreadsheet data.

Give these add-ins a try and you will see like most laborious tasks are
accomplished in a couple of steps saving you hours of time and energy:
https://www.ablebits.com/excel-suite/


SUPPORTED MICROSOFT EXCEL VERSIONS
--------------------------------------------------------------------------------
Ablebits Add-ins work with the following Microsoft Excel versions:
- Microsoft Excel from Microsoft 365 / Office 365 (desktop, 32- and 64-bit)
- Microsoft Excel 2019 (32- and 64-bit)
- Microsoft Excel 2016 (32- and 64-bit)
- Microsoft Excel 2013 (32- and 64-bit)
- Microsoft Excel 2010 (32- and 64-bit)

Please note, Compare Sheets and Compare Multiple Sheets are compatible only 
with Excel starting from version 2013. Excel 2010 is not supported by 
these tools.

All Office releases of the Office Insider update channels are not supported.


SUPPORTED MICROSOFT WINDOWS VERSIONS
--------------------------------------------------------------------------------
Ablebits Add-ins support the following Microsoft Windows versions:
- Windows 10 (32- and 64-bit)
- Windows 8 (32- and 64-bit)
- Windows 7 (32- and 64-bit)
- Windows Server 2008, 2012, 2016, 2019

.NET Framework 4.0 must be installed.


HOW TO INSTALL - PERSONAL EDITION
--------------------------------------------------------------------------------
- Close all Microsoft Excel windows.
- Run setup.exe and follow the instructions of the installation wizard.
- Open Microsoft Excel to see the add-ins running.


HOW TO INSTALL - BUSINESS EDITION
--------------------------------------------------------------------------------
- Close all Microsoft Excel windows.
- Run setup.exe and follow the instructions of the installation wizard.
- Open Microsoft Excel to see the add-ins running.

The edition labeled as "Business" supports several scenarios of corporate 
deployment including silent msi-based installation, deployment via GPO or 
SCCM; it is a per-user installation. Detailed information about corporate 
deployment you can find here: 
https://www.ablebits.com/support/corporate-deployment.php


HOW TO INSTALL - "FOR TERMINAL SERVER USERS" EDITION
--------------------------------------------------------------------------------
The edition marked as "for Terminal Server users" supports interactive
installation on terminal servers as well as silent deployment via GPO and SCCM.

INSTALLATION

The "for Terminal Server users" edition should be installed for all users; it
is a per-server deployment. You can run the interactive installation or deploy
the add-in using GPO or SCCM.

Please note, by default the Ultimate Suite is disabled for all users after installation.

To enable the Ultimate Suite for the selected users set the following values
in the users' registry:

  [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\Addins\Ablebits.Suite]
  "LoadBehavior"=dword:00000003

To activate the Ultimate Suite for the selected users add the following string
values to the users' registry:

  [HKEY_CURRENT_USER\SOFTWARE\Ablebits\Ultimate Suite for Microsoft Excel]
  "ProductKey"="Your license key"
  "RegistrationName"="Your registration name"


UNINSTALLATION

Please remove the following keys from the users' registry after uninstallation
the "for Terminal Server users" edition: 
  [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Options]
  "OPENx"="\"C:\\Program Files (x86)\\Ablebits\\Ultimate Suite for Microsoft Excel\\AblebitsSuiteForExcel.xlam\""
  "OPENx"="/A \"AblebitsUDF\""

NOTE! The x indexes in the OPENx values may vary depending on the xla/xlam
add-ins the current user has installed. However, the user can remove 
these xlam/udf add-ins interactively via the "File \ Options \ Add-ins"
dialog box.


TRIAL VERSION LIMITATIONS
--------------------------------------------------------------------------------
Ablebits Add-ins can be used for free during the trial period indicated on our
website. If you find Ablebits Add-ins useful and if you would like to continue
using them, please buy a license. You can order the license online on our
web-site: https://www.ablebits.com/purchase/


SUPPORT AND WEB-RESOURCES
--------------------------------------------------------------------------------
We offer free email support. Please contact us at 
https://www.ablebits.com/support/

We are highly interested in your feedback. Your comments, criticism, suggestions 
and bug reports will be highly appreciated. 

Please visit our web-site for the latest versions and upgrades:
https://www.ablebits.com/downloads/

Best regards,
Ablebits Team

---------------------------------------------------------------------------
Copyright (C) Ablebits.com
