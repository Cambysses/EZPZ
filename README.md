# EZPZ

**NOTE**: This is a redacted version which removes critical sensitive information. This will not run as expected, it exists only for the purposes of my Github portfiolio.

EZPZ Edition is a PowerShell-based GUI utility meant to help with various IT tasks for Maritime Travel.

## Prerequisites

The following modules are required to run EZPZ:
* AnyBox
* SQLPS
* ActiveDirectory

## Installation

Simply copy the Github files to your desktop, then right-click **EZPZ.ps1** and click **Run with PowerShell**. As long as the required modules have been installed, it should work. At first launch you will be asked to enter credentials. At the time of writing, EZPZ is designed to use the **REDACTED@maritimetravel.ca** service account, however this will be changed in the future to support any given admin account.

*EZPZ will not work with any admin account which requires Microsoft MFA, this is why the REDACTED account is used.*

## Features

### General

* Create Employee
* Disable Employee
* Create Distribution Group
* Reset AD Password
* Unlock AD Account

### Email Tools

* Set Mailbox Permissions
* Set Mailbox Type

### Diagnostics & Fixes

* Get Computer Uptime
* Rebuild Search Index
* Get Public IP
* Fix Smartpoint
* Fix Printer

### Miscellaneous

* Run AD Azure Sync
* Rename PC
* Send Galileo Vacations Mergeback
* Execute Custom PS Command
