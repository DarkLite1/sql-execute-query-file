# Description
A PowerShell script designed to execute .SQL files on an SQL database. 

This script will execute all statements found in a .SQL file on the specified computer and database. All required parameters are read from a [.JSON file](Example.json). 

The script [SQL Backup and restore database](https://github.com/DarkLite1/sql-backup-and-restore-database) in combination with this script can be a very powerful tool. In case a database needs to be migrated after an application update for example, it is often required to first take a backup, then do a restore on another computer and then execute some SQL statements to update the restored database to the desired state.

A scheduled task can be created to run the script in an automated way.

# How does it work?
- Each .SQL file with its statements is executed against the specified computer and database.
- An e-mail is sent containing the results in an excel file.

# Best practices
- Use explicit file names for the .SQL files in the .JSON input file that clearly explains what the file does. This will simplify re-using and organizing the different input files.
- Read the [help section](SQL%20Execute%20query%20file.ps1) in the script for an explanation of the input file parameters.

# Requirements
All required modules can be found in the [repository](https://github.com/DarkLite1?tab=repositories) and are free to use and download.