
# DisabledComputersManager

DESCRIPTION: 
- Manage disabled computers

> NOTES: "v1.0" was completed in 2011. DisabledComputersManager was written to work in on-premises Active Directory environments. The purpose of DisabledComputersManager was/is to organize disabled computers to support a "clean AD".

## Requirements:

Operating System Requirements:
- Windows Server 2003 or higher (32-bit)
- Windows Server 2008 or higher (32-bit)

Additional software requirements:
Microsoft .NET Framework v3.5

Active Directory requirements:
One of following domain functional levels
- Windows Server 2003 domain functional level
- Windows Server 2008 domain functional level

Additional requirements:
Domain administrative access is required to perform operations by DisabledComputersManager


## Operation and Configuration:

Command-line parameters:
- run (Required parameter)

Configuration file: configDisabledComputersManager.txt
- Located in the same directory as DisabledComputersManager.exe

Configuration file parameters:

DisabledHoldPeriod: Number of days to wait before removing, from Active Directory, disabled accounts in the OU specified using the DisabledComputersLocation parameter
- If not specified, the default is 7 days

DisabledComputersLocation: Specifies an OU location in Active Directory to place disabled computers; The OU location specified must already be present

Exclude: Exclude one or more computers by specifying the desired computer on a separate line

ExcludePrefix: Exclude one or more computers using a prefix that will match the desired computer(s)

ExcludeSuffix: Exclude one or more computers using a suffix that will match the desired computer(s)

Output:
- Located in the Log directory inside the installation directory; log files are in tab-delimited format
- Path example: (InstallationDirectory)\Log\

Additional detail:
- DisabledComputersManager will act on any disabled computer except disabled computers excluded from being processed. This includes disabled computers that were disabled manually or by other automated processes. Excluded user accounts include those specified in the configuration file and accounts that are in the OU specified using DisabledComputersLocation parameter in the configuration file.
- DisabledComputersManager will automatically exclude all Domain Controllers in the domain.
- DisabledComputersManager is built to perform the following operations on a disabled computer:
    - Remove all group membership except for the disabled computer’s Primary group.
    - Move the disabled computer to an OU specified in the configuration file
    - Add the disabled computer to a group created specifically to trigger the removal of the disabled computer from Active Directory based upon the DisabledHoldPeriod parameter
