# Invoke-UserSimulator
Simulates common user behaviour on local and remote Windows hosts.

Invoke-UserSimulator is a tool developed with the aim of improving the realism of penetration testing labs (or other lab environments) to more accurately mirror a real network with users that create various types of traffic. Currently supported user behaviours the tool simulates are:

**Internet Explorer Browsing -** Creates an IE process and browses to a psuedo-random URL, then spiders the page for additional links to browse to. Simulates a user browsing the internet and creating web traffic on the network. 

**Mapping Shares -** Generates a random share name, and attempts to map it to the "K" drive. Creates LLMNR traffic on the network, allowing capturing network credentials via MitM attacks (Responder).

**Opening Emails -** Creates and Outlook COM object and iterates through any unread mail of the logged in user. Downloads and executes any attachments, and browses to any embedded links in IE.

The script can be run on a local server, or numerous remote hosts at once. For running on remote hosts, the script includes a configuration function to preconfigure Remote Desktop Users and various 

### Requirements:
**Windows -** The tool should work with any recent versions of Microsoft Windows (tested on Windows 7 through Server 2016). There is heavy use of PowerShell remoting, so when working with Windows 7 machines, some additional configuration will be required. 

**Microsoft Office -** When running the tool with -All or -Email flags, you'll need to have Outlook installed and configured to properly receive mail. Your users will also need to have working email addresses. If you plan on sending Macro phishing payloads, be sure the rest of the Office suite is installed as well.

### Arguments:
-Standalone
Define if the script should run as a standalone script on the localhost or on remote systems.

-ConfigXML [filepath]
The configuration xml file to use for host configuration and when running on remote hosts.

-ALL
Run all script simulation functions (IE, Shares, Email).

-IE
Run the Internet Explorer simulation.

-Shares
Run the mapping shares simulation.

-Email
Run the opening email simulation.

### Examples:
Import the script modules:

`PS>Import-Module .\Invoke-UserSimulator.ps1`

Run only the Internet Explorer function on the local host:

`PS>Invoke-UserSimulator -StandAlone -IE`

Configure remote hosts prior to running the script remotely:

`PS>Invoke-ConfigureHosts -ConfigXML .\config.xml`

Run all simulation functionality on remote hosts configured in the config.xml file:

`PS>Invoke-UserSimulator -ConfigXML .\config.xml -All`

*View the sample.xml file for an example of the file ConfigXML takes.*

### Walkthrough:
You can view a video walkthrough of the tool here: https://www.youtube.com/watch?v=lsC8mBKRZrs
