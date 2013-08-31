AD-Copmuter-Rename-and-Join
===========================

VB Script - Disable IPv6, Change Computer Name, Join Domain and Place in Specified OU

### Installation

1. Clone this script from github or copy the files manually to your prefered directory.
2. Edit the following AD variables:
 - strDomainUser="user@domain.com"
 - strDomainPasswd="password"
 - strDomainName="DOMAIN"
 - VLAN1OU="Faculty"
 - VLAN2OU="Students"
 - VLAN1strOU = "OU=Faculty,OU=Workstations,DC=DOMAIN,DC=COM"
 - VLAN2strOU = "OU=Faculty,OU=Workstations,DC=DOMAIN,DC=COM"
3. Fill in VLAN for example (192.168.228.0,10.0.226.0):
 - VLAN1="228"
 - VLAN2="226"
4. Use srvany.exe to create a service for it and set it for delayed start.

#### It will need 2 restarts to rename and join the computer to active directory.

by [RaveMaker][RaveMaker].
[RaveMaker]: http://ravemaker.net
