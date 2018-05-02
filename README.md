# bo-owner-swap
A Powershell script to change the owner of all recurring instances owned by a single user in SAP Business Objects to another owner. It has been tested on versions 4.1 and 4.2.
This has been useful for us in our company to allow administrators to reschedule recurring reports on behalf of business users. The original owners do not have visibility or ownership to these reports unless the owner is switched back to them.
- Swap out the values for Username, Password, and CMS with the values for your Business Objects platform
- Run the powershell script. It will search for ALL recurring instances owned by $selectedOwner in the script.
- Provide the value for an existing username in the Business Objects platform. It will swap out the owner for all recurring instances

*NOTE: This script requires Powershell 5.0+ 
