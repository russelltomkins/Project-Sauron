Welcome to Project Sauron

The purpose of these project is to provide organisations without access to expenseive SIEM platforms to export Windows event log audit data from multiple Windows machines to a central location using built-in Windows functionality. The solution is also ideal for deployment in UAT/DEV/TEST environments that aren't currently covered by production SEM/SIEM deployments.

The catalyst for this project and primary working example was to provide a mechanism to allow Domain Controllers to centrally store and archive the large number of audit events they generate for archival and lookup purposes.


The 4 core scripts can be used to build your own solutions as well.
Custom View Creation - Create a custom view tree that allows you to easily extract specific events 
Manifest Creation - Creates an event channel manifest file for .dll compilation to create dedicated event channels (logs) for storage of events in management .evtx files
Event Channel Preparation - Enables the custom event channels, configures their default size and enables auto-archive.
Subscription Creation - Creates the windows event collection subscription files to forward and store events in the appproiate log file.

Getting Started - DC Events 
Some people will happily just use the pre-provided solution and thats cool. Check out the latest release for pre-compiled Custom Views, Event Channel manifest and DLL that can quickly be used.

Refer to the following blog post for more details
http://blogs.technet.microsoft.com/russellt/2017/03/23/project-sauron-part-1

1. Create or use an existing import csv to definie the custom event channels and xPath queries
2. Compile a new or reuse an existing .manifest and .dll file to define the custom event channels
3. Load the custom events channel .manifest and .dll into your Windows Event Collector
4. Prepare the event channels 
5. Load your the correspondign WEC subscriptions into the central Windows Event Collector Server
6. Configure the machines to pull subscriptions from the WEC Subscription server
7. Begin leveraging your new centralised event logs.

Domain Controller Event Data Sources
Account Management https://technet.microsoft.com/en-us/library/dd941622(v=ws.10).aspx
Audit Security Group Management https://technet.microsoft.com/en-us/library/dd772663(v=ws.10).aspx
Audit User Account Management https://technet.microsoft.com/en-us/library/dd772693(v=ws.10).aspx
Audit Security Group Management https://technet.microsoft.com/en-us/library/dd772663
Audit Other Account Management Events https://technet.microsoft.com/en-us/library/dd941586(v=ws.10).aspx
Audit Distribution Group Management https://technet.microsoft.com/en-us/library/dd772713(v=ws.10).aspx
Audit Computer Account Management https://technet.microsoft.com/en-us/library/dd772717(v=ws.10).aspx
Windows security audit events https://www.microsoft.com/en-us/download/details.aspx?id=50034

Contribute
Got an idea for a new Channel/Subscription/View? Leave a comment on the repository
