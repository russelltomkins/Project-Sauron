
Welcome to Project Sauron

For an introduction to Project Sauron and a quick-start using a Domain Controller example, refer to the following blog post.
https://blogs.technet.microsoft.com/russellt/2017/05/09/project-sauron-introduction/


The 4 core scripts can be used to build your own solutions as well.
Create-CustomView.ps1 - Create a custom view tree that allows you to easily extract specific events 
Create-Manifest.ps1 - Creates an event channel manifest file for .dll compilation to create dedicated event channels (logs) for storage of events in management .evtx files
Prepare-EventChannel.ps1 - Enables the custom event channels, configures their default size and enables auto-archive.
Create-Subscriptions.ps1 - Creates the windows event collection subscription files to forward and store events in the apppropriate log file.

Want to create your own?

1. Create a csv to define the custom event channels and xPath queries
2. Compile a new .manifest and .dll file to define the custom event channels from your master input csv.
3. Load the custom events channel .manifest and .dll into your Windows Event Collector using wevtutil.exe um <name.man>
4. Prepare the event channels 
5. Create and import your WEC subscriptions using the master input csv.
6. Configure the machines to pull subscriptions from the WEC Subscription server
7. Begin leveraging your new centralised event logs.


Contribute
Got an idea for a new Channel/Subscription/View? Leave a comment on the repository
