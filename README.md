# PSHeap
Various pwsh functions and clases
## Classes
### V_WmiClass
This is exotic "wmi path" powershell class wich supports wildcards in wmi path and inspired by [System.Management.ManagementPath] class. Wildcard path may be 
resolved into array of matching existing pathes with .Resolve() method. Resolved paths may be accessed by .ResolvedPaths property. Security behavior of .Resolve() 
method is obvious: this method enumerates only permitted objects and ignores "access denied" errors. Recursive resolving (namespace path part only) may be executed 
as .Resolve($true).