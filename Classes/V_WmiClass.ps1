# Copyright (c) VCUDACHI, 2022.
# Licensed under the MIT License.

<#
This is exotic "wmi path" powershell class wich supports wildcards in wmi path and inspired by [System.Management.ManagementPath] class. Wildcard path may be 
resolved into array of matching existing pathes with .Resolve() method. Resolved paths may be accessed by .ResolvedPaths property. Security behavior of .Resolve() 
method is obvious: this method enumerates only permitted objects and ignores "access denied" errors. Recursive resolving (namespace path part only) may be executed 
as .Resolve($true).

Example 1:Get all Win32 classes of root\cimv2 namespace:
$vp = [V_WmiPath]::New('root\cimv2:Win32_*')
$ResolvedCount = $vp.Resolve()
$vp.ResolvedPaths.Path

Example 2:Get all namespaces on local machine recursivly:
$ns = [V_WmiPath]::GetAllNamespaces('.')

Example 3:Get all classes from all namespaces on local machine recursivly (this task is memory and time consuming (~100MiB and ~200 seconds on regular machine):
$vp = [V_WmiPath]::New('*:*')
$ResolvedCount = $vp.Resolve($true)

Example 4:Do some instance searches in Win32_IP4RouteTable class:
#First, lets look at instance naming format for Win32_IP4RouteTable class:
[V_WmiPath]::ExplainInstance([System.Management.ManagementPath]::New('\\.\root\cimv2:Win32_IP4RouteTable'))
#Result is: \\.\root\cimv2:Win32_IP4RouteTable.Destination="*",InterfaceIndex=*,Mask="*",NextHop="*"

Example 4A: All instances
$vp = [V_WmiPath]::New('\\.\root\cimv2:Win32_IP4RouteTable.Destination="*",InterfaceIndex=*,Mask="*",NextHop="*"')
$ResolvedCount = $vp.resolve()
[wmi[]]($vp.ResolvedPaths) | FT

Example 4B: Get only where InterfaceIndex=1
$vp = [V_WmiPath]::New('\\.\root\cimv2:Win32_IP4RouteTable.Destination="*",InterfaceIndex=1,Mask="*",NextHop="*"')
$ResolvedCount = $vp.resolve()
[wmi[]]($vp.ResolvedPaths) | FT

Example 4C: Get only broadcast destinations:
$vp = [V_WmiPath]::New('\\.\root\cimv2:Win32_IP4RouteTable.Destination="*.*.*.255",InterfaceIndex=*,Mask="*",NextHop="*"')
$ResolvedCount = $vp.resolve()
[wmi[]]($vp.ResolvedPaths) | FT

Example 4D: Advanced filter by route age:
$vp = [V_WmiPath]::New('\\.\root\cimv2:Win32_IP4RouteTable.Destination="*",InterfaceIndex=*,Mask="*",NextHop="*"')
$vp.SetAdvancedWQLFilter('WHERE Age < 100000')
$ResolvedCount = $vp.resolve()
[wmi[]]($vp.ResolvedPaths) | FT

Example 5: Base filtering: show only Win32 classes in cimv2:
$vp = [V_WmiPath]::New('\\.\root\cimv2:*')
$vp.SetBaseFilter($true,$true,$true,$false)
$ResolvedCount = $vp.resolve()
[wmi[]]($vp.ResolvedPaths) | FT

##########################################################
#TODO 0.Hierarchy view for resolved paths
#TODO 1.Impliment direct parsing/structuring of path string
##########################################################
#! Powershell version requrements:
#! Minimal:      5.1
#! Recommended:  7.2+
##########################################################
#>
Class V_WmiPath {
    #region begin ABOUT
    static [PSCustomObject] About() {
        Return [PSCustomObject]@{
            Copyright = '(c) VCUDACHI, 2022'
            License   = 'MIT'
            GitHub    = 'https://github.com/vcudachi/PSHeap'
            Name      = '[V_WmiPath]'
            Version   = [version]'0.1.4'
            TimeStamp = Get-Date '2022-06-19 16:46'
            Platform  = 'Windows'
            TestedOn  = @('7.2.4', '5.1')
        }
    }
    #endregion

    #region begin INTERFACE
    #region begin Properties:
    # .WPath                 : get/set wildcard path
    # .WRelativePath         : get/set wildcard relative part of path
    # .WRootPath             : get/set wildcard root part of path
    # .Server                : get/set server
    # .WNamespacePath        : get/set wildcard namespace part of path
    # .WClassName            : get/set wildcard name part of path
    # .WInstance             : get/set wildcard instance name of path
    # .isNamespace           : get 
    # .isClass               : get 
    # .isInstance            : get 
    # .isSingleton           : get 
    # .WildcardInNamespace   : get 
    # .WildcardInClass       : get 
    # .WildcardInInstance    : get 
    # .ResolvedPaths         : get array of resolved from wildcard path paths
    #endregion

    #region begin Static methods:
    # ::About
    # ::TestPath
    # ::GetChildClasses
    # ::GetChildInstances
    # ::GetChildNamespaces
    # ::GetAllNamespaces
    # ::CheckWNamespace
    #endregion

    #region begin Methods:
    # GetBaseFilter
    # SetBaseFilter
    # GetAdvancedWQLFilter
    # SetAdvancedWQLFilter
    # ToString
    # Clone
    # Resolve
    #endregion

    #region begin May be constructed from:
    # $null
    # [string]
    # [System.Management.ManagementPath]
    # [wmiclass]
    # [wmi]
    # [cimclass]
    # [ciminstance]
    #endregion
    #endregion

    #region begin Comment
    <#
    CimType        NetType
    -------        -------
    Unknown
    Boolean        System.Boolean
    UInt8          System.Byte
    SInt8          System.SByte
    UInt16         System.UInt16
    SInt16         System.Int16
    UInt32         System.UInt32
    SInt32         System.Int32
    UInt64         System.UInt64
    SInt64         System.Int64
    Real32         System.Single
    Real64         System.Double
    Char16         System.Char
    DateTime
    String         System.String
    Reference      Microsoft.Management.Infrastructure.CimInstance
    Instance       Microsoft.Management.Infrastructure.CimInstance
    BooleanArray   System.Boolean[]
    UInt8Array     System.Byte[]
    SInt8Array     System.SByte[]
    UInt16Array    System.UInt16[]
    SInt16Array    System.Int64[]
    UInt32Array    System.UInt32[]
    SInt32Array    System.Int32[]
    UInt64Array    System.UInt64[]
    SInt64Array    System.Int64[]
    Real32Array    System.Single[]
    Real64Array    System.Double[]
    Char16Array    System.Char[]
    DateTimeArray
    StringArray    System.String[]
    ReferenceArray Microsoft.Management.Infrastructure.CimInstance[]
    InstanceArray  Microsoft.Management.Infrastructure.CimInstance[]    
    #>
    #endregion

    #region begin HIDDEN PROPERTIES:
    #ROP = Read/Only Property
    #RWP = Read/Write Property

    #region begin Flags:
    hidden [bool]$ROP_IsNamespace
    hidden [bool]$ROP_IsClass
    hidden [bool]$ROP_IsInstance
    hidden [bool]$ROP_IsSingleton
    hidden [bool]$ROP_WildcardInNamespace
    hidden [bool]$ROP_WildcardInClassName
    hidden [bool]$ROP_WildcardInInstance
    #endregion
    #region begin Regex:
    hidden static [string]$RgxCN = '(?<ComputerName>(?:\.|(?:\w|\w[\w-]*\w)(?:\.\w|\.\w[\w-]*\w)*))'
    hidden static [string]$RgxNS = '(?<Namespace>(?:[A-Za-z\*\?]|[A-Za-z\*\?\[][A-Za-z0-9\*\?\[\]_-]*[A-Za-z0-9\*\?\]-])(?:(?:\\|/)[A-Za-z\*\?]|(?:\\|/)[A-Za-z\*\?\[][A-Za-z0-9\*\?\[\]_-]*[A-Za-z0-9\*\?\]-])*)'
    hidden static [string]$RgxCL = '(?<ClassName>(?:[A-Za-z\*\?]|(?:__|[A-Za-z\*\?\[])[A-Za-z0-9\*\?\[\]_-]*[A-Za-z0-9\*\?\]]))'
    hidden static [string]$RgxKY = '(?<Key>[\w-]+=(?:".+?"|[^=,]+))'
    hidden static [string]$RgxIN = "(?<Instance>$([V_WmiPath]::RgxKY)(?:,$([V_WmiPath]::RgxKY))*)"
    hidden static [string]$RgxSI = '(?<Singleton>@)'
    #endregion
    #region begin Props:
    hidden [string]$RWP_WPath
    hidden [string]$RWP_WRelativePath
    hidden [string]$RWP_WRootPath
    hidden [string]$RWP_Server
    hidden [string]$RWP_WNamespace
    hidden [string]$RWP_WClassName
    hidden [string]$RWP_WInstance
    hidden [System.Collections.ArrayList]$WNamespaceParts = [System.Collections.ArrayList]::New()
    #endregion
    #region begin Resolve:
    hidden [bool[]]$RWP_BaseFilter = @($false, $false, $false, $false)
    hidden [string]$RWP_AdvancedFilter = [string]::Empty
    hidden [System.Collections.ArrayList]$ROP_ResolvedPaths = [System.Collections.ArrayList]::New()
    #endregion

    #endregion

    #region begin CONSTRUCTORS
    V_WmiPath() {
        $this.Init()
        $this.SetWPath('')
    }
    V_WmiPath([string]$StringPath) {
        $this.Init()
        $this.SetWPath($StringPath)
    }
    V_WmiPath([System.Management.ManagementPath]$Path) {
        $this.Init()
        $this.WPath = $Path.Path
    }
    V_WmiPath([System.Management.ManagementClass]$WmiClass) {
        $this.Init()
        $this.WPath = $WmiClass.__PATH
    }
    V_WmiPath([System.Management.ManagementObject]$WmiObject) {
        $this.Init()
        $this.WPath = $WmiObject.__PATH
    }
    V_WmiPath([Microsoft.Management.Infrastructure.CimClass]$CimClass) {
        $this.Init()
        $this.WPath = ('\\{0}\{1}:{2}' -f $CimClass.CimSystemProperties.ServerName, $CimClass.CimSystemProperties.Namespace, $CimClass.CimSystemProperties.ClassName)
    }
    V_WmiPath([Microsoft.Management.Infrastructure.CimInstance]$CimInstance) {
        $this.Init()
        $WPath = ('\\{0}\{1}:{2}' -f $CimInstance.CimSystemProperties.ServerName, $CimInstance.CimSystemProperties.Namespace, $CimInstance.CimSystemProperties.ClassName)
        If (($CimInstance.CimClass.CimClassQualifiers | Where-Object { $_.Name -eq 'Singleton' }).Value) {
            $this.WPath = $WPath + '=@'
        }
        Else {
            $keys = $CimInstance.CimInstanceProperties.Where({ $_.Flags.HasFlag([Microsoft.Management.Infrastructure.CimFlags]::Key) })
            If ($keys) {
                $this.WPath = $WPath + '.' + ("$($keys.Name)=`"$($keys.Value)`"" -join ',')
            }
            Else {
                $this.WPath = $WPath
            }
        }
    }
    #endregion

    #region begin INIT
    #Constructors helper method Init()
    hidden [void] Init() {
        $ScriptProperties = [System.Collections.ArrayList]::New()
        #PATH
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WPath', {
                    Return $this.RWP_WPath
                }, {
                    $this.SetWPath($args[0])
                })
        )
        #RELATIVEPATH
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WRelativePath', {
                    Return $this.RWP_WRelativePath
                }, {
                    $this.SetWRelativePath($args[0])
                })
        )
        #ROOTPATH
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WRootPath', {
                    Return $this.RWP_WRootPath
                }, {
                    $this.SetWRootPath($args[0])
                })
        )
        #SERVER
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('Server', {
                    Return $this.RWP_Server
                }, {
                    $this.SetServer($args[0])
                })
        )
        #NAMESPACEPATH
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WNamespacePath', {
                    Return $this.RWP_WNamespace
                }, {
                    $this.SetWNamespacePath($args[0])
                })
        )
        #CLASSNAME
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WClassName', {
                    Return $this.RWP_WClassName
                }, {
                    $this.SetWClassName($args[0])
                })
        )
        #INSTANCE
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WInstance', {
                    Return $this.RWP_WInstance
                }, {
                    $this.SetWInstance($args[0])
                })
        )
        #ISNAMESPACE
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('isNamespace', {
                    Return $this.ROP_IsNamespace
                })
        )
        #ISCLASS
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('isClass', {
                    Return $this.ROP_IsClass
                })
        )
        #ISINSTANCE
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('isInstance', {
                    Return $this.ROP_IsInstance
                })
        )
        #ISSINGLETON
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('isSingleton', {
                    Return $this.ROP_IsSingleton
                })
        )
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WildcardInNamespace', {
                    Return $this.ROP_WildcardInNamespace
                })
        )
        #ISCLASS
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WildcardInClassName', {
                    Return $this.ROP_WildcardInClassName
                })
        )
        #ISINSTANCE
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('WildcardInInstance', {
                    Return $this.ROP_WildcardInInstance
                })
        )
        #RESOLVEDPATHS
        $null = $ScriptProperties.Add(
            [psscriptproperty]::New('ResolvedPaths', {
                    If ($this.ROP_ResolvedPaths.Count -eq 0) {
                        #$null = $this.Resolve()
                        Return $null
                    }
                    Return ([System.Management.ManagementPath[]]$this.ROP_ResolvedPaths)
                })
        )
        $ScriptProperties | ForEach-Object {
            $null = $this.PSObject.Properties.Add($_)
        }
        
    }
    #endregion

    #region begin STATIC METHODS

    #Test if Path exists on system
    static [bool] TestPath([System.Management.ManagementPath]$Path) {
        If ($Path.IsInstance -or $Path.IsSingleton) {
            Try { $Instance = [wmi]$Path }
            Catch { $Instance = $null }
            If ($Instance) { Return $true } Else { Return $false }
        }
        ElseIf ($Path.IsClass) {
            Try { 
                $Class = [wmiclass]$Path
                $null = $Class.path
            }
            Catch { $Class = $null }
            If ($Class) { Return $true } Else { Return $false }
        }
        Else {
            Try { $Whatever = [wmi]$Path }
            Catch { $Whatever = $null }
            If ($Whatever) { Return $true } Else { Return $false }
        }
    }

    #Get all child classes in namespace of Path provided
    static [System.Management.ManagementPath[]] GetChildClasses([System.Management.ManagementPath]$Path) {
        $NSPath = $Path.Clone()
        $NSPath.RelativePath = ''
        $WQL = 'SELECT * FROM meta_class'
        Try {
            Try {
                $RawResult = [wmisearcher]::New($NSPath, $WQL).Get()
            }
            Catch {
                If ($_.Exception.InnerException.ErrorCode -eq 'AccessDenied') {
                    $RawResult = $null
                }
                Else {
                    throw [System.ArgumentException]::New('[wmisearcher] thrown an error', $_.Exception)
                }
            }
            If ($RawResult.Count -gt 0) {
                $Buffer = [Object[]]::New($RawResult.Count)
                $null = $RawResult.CopyTo($Buffer, 0)
                $null = $RawResult.Dispose()
                Return [System.Management.ManagementPath[]]$Buffer.Path
            }
            Else {
                If ([V_WmiPath]::TestPath($NSPath)) {
                    Return $null
                }
                Else {
                    throw [System.ArgumentException]::New('Namespace not exists')
                }  
            }
        }
        Catch {
            throw [System.Exception]::New('Unable to get child classes', $_.Exception)
        }
    }

    #Get instances of class filtered by filter in Path provided
    static [System.Management.ManagementPath[]] GetChildInstances([System.Management.ManagementPath]$Path, [string]$AdvancedFilter) {
        If ($Path.IsClass -or $Path.IsInstance -or $Path.IsSingleton) {
            $NSPath = $Path.Clone()
            $NSPath.RelativePath = ''
            If ($AdvancedFilter -eq [string]::Empty) {
                $WQL = "SELECT * FROM $($Path.ClassName)"
            }
            Else {
                $WQL = "SELECT * FROM $($Path.ClassName) $AdvancedFilter"
            }
            
            Try {
                Try {
                    $RawResult = [wmisearcher]::New($NSPath, $WQL).Get()
                }
                Catch {
                    If ($_.Exception.InnerException.ErrorCode -eq 'AccessDenied') {
                        $RawResult = $null
                    }
                    Else {
                        throw [System.ArgumentException]::New('[wmisearcher] thrown an error', $_.Exception)
                    }
                }
                If ($RawResult.Count -gt 0) {
                    $Buffer = [Object[]]::New($RawResult.Count)
                    $null = $RawResult.CopyTo($Buffer, 0)
                    $null = $RawResult.Dispose()
                    Return [System.Management.ManagementPath[]]$Buffer.__PATH
                }
                Else {
                    If ([V_WmiPath]::TestPath($NSPath)) {
                        Return $null
                    }
                    Else {
                        throw [System.ArgumentException]::New('Namespace not exists')
                    }  
                }
            }
            Catch {
                throw [System.Exception]::New('Unable to get child instances', $_.Exception)
            }
        }
        Else {
            throw [System.ArgumentException]::New('Not a class or instance or singleton')
        }
    }

    #Get instances of class in Path provided
    static [System.Management.ManagementPath[]] GetChildInstances([System.Management.ManagementPath]$Path) {
        Return [V_WmiPAth]::GetChildInstances($Path, [string]::Empty)
    }

    #Extract template of instance as wildcard string from class.
    # Example:  [V_WmiPath]::ExplainInstance([System.Management.ManagementPath]::New('\\.\root\cimv2:Win32_IP4RouteTable'))
    # Result:   \\.\root\cimv2:Win32_IP4RouteTable.Destination="*",InterfaceIndex=*,Mask="*",NextHop="*"
    static [string] ExplainInstance([System.Management.ManagementPath]$Path) {
        If ($Path.isClass) {
            $Class = [wmiclass]::New($Path)
            $Propereties = $Class.Properties | Where-Object { $_.Qualifiers.Where({ $_.Name -eq 'key' }) }
            Return ($Path.Path + '.' + (
                ($Propereties | ForEach-Object { If ($_.Type -in @(
                                'UInt8', 'SInt8', 'UInt16', 'SInt16', 'UInt32', 'SInt32', 'UInt64', 'SInt64', 'Real32', 'Real64'
                            )) { "$($_.Name)=*" } Else { "$($_.Name)=`"*`"" } }) -join ','
                ))
        }
        Else {
            throw [System.ArgumentException]::New('Not a class')
        }
    }

    #Get all child namespaces in namespace of Path provided
    static [System.Management.ManagementPath[]] GetChildNamespaces([System.Management.ManagementPath]$Path) {
        $NSPath = $Path.Clone()
        $NSPath.RelativePath = ''
        $WQL = 'SELECT * FROM __NAMESPACE'
        Try {
            Try {
                $RawResult = [wmisearcher]::New($NSPath, $WQL).Get()
            }
            Catch {
                If ($_.Exception.InnerException.ErrorCode -eq 'AccessDenied') {
                    $RawResult = $null
                }
                Else {
                    throw [System.ArgumentException]::New('[wmisearcher] thrown an error', $_.Exception)
                }
            }
            If ($RawResult.Count -gt 0) {
                $Buffer = [Object[]]::New($RawResult.Count)
                $null = $RawResult.CopyTo($Buffer, 0)
                $Result = [System.Management.ManagementPath[]]::New($RawResult.Count)
                For ($i = 0; $i -lt $Result.Count; $i++) {
                    $Result[$i] = [System.Management.ManagementPath]::New($Buffer[$i].Path.NamespacePath + '\' + $Buffer[$i].Path.RelativePath.Substring(18, ($Buffer[$i].Path.RelativePath.Length - 19)))
                }
                $null = $RawResult.Dispose()
                $null = $Buffer.Clear()
                Return $Result
            }
            Else {
                If ([V_WmiPath]::TestPath($NSPath)) {
                    Return $null
                }
                Else {
                    throw [System.ArgumentException]::New('Namespace not exists')
                }  
            }
        }
        Catch {
            throw [System.Exception]::New('Unable to get child namespaces', $_.Exception)
        }
    }

    #Get all namespaces recursivly. It takes ~55 seconds to enumerate all on a regular machine
    static [System.Management.ManagementPath[]] GetAllNamespaces([string]$Server) {
        If ($Server -match [V_WmiPath]::RgxCN) {
            $Paths = [System.Collections.ArrayList]::New()
            $Nick = 0
            $null = $Paths.Add([System.Management.ManagementPath]::New("\\$Server\ROOT"))
            $AddedCount = 1
            Do {
                $n = 0
                ForEach ($Path in $Paths.GetRange($Nick, $AddedCount).Clone()) {
                    ForEach ($Result in [V_WmiPath]::GetChildNamespaces($Path)) {
                        $null = $Paths.Add($Result)
                        $n++
                    }
                }
                $Nick = $Nick + $AddedCount
                $AddedCount = $n
            } While ($AddedCount -gt 0)
            
            $Result = [System.Management.ManagementPath[]]::New($Paths.Count)
            For ($i = 0; $i -lt $Result.Count; $i++) {
                $Result[$i] = $Paths[$i]
            }
            $null = $Paths.Clear()
            Return $Result
        }
        Else {
            throw [System.ArgumentException]::New('Server name is not in right format')
        }
    }

    static [System.Management.ManagementPath[]] GetAllNamespaces() {
        Return [V_WmiPath]::GetAllNamespaces('.')
    }

    #This method checks if $WNamespace contains valid namespace (starts with root or empty)
    #Example:  [V_WmiPath]::CheckWNamespace('root\cimv2:Win32_Process')
    #Result:   true
    #Example:  [V_WmiPath]::CheckWNamespace('\cimv2:Win32_Process')
    #Result:   false
    static [bool] CheckWNamespace([string]$WNamespace) {
        If ($WNamespace -eq [string]::Empty) {
            Return $true
        }
        Else {
            $pos1 = $WNamespace.IndexOf('\')
            $pos2 = $WNamespace.IndexOf('/')
            If ($pos1 -ge 0 -and $pos2 -ge 0) {
                $pos = [Math]::Min($WNamespace.IndexOf('\'), $WNamespace.IndexOf('/'))
            }
            ElseIf ($pos1 -ge 0) {
                $pos = $pos1
            }
            ElseIf ($pos2 -ge 0) {
                $pos = $pos2
            }
            Else {
                $pos = -1
            }
            If ($pos -eq -1) {
                $candidate = $WNamespace
            }
            Else {
                $candidate = $WNamespace.Substring(0, $pos)
            }
            Return ('root' -like $candidate)
        }
    }
    
    static hidden [bool] PassesFilter ([string]$ClassName, [object[]]$Filter) {
        #FILTER DESCRIPTION:
        <#
        $Filter[0] -> ExcludeSystem -> (-notlike '__*')
        $Filter[1] -> ExcludeCIM    -> (-notlike 'CIM_*')
        $Filter[2] -> ExcludeMSFT   -> (-notlike 'MSFT_*')
        $Filter[3] -> ExcludeWin32  -> (-notlike 'Win32_*')
        #>
        Return ( -not (
            ($Filter[0] -and $ClassName -like '__*') `
                    -or ($Filter[1] -and $ClassName -like 'CIM_*') `
                    -or ($Filter[2] -and $ClassName -like 'MSFT_*') `
                    -or ($Filter[3] -and $ClassName -like 'Win32_*')
            ))
    }
    #endregion

    #region begin METHODS
    hidden [void] SetWPath([string]$WPath) {
        If ($WPath.StartsWith('\\') -or $WPath.StartsWith('//')) {
            $pattern = "^(?:(?:\\\\|//)$([V_WmiPath]::RgxCN))(?:(?:(?:\\|/)$([V_WmiPath]::RgxNS)):$([V_WmiPath]::RgxCL)(?:\.$([V_WmiPath]::RgxIN)|=$([V_WmiPath]::RgxSI)){0,1}|(?:(?:\\|/|)$([V_WmiPath]::RgxNS))|$([V_WmiPath]::RgxCL)(?:\.$([V_WmiPath]::RgxIN)|=$([V_WmiPath]::RgxSI)){0,1}){0,1}$"
        }
        Else {
            $pattern = "^(?:(?:(?:\\|/|)$([V_WmiPath]::RgxNS)):$([V_WmiPath]::RgxCL)(?:\.$([V_WmiPath]::RgxIN)|=$([V_WmiPath]::RgxSI)){0,1}|(?:(?:\\|/|)$([V_WmiPath]::RgxNS))|$([V_WmiPath]::RgxCL)(?:\.$([V_WmiPath]::RgxIN)|=$([V_WmiPath]::RgxSI)){0,1}){0,1}$"
        }
        If ($WPath -match $pattern) {
            If ([V_WmiPath]::CheckWNamespace($matches['Namespace'])) {
                $this.ROP_ResolvedPaths.Clear()
                $this.RWP_Server = If ($matches['ComputerName']) { $matches['ComputerName'] }Else { '.' }
                $this.RWP_WNamespace = If ($matches['Namespace']) { $matches['Namespace'] } Else { [System.Management.ManagementPath]::DefaultPath.NamespacePath }
                $this.RWP_WClassName = $matches['ClassName']
                $this.RWP_WInstance = $matches['Instance'] + $matches['Singleton']
                $this.UpdateFlags()
                $this.UpdateRelativePath()
                $this.UpdateRootPath()
                $this.UpdateNamespaceParts()
                $this.UpdatePath()
            }
            Else {
                throw [System.ArgumentException]::New('Namespace part is not like "ROOT\*" or empty')
            }
        }
        Else {
            throw [System.ArgumentException]::New('Path is not in right format')
        }
    }

    hidden [void] SetWRelativePath([string]$WRelativePath) {
        $pattern = "^(?:$([V_WmiPath]::RgxCL)(?:\.$([V_WmiPath]::RgxIN)|=$([V_WmiPath]::RgxSI)){0,1}){0,1}$"
        If ($WRelativePath -match $pattern) {
            $this.ROP_ResolvedPaths.Clear()
            $this.RWP_WClassName = $matches['ClassName']
            $this.RWP_WInstance = $matches['Instance'] + $matches['Singleton']
            $this.UpdateFlags()
            $this.UpdateRelativePath()
            $this.UpdatePath()
        }
        Else {
            throw [System.ArgumentException]::New('Relative path is not in right format')
        }
    }

    hidden [void] SetWRootPath([string]$WRootPath) {
        $pattern = "^(?:(?:\\\\|//)$([V_WmiPath]::RgxCN))(?:(?:\\|/)$([V_WmiPath]::RgxNS))$"
        If ($WRootPath -match $pattern) {
            If ([V_WmiPath]::CheckWNamespace($matches['Namespace'])) {
                $this.ROP_ResolvedPaths.Clear()
                $this.RWP_Server = If ($matches['ComputerName']) { $matches['ComputerName'] }Else { '.' }
                $this.RWP_WNamespace = If ($matches['Namespace']) { $matches['Namespace'] } Else { [System.Management.ManagementPath]::DefaultPath.NamespacePath }
                $this.UpdateFlags()
                $this.UpdateRootPath()
                $this.UpdateNamespaceParts()
                $this.UpdatePath()
            }
            Else {
                throw [System.ArgumentException]::New('Namespace part is not like "ROOT\*" or empty')
            }
        }
        Else {
            throw [System.ArgumentException]::New('Root path is not in right format')
        }
    }

    hidden [void] SetServer([string]$Server) {
        $pattern = "^(?:(?:\\\\|//)$([V_WmiPath]::RgxCN)(?:\\|/|)){0,1}$"
        If ($Server -match $pattern) {
            $this.ROP_ResolvedPaths.Clear()
            $this.RWP_Server = If ($matches['ComputerName']) { $matches['ComputerName'] }Else { '.' }
            $this.UpdateRootPath()
            $this.UpdatePath()
        }
        Else {
            throw [System.ArgumentException]::New('Server name is not in right format')
        }
    }

    hidden [void] SetWNamespacePath([string]$WNamespacePath) {
        $pattern = "^(?:(?:\\|/|)$([V_WmiPath]::RgxNS)(?:\\|/|)){0,1}$"
        If ($WNamespacePath -match $pattern) {
            If ([V_WmiPath]::CheckWNamespace($matches['Namespace'])) {
                $this.ROP_ResolvedPaths.Clear()
                $this.RWP_WNamespace = If ($matches['Namespace']) { $matches['Namespace'] } Else { [System.Management.ManagementPath]::DefaultPath.NamespacePath }
                $this.UpdateFlags()
                $this.UpdateRootPath()
                $this.UpdateNamespaceParts()
                $this.UpdatePath()
            }
            Else {
                throw [System.ArgumentException]::New('Namespace part is not like "ROOT\*" or empty')
            }
        }
        Else {
            throw [System.ArgumentException]::New('Namespace is not in right format')
        }
    }

    hidden [void] SetWClassName([string]$WClassName) {
        $pattern = "^$([V_WmiPath]::RgxCL){0,1}$"
        If ($WClassName -match $pattern) {
            $this.ROP_ResolvedPaths.Clear()
            $this.RWP_WClassName = $matches['ClassName']
            $this.UpdateFlags()
            $this.UpdateRelativePath()
            $this.UpdatePath()
        }
        Else {
            throw [System.ArgumentException]::New('Classname is not in right format')
        }
    }

    hidden [void] SetWInstance([string]$WInstance) {
        $pattern = "^(?:$([V_WmiPath]::RgxIN)|$([V_WmiPath]::RgxSI)){0,1}$"
        If ($WInstance -match $pattern) {
            $this.ROP_ResolvedPaths.Clear()
            $this.RWP_WInstance = $matches['Instance'] + $matches['Singleton']
            $this.UpdateFlags()
            $this.UpdateRelativePath()
            $this.UpdatePath()
        }
        Else {
            throw [System.ArgumentException]::New('Instance name is not in right format')
        }
    }

    hidden [void] UpdateFlags() {
        #Namespace
        If ($this.RWP_WNamespace -ne [string]::Empty) {
            $this.ROP_WildcardInNamespace = [WildcardPattern]::ContainsWildcardCharacters($this.RWP_WNamespace)
            $this.ROP_IsNamespace = ($this.RWP_WClassName -eq [string]::Empty)
        }
        Else {
            $this.ROP_WildcardInNamespace = $false
            $this.ROP_IsNamespace = $false
        }
        #Class
        If ($this.RWP_WClassName -ne [string]::Empty) {
            $this.ROP_WildcardInClassName = [WildcardPattern]::ContainsWildcardCharacters($this.RWP_WClassName)
            $this.ROP_IsClass = ($this.RWP_WInstance -eq [string]::Empty)
        }
        Else {
            $this.ROP_WildcardInClassName = $false
            $this.ROP_IsClass = $false
        }
        #Instance or Singleton
        If ($this.RWP_WInstance -ne [string]::Empty) {
            $this.ROP_WildcardInInstance = [WildcardPattern]::ContainsWildcardCharacters($this.RWP_WInstance)
            $this.ROP_IsInstance = $true
            $this.ROP_IsSingleton = ($this.RWP_WInstance -eq '@')
        }
        Else {
            $this.ROP_WildcardInInstance = $false
            $this.ROP_IsInstance = $false
            $this.ROP_IsSingleton = $false
        }
    }

    hidden [void] UpdateRelativePath() {
        If ($this.ROP_IsInstance) {
            If ($this.ROP_IsSingleton) {
                $this.RWP_WRelativePath = $this.RWP_WClassName + '=@'
            }
            Else {
                $this.RWP_WRelativePath = $this.RWP_WClassName + '.' + $this.RWP_WInstance
            }
        }
        Else {
            $this.RWP_WRelativePath = $this.RWP_WClassName
        }
    }

    hidden [void] UpdateRootPath() {
        $this.RWP_WRootPath = '\\{0}\{1}' -f $this.RWP_Server, $this.RWP_WNamespace
    }

    hidden [void] UpdateNamespaceParts() {
        $this.WNamespaceParts.Clear()
        $NamespaceParts = $this.RWP_WNamespace.split('/').Split('\') | Where-Object { $_ -ne [string]::Empty }
        ForEach ($NSPart in $NamespaceParts) {
            $null = $this.WNamespaceParts.Add(
                [PSCustomObject]@{
                    Value      = $NSPart
                    isWildcard = [WildcardPattern]::ContainsWildcardCharacters($NSPart)
                }
            )
        }
    }

    hidden [void] UpdatePath() {
        $this.RWP_WPath = $this.RWP_WRootPath
        If ($this.RWP_WRelativePath) {
            $this.RWP_WPath = $this.RWP_WPath + ':' + $this.RWP_WRelativePath
        }
    }

    #Get BaseFilter as Hashtable
    [hashtable] GetBaseFilter() {
        Return @{
            ExcludeSystem = $this.RWP_BaseFilter[0]
            ExcludeCIM    = $this.RWP_BaseFilter[1]
            ExcludeMSFT   = $this.RWP_BaseFilter[2]
            ExcludeWin32  = $this.RWP_BaseFilter[3]
        }
    }

    #Set BaseFilter from bools
    [void] SetBaseFilter([bool]$ExcludeSystem, [bool]$ExcludeCIM, [bool]$ExcludeMSFT, [bool]$ExcludeWin32) {
        $this.RWP_BaseFilter[0] = $ExcludeSystem
        $this.RWP_BaseFilter[1] = $ExcludeCIM
        $this.RWP_BaseFilter[2] = $ExcludeMSFT
        $this.RWP_BaseFilter[3] = $ExcludeWin32
    }

    #Set BaseFilter from hashtable
    [void] SetBaseFilter([hashtable]$HT) {
        $this.RWP_BaseFilter[0] = [bool]$HT['ExcludeSystem']
        $this.RWP_BaseFilter[1] = [bool]$HT['ExcludeCIM']
        $this.RWP_BaseFilter[2] = [bool]$HT['ExcludeMSFT']
        $this.RWP_BaseFilter[3] = [bool]$HT['ExcludeWin32']
    }

    #Get AdvancedFilter string
    [string] GetAdvancedWQLFilter() {
        Return $this.RWP_AdvancedFilter
    }

    #Set AdvancedFilter from string
    [void] SetAdvancedWQLFilter([string]$Filter) {
        $null = $this.ROP_ResolvedPaths.Clear()
        If ($Filter -ne [string]::Empty) {
            If ($Filter.StartsWith('WHERE')) {
                $this.RWP_AdvancedFilter = $Filter
            }
            Else {
                $this.RWP_AdvancedFilter = "WHERE $Filter"
            }
        }
        Else {
            $this.RWP_AdvancedFilter = [string]::Empty
        }
    }

    #Returns path string. May contain wildcards
    [string] ToString() {
        Return $this.RWP_WPath
    }

    #Partial clone
    [V_WmiPath] Clone() {
        Return [V_WmiPath]::New($this.RWP_WPath)
    }

    #Resolves wildcard-path into exact existing paths. If Recurse, searches throught full namespace hierarchy. Recurse is time-consuming.
    [int] Resolve([bool]$Recurse) {
        $null = $this.ROP_ResolvedPaths.Clear()

        Try {
            #NAMESPACES
            If ($this.ROP_WildcardInNamespace) {
                If ($Recurse) {
                    #Search namespaces recursivly
                    $Pattern = $this.RWP_WRootPath
                    ForEach ($NSPath in [V_WmiPath]::GetAllNamespaces($this.RWP_Server)) {
                        If ($NSPath.Path -like $Pattern) {
                            $null = $this.ROP_ResolvedPaths.Add($NSPath)
                        }
                    }
                }
                Else {
                    #Search throught namespaces
                    $null = $this.ROP_ResolvedPaths.Add([System.Management.ManagementPath]::New("\\$($this.RWP_Server)\root"))

                    If ($this.WNamespaceParts.Count -gt 1) {
                        For ($i = 1; $i -lt $this.WNamespaceParts.Count; $i++) {
                            If ($this.WNamespaceParts[$i].isWildcard) {
                                $Nick = $this.ROP_ResolvedPaths.Count
                                ForEach ($Path in $this.ROP_ResolvedPaths.Clone()) {
                                    $Pattern = $Path.NamespacePath + '\' + $this.WNamespaceParts[$i].Value
                                    ForEach ($result in ([V_WmiPath]::GetChildNamespaces($Path) | Where-Object { $_.NamespacePath -like $Pattern })) {
                                        $null = $this.ROP_ResolvedPaths.Add($result)
                                    }
                                }
                                $null = $this.ROP_ResolvedPaths.RemoveRange(0, $Nick)
                            }
                            Else {
                                ForEach ($Path in $this.ROP_ResolvedPaths) {
                                    $Path.NamespacePath = $Path.NamespacePath + '\' + $this.WNamespaceParts[$i].Value
                                }
                            }
                        }
                    }
                }
            }
            Else {
                #There is single namespace
                $null = $this.ROP_ResolvedPaths.Add([System.Management.ManagementPath]::New($this.RWP_WRootPath))
            }

            #CLASSNAME
            If ($this.ROP_WildcardInClassName) {
                #Search throught classnames
                $Nick = $this.ROP_ResolvedPaths.Count
                ForEach ($Path in $this.ROP_ResolvedPaths.Clone()) {
                    $Pattern = $this.RWP_WClassName
                    ForEach ($result in ([V_WmiPath]::GetChildClasses($Path) | Where-Object { [V_WmiPath]::PassesFilter($_.ClassName, $this.RWP_BaseFilter) } | Where-Object { $_.ClassName -like $Pattern })) {
                        $null = $this.ROP_ResolvedPaths.Add($result)
                    }
                }
                $null = $this.ROP_ResolvedPaths.RemoveRange(0, $Nick)
            }
            ElseIf ($this.ROP_IsClass -or $this.ROP_IsInstance -or $this.ROP_IsSingleton) {
                #There is a single classname
                If ([V_WmiPath]::PassesFilter($this.RWP_WClassName, $this.RWP_BaseFilter)) {
                    ForEach ($Path in $this.ROP_ResolvedPaths) {
                        $Path.ClassName = $this.RWP_WClassName
                    }
                    $ii = [System.Collections.ArrayList]::New()
                    For ($i = 0; $i -lt $this.ROP_ResolvedPaths.Count; $i++) {
                        If (-not [V_WmiPath]::TestPath($this.ROP_ResolvedPaths[$i])) {
                            $null = $ii.Add($i)
                        }
                    }
                    $null = $ii.Reverse()
                    ForEach ($i in $ii) {
                        $null = $this.ROP_ResolvedPaths.RemoveAt($i)
                    }
                }
                Else {
                    $null = $this.ROP_ResolvedPaths.Clear()
                }
            }

            #INSTANCE OR SINGLETON
            If ($this.ROP_WildcardInInstance) {
                #Search throught instances
                $Nick = $this.ROP_ResolvedPaths.Count
                ForEach ($Path in $this.ROP_ResolvedPaths.Clone()) {
                    $Pattern = $Path.ClassName + '.' + $this.RWP_WInstance
                    ForEach ($result in ([V_WmiPath]::GetChildInstances($Path, $this.RWP_AdvancedFilter) | Where-Object { $_.RelativePath -like $Pattern })) {
                        $null = $this.ROP_ResolvedPaths.Add($result)
                    }
                }
                $null = $this.ROP_ResolvedPaths.RemoveRange(0, $Nick)
            }
            ElseIf ($this.ROP_IsInstance -or $this.ROP_IsSingleton) {
                If ($this.ROP_IsInstance) {
                    #There is a single instance
                    ForEach ($Path in $this.ROP_ResolvedPaths) {
                        $Path.RelativePath = $Path.ClassName + '.' + $this.RWP_WInstance
                    }
                }
                ElseIf ($this.ROP_IsSingleton) {
                    #There is a singleton
                    ForEach ($Path in $this.ROP_ResolvedPaths) {
                        $Path.RelativePath = $Path.ClassName + '=@'
                    }
                }
                $ii = [System.Collections.ArrayList]::New()
                For ($i = 0; $i -lt $this.ROP_ResolvedPaths.Count; $i++) {
                    If (-not [V_WmiPath]::TestPath($this.ROP_ResolvedPaths[$i])) {
                        $null = $ii.Add($i)
                    }
                }
                $null = $ii.Reverse()
                ForEach ($i in $ii) {
                    $null = $this.ROP_ResolvedPaths.RemoveAt($i)
                }
                If ($this.RWP_AdvancedFilter) {
                    $ii = [System.Collections.ArrayList]::New()
                    For ($i = 0; $i -lt $this.ROP_ResolvedPaths.Count; $i++) {
                        $InstancePropsString = $this.ROP_ResolvedPaths[$i].RelativePath.Replace("$($this.ROP_ResolvedPaths[$i].ClassName).", '')
                        $InstanceProps = [System.Collections.ArrayList]::New()
                        While ($InstancePropsString -match "^$([V_wmiPath]::RgxIN)$") {
                            $null = $InstanceProps.Add($matches['Key'])
                            $InstancePropsString = $InstancePropsString.Replace($matches['Key'], '') -replace ',$', ''
                        }
                        $LocalAdvancedFilter = $this.RWP_AdvancedFilter + ' AND ' + ($InstanceProps -join ' AND ')
                        $WQL = "SELECT * FROM $($this.ROP_ResolvedPaths[$i].ClassName) $LocalAdvancedFilter"
                        $NSPath = $this.ROP_ResolvedPaths[$i].Clone()
                        $NSPath.RelativePath = ''
                        If (([wmisearcher]::New($NSPath, $WQL).Get()).Count -eq 0) {
                            $null = $ii.Add($i)
                        }
                    }
                    $null = $ii.Reverse()
                    ForEach ($i in $ii) {
                        $null = $this.ROP_ResolvedPaths.RemoveAt($i)
                    }
                }
            }
        }
        Catch {
            $null = $this.ROP_ResolvedPaths.Clear()
        }

        Return $this.ROP_ResolvedPaths.Count
    }

    #Resolve with non-recurse
    [int] Resolve() {
        Return $this.Resolve($false)
    }
    #endregion
}