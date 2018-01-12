<#
Tools for extracting .NET COM information from .NET assemblies.

This file is part of DotNetInteropDemos
Copyright (C) James Forshaw 2017

DotNetInteropDemos is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

DotNetInteropDemos is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with DotNetInteropDemos.  If not, see <http://www.gnu.org/licenses/>.

#>

function Get-Assembly {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [string]$Name
    )

    PROCESS {
        $asm = $null
        $exists = Test-Path $Name
        if ($exists) {
            $path = Resolve-Path $Name
            $asm = [Reflection.Assembly]::LoadFrom($path);
        } else {
            $asm = [Reflection.Assembly]::LoadWithPartialName($Name);
        }
        Write-Output $asm
    }
}


function Get-IsGenericType {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )

    $Type.IsGenericType -or $Type.IsConstructedGenericType -or $Type.IsGenericTypeDefinition
}

function Get-COMTypesFromAssembly {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Reflection.Assembly]$Assembly
    )
    PROCESS {
        $Assembly.GetTypes() | Where-Object {[System.Runtime.InteropServices.Marshal]::IsTypeVisibleFromCom($_)}
    }
}

function Get-IsConstructableCOMType {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )
    $Type.IsPublic -and $Type.IsClass -and !$Type.IsAbstract -and ($Type.GetConstructor([Type[]]@()) -ne $null)
}

function Get-IsIgnorableCOMType {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )
    [Exception].IsAssignableFrom($Type) -or [Attribute].IsAssignableFrom($Type) -or $Type.IsAbstract
}

function Get-IsDispatchInterface {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )

    $t = [System.Runtime.InteropServices.InterfaceTypeAttribute]
    $attr = $Type.GetCustomAttributes($t, $false)
    if ($attr.Length -gt 0) {
        return $attr[0].Value -eq "InterfaceIsIDispatch" -or $attr[0].Value -eq "InterfaceIsDual"
    }

    return $true
}

function Get-IsCOMImport {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )

    #$t = [System.Runtime.InteropServices.ComImportAttribute]
    #$attr = $Type.GetCustomAttributes($t, $false)
    #$attr.Length -gt 0
    return ($Type.Attributes -band "Import") -eq "Import"
}

function Get-ClassInterface {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )
    $t = [System.Runtime.InteropServices.ClassInterfaceAttribute]
    $attr = $Type.GetCustomAttributes($t, $false)
    if ($attr.Length -eq 0) {
        return [System.Runtime.InteropServices.ClassInterfaceType]::AutoDispatch
    }

    return $attr[0].Value
}

function Get-DefaultCOMInterface
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )

    if (!$Type.IsClass) {
        return $null
    }

    if ($Type -eq [System.Object]) {
        return $null
    }

    $attr = $Type.GetCustomAttributes([System.Runtime.InteropServices.ComDefaultInterfaceAttribute], $false)
    if ($attr.Count -gt 0) {
        return $attr[0].Value
    }

    
    foreach($intf in $Type.GetInterfaces()) {
        $generic = Get-IsGenericType $intf
        if ($generic) {
            continue
        }
        if ([System.Runtime.InteropServices.Marshal]::IsTypeVisibleFromCom($intf) -and !$intf.IsAssignableFrom($Type.BaseType)) {
            return $intf
        }
    }

    return Get-DefaultCOMInterface $Type.BaseType
}

# Based on GetDefaultInterfaceForClassInternal from coreclr
function Get-IsDispatchClass {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )

    if (Get-IsCOMImport $Type) 
    {
        $iftype = [System.Runtime.InteropServices.ClassInterfaceType]::None
    }
    else
    {
        $iftype = Get-ClassInterface $Type
    }

    if ($iftype -eq "AutoDispatch" -or $iftype -eq "AutoDual") {
        return $true
    }

    $intf = Get-DefaultCOMInterface $Type
    # This seems different to definition in coreclr which
    # indicates the type should be IUnknown.
    if ($intf -eq $null) {
        return $true
    }

    return Get-IsDispatchInterface $intf
}

function Get-IsDispatch {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )

    if ($Type.IsInterface) {
        return Get-IsDispatchInterface $Type
    } elseif ($Type.IsClass) {
        return Get-IsDispatchClass $Type
    }
    return $false
}

function Get-IsCOMRegistered
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]
        [Type]$Type
    )

    $guid = $Type.GUID.ToString("B")
    $key = $null

    if ($Type.IsInterface) {
        $key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("Interfaces\$guid")
    } elseif ($Type.IsClass) {
        $key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("CLSID\$guid")
    }

    if($key -ne $null) {
        $key.Close()
        return $true
    }
    return $false
}

function Get-IsMemberCOMVisible {
    param([System.Reflection.MemberInfo]$Member)
    $type = [System.Runtime.InteropServices.ComVisibleAttribute]
    $attr = $Member.GetCustomAttributes($t, $false)
    if ($attr.Count -gt 0) {
        return $attr.Value
    }
    return $true
}

function Get-COMMethods {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Type]$Type
    )

    $methods = @()

    if (!$Type.IsInterface -and !$Type.IsClass) {
        return @();
    }

    if ($Type.IsClass) {
        $classintf = Get-ClassInterface $Type
        if ($classintf -eq "None") {
            # If class is None then check for a default interface, if $null
            # then the class is AutoDispatch
            $intf = Get-DefaultCOMInterface $Type
            if ($intf -ne $null) {
                $Type = $intf
            }
        }
    }

    $mis = $Type.GetMethods()
    [array]::Reverse($mis)
    $namemap = @{}

    foreach($mi in $mis) {
        #$visible = Get-IsMemberCOMVisible $mi
        #if (!$visible) {
        #    continue;
        #}
        $name = $mi.Name
        if ($namemap.ContainsKey($name.ToLower())) {
            $index = $namemap[$name.ToLower()]            
            $namemap[$name.ToLower()] = $index + 1
            $name = $name + "_$index"
        } else {
            $namemap.Add($name.ToLower(), 2)
        }

        if ($mi.IsSpecialName) {
            continue
        }

        $props = @{
            Method=$mi;
            Name=$name;
            }

        $obj = New-Object –TypeName PSObject –Prop $props
        $methods += @($obj)
    }

    return $methods
}

function Get-COMProperties {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Type]$Type
    )

    $properties = @()

    if (!$Type.IsInterface -and !$Type.IsClass) {
        return @();
    }

    if ($Type.IsClass) {
        $classintf = Get-ClassInterface $Type
        if ($classintf -eq "None") {
            # If class is None then check for a default interface, if $null
            # then the class is AutoDispatch
            $intf = Get-DefaultCOMInterface $Type
            if ($intf -ne $null) {
                $Type = $intf
            }
        }
    }

    $pis = $Type.GetProperties()
    [array]::Reverse($pis)
    $namemap = @{}

    foreach($pi in $pis) {
        #$visible = Get-IsMemberCOMVisible $pi
        #if (!$visible) {
        #    continue;
        #}
        $name = $pi.Name
        if ($namemap.ContainsKey($name.ToLower())) {
            $index = $namemap[$name.ToLower()]
            $namemap[$name.ToLower()] = $index + 1
            $name = $name + "_$index"
        } else {
            $namemap.Add($name.ToLower(), 2)
        }

        $props = @{
            Property=$pi;
            Name=$name;
            }

        $obj = New-Object –TypeName PSObject –Prop $props
        $properties += @($obj)
    }

    return $properties
}


function Format-COMType {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Type]$Type,
        [bool]$NoMethods
    )
    PROCESS {
        $class_intf = Get-ClassInterface $Type
        $def_intf = $null
        if ($class_intf -eq "None") {
            $def_intf = Get-DefaultCOMInterface $Type
        }
        $methods = @()
        $properties = @()
        if (!$NoMethods) {
            $methods = Get-COMMethods $Type
            $properties = Get-COMProperties $Type
        }
        $props = @{
            AssemblyName=$Type.Assembly.GetName().Name;
            IsConstructable=Get-IsConstructableCOMType $Type;
            IsIgnorable=Get-IsIgnorableCOMType $Type;
            IsInterface=$Type.IsInterface;
            IsClass=$Type.IsClass;
            IsDispatch=Get-IsDispatch $Type;
            IsImport=Get-IsCOMImport $Type;
            DefaultInterface=$def_intf;
            IsRegistered=Get-IsCOMRegistered $Type;
            ClassInterface=$class_intf;
            Methods=$methods;
            Properties=$properties;
            Guid=$Type.GUID;
            Type=$Type}
        $obj = New-Object –TypeName PSObject –Prop $props
        Write-Output $obj
    }
}

<#
.SYNOPSIS
Get a list COM accessible types from an assembly.
.DESCRIPTION
This cmdlet enumerates the COM accessible types in a list of assemblies, generates information about the type
and returns the result as a list of PSObject's containing the type properties.
.PARAMETER AssemblyName
Specify a list of assembly names to load and inspect. Can be a short/long assembly name or a path to a file.
.PARAMETER Assembly
Specify a list of loaded assemblies.
.PARAMETER NoMethods
Set to not enumerate method and property information.
.INPUTS
string[] - AssemblyName from pipeline.
.OUTPUTS
PSObject
.EXAMPLE
@("mscorlib", "System") | Get-ComTypes
Get all COM accessible types from the mscorlib and System assemblies.
.EXAMPLE
@("mscorlib", "System") | Get-ComTypes | Where-Object {$_.IsConstructable -and !$_.IsRegistered}
Get all COM accessible types from the mscorlib and System assemblies which are constructable by a COM client but not yet registered.
#>
function Get-ComTypes {
    [CmdletBinding(DefaultParameterSetName="FromName")]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ParameterSetName="FromName")]
        [string[]]$AssemblyName,
        [parameter(Mandatory=$true, Position=0, ParameterSetName="FromAssembly")]
        [System.Reflection.Assembly[]]$Assembly,
        [switch]$NoMethods
    )
    BEGIN {
        $created = New-Object System.Collections.Generic.HashSet[string]
    }
    PROCESS {
        if ($Assembly -eq $null) {
            $asm = $AssemblyName | Get-Assembly
        } else {
            $asm = $Assembly
        }

        $types = $asm | Where-Object { $created.Add($_.Fullname) } | Get-COMTypesFromAssembly | Format-COMType -NoMethods $NoMethods
        if ($types.Count -gt 0) {
            Write-Output $types
        }
    }
}

function Format-RegistryKey {
   param(
        [parameter(Mandatory=$true, Position=0)]
        [string]$BasePath,
        [parameter(Mandatory=$true, Position=1)]
        [System.Text.StringBuilder]$Builder,
        [parameter(Mandatory=$true, Position=2)]
        [string]$Name,
        [string]$Default,
        [Hashtable]$Values = @{}
    )

    $Builder.AppendFormat("[{0}\{1}]", $BasePath, $Name).AppendLine() | Out-Null
    if ($Default.Length -gt 0) {
        $Builder.AppendFormat("@=""{0}""", $Default).AppendLine() | Out-Null
    }

    if ($Values.Count -gt 0) {
        foreach($pair in $Values.GetEnumerator()) {
            $Builder.AppendFormat("""{0}""=""{1}""", $pair.Name, $pair.Value).AppendLine() | Out-Null
        }
    }

    $Builder.AppendLine() | Out-Null
}

<#
.SYNOPSIS
Output a registry file for use with regedit or reg.exe to register a list of .NET COM classes.
.DESCRIPTION
This cmdlet accepts a list of COM accessible types and generates a registry file which can be used to register those types.
.PARAMETER Type
Specify a list of types. Any types which are not COM accessible and constructable are ignored.
.PARAMETER Output
Specify the output file.
.PARAMETER CodeBase
Specify to emit the CodeBase value for the registration.
.PARAMETER LocalMachine
Specify to register the types in the Local Machine hive. The default is the Current User hive.
.PARAMETER Version2
Specify to register the types with the v2 runtime. Default is the v4 runtime.
.PARAMETER FakeClsid
Specify to generate a new random CLSID when generating the registry file. This can be useful to allow you to use objects which have already been registered.
.PARAMETER ClsidList
Specify a list of CLSIDs to use for the registrations.
.INPUTS
Type[] - COM accessible types.
.OUTPUTS
None
.EXAMPLE
Get-ComTypes "System" | Out-ComTypeRegistry -Output system_types.reg
Get all COM types from the System assembly and generate a registry file to register them.
#>
function Out-ComTypeRegistry {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName = $true)]
        [System.Type[]]$Type,
        [parameter(Mandatory=$true, Position=1)]
        [string]$Output,
        [switch]$CodeBase,
        [switch]$LocalMachine,
        [switch]$Version2,
        [switch]$FakeClsid,
        [string[]]$ClsidList
    )
    BEGIN {
        "REGEDIT4`r`n" | Set-Content $Output
        $created = New-Object System.Collections.Generic.HashSet[string]
        $clsid_index = 0
    }
    PROCESS {
        $builder = [System.Text.StringBuilder]::new()
        foreach($type in $Type) {
            $fullname = $type.FullName
            $valid = Get-IsConstructableCOMType $type[0]
            if (!$valid) {
                $PSCmdLet.WriteWarning("Type $fullname is not a constructable COM object")
                continue
            }

            if (!$created.Add($fullname)) {
                $PSCmdLet.WriteWarning("Type $fullname has already been emitted")
                continue
            }

            $base = "HKEY_CURRENT_USER"
            if ($LocalMachine) {
                $base = "HKEY_LOCAL_MACHINE"
            }
            $base += "\Software\Classes"

            $clsid = $type.GUID
            if ($FakeClsid) {
              $clsid = [Guid]::NewGuid()
            } elseif ($ClsidList.Count -gt 0) {
                if ($clsid_index -eq $ClsidList.Count) {
                  $PSCmdLet.WriteWarning("Used up all the CLSIDs, using a fake one")
                  $clsid = [Guid]::NewGuid()
                } else {
                  $clsid = [Guid]::Parse($ClsidList[$clsid_index++])
                }
            }

            $guid = $clsid.ToString("B").ToUpper()

            $version = "v4.0.30319"
            if ($Version2) {
                $version = "v2.0.50727"
            }

            Format-RegistryKey $base $builder $fullname -Default $fullname
            Format-RegistryKey $base $builder "$fullname\CLSID" -Default $guid
            Format-RegistryKey $base $builder "CLSID\$guid" -Default $fullname

            $regvalues = @{ThreadingModel="Both";Class=$fullname;Assembly=$type.Assembly.FullName;RuntimeVersion=$version}
            if ($CodeBase) {
                $regvalues.Add("CodeBase", $type.Assembly.CodeBase)
            }

            Format-RegistryKey $base $builder "CLSID\$guid\InprocServer32" -Default "mscoree.dll" -Values $regvalues
            Format-RegistryKey $base $builder "CLSID\$guid\ProgID" -Default $fullname
            Format-RegistryKey $base $builder "CLSID\$guid\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
        }
        $builder | Add-Content $Output
    }
}

<#
.SYNOPSIS
Output a manifest file for use with ActCtx in JScript/VBScript/VBA for a list of .NET COM types.
.DESCRIPTION
This cmdlet accepts a list of COM accessible types and generates a manifest file which can be used to used with the ActCtx object in JScript/VBScript.
.PARAMETER Type
Specify a list of types. Any types which are not COM accessible and constructable are ignored.
.PARAMETER Output
Specify the output file.
.PARAMETER ManifestString
Specify to emit the manifest as a JScript string which can be used with the ManifestText property.
.PARAMETER Version2
Specify to register the types with the v2 runtime. Default is the v4 runtime.
.PARAMETER FakeClsid
Specify to generate a new random CLSID when generating the manifest. This can be useful to allow you to use objects which have already been registered.
.PARAMETER ClsidList
Specify a list of CLSIDs to use for the registrations.
.INPUTS
Type[] - COM accessible types.
.OUTPUTS
None
.EXAMPLE
Get-ComTypes "System" | Out-ComTypeManifest -Output system_types.manifest
Get all COM types from the System assembly and generate a manifest file.
#>
function Out-ComTypeManifest {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName = $true)]
        [System.Type[]]$Type,
        [parameter(Mandatory=$true, Position=1)]
        [string]$Output,
        [switch]$ManifestString,
        [switch]$Version2,
        [switch]$FakeClsid,
        [string[]]$ClsidList
    )
    BEGIN {
        $assembly = $null
        $xmldoc = [System.Xml.XmlDocument]::new()
        $ns = "urn:schemas-microsoft-com:asm.v1"
        if ($ManifestString) {
            $enc = "UTF-16"
        } else {
            $enc = "UTF-8"
        }
        $xmldec = $xmldoc.CreateXmlDeclaration("1.0", $enc, "yes")
        $xmldoc.AppendChild($xmldec) | Out-Null
        $root = $xmldoc.CreateElement("assembly", $ns)
        $xmldoc.AppendChild($root) | Out-Null        
        $root.SetAttribute("manifestVersion", "1.0")
        $created = New-Object System.Collections.Generic.HashSet[string]
        $clsid_index = 0
    }
    PROCESS {
        foreach($type in $Type) {
            $fullname = $type.FullName

            $valid = Get-IsConstructableCOMType $type[0]
            if (!$valid) {
                $PSCmdLet.WriteWarning("Type $fullname is not a constructable COM object")
                continue
            }

            if (!$created.Add($fullname)) {
                $PSCmdLet.WriteWarning("Type $fullname has already been emitted")
                continue
            }

            if ($FakeClsid) {
                $guid = [Guid]::NewGuid().ToString("B").ToUpper()
            } elseif ($ClsidList.Count -gt 0) {
                if ($clsid_index -eq $ClsidList.Count) {
                  $PSCmdLet.WriteWarning("Used up all the CLSIDs, using a fake one")
                  $guid = [Guid]::NewGuid().ToString("B").ToUpper()
                } else {
                  $guid = [Guid]::Parse($ClsidList[$clsid_index++]).ToString("B").ToUpper();
                }
            } else {
                $guid = $type.GUID.ToString("B").ToUpper()
            }

            $asmname = $type.Assembly.GetName()
            if ($assembly -eq $null) {
                # Output assembly identity
                $identity = $xmldoc.CreateElement("assemblyIdentity", $ns)
                $identity.SetAttribute("name", $asmname.Name)
                $identity.SetAttribute("version", $asmname.Version)
                $identity.SetAttribute("publicKeyToken", [System.BitConverter]::ToString($asmname.GetPublicKeyToken()).Replace("-", ""))
                $root.AppendChild($identity) | Out-Null
                $assembly = $asmname.FullName
            } elseif ($assembly -ne $type.Assembly.FullName) {
               $PSCmdLet.WriteWarning("Type $fullname in a different assembly, manifest might not work")
            }
            
            $version = "v4.0.30319"
            if ($Version2) {
                $version = "v2.0.50727"
            }

            $clrclass = $xmldoc.CreateElement("clrClass", $ns)
            $clrclass.SetAttribute("clsid", $guid)
            $clrclass.SetAttribute("progid", $fullname)
            $clrclass.SetAttribute("threadingModel", "Both")
            $clrclass.SetAttribute("name", $fullname)
            $clrclass.SetAttribute("runtimeVersion", $version)
            $root.AppendChild($clrclass) | Out-Null
        }
    }
    END {
        $xml = $xmldoc.OuterXml
        if ($ManifestString) {
            "var manifest = '$xml';" | Set-Content $Output -Encoding UTF8
        } else {
            $xml | Set-Content $Output -Encoding UTF8
        }
    }
}

function Get-PInvokeMethodsFromAssembly {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Reflection.Assembly]$Assembly
    )
    PROCESS {
        $methods = @()
        $types = $Assembly.GetTypes()
        foreach($type in $types) {
            $public = $type.GetMethods("Static, Public") | Where-Object {$_.Attributes -match "PinvokeImpl"}
            $private = $type.GetMethods("Static, NonPublic") | Where-Object {$_.Attributes -match "PinvokeImpl"}
            $methods += $public
            $methods += $private
        }
        Write-Output $methods
    }
}

function Get-DefaultDllImportSearchPaths {
    param(
        [System.Reflection.ICustomAttributeProvider]$Provider
    )
    $type = [System.Runtime.InteropServices.DefaultDllImportSearchPathsAttribute]
    if ($Provider.IsDefined($type, $false)) {
        return $Provider.GetCustomAttributes($type, $false)[0].Paths
    }
    return 0
}

function Format-PInvokeMethod {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [System.Reflection.MethodInfo]$Method
    )
    PROCESS {
        $dllimp_type = [System.Runtime.InteropServices.DllImportAttribute]
        $dllimport = $Method.GetCustomAttributes($dllimp_type, $false)
        if ($dllimport.Count -eq 0) {
            $PSCmdlet.WriteWarning("Method $Method does not have a valid DllImport attribute")
            return
        }

        $is_public = $Method.IsPublic -and $Method.DeclaringType.IsPublic
        $dll_default = Get-DefaultDllImportSearchPaths($Method)
        if ($dll_default -eq 0) {
            $dll_default = Get-DefaultDllImportSearchPaths($Method.DeclaringType.Assembly)
            if ($dll_default -eq 0) {
                $dll_default = [System.Runtime.InteropServices.DllImportSearchPath]::LegacyBehavior
            }
        }

        $props = @{
            AssemblyName=$Method.DeclaringType.Assembly.GetName().Name;
            Type=$Method.DeclaringType;
            Method=$Method;
            DllName=$dllimport[0].Value;
            DllImport=$dllimport[0];
            IsPublic=$is_public;
            DllSearchPaths=$dll_default;
        }
        $obj = New-Object –TypeName PSObject –Prop $props
        Write-Output $obj
    }
}


<#
.SYNOPSIS
Get a list pinvoke methods from an assembly.
.DESCRIPTION
This cmdlet enumerates the types in a list of assemblies, finds any pinvoke methods and formats them into an object.
.PARAMETER AssemblyName
Specify a list of assembly names to load and inspect. Can be a short/long assembly name or a path to a file.
.PARAMETER Assembly
Specify a list of loaded assemblies.
.INPUTS
string[] - AssemblyName from pipeline.
.OUTPUTS
PSObject
.EXAMPLE
@("mscorlib", "System") | Get-PInvokeMethods
Get all PInvoke methods from the mscorlib and System assemblies.
#>
function Get-PInvokeMethods {
    [CmdletBinding(DefaultParameterSetName="FromName")]
    param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ParameterSetName="FromName")]
        [string[]]$AssemblyName,
        [parameter(Mandatory=$true, Position=0, ParameterSetName="FromAssembly")]
        [System.Reflection.Assembly[]]$Assembly
    )
    BEGIN {
        $created = New-Object System.Collections.Generic.HashSet[string]
    }
    PROCESS {
        if ($Assembly -eq $null) {
            $asm = $AssemblyName | Get-Assembly
        } else {
            $asm = $Assembly
        }

        $methods = $asm | Where-Object { $created.Add($_.Fullname) } | Get-PInvokeMethodsFromAssembly | Format-PInvokeMethod
        Write-Output $methods
    }
}
