# VBScript WMI Wrapper

Accessing Windows Management Instrumentation (*WMI*) through `WMIC` can at times be restricted by system administrators. However, those same administrators always forget to restrict the use of VBScript through Windows Shell Host (*Wsh*). When a script gets ran through Wsh and uses system APIs (such as WMI API), the system usually allows it. In case your system administrator forgets about that, this tool is for you!

# Installation

`WMI.vbs` is a console program. Copy it onto your executable path, and run `cscript WMI.vbs` to get started

# Usage
The most simple command way to use WMI.vbs is by supplying it a Query:

    cscript WMI.vbs /Query:"Select * From Win32_UserAccount"

The tool also allows the use of different computers on the network, as well as different WMI namespaces:

    cscript WMI.vbs /Query:"Select * From Win32_Process" /Computer:127.0.0.1 /Namespace:"\root\cimv2"


