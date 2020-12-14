1. Cài đặt winservice.

Navigate to the installutil.exe in your .net folder (for .net 4 it's C:\Windows\Microsoft.NET\Framework\v4.0.30319 for example) and use it to install your service, like this:

"C:\Windows\Microsoft.NET\Framework\v4.0.30319\installutil.exe" "c:\myservice.exe"

If it is the x64 compiled service, use "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\installutil.exe"