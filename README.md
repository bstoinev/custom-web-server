# custom-web-server

This is an exerpt of a web server, designated to serve requests on the localhost interface, consumed by another web site.

One of the functions is to provision the web application with data which is unavailable through javascript - report the OS and other products version, read Windows registry, etc. See AmbientReportCommand.cs 

Also, the custom-web-server accepts commands to execute local to the OS actions, such as starting applications in the localy logged in user session. Check LaunchProcessCommand.cs and LaunchProcessCommandWin32Helper.cs

The custm-web-server is loaded in a separate AppDomain and make use of shadow copies to hot-update the executing files. See KernelBootstrapper.cs

