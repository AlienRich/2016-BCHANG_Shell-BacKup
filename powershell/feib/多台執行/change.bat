@echo off
    ::主機IP清單
      set ip_addr=10.48.14.59 10.48.14.29 10.48.14.53 10.48.14.87 10.48.14.45 10.48.14.115 10.48.14.19 10.48.14.77 10.48.14.65 10.48.14.41 10.48.14.111 10.48.14.21 10.48.14.79 10.48.14.43 10.48.14.113 10.48.14.85 10.48.14.122 10.48.14.121

    if exist %~dp0過版_rpt.bat for %%L in (%ip_addr%) do call %~dp0過版_rpt.bat %%L
    if exist %~dp0過版_exe.bat for %%L in (%ip_addr%) do call %~dp0過版_exe.bat %%L
exit