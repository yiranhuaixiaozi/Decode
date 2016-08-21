copy .\MSCOMM32.ocx %systemroot%\system32\MSCOMM32.ocx /y
regsvr32 %systemroot%\system32\MSCOMM32.ocx /s
regedit /s .\MscommLicence.reg
