Get-Process msedge, msedgedriver -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item "C:\Users\consvla\tep\DevToolsActivePort" -ErrorAction SilentlyContinue
Remove-Item "C:\Users\consvla\tep\lockfile" -ErrorAction SilentlyContinue
