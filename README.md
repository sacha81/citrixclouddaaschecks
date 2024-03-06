The script **Invoke-Citrix_HealthCheck.ps1** is designed to perform health checks on a Citrix Cloud Desktop as a Service (DaaS) environment. Here's an overview of its functionality:

Variable Setup: The script starts by defining various variables, including the names of delivery groups, whether to show only error VDI, whether to perform advanced checks, maximum uptime for VDAs, Citrix Cloud API endpoints, client credentials, paths to client secrets, email settings, and more.

HTML Result Generation: The script generates an HTML report with the health status of the Citrix environment. It creates headers, footers, and tables to organize the data neatly.

Functions: The script defines several functions to facilitate HTML report generation, including WriteHtmlHeader, WriteHtmlFooter, Out-Html, Get-HTMLStyle, Get-HTMLHeader, Get-HTMLBodyHeader, Get-HTMLTableHeader, Get-HTMLFooter, WriteTableHeader, WriteTableFooter, and WriteTableRow.

Main Functionality:
- It queries the Citrix DaaS environment via its REST API using the provided credentials.
- It retrieves information about virtual desktop instances (VDIs) within specified delivery groups.
- It checks various attributes of each VDI, such as power state, health score, maintenance mode, uptime, sessions, last connection time, registration state, version, associated user names, allocation type, session support, tags, and hosted information.
- It generates an HTML report summarizing the health status of the Citrix environment, including infrastructure health score and detailed information about each VDI. Optionally, it sends an email with the HTML report attached to specified recipients.
Overall, the script provides administrators with a comprehensive overview of the health status of their Citrix DaaS environment, enabling them to identify any issues and take appropriate actions to maintain optimal performance and reliability.
