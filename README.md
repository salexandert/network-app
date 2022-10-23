# network-app
Network Automation Utilities 


Program Made By Stephen Twait
Currently in progress and not ready for production

This program will:
* open a file called networkdevices.txt that contains a list of device ip addresses one per line
* Then look for a file called creds.txt and to use for credentials admin,pass
* check ping connectivity to all IP's
* Attempt to Connect to each pingable device via SSH then Telnet and it will compare the running-config file to the startup-config file on that device Then
* Save any differences to a file called [ip_start_run_diff_datetime.txt]
* then Compare a locally stored config file(if found) with the format of [ip_hostname_config_rev.txt]
* with the running-config file running on a device and save any differences to a file called [ip_file_run_diff_datetime.txt]
* Then save the full running config to a file named [ip_hostname_running-config.txt]
Don't forget to configure Telnet access on the router!