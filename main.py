# Program Made By Stephen Twait
# Currently in progress and not ready for production

# This program will:
# * open a file called networkdevices.txt that contains a list of device ip addresses one per line
# * Then look for a file called creds.txt and to use for credentials admin,pass
# * check ping connectivity to all IP's
# * Attempt to Connect to each pingable device via SSH then Telnet and it will compare the running-config file to the startup-config file on that device Then
# * Save any differences to a file called [ip_start_run_diff_datetime.txt]
# * then Compare a locally stored config file(if found) with the format of [ip_hostname_config_rev.txt]
# * with the running-config file running on a device and save any differences to a file called [ip_file_run_diff_datetime.txt]
# * Then save the full running config to a file named [ip_hostname_running-config.txt]
# Don't forget to configure Telnet access on the router!

# username admin privilege 15 password 0 Water!123
# line vty 0 4
# privilege level 15
# login local
# transport input telnet ssh


import difflib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os.path
import os
import datetime
from platform import system as system_name  # Returns the system/OS name
from os import system as system_call       # Execute a shell command
from netmiko import ConnectHandler
from subprocess import call


ip_file = b'networkdevices.txt'
credentials_file = b'credentials.txt'

ip_list = []

send_report = True


def file_validity(file):
    if os.path.isfile(file) is True:
        print("\nFile: {} was found...".format(file))
    else:
        print("\nFile {} does not exist! Please check and try again!".format(file))
        # input("please read scrollback and press enter to close")


# Checking IP validity of all IP's in ip_file saves to ip_list
def ip_validity():
    global ip_list
    try:
        # Open user selected file for reading (IP addresses file)
        selected_ip_file = open(ip_file, 'r')

        # Starting from the beginning of the file
        selected_ip_file.seek(0)

        # Reading each line (IP address) in the file
        ip_list_wnl = selected_ip_file.readlines()

        for ips in ip_list_wnl:
            ip_list.append(ips.rstrip("\n"))

        # Closing the file
        selected_ip_file.close()

    except IOError:
        print("\nFile {} does not exist! Please check and try again!\n".format(ip_file))
        # input("please read scrollback and press enter to close")

    # Checking octets
    for ip in ip_list:
        a = ip.split('.')
        if (len(a) == 4) and (1 <= int(a[0]) <= 223) and (int(a[0]) != 127) and (int(a[0]) != 169 or int(a[1]) != 254) and (0 <= int(a[1]) <= 255 and 0 <= int(a[2]) <= 255 and 0 <= int(a[3]) <= 255):
            break
        else:
            print("\nThe IP address {} is INVALID! Please fix in networkdevices.txt!\n".format(ip))
            # input("please read scrollback and press enter to close")
            continue


def ping(ip):
    """
    Returns True if host (str) responds to a ping request.
    Remember that some hosts may not respond to a ping request even if the host name is valid.
    """

    # Ping parameters as function of OS
    parameters = "-n 1" if system_name().lower() == "windows" else "-c 1"

    # Pinging
    return system_call("ping " + parameters + " " + ip) == 0


for ip in reachable:
    print('pinging {} was successful!!'.format(ip))
for ip in unreachable:
    print('pining {} was a failure....'.format(ip))


def get_creds():
    try:
            # Define SSH parameters
        with open(credentials_file, 'r') as selected_user_file:
            usernames = []
            passwords = []

            # Starting from the beginning of the file
            selected_user_file.seek(0)

            for i in selected_user_file.readlines():
                usernames.append(i.split(',')[0])
                passwords.append(i.split(',')[1].rstrip("\n"))

            return usernames, passwords
    except:
        print('something is wrong with the credentials file')


# Using Netmiko to connect to the device and extract the running configuration
def diff_function(each_device, device_types, vendor, usernames, passwords, command):
    # making the send report bool global
    global send_report
    for transport_type in device_types:
        credsworked = False
        for username, password in zip(usernames, passwords):
            try:
                session = ConnectHandler(device_type=transport_type, ip=each_device,
                                         username=username, secret=password, password=password, global_delay_factor=3)
                print('Username: [{}] and Password: [{}] using [{}] worked!'.format(
                    username, password, transport_type))
                credsworked = True
                break
            except:
                print('Username: [{}] and Password: [{}] using [{}] did not work'.format(
                    username, password, transport_type))
                pass
        if credsworked is True:
            break
    if credsworked is True:

        session.enable()
        session_output = session.send_command(command)
        cmd_output = session_output
        session.disconnect()

        # Defining the file from yesterday, for comparison.
        device_cfg_old = os.path.join(each_device, 'cfgfiles' + vendor + '_' + (
            datetime.date.today() - datetime.timedelta(days=1)).isoformat() + '.txt')

        # Create path and filename to device_cfg_new
        device_cfg_new = os.path.join(each_device, 'cfgfiles' + vendor +
                                      datetime.date.today().isoformat() + '.txt')
        # if path  not found make path
        if not os.path.isdir(os.path.dirname(device_cfg_new)):
            os.makedirs(os.path.dirname(device_cfg_new))

        with open(device_cfg_new, 'w') as device_cfg_new:
            device_cfg_new.write(cmd_output)

        # Defining the differences file as diff_file + the current date and time.
        diff_file_date = os.path.join(each_device, 'cfgfiles' + vendor +
                                      'diff_file_' + datetime.date.today().isoformat() + '.txt')
        if not os.path.isdir(os.path.dirname(diff_file_date)):
            os.makedirs(os.path.dirname(diff_file_date))

        # The same for the final report file.
        report_file_date = os.path.join(
            each_device, 'cfgfiles' + vendor + 'report_' + datetime.date.today().isoformat() + '.txt')
        if not os.path.isdir(os.path.dirname(report_file_date)):
            os.makedirs(os.path.dirname(report_file_date))

        # Opening the old config file, the new config file for reading and a new
        # file to write the differences.
        if os.path.isfile(device_cfg_old):
            with open(device_cfg_old, 'r') as old_file, open(device_cfg_new, 'r') as new_file, open(diff_file_date, 'w') as diff_file:
                # Using the ndiff() method to read the differences.
                diff = difflib.ndiff(old_file.readlines(), new_file.readlines())
                # Writing the differences to the new file.
                diff_file.write(''.join(list(diff)))

            # Opening the new file, reading each line and creating a list where each
            # element is a line in the file.
            with open(str(diff_file_date), 'r') as diff_file:
                # Creating the list of lines.
                diff_list = diff_file.readlines()
                # print diff_list

            # Interating over the list and extracting the differences by type. Writing all the differences to the report file.
            # Using try/except to catch and ignore any IndexError exceptions that might occur.
            try:
                with open(str(report_file_date), 'a') as report_file:
                    for index, line in enumerate(diff_list):
                        if line.startswith('- ') and diff_list[index + 1].startswith(('?', '+')) is False:
                            report_file.write('\nWas in old version, not there anymore: ' +
                                              '\n\n' + line + '\n-------\n\n')
                        elif line.startswith('+ ') and diff_list[index + 1].startswith('?') is False:
                            report_file.write('\nWas not in old version, is there now: ' + '\n\n' + '...\n' +
                                              diff_list[index - 2] + diff_list[index - 1] + line + '...\n' + '\n-------\n')
                        elif line.startswith('- ') and diff_list[index + 1].startswith('?') and diff_list[index + 2].startswith('+ ') and diff_list[index + 3].startswith('?'):
                            report_file.write('\nChange detected here: \n\n' + line +
                                              diff_list[index + 1] + diff_list[index + 2] + diff_list[index + 3] + '\n-------\n')
                        elif line.startswith('- ') and diff_list[index + 1].startswith('+') and diff_list[index + 2].startswith('? '):
                            report_file.write('\nChange detected here: \n\n' + line +
                                              diff_list[index + 1] + diff_list[index + 2] + '\n-------\n')
                        else:
                            pass
            except IndexError:
                pass

            # Reading the report file and writing to the master file.
            with open(str(report_file_date), 'r') as report_file, open(os.path.join('cfgfiles' + 'master_report_' + datetime.date.today().isoformat() + '.txt'), 'a') as master_report:
                if len(report_file.readlines()) < 1:
                    send_report = False
                    # Adding device as first line in report.
                    master_report.write('\n\n*** Device: ' + each_device + ' ***\n')
                    master_report.write('\n' + 'No Configuration Changes Recorded On ' +
                                        datetime.datetime.now().isoformat() + '\n\n\n')
                else:
                    # Appending the content to the master report file.
                    # report_file.seek(0)
                    master_report.write('\n\n*** Device: ' + each_device + ' ***\n\n')
                    master_report.write(report_file.read())
        else:
            print("no previous config file found saving todays With yestedays date for re-run functiality test")
            with open(device_cfg_old, 'w') as device_cfg_old:
                device_cfg_old.write(cmd_output)
    else:
        print("No Valid Credentials found for IP [{}]".format(ip))


def send_email(pingable, not_pingable):
    try:

        fromaddr = 'DoNotReply@7seaswater.com'
        toaddr = 'stwait@7seaswater.com'
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = 'Daily Configuration Change Report'

        # Checking whether any changes were recorded and building the email body.
        with open(os.path.join('cfgfiles' + 'master_report_' + datetime.date.today().isoformat() + '.txt'), 'r') as master_report:
            master_report.seek(0)

            body = ('\n' + "IP's that were pingable: \n" + "\n".join(pingable) + " \nIP's that were unpingable: \n" + "\n".join(not_pingable) +
                    master_report.read() + '\n****************\n\nReport Generated: ' +
                    datetime.datetime.now().isoformat() + '\n\nEnd Of Report\n')
            msg.attach(MIMEText(body, 'plain'))

            # Sending the email.
            server = smtplib.SMTP('smtp.7seaswater.com', 25)
            server.helo()
            text = msg.as_string()
            server.sendmail(fromaddr, toaddr, text)
    except:
        print('could not send email')
        pass


# Start Running Program
file_validity(ip_file)
file_validity(credentials_file)
ip_validity()
usernames, passwords = get_creds()

# Clearing master report
open(os.path.join('cfgfiles' + 'master_report_' +
                  datetime.date.today().isoformat() + '.txt'), 'w').close()

pingable = []
not_pingable = []
for ip in ip_list:
    if ping(ip) is True:
        pingable.append(ip)
        print("\n")
        print("for Device [{}] ....".format(ip))
        diff_function(ip, ['cisco_ios', 'cisco_ios_telnet'],
                      'cisco', usernames, passwords, 'show running')
    else:
        not_pingable.append(ip)


if send_report is True:
    print('Changes Detected, Sending Email Report')
    send_email(pingable, not_pingable)
else:
    print('No Changes Detected, Not sending Email Report')


# # Creating threads
# def create_start_run_threads():
#     threads = []
#     for ip in ip_list:
#         # args is a tuple with a single element
#         th = threading.Thread(target=compare_run_start, args=(ip,))
#         th.start()
#         threads.append(th)
#     for th in threads:
#         th.join()
#
#
# def create_run_file_threads():
#     threads = []
#     for ip in ip_list:
#         # args is a tuple with a single element
#         th = threading.Thread(target=compare_run_file, args=(ip,))
#         th.start()
#         threads.append(th)
#     for th in threads:
#         th.join()
#
#
# # Calling threads creation function
# create_start_run_threads()
# create_run_file_threads()
