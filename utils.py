import os
from colorama import Fore, Style
from ruamel.yaml import YAML
from constants import INTERFACE_NAME_RE, NORMALIZED_INTERFACES
import collections
from nornir.core.filter import F

def check_hosts_in_inventory(nr, hosts, summary_data):

    directory = "./Show Tech/Configs"
    site = summary_data['site']
    new_hosts = {}
    for root, dirs, files in os.walk(directory):
        for f in files:
            new_host_data = {}
            hostname = None
            ip = None            
            
            with open(f"{root}/{f}", 'r', errors='replace') as a:
                config_lines = a.readlines()
            
            for index, line in enumerate(config_lines):
                if line.startswith('hostname'):
                    # print(line)
                    hostname = line.split()[1]
                    # print(f"Hostname [{hostname}]")
                    
                    if hostname not in f:
                        os.rename(f"{root}/{f}", f"{root}/{hostname}.cfg")
                        print(Fore.RED + f"\nSwitch {hostname} found in config file {f} but is not in the filename,\
                         \nThe script resolves this by renaming config files to hostname found in config name but please investigate why." + Fore.RESET)
                    
                # This is not perfect below, create logic for perfect end of file searching, something like, while index + range <= len(config_lines)
                elif 'vlan260' in line.lower():
                    # print(len(config_lines) - index)
                    if (len(config_lines) - index) > 3:
                        for i in range(1, 3):
                            if 'address' in config_lines[index + i].lower():
                                # print(config_lines[index + i])
                                ip = config_lines[index + i].split()[2]
               
                elif 'vlan261' in line.lower():
                    if (len(config_lines) - index) > 3:
                        for i in range(1, 3):
                            if 'address' in config_lines[index + i].lower():
                                # print(config_lines[index + i])
                                ip = config_lines[index + i].split()[2]   
                
                elif 'vlan600' in line.lower():
                    if (len(config_lines) - index) > 3:
                        for i in range(1, 3):
                            if 'address' in config_lines[index + i].lower():
                                # print(config_lines[index + i])
                                ip = config_lines[index + i].split()[2]                
            
            if hostname is None:
                print(f"No Hostname found in file {f}")
            else:
                if site[:7] not in hostname and not hostname.startswith('pep-'):
                    print(f"{Fore.RED + Style.BRIGHT}\nSite Name {site} does not match device name {hostname} found in config files. This is probably due to mismatch between show tech's site name found on summary page.")
                    print(f"Please Resolve, then restart!\n {Fore.RESET}")
                    exit()
                if hostname not in hosts.keys():
                    if ip is not None:
                        new_host_data = {}
                        new_hosts[hostname] = new_host_data
                        new_host_data['hostname'] = ip
                        new_host_data["groups"] = [summary_data['site'], 'cisco_ios']

                    else:
                        print(Fore.RED + f"Couldn't find IP for Host {hostname} in config file, add manually")
                    # path for configuration masters
            
            master_config = os.path.join(f"change_control/master_config_files/{summary_data['site']}/{hostname}.cfg")
            
            if not os.path.isdir(os.path.dirname(master_config)):
                os.makedirs(os.path.dirname(master_config))
                # copying current config to master config if not found.
            
            if not os.path.isfile(master_config):
                print(f"master config not found for {hostname} creating one from show run cmd_output")
                with open(master_config, "w") as f:
                    for line in config_lines:
                        f.write(line)
    print("\n")
    new_host_found = False

    for host in new_hosts:
        if host not in hosts.keys():
            new_host_found = True
            print(Fore.BLUE + f"Switch {host} wasn't in inventory, adding to site {site}" + Fore.RESET)
            hosts[host] = new_hosts[host]
    
    if new_host_found:
        write_hosts_data_to_inventory_file(hosts)
    
    return new_host_found

  
def update_interfaces_from_config_files(hosts):
    
    configs_dir = "./Show Tech/Configs"
    # configs_dir = directory_check(directory=configs_dir, func='CONFIGS')
    hosts_not_found = hosts.copy()
    no_config = []
    
    if not os.path.isdir(os.path.dirname(configs_dir)):
        print(f"{configs_dir}  is no good")
    
    for root, dirs, files in os.walk(configs_dir):
        for config_file in files:
            for host in hosts:
                host_data = hosts[host]
                
                if host.lower() not in config_file.lower():
                    continue
                
                # print(hosts_not_found.keys())
                del hosts_not_found[host]
                with open(f"{root}/{config_file}", "r") as f:
                    config_lines = f.readlines()
    
                host_interfaces = collections.OrderedDict()
                vlans = {}
                description = ""
                host_data["interfaces"] = host_interfaces
                host_data['vlans'] = vlans
                
                for index, line in enumerate(config_lines):
                    try:
                        if line.startswith('vlan'):
                            vlan = line.lstrip('vlan').strip()
                            vlan_name = config_lines[index + 1].lstrip(' name ').strip()
                            vlans[vlan] = {'name': vlan_name}
                        if line.startswith("interface"):
                            interface_data = {}
                            interface = normalize_interface_name(line.strip().lstrip("interface ").lower())
                            
                            host_interfaces[interface] = interface_data
                            
                            interface_config_lines = []
                            interface_data["config"] = interface_config_lines
                            
                            for i in range(1, 6):
                                if config_lines[index + i].startswith(("!", "end", "exit")):
                                    # print(f'breaking index {index + i}')
                                    break
                                elif config_lines[index + i].strip().startswith("description"):
                                    description = config_lines[index + i].strip().lstrip("description")
                                    interface_data["description"] = description
                                    continue
                                elif config_lines[index + i].strip().startswith("switchport access vlan "):
                                    vlan = config_lines[index + i].strip().lstrip("switchport access vlan ")
                                    interface_data['vlan'] = vlan
                                elif not config_lines[index + i].strip() == "":
                                    interface_config_lines.append(config_lines[index + i].strip())
                            
                            if len(interface_config_lines) == 0:
                                interface_data.pop("config", None)
                            
                            if len(host_interfaces[interface]) == 0:
                                del host_interfaces[interface]
                    
                    except IndexError:
                        pass
    
    if len(hosts_not_found) > 0:
        for host in hosts_not_found:
            print(Fore.RED + f"A Config file for {host} wasn't found" + Fore.RESET)


def file_check(filename):
    valid_file = False
    # output_filename = f"{summary_data['bt_location']}Q_{city[:8].capitalize()}_{state.upper()}_IP_Address_Assignments_{strftime('Y%Y-M%m-D%d_H%H-M%M')}.xlsx"
    while valid_file is not True:
        if filename is None:
            match_object = "Address_Assignments"
            print(match_object)
            for root, dirs, files in os.walk('./'):
                for f in files:
                    if match_object in f and f.endswith('xlsx'):
                        r1 = input(f"Use Parameters File '{f}'? [y/n/q][default y]: ").lower().strip()
                        
                        if r1 == 'q':
                            exit()
                        
                        elif r1 == 'n':
                            continue
                        
                        else:
                            filename = f

                            valid_file = True

                            break
                        
                        
        else:
        
            r1 = input(f"Use Parameters File '{filename}'? [y/n/q][default y]: ").lower()
            if r1 == 'q':
                exit()
            
            elif r1 == 'n':
                r2 = input("Enter the Parameters Filename: ")
                filename = r2
            
            if not os.path.isfile(filename):
                print(Fore.RED + f"{filename} not found" + Fore.RESET)
            
            else:
                print(Fore.BLUE + Style.BRIGHT + f"{filename} is valid, continuing...\n" + Fore.RESET)
                valid_file = True
    
    return filename


def directory_check(directory, func):
    valid_directory = False
    while valid_directory is not True:
        r1 = input(f"\nLook in directory '.\{directory}' for {func} files? [y/n/q][default y]: ").lower()
        if r1 == 'n':
            r2 = input("Enter the Directory: ")
            directory = r2
            if os.path.isdir(directory):
                valid_directory = True
        elif r1 == 'q':
            exit()
        
        if not os.path.isdir(directory):
            print(Fore.RED + f"{directory} not found" + Fore.RESET)
        else:
            print(Fore.BLUE + f"{directory} is valid, continuing...\n" + Fore.RESET)
            valid_directory = True
    return directory


def normalize_interface_name(interface_name: str):
    """ Normalizes interface name

    For example, GigabitEthernet1 is converted to Gi0/1,
    TenGigabitEthernet1/1  is converted to Te1/1
    """
    match = INTERFACE_NAME_RE.search(interface_name)
    if match:
        int_type = match.group("interface_type")
        # print(int_type)
        normalized_int_type = int_type[0:2:]
        int_num = match.group("interface_num")
        # print(int_num)
        return normalized_int_type + int_num
    raise ValueError(
        f"Does not recognize {interface_name} as an interface name")


def write_hosts_data_to_inventory_file(hosts):
    path = "./inventory/hosts.yaml"
    groups_path = "./inventory/groups.yaml"
    
    yaml = YAML()
    yaml.preserve_quotes = True
    with open(path) as f:
        parsed_yaml = yaml.load(f)
    
    with open(groups_path) as f:
        g_parsed_yaml = yaml.load(f)

    # for host in hosts:
    #     print(type(host))
    #     if "groups" in hosts[host]:
    #         print(hosts[host]['groups'])
    site = None
    for host in hosts:
        if host not in parsed_yaml:
            parsed_yaml[host] = hosts[host]
            site = hosts[host]['groups'][0]

    if site not in g_parsed_yaml:
        # print(g_parsed_yaml)
        site_data = {}
        site_data['data'] = {}
        site_data['data']['site'] = site
        g_parsed_yaml[site] = site_data

    with open(path, "w") as f:
        yaml.dump(parsed_yaml, f)
    
    with open(groups_path, "w") as f:
        yaml.dump(g_parsed_yaml, f)


def normalize_interface_type(interface_type: str) -> str:
    """Normalizes interface type

    For example, G is converted to GigabitEthernet, Te is converted to TenGigabitEthernet
    """
    int_type = interface_type.strip().lower()
    for norm_int_type in NORMALIZED_INTERFACES:
        if norm_int_type.lower().startswith(int_type):
            return norm_int_type
    return int_type



# def fancy_pandas():
#     df = pd.DataFrame([vars(s) for s in self.conversions if s.symbol == asset])

#     # Conversions sheet Create row 1 column names
#     for column in df.columns:
#         column_index = df.columns.get_loc(column)
#         sheet.cell(row=3, column=column_index + 1, value=column)
    
#     # Conversions sheet Create rows
#     row = 5
#     for index, session in df.iterrows():
#         for column in df.columns:
#             column_index = df.columns.get_loc(column)
#             sheet.cell(row=row, column=column_index + 1, value=str(session[column]))
#         row += 1