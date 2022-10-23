
import ipaddress
import os
import re
from colorama import Fore, Style, init
from nornir import InitNornir
from nornir.core.filter import F
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import inspect
import collections
from ruamel.yaml import YAML
from constants import NORMALIZED_INTERFACES, INTERFACE_NAME_RE
from utils import write_hosts_data_to_inventory_file, normalize_interface_name, file_check, update_interfaces_from_config_files, parse_switch_configs, directory_check, generate_portmap_data
from time import strftime
import logging.config

logging.basicConfig(
    filename="nornir_python.log",
    filemode="a",
    format="%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d_%H-%M-%S",
    level=logging.INFO,
)

# colorma issues
init()


if __name__ == "__main__":




    options = [
    "Populate Inventory from IP Scan",
    "Populate Inventory from config files",
    "Compare Inventory to Config Files"
    ]


    response = ''

    first_run = True
    while response != 'q':
       
        # Main Menu Options
        num = 1
        for option in options:
            print(f"{Fore.BLUE}({Fore.RESET}{num}{Fore.BLUE}){Fore.RESET}{Fore.GREEN + Style.BRIGHT} {option} {Fore.RESET}")
            num += 1
        
        print(f"{Fore.BLUE}({Fore.RESET}Q{Fore.BLUE}){Fore.RESET}{Fore.GREEN + Style.BRIGHT} Quit {Fore.RESET}")

        response = input("\n Please choose a option, [1-9][Q] to quit: ").lower()
        
        if first_run is True and response != 'q':
            first_run = False 
            
            summary_data = {}

            # Load Parameters File
            template_filename = "Network_Information_Template.xlsx"
            
            filename = None
            match_object = "_Network_Information_Y"
            filename = file_check(filename, match_object)

            # Create Network Information workbook from template if filename is template
            if filename == "Network_Information_Template.xlsx":

                site = input("Please enter a descriptive site name. ex. Tampa_FFA: ")
                
                wb = load_workbook(template_filename)

                # Save output file
                common_sheet = wb['Common']
                common_sheet['B6'] = site
                filename = f"{site}_Network_Information_{strftime('Y%Y-M%m-D%d_H%H-M%M')}.xlsx"
                wb.save(filename)

            # Get data from workbook
            wb = load_workbook(filename)
            common_sheet = wb['Common']
            site = common_sheet['B6'].value
            summary_data['site'] = site
            summary_data['filename'] = filename
            
            # load networks from workbook
            networks = []
            summary_data['networks'] = networks
            networks_sheet = wb['Networks']

            for row in networks_sheet.iter_rows(min_row=3, max_col=6, values_only=True):
                if row[2]: 
                    network = {}
                    network['name'] = row[0]
                    network_obj = ipaddress.ip_network(f"{row[1]}/{row[2]}", strict=False)
                    network['network_obj'] = network_obj
                    network['vlan'] = row[3]
                    network['gateway'] = row[4]
                    networks.append(network)
            
            # Parse Switch Configs to get networks, interfaces, ect 
            hosts = parse_switch_configs(summary_data)
            
            # Reconcile switches found in configs vs Network Information
            switches_sheet = wb['Switches']
            row_num = 3
            for host in hosts:
                host_found = False
                for row in switches_sheet.iter_rows(min_row=3, max_col=4):
                    row_num = row[0].row
                    if host.lower() in row[0].value:
                        print(f"{host} found in configs and network information workbook")
                        host_found = True
                
                if host_found is False:
                    row_num += 1
                    print(f"Adding {host} to Network Information Workbook")
                    switches_sheet.cell(row=row_num, column=1, value=host)
                    switches_sheet.cell(row=row_num, column=2, value=hosts[host]['hostname'])
            
            # Reconcile networks found in configs vs Network Information 
            networks_sheet = wb['Networks']
            row_num = 3
            for network in networks:
                network_found = False
                for row in networks_sheet.iter_rows(min_row=3, max_col=6):
                    row_num = row[0].row
                    if network['vlan'] == row[3].value:
                        network['name'] = row[0].value
                        network_found = True
                
                if network_found is False:
                    row_num += 1
                    print(f"Adding {network['vlan']} to Network Information Workbook")
                    network_obj = network['network_obj']
                    networks_sheet.cell(row=row_num, column=1, value=network['name'])
                    networks_sheet.cell(row=row_num, column=2, value=str(network_obj.network_address))
                    networks_sheet.cell(row=row_num, column=3, value=str(network_obj.prefixlen))
                    networks_sheet.cell(row=row_num, column=4, value=network['vlan'])
                    

            for network in networks:
                network['address'] = str(network['network_obj'].network_address)
                network['netmask'] = str(network['network_obj'].netmask)
                
                if 'gateway' not in network:
                    network['gateway'] = str(network['network_obj'][1])
                
                sheet_found = False
                
                for sheetname in wb.sheetnames:
                    if str(network['vlan']) in sheetname:
                        # print(f"{vlan['id']} is in {sheetname}")
                        sheet_found = True
                        sheet = wb[sheetname]

                        sheet['B1'] = f"{network['name']} / VLAN {network['vlan']} / {str(network_obj.network_address)}"
                        sheet['B2'] = network['gateway']
                        sheet['B3'] = network['netmask']
                        
                        row = 4

                        print(Fore.BLUE + Style.BRIGHT + f"\nCreating IP Address Sheet for VLAN {network['vlan']}{Fore.RESET}")
                        print(Fore.BLUE + Style.BRIGHT + f" Name {network['name']} Network {network['network_obj'].with_prefixlen}" + Fore.RESET)
                         
                        sheet.title = f"VLAN {network['vlan']} {network['name']}"[:30]
                        # print(f"{ipaddress.ip_network(str(network['address'])).num_addresses} usable addresses" + Fore.RESET)
                        # print(str(network['address']))
                        for host in network['network_obj'].hosts():
                            row += 1
                            sheet[f'A{row}'] = str(host)   
                
                if sheet_found is False:
                    print(Fore.RED + f"\nVLAN {network['vlan']} - {network['name']} {str(network['address'])} was not found in any sheet, creating" + Fore.RESET)
                    template_sheet = wb['VLAN Template']
                    sheet = wb.copy_worksheet(template_sheet)
                    sheet['B1'] = f"{network['name']} / VLAN {network['vlan']} / {str(network['address'])}"
                    sheet['B2'] = network['gateway']
                    sheet['B3'] = network['netmask']
                    row = 4

                    print(Fore.BLUE + f"Writing ip's for VLAN {network['vlan']} - {network['name']} {str(network['address'])}" + Fore.RESET)
                    sheet.title = f"VLAN {network['vlan']} {network['name']}"[:30]
                    # print(f"{ipaddress.ip_network(str(network['address'])).num_addresses} usable addresses" + Fore.RESET)
                    # print(str(network['address']))
                    for host in network['network_obj'].hosts():
                        row += 1
                        sheet[f'A{row}'] = str(host)
            
            wb.save(filename)
        
        if response == 'q':
            exit()
        
        elif response == '1':
            update_interfaces_from_config_files(hosts)
            generate_portmap_data(hosts, summary_data)
        
        else:
            print(f"\n{Fore.RED + Style.BRIGHT}The Input [{response}] is Invalid!{Fore.RESET}") 
