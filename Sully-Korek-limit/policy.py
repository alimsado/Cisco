import paramiko
from openpyxl import Workbook

# Function to SSH into router and execute command
def ssh_command(router_ip, command, username, password):
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh_client.connect(router_ip, username=username, password=password)

    stdin, stdout, stderr = ssh_client.exec_command(command)
    output = stdout.read().decode()
    
    ssh_client.close()
    return output

# Function to extract router name from command output
def extract_router_name(command_output):
    lines = command_output.split('\n')
    for line in lines:
        if 'hostname' in line:
            return line.split()[1]
    return None

# Function to write results to Excel sheet
def write_to_excel(results):
    wb = Workbook()
    ws = wb.active
    ws.append(["Router Name", "IP Address", "Command Output"])
    for result in results:
        ws.append(result)
    wb.save("router_results.xlsx")

# Main function to SSH into the specified IP address, extract router name, run command, and record results
def main():
    results = []
    username = 'AliDoski'
    password = 'Al!$osk!22'
    router_ip = '192.168.14.15'

    # Get router name
    router_name_command = "show running-config | sec hostname"
    router_name_output = ssh_command(router_ip, router_name_command, username, password)
    router_name = extract_router_name(router_name_output)

    # Run command and record output
    command = "show running | sec service-policy"
    command_output = ssh_command(router_ip, command, username, password)

    results.append((router_name, router_ip, command_output))
    
    write_to_excel(results)

if __name__ == "__main__":
    main()
