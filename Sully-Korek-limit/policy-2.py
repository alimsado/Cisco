import paramiko
from openpyxl import Workbook

# Function to SSH into router and execute command
def ssh_command(router_ip, command, username, password):
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        ssh_client.connect(router_ip, username=username, password=password, timeout=5)
        stdin, stdout, stderr = ssh_client.exec_command(command)
        output = stdout.read().decode()
        ssh_client.close()
        return output
    except paramiko.AuthenticationException:
        print(f"Authentication failed for {router_ip}")
    except paramiko.SSHException as ssh_exception:
        print(f"Unable to establish SSH connection to {router_ip}: {ssh_exception}")
    except Exception as e:
        print(f"Error occurred while connecting to {router_ip}: {e}")
    return None

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

# Main function to loop through IP ranges, SSH into routers, execute command, and record results
def main():
    results = []
    username = 'AliDoski'
    password = 'Al!$osk!22'

    ip_ranges = ["10.42.11.", "192.168.14."]
    for ip_range in ip_ranges:
        for i in range(1, 255):
            router_ip = ip_range + str(i)
            command = "show running | sec service-policy"
            command_output = ssh_command(router_ip, command, username, password)
            if command_output is not None:
                router_name_command = "show running-config | sec hostname"
                router_name_output = ssh_command(router_ip, router_name_command, username, password)
                router_name = extract_router_name(router_name_output)
                results.append((router_name, router_ip, command_output))
    
    write_to_excel(results)

if __name__ == "__main__":
    main()
