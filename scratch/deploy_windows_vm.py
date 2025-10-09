#!/usr/bin/env python3
"""
SYNOPSIS:
    Script Name: deploy_windows_vm.py
    Author: hardev@nutanix.com + Co-Pilot
    Date: October 2025
    Version: 1.0
    Purpose:
    A script to deploy an AHV Windows VM using Nutanix API v4 REST calls.
    Lists available images and networks for user selection.
    Reads configuration from vars.txt file.

NB:
    This script is provided "AS IS" without warranty of any kind.
    Use of this script is at your own risk.
    The author(s) make no representations or warranties, express or implied,
    regarding the script's functionality, fitness for a particular purpose,
    or reliability.

    By using this script, you agree that you are solely responsible
    for any outcomes, including loss of data, system issues, or
    other damages that may result from its execution.
    No support or maintenance is provided.

NOTES:
    You may copy, edit, customize and use as needed.
    Test thoroughly in a safe environment before deploying to production systems.
"""

import os
import sys
import json
import uuid
import requests
from datetime import datetime
from base64 import b64encode
import urllib3

# Disable SSL warnings for self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =============================================================================
# VM CONFIGURATION VARIABLES - MODIFY THESE AS NEEDED
# =============================================================================

# VM Basic Configuration
VM_NAME = "Windows-VM-" + datetime.now().strftime("%Y%m%d-%H%M%S")
VM_DESCRIPTION = "Windows VM deployed via Python script using Nutanix v4 API"

# CPU and Memory Configuration
NUM_VCPUS = 4                    # Number of virtual CPUs
NUM_CORES_PER_VCPU = 1          # Cores per vCPU (usually 1)
MEMORY_SIZE_MIB = 8192          # Memory in MiB (8192 = 8GB)

# Disk Configuration
BOOT_DISK_SIZE_GIB = 80         # Boot disk size in GiB
BOOT_DISK_BUS = "SCSI"          # Bus type: SCSI, IDE, or SATA

# VM Hardware Configuration
MACHINE_TYPE = "PC"             # Machine type: PC or Q35
BIOS_TYPE = "UEFI"              # BIOS type: UEFI or LEGACY
SECURE_BOOT = True              # Enable secure boot (Windows requirement)
TPM_ENABLED = True              # Enable TPM 2.0 (Windows 11 requirement)

# Boot Configuration
BOOT_ORDER = ["DISK", "CDROM", "NETWORK"]  # Boot device priority

# =============================================================================
# SCRIPT FUNCTIONS
# =============================================================================

def read_vars_file():
    """Read configuration variables from vars.txt file"""
    vars_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'files', 'vars.txt')
    
    if not os.path.exists(vars_file):
        print(f"Error: vars.txt file not found at {vars_file}")
        print("Expected format:")
        print("baseUrl=https://xx.xx.xx.xx:9440/api")
        print("username=xxxxxxx")
        print("password=xxxxxxx")
        sys.exit(1)
    
    config = {}
    with open(vars_file, 'r') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                key, value = line.split('=', 1)
                config[key.strip()] = value.strip()
    
    required_keys = ['baseUrl', 'username', 'password']
    for key in required_keys:
        if key not in config:
            print(f"Error: Missing required configuration: {key}")
            sys.exit(1)
    
    return config

def create_auth_header(username, password):
    """Create basic authentication header"""
    credentials = f"{username}:{password}"
    encoded_credentials = b64encode(credentials.encode()).decode()
    return {'Authorization': f'Basic {encoded_credentials}'}

def make_api_request(method, url, headers, data=None):
    """Make API request with error handling"""
    try:
        if method.upper() == 'GET':
            response = requests.get(url, headers=headers, verify=False, timeout=30)
        elif method.upper() == 'POST':
            response = requests.post(url, headers=headers, json=data, verify=False, timeout=30)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
        
        response.raise_for_status()
        return response.json() if response.content else {}
    
    except requests.exceptions.RequestException as e:
        print(f"API request failed: {e}")
        if hasattr(e, 'response') and e.response is not None:
            try:
                error_details = e.response.json()
                print(f"Error details: {json.dumps(error_details, indent=2)}")
            except:
                print(f"Error response: {e.response.text}")
        sys.exit(1)

def list_images(base_url, headers):
    """List available disk images"""
    print("\n🖼️  Fetching available images...")
    
    url = f"{base_url}/vmm/v4.1/content/images"
    images_data = make_api_request('GET', url, headers)
    
    if 'data' not in images_data or not images_data['data']:
        print("No images found!")
        return []
    
    images = []
    print("\nAvailable Images:")
    print("-" * 80)
    print(f"{'Index':<6} {'Name':<30} {'Type':<12} {'Size (GB)':<12} {'ExtID'}")
    print("-" * 80)
    
    for idx, image in enumerate(images_data['data']):
        name = image.get('name', 'Unknown')
        image_type = image.get('type', 'Unknown')
        size_bytes = image.get('sizeInBytes', 0)
        size_gb = round(size_bytes / (1024**3), 2) if size_bytes else 0
        ext_id = image.get('extId', 'Unknown')
        
        images.append({
            'name': name,
            'type': image_type,
            'size_gb': size_gb,
            'ext_id': ext_id,
            'data': image
        })
        
        print(f"{idx:<6} {name:<30} {image_type:<12} {size_gb:<12} {ext_id}")
    
    return images

def list_networks(base_url, headers):
    """List available networks"""
    print("\n🌐 Fetching available networks...")
    
    url = f"{base_url}/networking/v4.1/config/subnets"
    networks_data = make_api_request('GET', url, headers)
    
    if 'data' not in networks_data or not networks_data['data']:
        print("No networks found!")
        return []
    
    networks = []
    print("\nAvailable Networks:")
    print("-" * 80)
    print(f"{'Index':<6} {'Name':<30} {'Type':<12} {'VLAN':<8} {'ExtID'}")
    print("-" * 80)
    
    for idx, network in enumerate(networks_data['data']):
        name = network.get('name', 'Unknown')
        subnet_type = network.get('subnetType', 'Unknown')
        vlan_id = network.get('vlanId', 'N/A')
        ext_id = network.get('extId', 'Unknown')
        
        networks.append({
            'name': name,
            'type': subnet_type,
            'vlan_id': vlan_id,
            'ext_id': ext_id,
            'data': network
        })
        
        print(f"{idx:<6} {name:<30} {subnet_type:<12} {str(vlan_id):<8} {ext_id}")
    
    return networks

def select_from_list(items, item_type):
    """Allow user to select from a list of items"""
    if not items:
        print(f"No {item_type} available!")
        return None
    
    while True:
        try:
            choice = input(f"\nSelect {item_type} by index (0-{len(items)-1}): ").strip()
            if choice.lower() in ['q', 'quit', 'exit']:
                print("Exiting...")
                sys.exit(0)
            
            index = int(choice)
            if 0 <= index < len(items):
                selected = items[index]
                print(f"Selected {item_type}: {selected['name']} (ExtID: {selected['ext_id']})")
                return selected
            else:
                print(f"Invalid index. Please enter a number between 0 and {len(items)-1}")
        except ValueError:
            print("Invalid input. Please enter a number or 'q' to quit.")

def create_vm_payload(vm_name, image, network):
    """Create the VM creation payload"""
    
    # Generate a unique UUID for the VM
    vm_uuid = str(uuid.uuid4())
    
    # Create boot device configuration
    boot_devices = []
    for boot_type in BOOT_ORDER:
        if boot_type == "DISK":
            boot_devices.append({"bootDeviceType": "DISK", "bootDeviceOrder": len(boot_devices)})
        elif boot_type == "CDROM":
            boot_devices.append({"bootDeviceType": "CDROM", "bootDeviceOrder": len(boot_devices)})
        elif boot_type == "NETWORK":
            boot_devices.append({"bootDeviceType": "NETWORK", "bootDeviceOrder": len(boot_devices)})
    
    payload = {
        "name": vm_name,
        "description": VM_DESCRIPTION,
        "numSockets": NUM_VCPUS,
        "numCoresPerSocket": NUM_CORES_PER_VCPU,
        "memorySizeBytes": MEMORY_SIZE_MIB * 1024 * 1024,  # Convert MiB to bytes
        "machineType": MACHINE_TYPE,
        "biosType": BIOS_TYPE,
        "isSecureBootEnabled": SECURE_BOOT,
        "isTpmEnabled": TPM_ENABLED,
        "bootConfig": {
            "bootDevices": boot_devices
        },
        "disks": [
            {
                "diskAddress": {
                    "busType": BOOT_DISK_BUS,
                    "index": 0
                },
                "diskSizeBytes": BOOT_DISK_SIZE_GIB * 1024 * 1024 * 1024,  # Convert GiB to bytes
                "storageContainer": {
                    "extId": image['ext_id']
                },
                "dataSourceReference": {
                    "extId": image['ext_id']
                }
            }
        ],
        "nics": [
            {
                "networkInfo": {
                    "subnet": {
                        "extId": network['ext_id']
                    }
                },
                "nicType": "NORMAL_NIC"
            }
        ],
        "guestOs": {
            "osType": "WINDOWS"
        },
        "apcConfig": {
            "isEnabled": False
        }
    }
    
    return payload

def deploy_vm(base_url, headers, payload):
    """Deploy the VM using the Nutanix v4 API"""
    print(f"\n🚀 Deploying VM: {payload['name']}")
    print("VM Configuration:")
    print(f"  - CPUs: {payload['numSockets']} x {payload['numCoresPerSocket']} cores")
    print(f"  - Memory: {payload['memorySizeBytes'] // (1024*1024)} MiB")
    print(f"  - Boot Disk: {payload['disks'][0]['diskSizeBytes'] // (1024*1024*1024)} GiB")
    print(f"  - Machine Type: {payload['machineType']}")
    print(f"  - BIOS: {payload['biosType']}")
    print(f"  - Secure Boot: {payload['isSecureBootEnabled']}")
    print(f"  - TPM: {payload['isTpmEnabled']}")
    
    # Create the VM
    url = f"{base_url}/vmm/v4.1/ahv/config/vms"
    
    print(f"\nSending POST request to: {url}")
    print(f"Payload size: {len(json.dumps(payload))} characters")
    
    try:
        response = make_api_request('POST', url, headers, payload)
        return response
    except Exception as e:
        print(f"Failed to create VM: {e}")
        return None

def main():
    """Main function"""
    print("=" * 80)
    print("🖥️  Nutanix AHV Windows VM Deployment Script")
    print("=" * 80)
    
    # Read configuration
    config = read_vars_file()
    base_url = config['baseUrl']
    
    # Create authentication header
    headers = create_auth_header(config['username'], config['password'])
    headers['Content-Type'] = 'application/json'
    headers['Accept'] = 'application/json'
    
    print(f"\n🔗 Connecting to Nutanix Prism Central: {base_url}")
    
    # Test connection
    try:
        test_url = f"{base_url}/vmm/v4.1/ahv/config/vms"
        make_api_request('GET', test_url, headers)
        print("✅ Connection successful!")
    except:
        print("❌ Connection failed! Please check your credentials and URL.")
        sys.exit(1)
    
    # List and select image
    images = list_images(base_url, headers)
    if not images:
        print("Cannot proceed without available images.")
        sys.exit(1)
    
    selected_image = select_from_list(images, "image")
    if not selected_image:
        sys.exit(1)
    
    # List and select network
    networks = list_networks(base_url, headers)
    if not networks:
        print("Cannot proceed without available networks.")
        sys.exit(1)
    
    selected_network = select_from_list(networks, "network")
    if not selected_network:
        sys.exit(1)
    
    # Confirm deployment
    print(f"\n📋 Deployment Summary:")
    print(f"VM Name: {VM_NAME}")
    print(f"Image: {selected_image['name']} ({selected_image['size_gb']} GB)")
    print(f"Network: {selected_network['name']}")
    print(f"CPUs: {NUM_VCPUS} x {NUM_CORES_PER_VCPU}")
    print(f"Memory: {MEMORY_SIZE_MIB} MiB")
    print(f"Boot Disk: {BOOT_DISK_SIZE_GIB} GiB")
    
    confirm = input(f"\nProceed with VM deployment? (y/N): ").strip().lower()
    if confirm not in ['y', 'yes']:
        print("Deployment cancelled.")
        sys.exit(0)
    
    # Create VM payload
    vm_payload = create_vm_payload(VM_NAME, selected_image, selected_network)
    
    # Deploy the VM
    result = deploy_vm(base_url, headers, vm_payload)
    
    if result:
        print("\n✅ VM deployment initiated successfully!")
        if 'data' in result:
            vm_data = result['data']
            print(f"VM ExtID: {vm_data.get('extId', 'Unknown')}")
            print(f"VM State: {vm_data.get('state', 'Unknown')}")
        
        print("\n📝 Next Steps:")
        print("1. Monitor VM creation progress in Prism Central")
        print("2. Power on the VM once creation is complete")
        print("3. Configure Windows installation (if using installation media)")
        print("4. Install Nutanix Guest Tools after OS installation")
        
        # Save deployment details
        deployment_info = {
            'timestamp': datetime.now().isoformat(),
            'vm_name': VM_NAME,
            'image_used': selected_image['name'],
            'network_used': selected_network['name'],
            'configuration': {
                'cpus': NUM_VCPUS,
                'memory_mib': MEMORY_SIZE_MIB,
                'disk_gib': BOOT_DISK_SIZE_GIB
            },
            'api_response': result
        }
        
        deployment_file = f"deployment_{VM_NAME}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(deployment_file, 'w') as f:
                json.dump(deployment_info, f, indent=2)
            print(f"📄 Deployment details saved to: {deployment_file}")
        except Exception as e:
            print(f"Warning: Could not save deployment details: {e}")
    else:
        print("❌ VM deployment failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()