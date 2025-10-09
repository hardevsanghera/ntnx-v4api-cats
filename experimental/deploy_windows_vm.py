#!/usr/bin/env python3
"""
SYNOPSIS:
==== UNDER CONSTRUCTION DO NOT USE
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
import argparse
from datetime import datetime
from base64 import b64encode
import urllib3

# Disable SSL warnings for self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =============================================================================
# SAFETY CONFIGURATION - CHANGE TO True TO ENABLE EXECUTION WITHOUT --do-it
# =============================================================================
ALLOW_EXECUTION_WITHOUT_DOIT = False  # Set to True to bypass --do-it requirement

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
BIOS_TYPE = "LEGACY"              # BIOS type: UEFI or LEGACY
SECURE_BOOT = False              # Enable secure boot (Windows requirement)
TPM_ENABLED = False              # Enable TPM 2.0 (Windows 11 requirement)

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
    """Make API request with error handling and debug output"""
    
    # Debug output - show request details
    print(f"\n🔍 API Request Debug Info:")
    print(f"📡 Method: {method.upper()}")
    print(f"🌐 URL: {url}")
    
    print(f"📋 Headers:")
    for key, value in headers.items():
        # Mask sensitive authorization header for security
        if key.lower() == 'authorization':
            masked_value = f"{value[:10]}...{value[-4:]}" if len(value) > 14 else "***"
            print(f"  {key}: {masked_value}")
        else:
            print(f"  {key}: {value}")
    
    if data is not None:
        print(f"📦 Payload:")
        print(json.dumps(data, indent=2))
    else:
        print(f"📦 Payload: None")
    
    print("-" * 80)
    
    try:
        if method.upper() == 'GET':
            response = requests.get(url, headers=headers, verify=False, timeout=30)
        elif method.upper() == 'POST':
            response = requests.post(url, headers=headers, json=data, verify=False, timeout=30)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
        
        # Debug output - show response details
        print(f"📨 Response Status: {response.status_code}")
        print(f"📏 Response Size: {len(response.content)} bytes")
        
        response.raise_for_status()
        
        response_data = response.json() if response.content else {}
        
        # Show summary of response data
        if response_data:
            if isinstance(response_data, dict):
                print(f"📄 Response Keys: {list(response_data.keys())}")
                if 'data' in response_data and isinstance(response_data['data'], list):
                    print(f"📊 Data Items Count: {len(response_data['data'])}")
            else:
                print(f"📄 Response Type: {type(response_data)}")
        else:
            print(f"📄 Response: Empty")
        
        print("✅ Request completed successfully")
        print("=" * 80)
        
        return response_data
    
    except requests.exceptions.RequestException as e:
        print(f"❌ API request failed: {e}")
        print(f"🔍 Response Status Code: {getattr(e.response, 'status_code', 'N/A')}")
        
        if hasattr(e, 'response') and e.response is not None:
            try:
                error_details = e.response.json()
                print(f"💥 Error details: {json.dumps(error_details, indent=2)}")
            except:
                print(f"💥 Error response: {e.response.text}")
        
        print("=" * 80)
        sys.exit(1)

def find_image_by_name(base_url, headers, image_name):
    """Find a specific image by name"""
    print(f"\n🔍 Searching for image: {image_name}")
    
    url = f"{base_url}/vmm/v4.1/content/images"
    images_data = make_api_request('GET', url, headers)
    
    if 'data' not in images_data or not images_data['data']:
        print("No images found!")
        return None
    
    # Search for exact match first, then case-insensitive
    for image in images_data['data']:
        if image.get('name', '') == image_name:
            print(f"✅ Found exact match: {image['name']}")
            return {
                'name': image.get('name', 'Unknown'),
                'type': image.get('type', 'Unknown'),
                'size_gb': round(image.get('sizeInBytes', 0) / (1024**3), 2),
                'ext_id': image.get('extId', 'Unknown'),
                'data': image
            }
    
    # Try case-insensitive match
    for image in images_data['data']:
        if image.get('name', '').lower() == image_name.lower():
            print(f"✅ Found case-insensitive match: {image['name']}")
            return {
                'name': image.get('name', 'Unknown'),
                'type': image.get('type', 'Unknown'),
                'size_gb': round(image.get('sizeInBytes', 0) / (1024**3), 2),
                'ext_id': image.get('extId', 'Unknown'),
                'data': image
            }
    
    # Show available images if not found
    print(f"❌ Image '{image_name}' not found!")
    print("\nAvailable images:")
    for idx, image in enumerate(images_data['data']):
        name = image.get('name', 'Unknown')
        print(f"  {idx}: {name}")
    
    return None

def find_network_by_name(base_url, headers, subnet_name):
    """Find a specific network/subnet by name"""
    print(f"\n🔍 Searching for subnet: {subnet_name}")
    
    url = f"{base_url}/networking/v4.1/config/subnets"
    networks_data = make_api_request('GET', url, headers)
    
    if 'data' not in networks_data or not networks_data['data']:
        print("No networks found!")
        return None
    
    # Search for exact match first, then case-insensitive
    for network in networks_data['data']:
        if network.get('name', '') == subnet_name:
            print(f"✅ Found exact match: {network['name']}")
            return {
                'name': network.get('name', 'Unknown'),
                'type': network.get('subnetType', 'Unknown'),
                'vlan_id': network.get('vlanId', 'N/A'),
                'ext_id': network.get('extId', 'Unknown'),
                'data': network
            }
    
    # Try case-insensitive match
    for network in networks_data['data']:
        if network.get('name', '').lower() == subnet_name.lower():
            print(f"✅ Found case-insensitive match: {network['name']}")
            return {
                'name': network.get('name', 'Unknown'),
                'type': network.get('subnetType', 'Unknown'),
                'vlan_id': network.get('vlanId', 'N/A'),
                'ext_id': network.get('extId', 'Unknown'),
                'data': network
            }
    
    # Show available networks if not found
    print(f"❌ Subnet '{subnet_name}' not found!")
    print("\nAvailable subnets:")
    for idx, network in enumerate(networks_data['data']):
        name = network.get('name', 'Unknown')
        print(f"  {idx}: {name}")
    
    return None

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
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description='Deploy a Windows VM on Nutanix AHV using v4 APIs',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Safety Notes:
  This script will deploy a real VM on your Nutanix cluster.
  Use --do-it parameter to confirm you want to proceed with deployment.
  
  Alternative: Set ALLOW_EXECUTION_WITHOUT_DOIT = True in the script to bypass this check.

Examples:
  python deploy_windows_vm.py --do-it                                    # Interactive mode
  python deploy_windows_vm.py --do-it --image-name "Windows2019"        # Specify image only
  python deploy_windows_vm.py --do-it --subnet "Production-VLAN100"     # Specify network only
  python deploy_windows_vm.py --do-it --image-name "Windows2019" --subnet "Prod-Net"  # Fully automated
  python deploy_windows_vm.py --dry-run --image-name "Windows2019"      # Dry run with specific image
  python deploy_windows_vm.py --help                                     # Show this help

Parameters:
  --image-name: Exact name of the disk image to use (case-insensitive fallback)
  --subnet:     Exact name of the subnet/network to use (case-insensitive fallback)
        """
    )
    
    parser.add_argument(
        '--do-it', 
        action='store_true',
        help='Required parameter to confirm VM deployment (safety measure)'
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true', 
        help='Show what would be deployed without actually creating the VM'
    )
    
    parser.add_argument(
        '--image-name',
        type=str,
        help='Name of the disk image to use (exact match). If not provided, will show interactive list.'
    )
    
    parser.add_argument(
        '--subnet',
        type=str,
        help='Name of the subnet/network to use (exact match). If not provided, will show interactive list.'
    )
    
    args = parser.parse_args()
    
    # Safety check - prevent accidental execution
    if not args.do_it and not ALLOW_EXECUTION_WITHOUT_DOIT:
        print("🛑 SAFETY CHECK: VM Deployment Script")
        print("=" * 50)
        print("This script will create a real VM on your Nutanix cluster.")
        print("To proceed, use one of these options:")
        print("")
        print("1. Add --do-it parameter:")
        print("   python deploy_windows_vm.py --do-it")
        print("")
        print("2. Set ALLOW_EXECUTION_WITHOUT_DOIT = True in the script")
        print("")
        print("3. Use --dry-run to see what would be deployed:")
        print("   python deploy_windows_vm.py --dry-run")
        print("")
        print("Use --help for more information.")
        sys.exit(1)
    
    if args.dry_run:
        print("🔍 DRY RUN MODE - No VM will actually be created")
        print("=" * 60)
    
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
    
    # Handle image selection
    selected_image = None
    if args.image_name:
        # Use specified image name
        selected_image = find_image_by_name(base_url, headers, args.image_name)
        if not selected_image:
            print(f"Cannot proceed - specified image '{args.image_name}' not found.")
            sys.exit(1)
    else:
        # Use interactive selection
        images = list_images(base_url, headers)
        if not images:
            print("Cannot proceed without available images.")
            sys.exit(1)
        
        selected_image = select_from_list(images, "image")
        if not selected_image:
            sys.exit(1)
    
    # Handle network selection
    selected_network = None
    if args.subnet:
        # Use specified subnet name
        selected_network = find_network_by_name(base_url, headers, args.subnet)
        if not selected_network:
            print(f"Cannot proceed - specified subnet '{args.subnet}' not found.")
            sys.exit(1)
    else:
        # Use interactive selection
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
    
    # Handle dry-run mode
    if args.dry_run:
        print(f"\n🔍 DRY RUN: VM configuration shown above would be deployed")
        print("No actual VM creation will occur in dry-run mode.")
        
        # Create VM payload to show what would be sent
        vm_payload = create_vm_payload(VM_NAME, selected_image, selected_network)
        
        print(f"\n📄 API Payload that would be sent:")
        print(json.dumps(vm_payload, indent=2))
        
        print(f"\n✅ Dry run completed successfully!")
        print("To actually deploy the VM, run with --do-it (without --dry-run)")
        sys.exit(0)
    
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