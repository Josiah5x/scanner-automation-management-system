import os
import argparse
from PIL import Image
import pythoncom
from pathlib import Path
from win32com.client import Dispatch

def list_scanners():
    """List all available WIA scanners"""
    manager = Dispatch("WIA.DeviceManager")
    devices = manager.DeviceInfos
    print("Available scanners:")
    for i in range(1, devices.Count + 1):
        device = devices.Item(i)
        # Check if the device is a scanner (Type = 1)
        if device.Type == 1:
            print(f"  Name: {device.Properties['Name'].Value}")
            print(f"  ID: {device.DeviceID}")
            print(f"  Description: {device.Properties['Description'].Value}")
            print("  ----------------")
list_scanners()
def scan_document(scanner_name=None):
    output_path= r"C:\Users\HP\Documents\Scanner_AG_Vision\output45.jpg"
    """
    Scan a document using WIA interface and save as image file
    
    Parameters:
        output_path: Path to save scanned image
        scanner_name: Name of specific scanner to use (None shows dialog)
    """
    # Create WIA objects
    wia = Dispatch("WIA.CommonDialog")

    selected_device = None
    if scanner_name:
        manager = Dispatch("WIA.DeviceManager")
        devices = manager.DeviceInfos
        for i in range(1, devices.Count + 1):
            device = devices.Item(i)
            if device.Type == 1 and device.Properties['Name'].Value == scanner_name:
                selected_device = device.Connect()
                break
        
        if not selected_device:
            print(f"Scanner '{scanner_name}' not found!")
            return None

    # Execute the scan
    print("Scanning...")
    
    img = None
    if selected_device is None:
        img = wia.ShowAcquireImage()  # Show scanning dialog
    else:
        img = wia.ShowTransfer(selected_device.Items[1])  # Transfer the scanned image
        
    if img is None:
        print("Scanning cancelled or failed")
        return None
    filename="output45.jpg"
    bmp_path = os.path.join(output_path, "temp_scan.bmp")
    jpg_path = os.path.join(output_path, filename)
    # Create output directory if it doesn't exist
    # output_dir = os.path.dirname(output_path)
    # output_dir = Path(output_path).resolve()
    # if output_dir:
    #     os.makedirs(output_dir, exist_ok=True)
    # Save the image
    if hasattr(img, "SaveFile"):  # WIA 2.0 compatible method
        # Scan (DON'T force JPEG for Canon)
        img.Transfer()

        # Save as BMP first
        img.SaveFile(bmp_path)
        # Convert to JPG
        Image.open(bmp_path).convert("RGB").save(jpg_path)

        # Remove temp BMP
        os.remove(bmp_path)

    else:  # Fallback for WIA 1.0 using PIL
        pil_img = Image.fromarray(img)
        pil_img.save(output_path)
    
    print(f"Scan completed. File saved to: {output_path}")
    return output_path

scan_document(scanner_name=None)

# def main():
#     parser = argparse.ArgumentParser(description="WIA Scanner Utility")
#     parser.add_argument("-L", "--list", action="store_true", 
#                         help="List available scanners")
#     parser.add_argument("-d", "--device", type=str,
#                         help="Name of scanner device to use")
#     parser.add_argument("-o", "--output", type=str, default="scan.png",
#                         help="Output file path (default: scan.png)")
    
#     args = parser.parse_args()
    
#     if args.list:
#         list_scanners()
#     else:
#         scan_document(output_path=args.output, scanner_name=args.device)

# if __name__ == "__main__":
#     main()