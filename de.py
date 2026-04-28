import os
from win32com.client import Dispatch
from PIL import Image
from pathlib import Path

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

def scan_document(out_path="scan.jpg", scanner_name=None):
    output_path =Path.cwd()/"scans"/out_path
    """
    Scan a document using WIA and save as image (Canon-safe)
    """
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
    print("Scanning...")
    if selected_device is None:
        img = wia.ShowAcquireImage()
    else:
        img = wia.ShowTransfer(selected_device.Items[1])
    if img is None:
        print("Scanning cancelled or failed")
        return None
    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    # 🔥 KEY FIX: always save as BMP first
    temp_bmp = os.path.splitext(output_path)[0] + "_temp.bmp"
    try:
        img.SaveFile(temp_bmp)  # Canon-safe
        # Convert to final format (jpg/png/etc.)
        Image.open(temp_bmp).convert("RGB").save(output_path)
        os.remove(temp_bmp)
        print(f"✅ Scan saved to: {output_path}")
        return output_path
    except Exception as e:
        print("❌ Error:", e)
        return None
scan_document(out_path="scan.jpg", scanner_name=None)