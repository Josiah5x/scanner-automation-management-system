import win32com.client
import random, os

names = ["Emma", "Liam", "Sophia", "Noah", "Olivia", "Smith", "Johnson", "Williams", "Brown", "Jones"]

def generate_random_name():
    # random.choice picks one random element from the provided list
    return f"{random.choice(names)}"

# print(f"{generate_random_name()}.jpg")
dirs = os.path.join(os.getcwd(), "scans")
file_path = os.path.join(dirs, "scan.jpg")
# print(folder)

def scan_from_printer():
    wia = win32com.client.Dispatch("WIA.CommonDialog")
    # ✅ Use a guaranteed valid Windows path
    # folder = r"C:\Users\Public\Documents\Scans"
    folder = dirs
    # ✅ Create folder if it doesn't exist
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, f"{generate_random_name()}.pdf")
    try:
        image = wia.ShowAcquireImage()
        if image:
            print("Saving to:", file_path)  # DEBUG
            image.SaveFile(file_path)
            print("Scan saved successfully")
        else:
            print("Scan cancelled")
    except Exception as e:
        print("Error:", e)


scan_from_printer()

