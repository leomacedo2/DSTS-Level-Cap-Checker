import os
import time
import struct
import subprocess
import sys

# Forces Python to work in the same directory as this script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

# CONFIGURATION
SAVE_FILE_ORIGINAL = "0001.bin"  # Assumes the save file is in the same folder
SAVE_FILE_DEC = "save_dec.bin"
AES_KEY = "33393632373736373534353535383833"

DIGIMON_SIZE = 336

def decrypt_save():
    """Decrypts the save file using OpenSSL."""
    cmd = [
        "openssl", "enc", "-d", "-aes-128-ecb",
        "-K", AES_KEY,
        "-in", SAVE_FILE_ORIGINAL,
        "-out", SAVE_FILE_DEC,
        "-nopad"
    ]
    # Runs the command silently; raises an error if OpenSSL is not found
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)

def analyze_and_display():
    """Reads the decrypted save and displays the monitoring dashboard."""
    if not os.path.exists(SAVE_FILE_DEC):
        print(f" ERROR: The file '{SAVE_FILE_DEC}' was not generated. OpenSSL failed.")
        return

    with open(SAVE_FILE_DEC, "rb") as f:
        data = f.read()

    os.system('cls' if os.name == 'nt' else 'clear')

    print("=" * 65)
    print("      🎮 DIGIMON LEVEL CAP CHECKER - REAL-TIME MONITOR")
    print("=" * 65)
    print(f" Status: Monitoring game save... | Last update: {time.strftime('%H:%M:%S')}")
    print("-" * 65)

    total_alerts = 0
    processed_offsets = set()
    alerts = {"PARTY": [], "BOX": [], "FARM": []}

    # 1. SCANNING PARTY AND BOX
    regions = [
        ("PARTY",   0x12C8 - 0x10, 0x10, 6),     
        ("BOX",     0x1AA8 - 0x10, 0x10, 999),   
    ]

    for location, start_offset, header_size, max_slots in regions:
        for i in range(max_slots):
            offset = start_offset + (i * DIGIMON_SIZE)
            
            if offset + DIGIMON_SIZE > len(data) or offset in processed_offsets:
                continue

            if data[offset] != 1:
                continue

            name_offset = offset + header_size
            name_bytes = bytearray()
            for b in data[name_offset : name_offset + 32]:
                if b == 0: break
                name_bytes.append(b)

            try:
                name = name_bytes.decode('ascii')
                if len(name) < 2 or not (65 <= ord(name[0]) <= 90): continue
            except: continue

            level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
            talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]

            if talent_raw >= 1000:
                talent = talent_raw // 1000
                if level >= talent:
                    alerts[location].append((name, level, talent))
                    total_alerts += 1
                    processed_offsets.add(offset)

    # 2. SCANNING FARM
    FARM_START = 0x50000
    FARM_END = 0x60000
    HEADER_FARM = 0x18

    for offset in range(FARM_START, FARM_END, 4):
        if offset + DIGIMON_SIZE > len(data) or offset in processed_offsets:
            continue

        if data[offset] == 1 and data[offset+1] == 0 and data[offset+2] == 0 and data[offset+3] == 0:
            name_offset = offset + HEADER_FARM
            
            if 65 <= data[name_offset] <= 90:
                name_bytes = bytearray()
                for b in data[name_offset : name_offset + 32]:
                    if b == 0: break
                    name_bytes.append(b)

                try:
                    name = name_bytes.decode('ascii')
                    if len(name) < 2: continue
                except: continue

                level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]

                if talent_raw >= 1000:
                    talent = talent_raw // 1000
                    if level >= talent:
                        alerts["FARM"].append((name, level, talent))
                        total_alerts += 1
                        processed_offsets.add(offset)

    # 3. DISPLAY RESULTS (Ordered by Location -> Name -> Level)
    for loc in ["PARTY", "BOX", "FARM"]:
        alerts[loc].sort(key=lambda x: (x[0], x[1]))
        
        for name, level, talent in alerts[loc]:
            print(f" [{loc:^7}] {name:<18} (Lv. {level:02d} / Limit {talent:02d}) <-- REACHED CAP!")

    print("-" * 65)
    if total_alerts == 0:
        print(" ✨ All Digimon are evolving normally.")
    else:
        print(f" ⚠️  WARNING: {total_alerts} Digimon(s) reached their Level Cap!")
    print("=" * 65)
    print("\n[Waiting for a new game save... Keep this window open on your 2nd monitor]")

def main():
    if not os.path.exists(SAVE_FILE_ORIGINAL):
        print(f"CRITICAL ERROR: Save file '{SAVE_FILE_ORIGINAL}' not found in the current folder!")
        print("Please place this script inside your game's save folder.")
        return

    last_mtime = 0
    print("Starting automatic monitoring...")

    while True:
        try:
            current_mtime = os.path.getmtime(SAVE_FILE_ORIGINAL)
            if current_mtime != last_mtime:
                last_mtime = current_mtime
                decrypt_save()
                analyze_and_display()
            
            time.sleep(2)
        except KeyboardInterrupt:
            print("\nMonitoring stopped by user.")
            break
        except subprocess.CalledProcessError:
            print("\n[!] ERROR: OpenSSL failed. Is it installed and added to your system PATH?")
            time.sleep(5)
        except Exception as e:
            print(f"\n[!] OPS! An error occurred: {e}")
            time.sleep(5)

if __name__ == "__main__":
    main()