"""
Interactive test for CoolSNAP EZ camera via PyVCAM.
Run with the camera connected to verify backend functions.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import types
_fake_eel = types.ModuleType("eel")
_fake_eel.expose = lambda fn: fn
_fake_eel.onLiveFrame = lambda *a: (lambda: None)
_fake_eel.onCamStatus = lambda *a: (lambda: None)
_fake_eel.onCamCaptureComplete = lambda *a: (lambda: None)
_fake_eel.onTimelapseProgress = lambda *a: (lambda: None)
sys.modules["eel"] = _fake_eel

from modules import coolsnap


def menu_connect():
    print("\n--- Connect ---")
    try:
        coolsnap.connect()
        info = coolsnap.get_camera_info()
        print(f"  Connected: {info['name']}")
        print(f"  Sensor: {info['sensor_size']}, {info['bit_depth']}-bit")
    except Exception as e:
        print(f"  ERROR: {e}")


def menu_disconnect():
    print("\n--- Disconnect ---")
    coolsnap.disconnect()
    print("  Disconnected.")


def menu_snap():
    print("\n--- Snap ---")
    try:
        frame = coolsnap.snap()
        print(f"  Shape: {frame.shape}, dtype: {frame.dtype}")
        print(f"  Min: {frame.min()}, Max: {frame.max()}, Mean: {frame.mean():.1f}")
    except Exception as e:
        print(f"  ERROR: {e}")


def menu_snap_save():
    print("\n--- Snap & Save ---")
    try:
        frame, path = coolsnap.snap_and_save()
        print(f"  Shape: {frame.shape}")
        print(f"  Saved: {path}")
    except Exception as e:
        print(f"  ERROR: {e}")


def menu_video():
    print("\n--- Record Video ---")
    try:
        n = int(input("  Number of frames [50]: ").strip() or "50")
        path = coolsnap.record_video_and_save(num_frames=n)
        print(f"  Saved: {path}")
    except Exception as e:
        print(f"  ERROR: {e}")


def menu_timelapse():
    print("\n--- Time-lapse ---")
    try:
        n = int(input("  Number of frames [5]: ").strip() or "5")
        iv = float(input("  Interval (sec) [3]: ").strip() or "3")
        path = coolsnap.timelapse_and_save(num_frames=n, interval_sec=iv)
        print(f"  Saved: {path}")
    except Exception as e:
        print(f"  ERROR: {e}")


def menu_exposure():
    print("\n--- Exposure ---")
    print(f"  Current: {coolsnap.get_exposure()} ms")
    v = input("  New value (ms) [enter to skip]: ").strip()
    if v:
        coolsnap.set_exposure(int(v))
        print(f"  Set to: {coolsnap.get_exposure()} ms")


def menu_binning():
    print("\n--- Binning ---")
    print(f"  Current: {coolsnap.get_binning()}x{coolsnap.get_binning()}")
    v = input("  New value (1/2/4/8) [enter to skip]: ").strip()
    if v:
        try:
            coolsnap.set_binning(int(v))
            print(f"  Set to: {coolsnap.get_binning()}x{coolsnap.get_binning()}")
        except Exception as e:
            print(f"  ERROR: {e}")


OPTIONS = [
    ("Connect", menu_connect),
    ("Disconnect", menu_disconnect),
    ("Snap (preview)", menu_snap),
    ("Snap & Save", menu_snap_save),
    ("Record Video", menu_video),
    ("Time-lapse", menu_timelapse),
    ("Exposure", menu_exposure),
    ("Binning", menu_binning),
]


def main():
    print("═══════════════════════════════════════════")
    print("  CoolSNAP EZ — Interactive Test")
    print("═══════════════════════════════════════════")
    while True:
        print()
        for i, (label, _) in enumerate(OPTIONS, 1):
            print(f"  {i}. {label}")
        print("  0. Quit")
        choice = input("\n> ").strip()
        if choice == "0":
            coolsnap.disconnect()
            break
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(OPTIONS):
                OPTIONS[idx][1]()
            else:
                print("  Invalid choice.")
        except ValueError:
            print("  Invalid choice.")


if __name__ == "__main__":
    main()
