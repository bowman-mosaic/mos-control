"""
Quick test script for the ctypes PVCAM wrapper.
Run with: python test_pvcam_raw.py
"""
from modules import pvcam_raw as pvc
import sys

def main():
    print("=== PVCAM ctypes wrapper test ===\n")

    print("[1] Loading pvcam64.dll ...")
    try:
        pvc._load_dll()
        print("    OK — DLL loaded\n")
    except RuntimeError as e:
        print(f"    FAILED: {e}")
        sys.exit(1)

    print("[2] Initializing PVCAM ...")
    pvc.init()
    print("    OK\n")

    print("[3] Enumerating cameras ...")
    n = pvc.cam_count()
    print(f"    Found {n} camera(s)")
    if n == 0:
        print("    No cameras — skipping remaining tests.")
        pvc.uninit()
        return

    name = pvc.cam_name(0)
    print(f"    Camera 0 name: '{name}'\n")

    print("[4] Opening camera ...")
    hcam = pvc.cam_open(name)
    print(f"    OK — handle = {hcam}\n")

    print("[5] Reading sensor info ...")
    w, h = pvc.sensor_size(hcam)
    bd = pvc.bit_depth(hcam)
    try:
        chip = pvc.chip_name(hcam)
    except Exception:
        chip = "(not available)"
    try:
        temp = pvc.sensor_temp_c(hcam)
    except Exception:
        temp = "(not available)"
    print(f"    Sensor:    {w} x {h}")
    print(f"    Bit depth: {bd}")
    print(f"    Chip name: {chip}")
    print(f"    CCD temp:  {temp}\n")

    print("[6] Test snap (continuous mode, 1 frame) ...")
    frame_bytes = pvc.setup_cont(hcam, exposure_ms=50, binning=1)
    print(f"    Frame size: {frame_bytes} bytes")

    buf = (pvc.uns16 * (frame_bytes * 2 // 2))()
    pvc.start_cont(hcam, buf, frame_bytes * 2)

    try:
        frame = pvc.poll_frame_numpy(hcam, w, h, binning=1, timeout_s=10)
        print(f"    Got frame: shape={frame.shape}, dtype={frame.dtype}")
        print(f"    Min={frame.min()}, Max={frame.max()}, Mean={frame.mean():.1f}\n")
    except TimeoutError as e:
        print(f"    TIMEOUT: {e}\n")
    finally:
        pvc.abort(hcam)

    print("[7] Closing camera ...")
    pvc.cam_close(hcam)
    print("    OK\n")

    print("[8] Uninitializing PVCAM ...")
    pvc.uninit()
    print("    OK\n")

    print("=== All tests passed! ===")


if __name__ == "__main__":
    main()
