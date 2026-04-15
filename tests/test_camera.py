"""Quick standalone CoolSNAP EZ diagnostic — no Flask, no threads."""
import sys, os, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import modules.pvcam_raw as pvc

print("=== CoolSNAP EZ Diagnostic ===\n")

print("1. Initializing PVCAM...")
pvc.init()
print("   OK")

print("2. Camera count:", pvc.cam_count())
name = pvc.cam_name(0)
print("   Camera 0:", name)

print("3. Opening camera...")
hcam = pvc.cam_open(name)
print("   Handle:", hcam)

w, h = pvc.sensor_size(hcam)
bd = pvc.bit_depth(hcam)
print(f"   Sensor: {w}x{h}, {bd}-bit")

try:
    temp = pvc.sensor_temp_c(hcam)
    print(f"   Temp: {temp:.1f} C")
except Exception as e:
    print(f"   Temp: (failed: {e})")

print("\n4. Setting up continuous acquisition (50ms, bin=1)...")
fb = pvc.setup_cont(hcam, 50, 1)
print(f"   Frame buffer: {fb} bytes")

n = 2
buf = (pvc.uns16 * (fb * n // 2))()
print(f"   Buffer allocated: {fb * n} bytes ({fb * n // 2} elements)")

print("5. Starting continuous acquisition...")
pvc.start_cont(hcam, buf, fb * n)
print("   Started OK")

print("6. Polling for frames (5 seconds)...")
t0 = time.monotonic()
frames = 0
while time.monotonic() - t0 < 5.0:
    status, arrived, buf_cnt = pvc.check_cont_status(hcam)
    if status >= pvc.FRAME_AVAILABLE:
        ptr = pvc.get_latest_frame(hcam)
        frame = pvc.frame_to_numpy(ptr, w, h, 1)
        frames += 1
        if frames <= 3:
            print(f"   Frame {frames}: shape={frame.shape}, "
                  f"min={frame.min()}, max={frame.max()}, mean={frame.mean():.0f}")
    else:
        if frames == 0 and (time.monotonic() - t0) > 1.0:
            print(f"   Still waiting... status={status} arrived={arrived} buf_cnt={buf_cnt}")
        time.sleep(0.002)

elapsed = time.monotonic() - t0
print(f"\n   Got {frames} frames in {elapsed:.1f}s = {frames/elapsed:.1f} fps")

if frames == 0:
    print("\n   *** NO FRAMES RECEIVED — camera may need power cycle ***")
    print("   Try: unplug FireWire cable, wait 5 sec, plug back in")

print("\n7. Aborting acquisition...")
pvc.abort(hcam)
print("   OK")

print("8. Closing camera...")
pvc.cam_close(hcam)
pvc.uninit()
print("   Done")
