"""Quick probe: test snap via start_live/poll_frame on CoolSNAP EZ."""

import time
from pyvcam import pvc
from pyvcam.camera import Camera

pvc.init_pvcam()
cam = next(Camera.detect_camera())
print(f"Camera: {cam.name}")
cam.open()
print(f"Sensor: {cam.sensor_size}, {cam.bit_depth}-bit")

print("\nStarting live mode (20ms exposure)...")
cam.start_live(exp_time=20)

print("Polling for frame...")
t0 = time.time()
for attempt in range(200):
    try:
        result = cam.poll_frame()
        frame = result[0]['pixel_data']
        fps = result[1]
        dt = time.time() - t0
        print(f"  Got frame in {dt:.3f}s (attempt {attempt+1})")
        print(f"  Shape: {frame.shape}, dtype: {frame.dtype}")
        print(f"  Min={frame.min()}, Max={frame.max()}, Mean={frame.mean():.1f}")
        print(f"  FPS: {fps}")
        break
    except RuntimeError:
        time.sleep(0.01)
else:
    print("  TIMEOUT: no frame after 200 attempts")

print("\nFinishing live mode...")
cam.finish()
cam.close()
pvc.uninit_pvcam()
print("Done.")
