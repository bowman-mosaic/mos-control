"""
Interactive Nikon Ti terminal controller.
Talks directly to the COM interface — no Flask, no web UI.
"""
import ctypes, sys, time

print("Initializing COM (MTA)...")
ctypes.windll.ole32.CoInitializeEx(None, 2)  # COINIT_MULTITHREADED

import comtypes.client

print("Creating Nikon Ti scope object...")
try:
    scope = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")
except Exception as e:
    print(f"FAILED to create scope: {e}")
    sys.exit(1)

print(f"Scope object: {scope}")
print()

# ── helpers ──
def dev(name):
    try:
        d = getattr(scope, name)
        print(f"  [{name}] object: {d}")
        return d
    except Exception as e:
        print(f"  [{name}] NOT AVAILABLE: {e}")
        return None

def read_val(obj):
    if obj is None:
        return None
    try:
        return obj.RawValue
    except Exception:
        try:
            return obj.Value
        except Exception:
            return "???"

# ── probe all devices ──
print("=== Probing devices ===")
devices = [
    "DiaShutter", "EpiShutter", "DiaLamp",
    "Nosepiece", "FilterBlockCassette1", "LightPathDrive",
    "ZDrive", "XDrive", "YDrive", "PFS",
]
found = {}
for name in devices:
    d = dev(name)
    if d is not None:
        found[name] = d
        try:
            mounted = d.IsMounted
            print(f"    IsMounted = {mounted}")
        except Exception:
            pass
        try:
            pos = read_val(d.Position)
            print(f"    Position = {pos}")
        except Exception:
            pass
        try:
            val = read_val(d.Value)
            print(f"    Value = {val}")
        except Exception:
            pass

print()
print("=== Interactive control ===")
print("Commands:")
print("  shutter open     / shutter close")
print("  lamp on          / lamp off")
print("  lamp <0-24>        set halogen intensity")
print("  z                  read Z position")
print("  z up <nm>        / z down <nm>")
print("  xy                 read XY position")
print("  nose <1-6>         set objective")
print("  filter <1-6>       set filter")
print("  status             re-read all states")
print("  quit")
print()

def show_status():
    if "DiaShutter" in found:
        s = found["DiaShutter"]
        try:
            print(f"  Shutter: Value={read_val(s.Value)}, Position={read_val(s.Position)}")
        except Exception as e:
            print(f"  Shutter: error reading - {e}")
    if "DiaLamp" in found:
        l = found["DiaLamp"]
        try:
            on = l.Enabled
            intensity = read_val(l.Value)
            print(f"  Lamp: Enabled={on}, Intensity={intensity}")
        except Exception as e:
            print(f"  Lamp: error reading - {e}")
    if "ZDrive" in found:
        z = found["ZDrive"]
        try:
            pos = read_val(z.Position)
            print(f"  Z: {pos} nm ({pos/1000:.1f} µm)" if pos else f"  Z: {pos}")
        except Exception as e:
            print(f"  Z: error - {e}")
    if "XDrive" in found and "YDrive" in found:
        try:
            x = read_val(found["XDrive"].Position)
            y = read_val(found["YDrive"].Position)
            print(f"  XY: ({x}, {y}) nm")
        except Exception as e:
            print(f"  XY: error - {e}")

show_status()
print()

while True:
    try:
        cmd = input("Ti> ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        break
    if not cmd:
        continue
    parts = cmd.split()

    try:
        if parts[0] == "quit":
            break

        elif parts[0] == "status":
            show_status()

        elif parts[0] == "shutter":
            s = found.get("DiaShutter")
            if not s:
                print("  DiaShutter not found")
                continue
            if len(parts) < 2:
                print(f"  Shutter state: Value={read_val(s.Value)}")
                continue
            if parts[1] == "open":
                print("  Calling DiaShutter.Open()...", end=" ", flush=True)
                s.Open()
                print("done")
                time.sleep(0.3)
                print(f"  State after: Value={read_val(s.Value)}")
            elif parts[1] == "close":
                print("  Calling DiaShutter.Close()...", end=" ", flush=True)
                s.Close()
                print("done")
                time.sleep(0.3)
                print(f"  State after: Value={read_val(s.Value)}")
            else:
                print("  Usage: shutter open|close")

        elif parts[0] == "lamp":
            l = found.get("DiaLamp")
            if not l:
                print("  DiaLamp not found")
                continue
            if len(parts) < 2:
                print(f"  Lamp: Enabled={l.Enabled}, Intensity={read_val(l.Value)}")
                continue
            if parts[1] == "on":
                print("  Calling DiaLamp.On()...", end=" ", flush=True)
                l.On()
                print("done")
            elif parts[1] == "off":
                print("  Calling DiaLamp.Off()...", end=" ", flush=True)
                l.Off()
                print("done")
            else:
                try:
                    v = int(parts[1])
                    print(f"  Setting intensity to {v}...", end=" ", flush=True)
                    l.Value.RawValue = v
                    print("done")
                except ValueError:
                    print("  Usage: lamp on|off|<0-24>")

        elif parts[0] == "z":
            z = found.get("ZDrive")
            if not z:
                print("  ZDrive not found")
                continue
            if len(parts) == 1:
                pos = read_val(z.Position)
                print(f"  Z = {pos} nm ({pos/1000:.1f} µm)" if pos else f"  Z = {pos}")
            elif parts[1] in ("up", "down") and len(parts) >= 3:
                nm = int(parts[2])
                delta = nm if parts[1] == "up" else -nm
                print(f"  Moving Z by {delta} nm...", end=" ", flush=True)
                z.MoveRelative(delta)
                print("done")
                time.sleep(0.5)
                pos = read_val(z.Position)
                print(f"  Z now = {pos} nm ({pos/1000:.1f} µm)" if pos else f"  Z = {pos}")
            else:
                print("  Usage: z | z up <nm> | z down <nm>")

        elif parts[0] == "xy":
            x = found.get("XDrive")
            y = found.get("YDrive")
            if not x or not y:
                print("  XDrive/YDrive not found")
                continue
            xp = read_val(x.Position)
            yp = read_val(y.Position)
            print(f"  X={xp} nm, Y={yp} nm")

        elif parts[0] == "nose" and len(parts) >= 2:
            n = found.get("Nosepiece")
            if not n:
                print("  Nosepiece not found")
                continue
            v = int(parts[1])
            print(f"  Setting nosepiece to {v}...", end=" ", flush=True)
            n.Position.RawValue = v
            print("done")

        elif parts[0] == "filter" and len(parts) >= 2:
            f = found.get("FilterBlockCassette1")
            if not f:
                print("  FilterBlockCassette1 not found")
                continue
            v = int(parts[1])
            print(f"  Setting filter to {v}...", end=" ", flush=True)
            f.Position.RawValue = v
            print("done")

        else:
            print("  Unknown command. Type 'status' or 'quit'")

    except Exception as e:
        print(f"  ERROR: {type(e).__name__}: {e}")

print("\nCleaning up...")
ctypes.windll.ole32.CoUninitialize()
print("Done.")
