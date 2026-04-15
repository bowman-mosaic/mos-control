"""
Automated Nikon Ti control test — no input() prompts.
Tests shutter (using IsOpened + Open/Close + SetPort) and lamp.
"""
import ctypes, time

ctypes.windll.ole32.CoInitializeEx(None, 2)
import comtypes.client

scope = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")
print("Scope created.\n")

# ── DiaShutter ──
print("=== DiaShutter ===")
s = scope.DiaShutter
print(f"  IsOpened.RawValue = {s.IsOpened.RawValue}")
print(f"  Value.RawValue    = {s.Value.RawValue}")

# Check what port is set
try:
    port = s.GetPort()
    print(f"  GetPort() = {port}")
except Exception as e:
    print(f"  GetPort() error: {e}")

# Try Open and check IsOpened
print("\n  >> Calling Open()...")
try:
    s.Open()
    time.sleep(1)
    print(f"     IsOpened after Open() = {s.IsOpened.RawValue}")
    print(f"     Value after Open()    = {s.Value.RawValue}")
except Exception as e:
    print(f"     Open() error: {e}")

# Try Close and check IsOpened
print("\n  >> Calling Close()...")
try:
    s.Close()
    time.sleep(1)
    print(f"     IsOpened after Close() = {s.IsOpened.RawValue}")
    print(f"     Value after Close()    = {s.Value.RawValue}")
except Exception as e:
    print(f"     Close() error: {e}")

# Try SetPort then Open
print("\n  >> Trying SetPort(0) then Open()...")
try:
    s.SetPort(0)
    time.sleep(0.3)
    s.Open()
    time.sleep(1)
    print(f"     IsOpened = {s.IsOpened.RawValue}")
except Exception as e:
    print(f"     Error: {e}")

print("\n  >> Trying SetPort(1) then Open()...")
try:
    s.SetPort(1)
    time.sleep(0.3)
    s.Open()
    time.sleep(1)
    print(f"     IsOpened = {s.IsOpened.RawValue}")
except Exception as e:
    print(f"     Error: {e}")

# Try scope-level
print("\n  >> scope.SetPosition('DiaShutter', 1)...")
try:
    scope.SetPosition("DiaShutter", 1)
    time.sleep(1)
    print(f"     IsOpened = {s.IsOpened.RawValue}")
except Exception as e:
    print(f"     Error: {e}")

print("\n  >> scope.SetPosition('DiaShutter', 0)...")
try:
    scope.SetPosition("DiaShutter", 0)
    time.sleep(1)
    print(f"     IsOpened = {s.IsOpened.RawValue}")
except Exception as e:
    print(f"     Error: {e}")

# ── DiaLamp ──
print("\n\n=== DiaLamp ===")
l = scope.DiaLamp
print(f"  Value.RawValue = {l.Value.RawValue}")
print(f"  LowerLimit     = {l.LowerLimit.RawValue}")
print(f"  UpperLimit     = {l.UpperLimit.RawValue}")

# All DiaLamp properties
for attr in ["IsOpened", "Enabled", "Switch", "Status"]:
    try:
        val = getattr(l, attr)
        if hasattr(val, 'RawValue'):
            print(f"  {attr}.RawValue = {val.RawValue}")
        else:
            print(f"  {attr} = {val}")
    except Exception as e:
        print(f"  {attr}: not available")

print("\n  >> Setting Value.RawValue = 12 (intensity)...")
try:
    l.Value.RawValue = 12
    time.sleep(0.3)
    print(f"     Value now = {l.Value.RawValue}")
except Exception as e:
    print(f"     Error: {e}")

print("  >> Calling On()...")
try:
    l.On()
    time.sleep(1)
    print(f"     Value after On() = {l.Value.RawValue}")
except Exception as e:
    print(f"     On() error: {e}")

print("  >> scope.SetPosition('DiaLamp', 12)...")
try:
    scope.SetPosition("DiaLamp", 12)
    time.sleep(1)
    print(f"     Value after SetPosition = {l.Value.RawValue}")
except Exception as e:
    print(f"     SetPosition error: {e}")

# ── Nosepiece (known working) ──
print("\n\n=== Nosepiece (verification) ===")
n = scope.Nosepiece
print(f"  Position = {n.Position.RawValue}")
target = 3 if n.Position.RawValue != 3 else 2
print(f"  >> Moving to {target}...")
n.Position.RawValue = target
time.sleep(2)
print(f"  Position after = {n.Position.RawValue}")
print(f"  >>> DID THE NOSEPIECE PHYSICALLY ROTATE? <<<")

# ── ZDrive ──
print("\n\n=== ZDrive ===")
z = scope.ZDrive
print(f"  Position = {z.Position.RawValue}")
print(f"  Limits   = {z.LowerLimit.RawValue} to {z.UpperLimit.RawValue}")
print("  >> Moving +5000 nm (5 µm)...")
try:
    z.MoveRelative(5000)
    time.sleep(1)
    print(f"     Position after = {z.Position.RawValue}")
except Exception as e:
    print(f"     Error: {e}")

print("\n=== SUMMARY ===")
print(f"  Nosepiece: {n.Position.RawValue}")
print(f"  ZDrive:    {z.Position.RawValue} nm")
print(f"  Shutter:   IsOpened={s.IsOpened.RawValue}, Value={s.Value.RawValue}")
print(f"  Lamp:      Value={l.Value.RawValue}")

print("\nDone.")
ctypes.windll.ole32.CoUninitialize()
