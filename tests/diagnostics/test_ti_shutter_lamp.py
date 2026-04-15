"""
Targeted test for DiaShutter and DiaLamp control methods.
"""
import ctypes, time

ctypes.windll.ole32.CoInitializeEx(None, 2)
import comtypes.client

scope = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")
print("Scope created.\n")

# ── DiaShutter ──
print("=== DiaShutter ===")
s = scope.DiaShutter
print(f"  All properties:")
for attr in dir(s):
    if attr.startswith("_"):
        continue
    try:
        val = getattr(s, attr)
        if hasattr(val, 'RawValue'):
            print(f"    .{attr}.RawValue = {val.RawValue}")
        elif callable(val):
            print(f"    .{attr}() [method]")
        else:
            print(f"    .{attr} = {val}")
    except Exception as e:
        print(f"    .{attr} -> {e}")

print(f"\n  Current Value.RawValue = {s.Value.RawValue}")

print("  Trying s.Value.RawValue = 0 (close)...")
try:
    s.Value.RawValue = 0
    time.sleep(0.5)
    print(f"    Value after = {s.Value.RawValue}")
    print("    >>> Did the shutter physically close? <<<")
except Exception as e:
    print(f"    Error: {e}")

input("  Press Enter to try opening...")

print("  Trying s.Value.RawValue = 1 (open)...")
try:
    s.Value.RawValue = 1
    time.sleep(0.5)
    print(f"    Value after = {s.Value.RawValue}")
    print("    >>> Did the shutter physically open? <<<")
except Exception as e:
    print(f"    Error: {e}")

input("  Press Enter to continue to lamp test...")

# ── DiaLamp ──
print("\n=== DiaLamp ===")
l = scope.DiaLamp
print(f"  All properties:")
for attr in dir(l):
    if attr.startswith("_"):
        continue
    try:
        val = getattr(l, attr)
        if hasattr(val, 'RawValue'):
            print(f"    .{attr}.RawValue = {val.RawValue}")
        elif callable(val):
            print(f"    .{attr}() [method]")
        else:
            print(f"    .{attr} = {val}")
    except Exception as e:
        print(f"    .{attr} -> {e}")

print(f"\n  Current Value.RawValue = {l.Value.RawValue}")

print("  Test 1: Set intensity first, then On()")
print("    Setting Value.RawValue = 12...")
try:
    l.Value.RawValue = 12
    time.sleep(0.3)
    print(f"    Value after = {l.Value.RawValue}")
except Exception as e:
    print(f"    Error: {e}")

print("    Calling On()...")
try:
    l.On()
    time.sleep(0.5)
    print(f"    Value after On() = {l.Value.RawValue}")
    print("    >>> Did the lamp physically turn on? <<<")
except Exception as e:
    print(f"    Error: {e}")

input("  Press Enter to try turning off...")

print("    Calling Off()...")
try:
    l.Off()
    time.sleep(0.5)
    print(f"    Value after Off() = {l.Value.RawValue}")
except Exception as e:
    print(f"    Error: {e}")

# ── Try scope-level SetPosition for shutter ──
print("\n=== scope.SetPosition for DiaShutter ===")
print("  Trying scope.SetPosition('DiaShutter', 0)...")
try:
    scope.SetPosition("DiaShutter", 0)
    time.sleep(0.5)
    print(f"    DiaShutter.Value = {s.Value.RawValue}")
except Exception as e:
    print(f"    Error: {e}")

print("  Trying scope.SetPosition('DiaShutter', 1)...")
try:
    scope.SetPosition("DiaShutter", 1)
    time.sleep(0.5)
    print(f"    DiaShutter.Value = {s.Value.RawValue}")
except Exception as e:
    print(f"    Error: {e}")

# ── Try scope-level SetPosition for DiaLamp ──
print("\n=== scope.SetPosition for DiaLamp ===")
print("  Trying scope.SetPosition('DiaLamp', 12)...")
try:
    scope.SetPosition("DiaLamp", 12)
    time.sleep(0.5)
    print(f"    DiaLamp.Value = {l.Value.RawValue}")
    print("    >>> Did the lamp change? <<<")
except Exception as e:
    print(f"    Error: {e}")

# ── Try ShowWindow ──
print("\n=== ShowWindow ===")
print("  Trying scope.ShowWindow(True, True)...")
try:
    scope.ShowWindow(True, True)
    print("  Window opened — try controlling from there!")
    input("  Press Enter when done...")
except Exception as e:
    print(f"  Error: {e}")
    print("  Trying scope.ShowWindow(1, 1)...")
    try:
        scope.ShowWindow(1, 1)
        print("  Window opened — try controlling from there!")
        input("  Press Enter when done...")
    except Exception as e2:
        print(f"  Error: {e2}")

print("\nDone.")
ctypes.windll.ole32.CoUninitialize()
