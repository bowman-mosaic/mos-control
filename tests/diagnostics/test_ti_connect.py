"""
Test Nikon Ti: probe MainController, try ShowWindow, and test actual control.
"""
import ctypes, sys, time

ctypes.windll.ole32.CoInitializeEx(None, 2)
import comtypes.client

scope = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")
print("Scope created.\n")

# ── Check MainController ──
print("=== MainController ===")
mc = scope.MainController
print(f"  Type: {type(mc)}")
for attr in dir(mc):
    if attr.startswith("_"):
        continue
    try:
        val = getattr(mc, attr)
        if hasattr(val, 'RawValue'):
            print(f"  .{attr}.RawValue = {val.RawValue}")
        elif callable(val):
            print(f"  .{attr}() [method]")
        else:
            print(f"  .{attr} = {val}")
    except Exception as e:
        print(f"  .{attr} -> {e}")

# ── Read nosepiece BEFORE ──
nose = scope.Nosepiece
pos_before = nose.Position.RawValue
print(f"\n=== Nosepiece test ===")
print(f"  Position BEFORE: {pos_before}")

# ── Try SetPosition (scope-level method) ──
print("\n=== Trying scope.SetPosition / SetPositionEx ===")
try:
    print(f"  scope.SetPosition exists: {scope.SetPosition}")
    print(f"  scope.GetPosition exists: {scope.GetPosition}")
except Exception as e:
    print(f"  Error: {e}")

# ── Try different approaches to move nosepiece ──
target = 2 if pos_before != 2 else 1
print(f"\n=== Attempting to move nosepiece to {target} ===")

print(f"  Method 1: nose.Position.RawValue = {target}")
try:
    nose.Position.RawValue = target
    time.sleep(1)
    print(f"    Result: Position = {nose.Position.RawValue}")
except Exception as e:
    print(f"    Error: {e}")

if nose.Position.RawValue == pos_before:
    print(f"  Method 2: scope.SetPosition('Nosepiece', {target})")
    try:
        scope.SetPosition("Nosepiece", target)
        time.sleep(1)
        print(f"    Result: Position = {nose.Position.RawValue}")
    except Exception as e:
        print(f"    Error: {e}")

if nose.Position.RawValue == pos_before:
    print(f"  Method 3: scope.SetPositionEx('Nosepiece', {target}, 0)")
    try:
        scope.SetPositionEx("Nosepiece", target, 0)
        time.sleep(1)
        print(f"    Result: Position = {nose.Position.RawValue}")
    except Exception as e:
        print(f"    Error: {e}")

# ── Try lamp ──
print("\n=== DiaLamp test ===")
lamp = scope.DiaLamp
print(f"  Value.RawValue = {lamp.Value.RawValue}")
print("  Calling lamp.On()...")
try:
    lamp.On()
    time.sleep(0.5)
    print(f"  Value after On(): {lamp.Value.RawValue}")
except Exception as e:
    print(f"  Error: {e}")

# ── Try ShowWindow ──
print("\n=== ShowWindow ===")
print("  Calling scope.ShowWindow(1)...")
try:
    scope.ShowWindow(1)
    print("  ShowWindow returned. Check if a Nikon controller window appeared!")
    print("  (If it did, try controlling the scope from THAT window)")
    input("  Press Enter after checking...")
except Exception as e:
    print(f"  Error: {e}")

# ── Try SendData ──
print("\n=== SendData / GetData ===")
# The Nikon Ti serial protocol uses command bytes
# Command to read nosepiece position: see Nikon Ti protocol docs
try:
    print(f"  scope.ReadVersion: {scope.ReadVersion()}")
except Exception as e:
    print(f"  ReadVersion error: {e}")

try:
    print(f"  scope.ReadProgramName: {scope.ReadProgramName()}")
except Exception as e:
    print(f"  ReadProgramName error: {e}")

try:
    print(f"  scope.ReadChecksum: {scope.ReadChecksum()}")
except Exception as e:
    print(f"  ReadChecksum error: {e}")

print("\n=== Done ===")
ctypes.windll.ole32.CoUninitialize()
