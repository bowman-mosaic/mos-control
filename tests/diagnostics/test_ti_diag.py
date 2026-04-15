"""
Deep diagnostic of the Nikon Ti COM interface.
Enumerates all methods/properties and checks real hardware connectivity.
"""
import ctypes, sys

print("Initializing COM...")
ctypes.windll.ole32.CoInitializeEx(None, 2)

import comtypes.client
import comtypes

print("Creating scope object...")
scope = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")

# ── List all COM interfaces the scope supports ──
print("\n=== Scope COM interfaces ===")
print(f"  Type: {type(scope)}")
print(f"  Repr: {repr(scope)}")

# ── Try to find Initialize / Connect / Open methods ──
print("\n=== Searching for init/connect methods on scope ===")
ti_iface = scope._com_interfaces_[0] if hasattr(scope, '_com_interfaces_') else None
if ti_iface:
    print(f"  Primary interface: {ti_iface}")
    for name in dir(ti_iface):
        if not name.startswith("_"):
            print(f"    .{name}")

print("\n=== All attributes on scope object ===")
for attr in dir(scope):
    if attr.startswith("_"):
        continue
    try:
        val = getattr(scope, attr)
        print(f"  .{attr} = {val}")
    except Exception as e:
        print(f"  .{attr} -> ERROR: {e}")

# ── Check IsMounted properly (read RawValue) ──
print("\n=== Device IsMounted (reading RawValue) ===")
devices = ["DiaShutter", "EpiShutter", "DiaLamp", "Nosepiece",
           "FilterBlockCassette1", "LightPathDrive",
           "ZDrive", "XDrive", "YDrive", "PFS"]

for name in devices:
    try:
        d = getattr(scope, name)
        mounted_param = d.IsMounted
        try:
            mounted = mounted_param.RawValue
        except Exception:
            mounted = "??"
        print(f"  {name}: IsMounted.RawValue = {mounted}")
    except Exception as e:
        print(f"  {name}: ERROR - {e}")

# ── Check if there's a "Connected" or "Status" property ──
print("\n=== Checking scope-level properties ===")
for prop in ["Connected", "IsConnected", "Status", "Initialized",
             "IsInitialized", "ControllerType", "SystemType",
             "Version", "SerialNumber", "Name"]:
    try:
        val = getattr(scope, prop)
        try:
            rv = val.RawValue
            print(f"  scope.{prop}.RawValue = {rv}")
        except Exception:
            print(f"  scope.{prop} = {val}")
    except AttributeError:
        pass
    except Exception as e:
        print(f"  scope.{prop} -> ERROR: {e}")

# ── Try QueryInterface for other interfaces ──
print("\n=== Enumerating type library ===")
try:
    import comtypes.typeinfo
    ti = scope._com_interfaces_[0]
    print(f"  Interface: {ti}")
    print(f"  IID: {ti._iid_}")
    # List all methods from the interface
    print("  Methods:")
    for name in dir(ti):
        if not name.startswith("_"):
            obj = getattr(ti, name, None)
            if callable(obj):
                print(f"    {name}()")
            else:
                print(f"    {name}")
except Exception as e:
    print(f"  Error: {e}")

# ── Inspect DiaShutter interface specifically ──
print("\n=== DiaShutter interface details ===")
shutter = scope.DiaShutter
print(f"  Type: {type(shutter)}")
si = shutter._com_interfaces_[0] if hasattr(shutter, '_com_interfaces_') else None
if si:
    print(f"  Interface: {si}")
    print(f"  Methods/props:")
    for name in dir(si):
        if not name.startswith("_"):
            print(f"    .{name}")

# ── Inspect DiaLamp interface ──
print("\n=== DiaLamp interface details ===")
lamp = scope.DiaLamp
print(f"  Type: {type(lamp)}")
li = lamp._com_interfaces_[0] if hasattr(lamp, '_com_interfaces_') else None
if li:
    print(f"  Interface: {li}")
    print(f"  Methods/props:")
    for name in dir(li):
        if not name.startswith("_"):
            print(f"    .{name}")

# ── Try DiaLamp specific properties ──
print("\n=== DiaLamp property reads ===")
for prop in ["Value", "Enabled", "On", "Off", "Switch", "Voltage",
             "Position", "RawValue", "LowerLimit", "UpperLimit"]:
    try:
        val = getattr(lamp, prop)
        try:
            rv = val.RawValue
            print(f"  lamp.{prop}.RawValue = {rv}")
        except Exception:
            print(f"  lamp.{prop} = {val}")
    except AttributeError:
        pass
    except Exception as e:
        print(f"  lamp.{prop} -> {type(e).__name__}: {e}")

# ── Try ZDrive specifics ──
print("\n=== ZDrive property reads ===")
zdrive = scope.ZDrive
for prop in ["Position", "Speed", "RawValue", "Value",
             "LowerLimit", "UpperLimit", "IsMounted"]:
    try:
        val = getattr(zdrive, prop)
        try:
            rv = val.RawValue
            print(f"  zdrive.{prop}.RawValue = {rv}")
        except Exception:
            print(f"  zdrive.{prop} = {val}")
    except AttributeError:
        pass
    except Exception as e:
        print(f"  zdrive.{prop} -> {type(e).__name__}: {e}")

print("\n=== Done ===")
ctypes.windll.ole32.CoUninitialize()
