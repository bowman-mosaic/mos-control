#!/usr/bin/env python3
"""
Diagnostic script: discovers the Nikon Ti COM interface structure.

Run with:  python test_ti_probe.py

This probes the main scope object to find how devices are accessed.
"""

import comtypes
import comtypes.client

comtypes.CoInitialize()


def probe_object(obj, label, depth=0):
    """Print all accessible attributes of a COM object."""
    indent = "  " * depth
    print(f"\n{indent}{'=' * 60}")
    print(f"{indent}  {label}")
    print(f"{indent}{'=' * 60}")

    attrs = []
    for name in dir(obj):
        if name.startswith("_"):
            continue
        try:
            val = getattr(obj, name)
            type_name = type(val).__name__
            # Truncate long values
            val_str = repr(val)
            if len(val_str) > 100:
                val_str = val_str[:100] + "..."
            attrs.append((name, val_str, type_name))
            print(f"{indent}  {name:35s} = {val_str}  ({type_name})")
        except Exception as ex:
            err = str(ex)
            if len(err) > 80:
                err = err[:80] + "..."
            attrs.append((name, f"ERROR: {err}", "error"))
            print(f"{indent}  {name:35s}   ERROR: {err}")

    return attrs


# ── Try the main scope ProgIDs ───────────────────────────────────────────────

progids_to_try = [
    "Nikon.TiScope.NikonTi",
    "Nikon.NikonTiS.NikonTiDevices",
    "Nikon.TiScope.MainController",
]

for progid in progids_to_try:
    print(f"\n\n{'#' * 60}")
    print(f"  Trying: {progid}")
    print(f"{'#' * 60}")
    try:
        obj = comtypes.client.CreateObject(progid)
        print(f"  --> Created successfully!  type={type(obj).__name__}")
        attrs = probe_object(obj, progid)

        # If any attribute returned a COM object, probe it one level deeper
        for name, val_str, type_name in attrs:
            if type_name not in ("str", "int", "float", "bool", "NoneType", "error"):
                if "ERROR" not in val_str and "POINTER" in type_name or "comtypes" in type_name.lower():
                    try:
                        sub_obj = getattr(obj, name)
                        if sub_obj is not None:
                            probe_object(sub_obj, f"{progid}.{name}", depth=1)
                    except Exception:
                        pass

    except Exception as e:
        print(f"  --> FAILED: {e}")

# ── Also try generating the type library wrapper ─────────────────────────────

print(f"\n\n{'#' * 60}")
print(f"  Attempting to generate type library info...")
print(f"{'#' * 60}")

try:
    obj = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")
    # Try to get the type library
    try:
        tlb = comtypes.client.GetModule(
            [comtypes.GUID("{00000000-0000-0000-0000-000000000000}"), 0, 0]
        )
    except Exception:
        pass

    # List COM interfaces on the object
    print(f"\n  Object type: {type(obj)}")
    print(f"  Object repr: {repr(obj)[:200]}")

    # Try QueryInterface for IDispatch
    try:
        from comtypes.automation import IDispatch
        disp = obj.QueryInterface(IDispatch)
        print(f"  IDispatch: {disp}")

        # Try GetTypeInfo
        try:
            ti = disp.GetTypeInfo(0, 0)
            print(f"  TypeInfo: {ti}")
            # Get type attributes
            ta = ti.GetTypeAttr()
            print(f"  TypeAttr: guid={ta.guid}, funcs={ta.cFuncs}, vars={ta.cVars}")

            # List functions
            for i in range(ta.cFuncs):
                fd = ti.GetFuncDesc(i)
                names = ti.GetNames(fd.memid)
                print(f"    func[{i}]: {names[0] if names else '?'} "
                      f"(invkind={fd.invkind}, params={fd.cParams})")
            # List variables
            for i in range(ta.cVars):
                vd = ti.GetVarDesc(i)
                names = ti.GetNames(vd.memid)
                print(f"    var[{i}]: {names[0] if names else '?'}")

        except Exception as e:
            print(f"  GetTypeInfo failed: {e}")
    except Exception as e:
        print(f"  IDispatch query failed: {e}")

except Exception as e:
    print(f"  Failed: {e}")

print("\n\nDone. Copy/paste this output so we can adjust the module.")
