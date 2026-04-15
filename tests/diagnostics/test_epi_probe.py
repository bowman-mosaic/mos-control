"""Quick probe of the EpiShutter interface to debug why it doesn't work."""
import comtypes
import comtypes.client
import traceback

comtypes.CoInitialize()
scope = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")
epi = scope.EpiShutter

print("=" * 60)
print("  EpiShutter interface dump")
print("=" * 60)
for attr in sorted(dir(epi)):
    if attr.startswith("_"):
        continue
    try:
        val = getattr(epi, attr)
        v_repr = repr(val)
        if len(v_repr) > 100:
            v_repr = v_repr[:100] + "..."
        print(f"  {attr:35s} = {v_repr}  ({type(val).__name__})")
    except Exception as e:
        print(f"  {attr:35s}   ERROR: {str(e)[:80]}")

print("\n" + "=" * 60)
print("  Trying IsMounted and Value")
print("=" * 60)

try:
    mounted = epi.IsMounted
    print(f"  IsMounted type: {type(mounted).__name__}")
    print(f"  IsMounted.RawValue: {mounted.RawValue}")
except Exception as e:
    print(f"  IsMounted read failed: {e}")

try:
    val = epi.Value
    print(f"  Value type: {type(val).__name__}")
    print(f"  Value.RawValue: {val.RawValue}")
except Exception as e:
    print(f"  Value read failed: {e}")

print("\n" + "=" * 60)
print("  Trying Open() / Close()")
print("=" * 60)

for method_name in ["Open", "Close"]:
    print(f"\n  Trying epi.{method_name}()...")
    try:
        fn = getattr(epi, method_name, None)
        if fn is None:
            print(f"    -> Method '{method_name}' does NOT exist!")
        else:
            print(f"    -> Method exists: {fn}")
            fn()
            print(f"    -> {method_name}() succeeded!")
    except Exception as e:
        print(f"    -> {method_name}() FAILED: {e}")
        traceback.print_exc()

print("\n" + "=" * 60)
print("  Trying Position / Value RawValue write")
print("=" * 60)

for prop_name in ["Value", "Position"]:
    for test_val in [1, 0]:
        print(f"\n  Trying epi.{prop_name}.RawValue = {test_val}...")
        try:
            p = getattr(epi, prop_name, None)
            if p is None:
                print(f"    -> Property '{prop_name}' does NOT exist!")
                break
            p.RawValue = test_val
            print(f"    -> Write succeeded! Now RawValue = {p.RawValue}")
        except Exception as e:
            print(f"    -> FAILED: {e}")

print("\n" + "=" * 60)
print("  Also checking DiaShutter for comparison")
print("=" * 60)

dia = scope.DiaShutter
for method_name in ["Open", "Close"]:
    print(f"\n  Trying dia.{method_name}()...")
    try:
        fn = getattr(dia, method_name, None)
        if fn is None:
            print(f"    -> Method '{method_name}' does NOT exist!")
        else:
            fn()
            print(f"    -> {method_name}() succeeded!")
    except Exception as e:
        print(f"    -> {method_name}() FAILED: {e}")

print("\nDone.")
