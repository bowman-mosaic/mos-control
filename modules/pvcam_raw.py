"""
Low-level ctypes wrapper around pvcam64.dll.

Replaces PyVCAM by calling the PVCAM C library directly, which means it works
with *any* installed PVCAM driver version — no Python-package / DLL mismatch.

C type cheat-sheet (from master.h):
    int16  -> ctypes.c_short   (16-bit signed)
    uns16  -> ctypes.c_ushort  (16-bit unsigned)
    int32  -> ctypes.c_int     (32-bit signed)
    uns32  -> ctypes.c_uint    (32-bit unsigned)
    rs_bool -> ctypes.c_ushort (PVCAM boolean — TRUE=1, FALSE=0)
    void*  -> ctypes.c_void_p  (generic pointer)
"""

import ctypes
import ctypes.wintypes
import numpy as np
from ctypes import (
    c_short, c_ushort, c_int, c_uint, c_void_p, c_char, c_char_p,
    byref, POINTER, Structure,
)

# ---------------------------------------------------------------------------
# Type aliases — mirrors master.h so you can read PVCAM docs side-by-side
# ---------------------------------------------------------------------------
int16   = c_short       # signed 16-bit
uns16   = c_ushort      # unsigned 16-bit
int32   = c_int         # signed 32-bit
uns32   = c_uint        # unsigned 32-bit
rs_bool = c_ushort      # PVCAM boolean (uns16 under the hood)

# ---------------------------------------------------------------------------
# Constants — generated from pvcam.h via pylablib reference
# ---------------------------------------------------------------------------
CAM_NAME_LEN   = 32
ERROR_MSG_LEN  = 255
PV_OK          = 1
PV_FAIL        = 0

# pl_cam_open() mode
OPEN_EXCLUSIVE = 0

# Exposure modes (pl_exp_setup_cont / pl_exp_setup_seq)
TIMED_MODE = 0

# Circular-buffer mode for pl_exp_setup_cont
CIRC_NONE      = 0
CIRC_OVERWRITE = 1
CIRC_NO_OVERWRITE = 2

# Camera status values returned by pl_exp_check_cont_status
READOUT_NOT_ACTIVE    = 0
EXPOSURE_IN_PROGRESS  = 1
READOUT_IN_PROGRESS   = 2
READOUT_COMPLETE      = 3
FRAME_AVAILABLE       = READOUT_COMPLETE

# Abort modes
CCS_NO_CHANGE      = 0
CCS_HALT           = 1
CCS_HALT_CLOSE_SHTR = 2

# Parameter attribute selectors for pl_get_param
ATTR_CURRENT  = 0
ATTR_COUNT    = 1
ATTR_TYPE     = 2
ATTR_MIN      = 3
ATTR_MAX      = 4
ATTR_DEFAULT  = 5

# Helper to build PARAM_* IDs the way pvcam.h does:
#   param_id = (type << 16) | (class_ << 24) | id_
#   but stored as signed int32 in Python
def _param_id(cls, typ, idx):
    v = (cls << 16) | (typ << 24) | idx
    return ctypes.c_int32(v).value      # wrap to signed int32

# --- Commonly used parameters ---
# Sensor geometry
PARAM_SER_SIZE     = _param_id(2, 6, 58)   # sensor width  (uns16)
PARAM_PAR_SIZE     = _param_id(2, 6, 57)   # sensor height (uns16)
PARAM_BIT_DEPTH    = _param_id(2, 1, 511)  # bits per pixel (int16)
PARAM_CHIP_NAME    = _param_id(2, 13, 129) # sensor name   (char*)
PARAM_TEMP         = _param_id(2, 1, 525)  # CCD temp ×100 (int16)

# Exposure
PARAM_EXP_TIME     = _param_id(3, 6, 1)    # exposure time (uns16, ms)
PARAM_EXP_RES      = _param_id(3, 9, 2)    # exposure resolution (enum)
PARAM_EXPOSURE_TIME = _param_id(3, 8, 8)   # 64-bit exposure (uns64)

# Binning
PARAM_BINNING_SER  = _param_id(3, 9, 165)  # serial binning (enum)
PARAM_BINNING_PAR  = _param_id(3, 9, 166)  # parallel binning (enum)


# ---------------------------------------------------------------------------
# Structures — must match the C struct layout byte-for-byte
# ---------------------------------------------------------------------------
class rgn_type(Structure):
    """Region-of-interest for acquisition.

    In C:
        typedef struct {
            uns16 s1;    // serial (X) start pixel
            uns16 s2;    // serial (X) end pixel
            uns16 sbin;  // serial binning factor
            uns16 p1;    // parallel (Y) start pixel
            uns16 p2;    // parallel (Y) end pixel
            uns16 pbin;  // parallel binning factor
        } rgn_type;
    """
    _fields_ = [
        ("s1",   uns16),
        ("s2",   uns16),
        ("sbin", uns16),
        ("p1",   uns16),
        ("p2",   uns16),
        ("pbin", uns16),
    ]


# ---------------------------------------------------------------------------
# Load the DLL
# ---------------------------------------------------------------------------
_dll = None

def _load_dll():
    global _dll
    if _dll is not None:
        return _dll
    try:
        _dll = ctypes.WinDLL("pvcam64")
    except OSError:
        try:
            _dll = ctypes.WinDLL("pvcam32")
        except OSError:
            raise RuntimeError(
                "Cannot find pvcam64.dll or pvcam32.dll. "
                "Is PVCAM installed?"
            )

    # -- Declare function signatures so ctypes can marshal arguments --
    # Each line reads like: "this C function returns <restype> and takes <argtypes>"

    # rs_bool pl_pvcam_init(void)
    _dll.pl_pvcam_init.restype  = rs_bool
    _dll.pl_pvcam_init.argtypes = []

    # rs_bool pl_pvcam_uninit(void)
    _dll.pl_pvcam_uninit.restype  = rs_bool
    _dll.pl_pvcam_uninit.argtypes = []

    # int16 pl_error_code(void)
    _dll.pl_error_code.restype  = int16
    _dll.pl_error_code.argtypes = []

    # rs_bool pl_error_message(int16 err_code, char* msg)
    _dll.pl_error_message.restype  = rs_bool
    _dll.pl_error_message.argtypes = [int16, c_char_p]

    # rs_bool pl_cam_get_total(int16* totl_cams)
    _dll.pl_cam_get_total.restype  = rs_bool
    _dll.pl_cam_get_total.argtypes = [POINTER(int16)]

    # rs_bool pl_cam_get_name(int16 cam_num, char* camera_name)
    _dll.pl_cam_get_name.restype  = rs_bool
    _dll.pl_cam_get_name.argtypes = [int16, c_char_p]

    # rs_bool pl_cam_open(char* camera_name, int16* hcam, int16 o_mode)
    _dll.pl_cam_open.restype  = rs_bool
    _dll.pl_cam_open.argtypes = [c_char_p, POINTER(int16), int16]

    # rs_bool pl_cam_close(int16 hcam)
    _dll.pl_cam_close.restype  = rs_bool
    _dll.pl_cam_close.argtypes = [int16]

    # rs_bool pl_get_param(int16 hcam, uns32 param_id, int16 param_attribute,
    #                      void* param_value)
    _dll.pl_get_param.restype  = rs_bool
    _dll.pl_get_param.argtypes = [int16, uns32, int16, c_void_p]

    # rs_bool pl_set_param(int16 hcam, uns32 param_id, void* param_value)
    _dll.pl_set_param.restype  = rs_bool
    _dll.pl_set_param.argtypes = [int16, uns32, c_void_p]

    # rs_bool pl_exp_setup_seq(int16 hcam, uns16 exp_total, uns16 rgn_total,
    #     rgn_type* rgn_array, int16 exp_mode, uns32 exposure_time,
    #     uns32* exp_bytes)
    _dll.pl_exp_setup_seq.restype  = rs_bool
    _dll.pl_exp_setup_seq.argtypes = [
        int16, uns16, uns16, POINTER(rgn_type), int16, uns32, POINTER(uns32)
    ]

    # rs_bool pl_exp_start_seq(int16 hcam, void* pixel_stream)
    _dll.pl_exp_start_seq.restype  = rs_bool
    _dll.pl_exp_start_seq.argtypes = [int16, c_void_p]

    # rs_bool pl_exp_setup_cont(int16 hcam, uns16 rgn_total,
    #     rgn_type* rgn_array, int16 exp_mode, uns32 exposure_time,
    #     uns32* exp_bytes, int16 buffer_mode)
    _dll.pl_exp_setup_cont.restype  = rs_bool
    _dll.pl_exp_setup_cont.argtypes = [
        int16, uns16, POINTER(rgn_type), int16, uns32, POINTER(uns32), int16
    ]

    # rs_bool pl_exp_start_cont(int16 hcam, void* pixel_stream, uns32 size)
    _dll.pl_exp_start_cont.restype  = rs_bool
    _dll.pl_exp_start_cont.argtypes = [int16, c_void_p, uns32]

    # rs_bool pl_exp_check_cont_status(int16 hcam, int16* status,
    #     uns32* bytes_arrived, uns32* buffer_cnt)
    _dll.pl_exp_check_cont_status.restype  = rs_bool
    _dll.pl_exp_check_cont_status.argtypes = [
        int16, POINTER(int16), POINTER(uns32), POINTER(uns32)
    ]

    # rs_bool pl_exp_get_latest_frame(int16 hcam, void** frame)
    _dll.pl_exp_get_latest_frame.restype  = rs_bool
    _dll.pl_exp_get_latest_frame.argtypes = [int16, POINTER(c_void_p)]

    # rs_bool pl_exp_abort(int16 hcam, int16 cam_state)
    _dll.pl_exp_abort.restype  = rs_bool
    _dll.pl_exp_abort.argtypes = [int16, int16]

    # rs_bool pl_exp_finish_seq(int16 hcam, void* pixel_stream, int16 hbuf)
    _dll.pl_exp_finish_seq.restype  = rs_bool
    _dll.pl_exp_finish_seq.argtypes = [int16, c_void_p, int16]

    return _dll


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------
def _check(ok, context="PVCAM"):
    """Raise RuntimeError with PVCAM's own error message if a call failed."""
    if ok == PV_OK:
        return
    dll = _load_dll()
    code = dll.pl_error_code()
    msg_buf = ctypes.create_string_buffer(ERROR_MSG_LEN)
    dll.pl_error_message(code, msg_buf)
    raise RuntimeError(f"{context}: [{code}] {msg_buf.value.decode('ascii', errors='replace')}")


# ---------------------------------------------------------------------------
# Public API — Pythonic wrappers around the raw C functions
# ---------------------------------------------------------------------------

def init():
    """Initialize the PVCAM library. Call once at startup."""
    dll = _load_dll()
    _check(dll.pl_pvcam_init(), "pl_pvcam_init")


def uninit():
    """Shut down the PVCAM library."""
    dll = _load_dll()
    _check(dll.pl_pvcam_uninit(), "pl_pvcam_uninit")


def cam_count():
    """Return the number of cameras detected."""
    dll = _load_dll()
    total = int16(0)
    _check(dll.pl_cam_get_total(byref(total)), "pl_cam_get_total")
    return total.value


def cam_name(index=0):
    """Get the name of camera at *index* (e.g. 'PM1394Cam00')."""
    dll = _load_dll()
    name_buf = ctypes.create_string_buffer(CAM_NAME_LEN)
    _check(dll.pl_cam_get_name(int16(index), name_buf), "pl_cam_get_name")
    return name_buf.value.decode("ascii")


def cam_open(name):
    """Open a camera by name. Returns the camera handle (int16)."""
    dll = _load_dll()
    hcam = int16(0)
    _check(
        dll.pl_cam_open(name.encode("ascii"), byref(hcam), int16(OPEN_EXCLUSIVE)),
        "pl_cam_open",
    )
    return hcam.value


def cam_close(hcam):
    """Close a previously opened camera."""
    dll = _load_dll()
    _check(dll.pl_cam_close(int16(hcam)), "pl_cam_close")


# ---------------------------------------------------------------------------
# Parameter access
# ---------------------------------------------------------------------------

def get_param_uns16(hcam, param_id, attr=ATTR_CURRENT):
    dll = _load_dll()
    val = uns16(0)
    _check(
        dll.pl_get_param(int16(hcam), uns32(param_id), int16(attr), byref(val)),
        f"pl_get_param(0x{param_id & 0xFFFFFFFF:08X})",
    )
    return val.value


def get_param_int16(hcam, param_id, attr=ATTR_CURRENT):
    dll = _load_dll()
    val = int16(0)
    _check(
        dll.pl_get_param(int16(hcam), uns32(param_id), int16(attr), byref(val)),
        f"pl_get_param(0x{param_id & 0xFFFFFFFF:08X})",
    )
    return val.value


def get_param_uns32(hcam, param_id, attr=ATTR_CURRENT):
    dll = _load_dll()
    val = uns32(0)
    _check(
        dll.pl_get_param(int16(hcam), uns32(param_id), int16(attr), byref(val)),
        f"pl_get_param(0x{param_id & 0xFFFFFFFF:08X})",
    )
    return val.value


def get_param_str(hcam, param_id, attr=ATTR_CURRENT, buflen=256):
    dll = _load_dll()
    buf = ctypes.create_string_buffer(buflen)
    _check(
        dll.pl_get_param(int16(hcam), uns32(param_id), int16(attr), buf),
        f"pl_get_param(0x{param_id & 0xFFFFFFFF:08X})",
    )
    return buf.value.decode("ascii", errors="replace")


def set_param_uns32(hcam, param_id, value):
    dll = _load_dll()
    val = uns32(value)
    _check(
        dll.pl_set_param(int16(hcam), uns32(param_id), byref(val)),
        f"pl_set_param(0x{param_id & 0xFFFFFFFF:08X})",
    )


# ---------------------------------------------------------------------------
# Convenience: sensor info
# ---------------------------------------------------------------------------

def sensor_size(hcam):
    """Return (width, height) of the full sensor in pixels."""
    w = get_param_uns16(hcam, PARAM_SER_SIZE)
    h = get_param_uns16(hcam, PARAM_PAR_SIZE)
    return (w, h)


def bit_depth(hcam):
    return get_param_int16(hcam, PARAM_BIT_DEPTH)


def chip_name(hcam):
    return get_param_str(hcam, PARAM_CHIP_NAME)


def sensor_temp_c(hcam):
    """CCD temperature in °C (PVCAM reports hundredths of a degree)."""
    raw = get_param_int16(hcam, PARAM_TEMP)
    return raw / 100.0


# ---------------------------------------------------------------------------
# Acquisition helpers
# ---------------------------------------------------------------------------

def make_region(hcam, binning=1):
    """Build a full-sensor rgn_type with the given binning."""
    w, h = sensor_size(hcam)
    r = rgn_type()
    r.s1   = 0
    r.s2   = w - 1
    r.sbin = binning
    r.p1   = 0
    r.p2   = h - 1
    r.pbin = binning
    return r


def setup_cont(hcam, exposure_ms, binning=1, circ_mode=CIRC_OVERWRITE):
    """Set up continuous (live) acquisition. Returns frame_size in bytes."""
    dll = _load_dll()
    region = make_region(hcam, binning)
    frame_bytes = uns32(0)
    _check(
        dll.pl_exp_setup_cont(
            int16(hcam),
            uns16(1),                  # 1 region
            byref(region),
            int16(TIMED_MODE),
            uns32(exposure_ms),
            byref(frame_bytes),
            int16(circ_mode),
        ),
        "pl_exp_setup_cont",
    )
    return frame_bytes.value


def start_cont(hcam, buffer, buffer_size):
    """Start continuous acquisition into *buffer* (a ctypes array)."""
    dll = _load_dll()
    _check(
        dll.pl_exp_start_cont(
            int16(hcam),
            ctypes.cast(buffer, c_void_p),
            uns32(buffer_size),
        ),
        "pl_exp_start_cont",
    )


def check_cont_status(hcam):
    """Poll continuous acquisition status. Returns (status, bytes_arrived, buf_cnt)."""
    dll = _load_dll()
    status = int16(0)
    arrived = uns32(0)
    buf_cnt = uns32(0)
    _check(
        dll.pl_exp_check_cont_status(
            int16(hcam), byref(status), byref(arrived), byref(buf_cnt),
        ),
        "pl_exp_check_cont_status",
    )
    return status.value, arrived.value, buf_cnt.value


def get_latest_frame(hcam):
    """Return a ctypes void pointer to the latest frame data."""
    dll = _load_dll()
    frame_ptr = c_void_p(0)
    _check(
        dll.pl_exp_get_latest_frame(int16(hcam), byref(frame_ptr)),
        "pl_exp_get_latest_frame",
    )
    return frame_ptr


def abort(hcam, mode=CCS_HALT):
    """Abort acquisition."""
    dll = _load_dll()
    _check(dll.pl_exp_abort(int16(hcam), int16(mode)), "pl_exp_abort")


def finish_seq(hcam, buffer):
    """Finish a sequence acquisition."""
    dll = _load_dll()
    _check(
        dll.pl_exp_finish_seq(int16(hcam), ctypes.cast(buffer, c_void_p), int16(0)),
        "pl_exp_finish_seq",
    )


# ---------------------------------------------------------------------------
# Single-frame (sequence) acquisition
# ---------------------------------------------------------------------------

def setup_seq(hcam, exposure_ms, binning=1):
    """Set up a single-frame sequence. Returns frame_size in bytes."""
    dll = _load_dll()
    region = make_region(hcam, binning)
    frame_bytes = uns32(0)
    _check(
        dll.pl_exp_setup_seq(
            int16(hcam),
            uns16(1),                  # 1 exposure
            uns16(1),                  # 1 region
            byref(region),
            int16(TIMED_MODE),
            uns32(exposure_ms),
            byref(frame_bytes),
        ),
        "pl_exp_setup_seq",
    )
    return frame_bytes.value


def start_seq(hcam, buffer):
    """Start a single-frame sequence into *buffer*."""
    dll = _load_dll()
    _check(
        dll.pl_exp_start_seq(int16(hcam), ctypes.cast(buffer, c_void_p)),
        "pl_exp_start_seq",
    )


# ---------------------------------------------------------------------------
# High-level: frame as numpy array
# ---------------------------------------------------------------------------

def frame_to_numpy(frame_ptr, width, height, binning=1):
    """Interpret a raw frame pointer as a 2-D numpy uint16 array.

    The DLL owns the memory — we copy immediately so it's safe to use
    after the next frame arrives.
    """
    bw = width // binning
    bh = height // binning
    n_pixels = bw * bh
    arr_type = (uns16 * n_pixels)
    raw = ctypes.cast(frame_ptr, POINTER(arr_type)).contents
    return np.ctypeslib.as_array(raw).reshape(bh, bw).copy()


def poll_frame_numpy(hcam, width, height, binning=1, timeout_s=5.0):
    """Poll until a frame is ready, return it as a numpy array.

    Uses continuous-mode status polling (the reliable path for FireWire cameras).
    """
    import time
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        status, _, _ = check_cont_status(hcam)
        if status >= FRAME_AVAILABLE:
            ptr = get_latest_frame(hcam)
            return frame_to_numpy(ptr, width, height, binning)
        time.sleep(0.002)
    raise TimeoutError(f"No frame received within {timeout_s}s")
