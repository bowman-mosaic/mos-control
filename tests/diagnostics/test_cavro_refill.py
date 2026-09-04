#!/usr/bin/env python3
"""
Bench scripts for Cavro refill (aspiration) speed vs cavitation.

Talks to the pump over RS-485 directly — do not run this while Mosaic Control
Center has the bus open.

The control software currently refills at the pump max:

    max_vel = min(counts_per_stroke, int(5800 / velocity_scale))
            = min(3000, 2900) = 2900 counts/s

On a 0.5 mL syringe that is ~0.48 mL/s (~29 mL/min). A full stroke finishes
in about 1 s, which is fast enough to nucleate bubbles on the pull.

Usage
-----
    python tests/diagnostics/test_cavro_refill.py ports

    python tests/diagnostics/test_cavro_refill.py max --port COM3 --source 4 --waste 3
    python tests/diagnostics/test_cavro_refill.py sweep --port COM3 --source 4 --waste 3
    python tests/diagnostics/test_cavro_refill.py sweep --port COM3 --ul-min max,20000,10000,5000
    python tests/diagnostics/test_cavro_refill.py jog --port COM3 --source 4 --waste 3
"""
from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from datetime import datetime

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "syringe_pump"))

from ftdi_serial import Serial as FtdiSerial  # noqa: E402
from syringe_pump_control import list_serial_ports  # noqa: E402
from tecan_cavro import TecanCavro  # noqa: E402

# Match modules/cavro.py
COUNTS_PER_STROKE = 3000
VELOCITY_SCALE = 2.0
V_CMD_LIMIT = 5800  # raw V operand the firmware will accept
DEFAULT_SYRINGE_ML = 0.5
DEFAULT_SLOPE = 14


# ── units ────────────────────────────────────────────────────────────────────

def software_max_counts(counts_per_stroke: int = COUNTS_PER_STROKE,
                        velocity_scale: float = VELOCITY_SCALE) -> int:
    return min(counts_per_stroke, int(V_CMD_LIMIT / velocity_scale))


def counts_per_ml(syringe_ml: float, counts_per_stroke: int = COUNTS_PER_STROKE) -> float:
    return counts_per_stroke / syringe_ml


def ulmin_to_counts(ul_min: float, cpm: float) -> int:
    ml_s = ul_min / 60000.0
    return max(3, round(ml_s * cpm))


def counts_to_ulmin(counts: float, cpm: float) -> float:
    return (counts / cpm) * 60000.0


def counts_to_ml_s(counts: float, cpm: float) -> float:
    return counts / cpm


def clamp_counts(counts: int, max_counts: int) -> int:
    return max(3, min(int(counts), max_counts))


def parse_ulmin_list(text: str, max_ulmin: float) -> list[float]:
    out = []
    for p in text.split(","):
        p = p.strip().lower()
        if not p:
            continue
        if p in ("max", "software", "software-max"):
            out.append(max_ulmin)
        else:
            out.append(float(p))
    return out


def parse_floats(text: str) -> list[float]:
    return [float(p.strip()) for p in text.split(",") if p.strip()]


def parse_ints(text: str) -> list[int]:
    return [int(p.strip()) for p in text.split(",") if p.strip()]


def fmt_rate(counts: int, cpm: float) -> str:
    return (f"{counts} cts/s  |  {counts_to_ml_s(counts, cpm):.4f} mL/s  |  "
            f"{counts_to_ulmin(counts, cpm):.0f} uL/min")


def expected_stroke_s(volume_ml: float, counts: int, cpm: float) -> float:
    if counts <= 0:
        return float("inf")
    return (volume_ml * cpm) / counts


# ── pump helpers ─────────────────────────────────────────────────────────────

class Bench:
    def __init__(self, pump: TecanCavro, source: int, waste: int,
                 empty_counts: int, slope: int):
        self.pump = pump
        self.source = source
        self.waste = waste
        self.empty_counts = empty_counts
        self.slope = slope

    def halt(self):
        try:
            self.pump.clear_batch()
        except Exception:
            pass
        try:
            self.pump.halt()
        except Exception:
            pass

    def home(self):
        print("Homing (plunger + valve)...")
        self.pump.home()
        print(f"  plunger {self.pump.volume_ml:.3f} mL  valve {self.pump.valve_position}")

    def empty(self, port: int | None = None):
        port = self.waste if port is None else port
        vol = self.pump.volume_ml
        if vol <= 0.002:
            return
        print(f"Emptying {vol:.3f} mL → port {port}...")
        self.pump.switch_valve(port, wait=True)
        self.pump.start_batch()
        self.pump.move_absolute_counts(
            0, velocity_counts=self.empty_counts, wait=False)
        self.pump.execute(wait=True)
        print(f"  plunger {self.pump.volume_ml:.3f} mL")

    def refill(self, target_counts: int, vel_counts: int,
               slope: int | None = None) -> float:
        slope = self.slope if slope is None else slope
        self.pump.switch_valve(self.source, wait=True)
        t0 = time.perf_counter()
        self.pump.start_batch()
        self.pump.move_absolute_counts(
            target_counts,
            velocity_counts=vel_counts,
            slope_code=slope,
            wait=False,
        )
        self.pump.execute(wait=True)
        dt = time.perf_counter() - t0
        print(f"  plunger {self.pump.volume_ml:.3f} mL  elapsed {dt:.2f}s")
        return dt


def connect(port: str, address: int, syringe_ml: float,
            valve_positions: int) -> TecanCavro:
    TecanCavro.instances.clear()
    print(f"Opening {port} at 9600 baud (address {address})...")
    serial = FtdiSerial(
        device_port=port,
        baudrate=9600,
        read_timeout=100,
        write_timeout=100,
        connect_settle_time=1,
    )
    pump = TecanCavro(
        serial,
        address=address,
        syringe_volume_ml=syringe_ml,
        counts_per_stroke=COUNTS_PER_STROKE,
        velocity_scale=VELOCITY_SCALE,
        total_valve_positions=valve_positions,
        slope_code=DEFAULT_SLOPE,
    )
    # Probe so we fail fast if the address is wrong / bus is busy.
    ready, status = pump.status()
    print(f"  pump ready={ready} status={status}  "
          f"pos={pump.volume_ml:.3f} mL  valve={pump.valve_position}")
    return pump


def print_software_context(syringe_ml: float):
    cpm = counts_per_ml(syringe_ml)
    max_c = software_max_counts()
    print()
    print("Software refill today (modules/cavro.py continuous / coordinated / dual-push):")
    print(f"  syringe           {syringe_ml} mL")
    print(f"  counts/stroke     {COUNTS_PER_STROKE}")
    print(f"  counts/mL         {cpm:.0f}")
    print(f"  refill velocity   {fmt_rate(max_c, cpm)}")
    print(f"  full-stroke time  {expected_stroke_s(syringe_ml, max_c, cpm):.2f} s")
    print(f"  slope code        {DEFAULT_SLOPE}  (1=slowest accel, 20=fastest)")
    print()


# ── ports ────────────────────────────────────────────────────────────────────

def cmd_ports(_args=None):
    ports = list_serial_ports()
    if not ports:
        print("No serial ports found.")
        return
    print("Serial ports:")
    for p in ports:
        print(f"  {p['device']:10s}  {p['description']}")


# ── sweep ────────────────────────────────────────────────────────────────────

def default_ulmin_list(max_ulmin: float) -> list[float]:
    candidates = [max_ulmin, 20000, 15000, 10000, 7500, 5000, 2500, 1000]
    out = []
    for v in candidates:
        v = round(v)
        if v <= 0:
            continue
        if out and abs(v - out[-1]) < 50:
            continue
        if v > max_ulmin + 1:
            continue
        out.append(v)
    return out


def cmd_sweep(args):
    syringe_ml = args.syringe
    cpm = counts_per_ml(syringe_ml)
    max_c = software_max_counts()
    max_ulmin = counts_to_ulmin(max_c, cpm)
    print_software_context(syringe_ml)

    if args.ml_s:
        rates_ulmin = [v * 60000.0 for v in parse_floats(args.ml_s)]
    elif args.ul_min:
        rates_ulmin = parse_ulmin_list(args.ul_min, max_ulmin)
    else:
        rates_ulmin = default_ulmin_list(max_ulmin)

    slopes = parse_ints(args.slope) if args.slope else [args.slope_code]
    if args.waste is None:
        args.waste = args.source
    volume_ml = args.volume if args.volume is not None else syringe_ml
    volume_ml = min(volume_ml, syringe_ml)
    target_counts = max(1, round(volume_ml * cpm))
    empty_counts = max_c

    trials = []
    for ul in rates_ulmin:
        vel = clamp_counts(ulmin_to_counts(ul, cpm), max_c)
        actual_ul = counts_to_ulmin(vel, cpm)
        for slope in slopes:
            for cyc in range(args.cycles):
                trials.append((ul, vel, actual_ul, slope, cyc))

    print(f"{len(trials)} trial(s): {volume_ml:.3f} mL pull, "
          f"source port {args.source}, waste port {args.waste}")
    for i, (ul, vel, actual_ul, slope, cyc) in enumerate(trials, 1):
        exp = expected_stroke_s(volume_ml, vel, cpm)
        extra = f"  cycle {cyc + 1}/{args.cycles}" if args.cycles > 1 else ""
        tag = "  <-- software max" if abs(actual_ul - max_ulmin) < 1 else ""
        print(f"  {i:2d}. {actual_ul:7.0f} uL/min   "
              f"slope {slope:2d}   ~{exp:.1f}s{extra}{tag}")
        if vel != ulmin_to_counts(ul, cpm) and ul > max_ulmin:
            print(f"      (requested {ul:.0f} uL/min, clamped to pump max)")

    if not args.yes:
        input("Enter to start (Ctrl+C aborts)...")

    pump = None
    bench = None
    log_path = args.log
    ensure_log(log_path)

    try:
        pump = connect(args.port, args.address, syringe_ml, args.valves)
        bench = Bench(pump, args.source, args.waste, empty_counts, args.slope_code)
        if not args.no_home:
            bench.home()
        else:
            bench.halt()
            time.sleep(0.3)
        bench.empty()

        results = []
        for i, (ul, vel, actual_ul, slope, cyc) in enumerate(trials, 1):
            print()
            print("=" * 64)
            tag = "  SOFTWARE MAX" if abs(actual_ul - max_ulmin) < 1 else ""
            print(f"Trial {i}/{len(trials)}  {fmt_rate(vel, cpm)}  slope={slope}{tag}")
            print("=" * 64)
            dt = bench.refill(target_counts, vel, slope=slope)
            exp = expected_stroke_s(volume_ml, vel, cpm)
            print(f"  expected ~{exp:.2f}s")
            if args.hold > 0:
                time.sleep(args.hold)

            verdict = "unscored"
            if args.no_prompt:
                print("  (no prompt — watch the syringe, then it will empty)")
            else:
                while True:
                    verdict = score_prompt()
                    if verdict == "quit":
                        break
                    if verdict != "retry":
                        break
                    print("  retrying same rate...")
                    bench.empty()
                    dt = bench.refill(target_counts, vel, slope=slope)
                    print(f"  expected ~{exp:.2f}s")
                    if args.hold > 0:
                        time.sleep(args.hold)
                if verdict == "quit":
                    print("Stopping.")
                    break

            row = {
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "port": args.port,
                "address": args.address,
                "source": args.source,
                "waste": args.waste,
                "syringe_ml": syringe_ml,
                "volume_ml": volume_ml,
                "ul_min": round(actual_ul, 1),
                "ml_s": round(counts_to_ml_s(vel, cpm), 5),
                "counts_s": vel,
                "slope": slope,
                "cycle": cyc + 1,
                "elapsed_s": round(dt, 3),
                "expected_s": round(exp, 3),
                "cavitation": verdict,
            }
            results.append(row)
            append_log(log_path, row)
            print(f"  scored: {verdict}  (logged {log_path})")
            bench.empty()

        print_summary(results)
    except KeyboardInterrupt:
        print("\nCtrl+C — halting pump.")
    finally:
        if bench is not None:
            bench.halt()
        if pump is not None:
            try:
                pump.serial.disconnect()
            except Exception:
                pass


def score_prompt() -> str:
    while True:
        raw = input("  Cavitation? [y]es [n]o [m]aybe [r]etry [q]uit > ").strip().lower()
        if raw in ("y", "yes"):
            return "yes"
        if raw in ("n", "no"):
            return "no"
        if raw in ("m", "maybe", ""):
            return "maybe"
        if raw in ("r", "retry"):
            return "retry"
        if raw in ("q", "quit"):
            return "quit"
        print("  Please enter y / n / m / r / q")


def ensure_log(path: str):
    if os.path.exists(path):
        return
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow([
            "timestamp", "port", "address", "source", "waste",
            "syringe_ml", "volume_ml", "ul_min", "ml_s", "counts_s",
            "slope", "cycle", "elapsed_s", "expected_s", "cavitation",
        ])


def append_log(path: str, row: dict):
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=[
            "timestamp", "port", "address", "source", "waste",
            "syringe_ml", "volume_ml", "ul_min", "ml_s", "counts_s",
            "slope", "cycle", "elapsed_s", "expected_s", "cavitation",
        ])
        w.writerow(row)


def print_summary(results: list[dict]):
    if not results:
        return
    print()
    print("Summary")
    print(f"  {'uL/min':>8}  {'mL/s':>8}  {'slope':>5}  {'t_s':>6}  cavitation")
    for r in results:
        print(f"  {r['ul_min']:8.0f}  {r['ml_s']:8.4f}  {r['slope']:5d}  "
              f"{r['elapsed_s']:6.2f}  {r['cavitation']}")
    clean = [r for r in results if r["cavitation"] == "no"]
    if clean:
        fastest = max(clean, key=lambda r: r["counts_s"])
        print()
        print("Fastest scored-clean refill:")
        print(f"  {fastest['ul_min']:.0f} uL/min  "
              f"({fastest['ml_s']:.4f} mL/s, slope {fastest['slope']})")
        print("  Use that as the refill velocity in the software "
              "(currently hard-coded to pump max).")


# ── jog ──────────────────────────────────────────────────────────────────────

JOG_HELP = """
Commands (standing at the bench):
  5000          refill at 5000 uL/min
  0.1 ml/s      refill at 0.1 mL/s
  max           refill at the software max (current, likely cavitates)
  e             empty syringe to waste
  h             home
  s 7           set slope code (1=slow accel, 14=software default, 20=fast)
  v 0.25        set pull volume in mL (default = full syringe)
  p 1           set source port
  w 2           set waste port
  ?             status
  q             quit
"""


def cmd_jog(args):
    syringe_ml = args.syringe
    cpm = counts_per_ml(syringe_ml)
    max_c = software_max_counts()
    print_software_context(syringe_ml)
    if args.waste is None:
        args.waste = args.source

    pump = None
    bench = None
    volume_ml = args.volume if args.volume is not None else syringe_ml

    try:
        pump = connect(args.port, args.address, syringe_ml, args.valves)
        bench = Bench(pump, args.source, args.waste, max_c, args.slope_code)
        if not args.no_home:
            bench.home()
        else:
            bench.halt()
            time.sleep(0.3)
        print(JOG_HELP)
        while True:
            try:
                raw = input("jog> ").strip()
            except EOFError:
                break
            if not raw:
                continue
            cmd, *rest = raw.split()
            low = cmd.lower()

            if low in ("q", "quit", "exit"):
                break
            if low in ("?", "status"):
                print(f"  plunger {pump.volume_ml:.3f} mL  valve {pump.valve_position}  "
                      f"source {bench.source}  waste {bench.waste}  "
                      f"slope {bench.slope}  volume {volume_ml:.3f} mL")
                print(f"  software max  {fmt_rate(max_c, cpm)}")
                continue
            if low in ("e", "empty"):
                bench.empty()
                continue
            if low in ("h", "home"):
                bench.home()
                continue
            if low in ("s", "slope") and rest:
                bench.slope = int(rest[0])
                print(f"  slope = {bench.slope}")
                continue
            if low in ("v", "vol", "volume") and rest:
                volume_ml = min(float(rest[0]), syringe_ml)
                print(f"  pull volume = {volume_ml:.3f} mL")
                continue
            if low in ("p", "source") and rest:
                bench.source = int(rest[0])
                print(f"  source port = {bench.source}")
                continue
            if low in ("w", "waste") and rest:
                bench.waste = int(rest[0])
                print(f"  waste port = {bench.waste}")
                continue
            if low == "help":
                print(JOG_HELP)
                continue

            try:
                if low == "max":
                    vel = max_c
                elif len(rest) >= 1 and rest[0].lower() in ("ml/s", "mls", "mlps"):
                    vel = clamp_counts(round(float(cmd) * cpm), max_c)
                elif low.endswith("ml/s"):
                    vel = clamp_counts(round(float(low.replace("ml/s", "")) * cpm), max_c)
                else:
                    vel = clamp_counts(ulmin_to_counts(float(cmd), cpm), max_c)
            except ValueError:
                print("  unknown command — try `help`")
                continue

            target = max(1, round(volume_ml * cpm))
            if pump.volume_ml > 0.002:
                print("  syringe not empty — emptying first")
                bench.empty()
            print(f"  refill {volume_ml:.3f} mL @ {fmt_rate(vel, cpm)}  slope={bench.slope}")
            print(f"  expected ~{expected_stroke_s(volume_ml, vel, cpm):.2f}s")
            bench.refill(target, vel)
            print("  look for bubbles, then `e` to empty or type another rate")
    except KeyboardInterrupt:
        print("\nCtrl+C — halting pump.")
    finally:
        if bench is not None:
            bench.halt()
        if pump is not None:
            try:
                pump.serial.disconnect()
            except Exception:
                pass


# ── argparse ─────────────────────────────────────────────────────────────────

def add_hw_args(p: argparse.ArgumentParser):
    p.add_argument("--port", required=True, help="COM port (see the `ports` command)")
    p.add_argument("--address", type=int, default=1, help="RS-485 address (C1=1)")
    p.add_argument("--syringe", type=float, default=DEFAULT_SYRINGE_ML,
                   help="Syringe volume in mL (software default 0.5)")
    p.add_argument("--valves", type=int, default=6, help="Valve positions (default 6)")
    p.add_argument("--source", type=int, default=1, help="Port to aspirate from")
    p.add_argument("--waste", type=int, default=None,
                   help="Port to empty to after each trial (default: same as source)")
    p.add_argument("--volume", type=float, default=None,
                   help="Pull volume in mL (default: full syringe)")
    p.add_argument("--slope-code", type=int, default=DEFAULT_SLOPE,
                   help="Default XCalibur slope / acceleration code")
    p.add_argument("--no-home", action="store_true",
                   help="Skip homing (only if the plunger position is already trusted)")


def main():
    parser = argparse.ArgumentParser(
        description="Standalone Cavro refill / cavitation bench tests "
                    "(does not use the Mosaic web app).")
    sub = parser.add_subparsers(dest="cmd")

    sub.add_parser("ports", help="List serial ports")

    p_sweep = sub.add_parser("sweep", help="Step through refill rates and score cavitation")
    add_hw_args(p_sweep)
    p_sweep.add_argument("--ul-min", default=None,
                         help="Comma-separated rates in uL/min, or 'max' for the "
                              "software refill. Default: software-max, then 20k..1k")
    p_sweep.add_argument("--ml-s", default=None,
                         help="Comma-separated rates in mL/s (overrides --ul-min)")
    p_sweep.add_argument("--slope", default=None,
                         help="Comma-separated slope codes to cross with each rate")
    p_sweep.add_argument("--cycles", type=int, default=1,
                         help="Repeats per rate/slope (cavitation can be intermittent)")
    p_sweep.add_argument("--hold", type=float, default=1.0,
                         help="Seconds to wait after refill before scoring")
    p_sweep.add_argument("--no-prompt", action="store_true",
                         help="Do not ask y/n; just run strokes while you watch")
    p_sweep.add_argument("--yes", action="store_true", help="Skip the start confirmation")
    p_sweep.add_argument("--log", default="cavro_refill_log.csv",
                         help="CSV path for scored results")

    p_max = sub.add_parser(
        "max",
        help="One refill at the software max (~29000 uL/min on a 0.5 mL syringe)")
    add_hw_args(p_max)
    p_max.add_argument("--cycles", type=int, default=1,
                       help="Repeats of the max-rate refill")
    p_max.add_argument("--hold", type=float, default=1.0,
                       help="Seconds to wait after refill before scoring")
    p_max.add_argument("--no-prompt", action="store_true",
                       help="Do not ask y/n; just run strokes while you watch")
    p_max.add_argument("--yes", action="store_true", help="Skip the start confirmation")
    p_max.add_argument("--log", default="cavro_refill_log.csv",
                       help="CSV path for scored results")
    p_max.add_argument("--slope", default=None,
                       help="Comma-separated slope codes (default: software 14)")

    p_jog = sub.add_parser("jog", help="Interactive: type a rate, watch, repeat")
    add_hw_args(p_jog)

    args = parser.parse_args()
    if args.cmd == "ports":
        cmd_ports(args)
    elif args.cmd == "sweep":
        cmd_sweep(args)
    elif args.cmd == "max":
        args.ul_min = "max"
        args.ml_s = None
        cmd_sweep(args)
    elif args.cmd == "jog":
        cmd_jog(args)
    else:
        parser.print_help()
        print_software_context(DEFAULT_SYRINGE_ML)


if __name__ == "__main__":
    main()
