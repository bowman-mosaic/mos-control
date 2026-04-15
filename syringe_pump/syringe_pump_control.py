#!/usr/bin/env python3
"""
Backend: Harvard Apparatus Model 22 Syringe Pump Control (model 980532).
Use this module from a separate frontend (GUI or script). It exposes HarvardPump
and list_serial_ports() for connection and control.
Protocol: MMD (diameter), MLM/ULM/MLH/ULH (rate), MLT (volume), RUN/REV/STP, CLV/CLT.
Serial: 9600 8N2 (2 stop bits).
"""
# For address 0: False = send "RUN" only (manual tutorial); True = send "0 RUN"
SEND_ADDRESS_FOR_ZERO = False

import serial
import serial.tools.list_ports
import time


def list_serial_ports():
    """
    Return list of available serial ports for use by frontend (e.g. dropdown).
    Each item is a dict: {"device": "COM3", "description": "USB Serial"}.
    """
    ports = list(serial.tools.list_ports.comports())
    return [{"device": p.device, "description": p.description or p.device} for p in ports]


class HarvardPump:
    """
    Control class for Harvard Apparatus Model 22 syringe pump (980532).
    Uses Model 22 RS-232 protocol (Publication 5385-003, Appendix C).
    """
    
    def __init__(self, port=None, baudrate=9600, address=1, timeout=1):
        """
        Initialize pump connection.
        
        Args:
            port: Serial port (e.g., 'COM3' on Windows, '/dev/ttyUSB0' on Linux)
                  If None, will attempt to auto-detect
            baudrate: Communication speed (9600; must match pump SET+START/STOP)
            address: Pump address 0-9 for daisy chain (single digit)
            timeout: Serial timeout in seconds
        """
        self.address = int(address) if address is not None else 0
        self.baudrate = baudrate
        self.timeout = timeout
        self._direction = 'INF'  # INF = RUN, WDR = REV
        
        if port is None:
            port = self._find_pump()
        
        try:
            self.serial = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_TWO,  # Model 22: 2 stop bits (Appendix D)
                timeout=timeout
            )
            print(f"Connected to pump on {port}")
            time.sleep(0.5)  # Give pump time to initialize
            self._clear_buffer()
        except serial.SerialException as e:
            raise RuntimeError(f"Error connecting to pump: {e}") from e
    
    def _find_pump(self):
        """Auto-detect pump serial port (used when port is None and running as script)."""
        ports = list(serial.tools.list_ports.comports())
        if not ports:
            raise RuntimeError("No serial ports found.")
        if len(ports) == 1:
            return ports[0].device
        # Multiple ports: require explicit port from frontend; when run as script, use first
        return ports[0].device
    
    def _clear_buffer(self):
        """Clear any residual data in the serial buffer."""
        self.serial.reset_input_buffer()
        self.serial.reset_output_buffer()
    
    def _send_command(self, command):
        """
        Send command to pump.
        
        Args:
            command: Command string (without address prefix or terminator)
        
        Returns:
            Response from pump
        """
        # Manual: for address 0 can omit address ("RUN" + CR). Else "n command" + CR.
        if not SEND_ADDRESS_FOR_ZERO and self.address == 0:
            full_command = (command + "\r") if command else "\r"
        else:
            full_command = f"{self.address}\r" if not command else f"{self.address} {command}\r"
        print(f"Full command: {full_command!r}")
        
        self._clear_buffer()
        self.serial.write(full_command.encode('ascii'))
        self.serial.flush()
        time.sleep(0.5)
        buf = bytearray()
        deadline = time.monotonic() + 2.0
        while time.monotonic() < deadline:
            n = self.serial.in_waiting
            if n:
                buf.extend(self.serial.read(n))
                if buf and buf[-1] in (ord(":"), ord(">"), ord("<"), ord("*")):
                    break
            time.sleep(0.05)
        response = buf.decode('ascii', errors='replace').strip()
        print(f"Response: {response!r}  |  hex: {buf.hex()}")
        return response
    
    def get_status(self):
        """Get current pump status."""
        response = self._send_command("")
        return response
    
    def set_diameter(self, diameter_mm):
        """
        Set syringe diameter in mm (Model 22: MMD). Rate is set to 0 after this.
        
        Args:
            diameter_mm: Syringe inner diameter (e.g., 14.5 for BD 10mL)
        """
        val = f"{float(diameter_mm):.2f}".rstrip("0").rstrip(".")
        command = f"MMD {val}"
        response = self._send_command(command)
        print(f"Syringe diameter set to {diameter_mm} mm")
        return response
    
    def set_rate(self, rate, units="ML/MIN"):
        """
        Set flow rate (Model 22: MLM/ULM/MLH/ULH).
        
        Args:
            rate: Flow rate value
            units: ML/MIN, ML/HR, UL/MIN, UL/HR
        """
        u = units.upper().replace("/", "")
        if u in ("MLMIN", "ML/MIN"):
            command = f"MLM {rate}"
        elif u in ("MLHR", "ML/HR"):
            command = f"MLH {rate}"
        elif u in ("ULMIN", "UL/MIN"):
            command = f"ULM {rate}"
        elif u in ("ULHR", "UL/HR"):
            command = f"ULH {rate}"
        else:
            raise ValueError(f"Unknown units: {units}")
        response = self._send_command(command)
        print(f"Rate set to {rate} {units}")
        return response
    
    def set_volume(self, volume, units="ML"):
        """
        Set target volume in mL (Model 22: MLT). Pump stops when reached.
        """
        command = f"MLT {volume}"
        response = self._send_command(command)
        print(f"Target volume set to {volume} mL")
        return response
    
    def set_direction(self, direction):
        """
        Set direction for next run(): INF = infuse (RUN), WDR = withdraw (REV).
        Model 22 has no DIR command; RUN/REV are sent when starting.
        """
        if direction.upper() not in ['INF', 'WDR']:
            raise ValueError("Direction must be 'INF' or 'WDR'")
        self._direction = direction.upper()
        print(f"Direction set to {self._direction}")
    
    def run(self):
        """Start the pump (RUN = infuse, REV = withdraw per set_direction)."""
        cmd = "REV" if self._direction == "WDR" else "RUN"
        response = self._send_command(cmd)
        print("Pump started")
        return response
    
    def stop(self):
        """Stop the pump."""
        response = self._send_command("STP")
        print("Pump stopped")
        return response
    
    def pause(self):
        """Pause the pump (can be resumed)."""
        response = self._send_command("STP")
        print("Pump paused")
        return response
    
    def clear_volume(self):
        """Clear volume accumulator (CLV). Use before volume dispense."""
        response = self._send_command("CLV")
        print("Volume accumulator cleared")
        return response
    
    def clear_target(self):
        """Clear target volume (CLT). Use for continuous run (no volume limit)."""
        response = self._send_command("CLT")
        print("Target cleared (continuous mode)")
        return response
    
    def is_running(self):
        """
        Check if pump is currently running.
        
        Returns:
            True if running, False if stopped
        """
        status = self.get_status()
        # Status typically contains info about running state
        # This may need adjustment based on your specific pump model
        return ">" in status or ":" in status  # Running pumps often show these
    
    def close(self):
        """Close serial connection."""
        if self.serial and self.serial.is_open:
            self.serial.close()
            print("Connection closed")


