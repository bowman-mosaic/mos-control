from typing import List, Tuple, Union, Optional
import time
import math
import logging
import codecs
from ftdi_serial import Serial, SerialReadTimeoutException
from motion import ml_to_counts, counts_to_ml

logger = logging.getLogger(__name__)


class TecanCavro:
    """
    This is a basic driver class for the Tecan Cavro XCalibur series of pumps that uses the OEM binary protocol for
    more robust error handling.

    To use the Cavro pump, create an instance of the TecanCavro class and pass it a serial connection and an address:

    >>> serial = Serial(baudrate=9600)
    >>> cavro = TecanCavro(serial, address=0)

    Multiple TecanCavro instances can share the same serial connection, which can be used for an RS485 network:

    >>> cavro2 = TecanCavro(serial, address=1)

    Most commands have a `batch` flag that will queue the command to be sent as a "multi-command" when executed:

    >>> cavro.move_absolute_ml(0.5, batch=True)
    >>> cavro.dispense_ml(2, from_port=1, to_port=2, batch=True)
    >>> cavro.execute()

    There are also class methods available to broadcast execute messages and home multiple pumps simultaneously:

    >>> TecanCavro.home_all(cavro, cavro2)
    >>> TecanCavro.home_all()  # this will home all pump instances if none are given
    >>> cavro.home(batch=True)
    >>> cavro2.home(batch=True)
    >>> TecanCavro.broadcast_execute(cavro, cavro2)  # this will broadcast to all pumps if none given
    >>> TecanCavro.wait_for_all(cavro, cavro2)  # this will wait for all pumps if none given

    Low level movements can be done with the `move_*` methods:

    >>> cavro.move_absolute_counts(1500)  # this will move the pump halfway down the syringe
    >>> cavro.move_absolute_ml(0.5)  # the `*_ml` methods allow movements in mL, calculated from syringe size
    >>> cavro.move_relative_counts(500)  # relative moves are also supported
    >>> cavro.move_relative_ml(0.1)

    A higher-level `dispense` method is provided to pump from one valve port to another, using multiple strokes:

    >>> cavro.dispense_ml(4.2, from_port=1, to_port=3)

    The current position of the pump can be queried in counts and mL:

    >>> print(cavro.position_counts)
    >>> print(cavro.volume_ml)

    Pump configuration values can be changed (warning: be careful!) this is useful for changing the baud rate:

    >>> cavro.configure_pump(TecanCavro.BAUD_38400)  # most changes require a power cycle of the pump
    """

    BROADCAST_ADDRESS = 46

    # Pump configuration values (these can be used with the configure_pump method)
    VALVE_NONE = 0          # NoVlv
    VALVE_3_PORT = 1        # 3-way
    VALVE_4_PORT = 2        # 4-way
    VALVE_3_PORT_DIST = 3   # 3dist
    VALVE_T = 5             # T-Vlv
    VALVE_6_PORT_DIST = 7   # 6dist
    VALVE_9_PORT_DIST = 8   # 9dist
    VALVE_DUAL_LOOP = 9     # 2loop

    BAUD_9600 = 41
    BAUD_38400 = 47

    instances: List['TecanCavro'] = []

    @classmethod
    def home_all(cls, *pumps: 'TecanCavro'):
        if len(pumps) == 0:
            pumps = cls.instances

        for pump in pumps:
            pump.home(batch=True)
            pump.send_batch()

        cls.broadcast_execute(*pumps)
        cls.wait_for_all(*pumps)

    @classmethod
    def broadcast_execute(cls, *pumps: 'TecanCavro'):
        if len(pumps) == 0:
            pumps = cls.instances

        serial_ports = []
        serial_pumps = []

        # search for different serial ports to broadcast on and send pump batch commands if needed
        for pump in pumps:
            if pump.serial not in serial_ports:
                serial_ports.append(pump.serial)
                serial_pumps.append(pump)

            pump.send_batch()

        # broadcast the execute command using one of the pumps connected to each serial port
        for pump in serial_pumps:
            pump.execute(broadcast=True)

    @classmethod
    def wait_for_all(cls, *pumps: 'TecanCavro'):
        for pump in pumps:
            pump.wait_for_ready()

    @classmethod
    def build_command_request(cls, address: int, sequence: int, data: str, execute: bool=True, retry: bool=False) -> bytes:
        address_bytes = (0x31 + address).to_bytes(1, byteorder='little')
        retry_sequence = 0b1000 if retry else 0
        sequence_bytes = (0b110000 + retry_sequence + sequence).to_bytes(1, byteorder='little')
        exec_bytes = b'R' if execute else b''

        msg = b'\x02' + address_bytes + sequence_bytes + data.encode() + exec_bytes + b'\x03'
        checksum = cls.build_checksum(msg)

        return msg + checksum

    @staticmethod
    def build_checksum(data: bytes) -> bytes:
        checksum = 0

        for byte in data:
            checksum ^= byte

        return checksum.to_bytes(1, byteorder='little')

    @staticmethod
    def dispense_steps(volume_ml: float, syringe_volume_ml: float) -> List[float]:
        volume = volume_ml
        steps = []

        while volume > 0:
            step_volume = min(syringe_volume_ml, volume)
            steps.append(step_volume)
            volume -= step_volume

        return steps

    @staticmethod
    def build_parameters(parameters: List[int]) -> str:
        return ','.join([str(p) for p in parameters])

    @staticmethod
    def ml_min_to_counts(ml_min: Optional[float], velocity_counts: Optional[float], counts_per_ml: float) -> Optional[float]:
        if velocity_counts is not None:
            return velocity_counts

        if ml_min is None:
            return None

        return ml_to_counts(ml_min, counts_per_ml) / 60

    def __init__(self, serial: Serial, address: int, syringe_volume_ml: float=1.0, counts_per_stroke: int=3000, velocity_scale: float=2.0,
                 velocity_counts: int=1400, dead_volume: int=50, total_valve_positions: int=3, distribution_valve: bool=True,
                 slope_code: int=14, wait_timeout: Optional[float]=3600):
        self.logger = logger.getChild(self.__class__.__name__)

        self.serial = serial
        self.address = address

        self._velocity_counts = velocity_counts
        self.velocity_scale = velocity_scale
        self.syringe_volume_ml = syringe_volume_ml
        self.counts_per_stroke = counts_per_stroke
        self.counts_per_ml = counts_per_stroke / syringe_volume_ml
        self.total_valve_positions = total_valve_positions
        self.distribution_valve = distribution_valve
        self.dead_volume = dead_volume
        self._slope_code = slope_code   # slope code for acceleration
        self.wait_timeout = wait_timeout

        self.batch_command = ''
        self.batch_mode = False

        self.sequence = 1

        self.instances.append(self)

    @property
    def plunger_home_speed(self) -> int:
        """ Get the current home speed in counts / s """
        if self.syringe_volume_ml >= 1.0:
            return 0  # full speed
        elif self.syringe_volume_ml >= 250:
            return 1  # half speed
        else:
            return 2  # third speed

    @property
    def position_counts(self) -> int:
        """ Get the current position in counts """
        return int(self.command_request('?'))

    @property
    def volume_ml(self) -> float:
        """ Get the current plunger volume in mL """
        return counts_to_ml(self.position_counts, self.counts_per_ml)

    @property
    def velocity_counts(self) -> int:
        """ Get the default velocity in counts / s """
        return int(self.command_request('?2'))

    @property
    def velocity_ml(self) -> float:
        """ Get the default velocity in mL / s """
        return counts_to_ml(self.velocity_counts, self.counts_per_ml)

    @property
    def valve_position(self) -> int:
        """ Get the current valve position """
        return int(self.command_request('?6'))

    @property
    def start_speed(self) -> int:
        """ Get the current start speed """
        return int(self.command_request('?1'))

    @property
    def cutoff_speed(self) -> int:
        """ Get the current cutoff speed """
        return int(self.command_request('?3'))

    @property
    def pump_configuration(self) -> str:
        """ Get the current pump configuration """
        return self.command_request('?76').decode()

    @property
    def valve_type(self) -> str:
        """
        Get information about the valve type

        :returns: A tuple containing (baud, can, com_type, valve_type, mem_mode)
        """
        baud, can, com_type, valve_type, mem_mode = self.pump_configuration.split('|')
        return valve_type

    def command_request_raw(self, command: str, parameters: List[int]=[], retry: bool=False, execute: bool=True,
                            retries: int=5, timeout: float=0.5, broadcast: bool=False, log: bool=True) -> Tuple[int, bytes]:
        """
        Perform a low-level command request (note: the simpler command_request method should be used when possible)

        :param command: The command str to send to the pump (without parameters)
        :param parameters: [Optional] A list of parameters for the command
        :param retry: [Optional] Retries the command on failure if True, defaults to False
        :param execute: [Optional] Append an execute command if True (only if needed), defaults to True
        :param retries: [Optional] Number of times to retry command on failure, defaults to 5
        :param timeout: [Optional] Length of time to wait before retrying command, defaults to 0.5s
        :param broadcast: [Optional] Broadcast this command to all connected pumps if True, defaults to False
        :param log: [Optional] Log command info if True, defaults to True
        :returns: A (status: int, response: bytes) tuple with the command response and status
        """
        if len(command) > 255:
            raise TecanCavroCommandTooLongError(f'Command too long (max 255 chars): {command}')

        self.sequence = (self.sequence + 1) % 7 + 1  # increment sequence number, keeping it between 1 and 7 (inclusive)
        address = self.address if not broadcast else self.BROADCAST_ADDRESS
        command_str = command + self.build_parameters(parameters)
        command_message = self.build_command_request(address, self.sequence, command_str, retry=retry, execute=execute and command != 'R')

        if log:
            self.logger.debug(f'Sending command: {command_str} [{codecs.encode(command_message, "hex_codec")}]')

        try:
            # broadcast messages don't have a response, so we can skip a bunch of stuff
            if broadcast:
                # write the same message 3 times to hopefully handle any transmission errors
                self.serial.write(command_message * 3)
                return 0, b''

            # send the request and wait for a response ending with 0x03
            response = self.serial.request(command_message, line_ending=b'\x03', timeout=timeout)

            try:
                response_stripped = response[response.index(b'\x02'):]  # the actual response starts at byte 0x02
            except ValueError:
                raise TecanCavroInvalidResponseError(f'Response invalid: {response}')

            response_checksum = self.serial.read(1)  # read the response checksum, which is the byte following 0x03
            checksum = self.build_checksum(response_stripped + b'\x03')  # calculate a new checksum for the response

            # raise an error if the calculated checksum doesn't match the response
            if response_checksum != checksum:
                raise TecanCavroResponseChecksumError(f'Response checksum does not match: {response_checksum} != {checksum}')

            status = response_stripped[2]
            data = response_stripped[3:]

            if log:
                self.logger.debug(f'Received response: status={status}, data={data} [{response}]')

            # raise a device error with a message based on the status code if it's not success
            if status & 0xf != TecanCavroDeviceError.SUCCESS:
                raise TecanCavroDeviceError.build_error(status & 0xf, data)

            return status, data

        except (SerialReadTimeoutException, TecanCavroResponseChecksumError, TecanCavroInvalidResponseError) as err:
            if retries <= 0:
                raise err

            return self.command_request_raw(command, parameters, retry=True, execute=execute, retries=retries - 1)

    def command_request(self, command: str, parameters: List[int]=[], retry: bool=False, batch: bool=False,
                        broadcast: bool=False, check_busy: bool=False) -> bytes:
        """
        Perform a command request

        :param command: The command str to send to the pump (without parameters)
        :param parameters: [Optional] A list of parameters for the command
        :param retry: [Optional] Retries the command on failure if True, defaults to False
        :param broadcast: [Optional] Broadcast this command to all connected pumps if True, defaults to False
        :param check_busy: [Optional] Check to see if pump is busy before executing command if True, defaults to False
        :returns: A (status: int, response: bytes) tuple with the command response and status
        """
        if batch:
            self.batch_command += command + self.build_parameters(parameters)
            return b''
        else:
            if check_busy and not self.ready():
                raise TecanCavroDeviceBusyError(f'Cannot execute command, device busy (address={self.address})')

            status, data = self.command_request_raw(command, parameters, retry=retry, broadcast=broadcast)
            return data

    def start_batch(self):
        """ Start a new set of commands to send in a single batch """
        self.batch_command = ''
        self.batch_mode = True

    def clear_batch(self):
        """ Clear the current batch commands """
        self.batch_command = ''
        self.batch_mode = False

    def send_batch(self, execute: bool=False):
        """ Send all batched commands """
        if self.batch_command == '':
            return

        self.command_request_raw(self.batch_command, execute=execute)
        self.batch_command = ''

    def configure(self, batch: bool=False):
        """
        Send pump configuration to pump

        :param batch: [Optional] Send command in a batch if True, defaults to False
        """
        self.command_request('k', [self.dead_volume], batch=batch)
        self.command_request('V', [self._velocity_counts], batch=batch)
        self.command_request('L', [self._slope_code], batch=batch)

    def execute(self, broadcast: bool = False, wait: bool = True, execute: bool = True):
        """
        Send an execute command

        :param broadcast: [Optional] Broadcast command to all connected pumps if True, defaults to False
        :param wait: [Optional] Wait for pump to be ready after sending command if True, defaults to True
        :param execute: [Optional] Append an execute command if True, defaults to True
        """
        if self.batch_command != '':
            self.send_batch()

        if self.batch_mode:
            self.batch_mode = False

        # having an execute flag cleans up code that would otherwise need an if statement to send the batch without executing
        if execute:
            self.command_request('R', broadcast=broadcast)

            if wait and not broadcast:
                self.wait_for_ready()

    def configure_pump(self, value: int):
        """
        Send raw pump configuration value

        :param value: Pump configuration value (available in datasheet)
        """
        self.command_request_raw('U', [value])

    def home(self, home_plunger: bool=True, home_valves: bool=True, wait: bool=True, batch: bool=False,
             valve_position: str='Z'):
        """
        Home the pump and/or valves

        :param home_plunger: [Optional] Home the plunger if True, defaults to True
        :param home_valves: [Optional] Home the valves if True, defaults to True
        :param wait: [Optional] Wait for homing to finish if True, defaults to True
        :param batch: [Optional] Send as a batch command if True, defaults to False
        :param valve_position: [Optional] Valve position to home to, defaults to "Z" (clockwise)
        """
        # valve_position: used 'Z' for CW and 'Y' for CCW
        self.configure(batch=batch)

        if home_plunger and home_valves:
            self.command_request(valve_position, [self.plunger_home_speed], batch=batch)
        elif home_plunger:
            self.command_request('W', [self.plunger_home_speed], batch=batch)
        elif home_valves:
            self.command_request('w', batch=batch)

        if wait and not batch:
            self.wait_for_ready()

    def halt(self) -> bytes:
        """ Send a halt command """
        return self.command_request('T')

    def status(self, log: bool=True) -> Tuple[bool, int]:
        """
        Get the pump status

        :param log: [Optional] Log the status command if True, defaults to True
        :return: A (ready: bool, status_code: int) tuple
        """
        status, data = self.command_request_raw('Q', log=log)
        ready = (status & 1 << 5) > 0  # pump is busy if bit 5 is 0
        status_code = status & 0xf  # strip off busy bit

        return ready, status_code

    def check_status(self):
        """ Check the current pump status and raise an error if it is not SUCCESS """
        ready, status = self.status()

        if status != TecanCavroDeviceError.SUCCESS:
            raise TecanCavroDeviceError.build_error(status)

    def ready(self, log: bool=True) -> bool:
        """
        Get the ready state of the pump

        :param log: [Optional] Log this command if True, defaults to True
        :return: True if pump ready for more commands
        """
        ready, status_code = self.status(log=log)
        return ready

    def wait_for_ready(self, poll_interval: float=0.1):
        """
        Wait for the pump to be ready

        :param poll_interval: [Optional] Status polling interval, defaults to 0.1s
        """
        start_time = time.time()

        while not self.ready(log=False):
            if self.wait_timeout is not None and (time.time() - start_time) > self.wait_timeout:
                raise TecanCavroReadyTimeout(f'Timeout while waiting for Cavro (address: {self.address})')

            time.sleep(poll_interval)

    def speed(self, velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None):
        """
        Change the default velocity

        :param velocity_ml: [Optional] The new velocity in mL / s
        :param velocity_counts: [Optional] The new velocity in counts / s
        """
        velocity = self.ml_min_to_counts(velocity_ml, velocity_counts, self.counts_per_ml)

        if velocity is None:
            raise TecanCavroInvalidSpeedError(f'One of `velocity_ml` or `velocity_counts` required')

        try:
            self.command_request('V', [int(velocity * self.velocity_scale)])
        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid velocity: {velocity_ml or velocity_counts}')
            else:
                raise err

        self._velocity_counts = velocity

    def loop_start(self):
        """ Start a batch command loop """
        self.command_request('g', batch=True)
        self.batch_mode = True

    def loop_end(self, iteration_count: int):
        """ End a batch command loop """
        self.command_request('G', [iteration_count], batch=True)

    def move_absolute_counts(self, position_counts: Union[float, int], velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None, slope_code: Optional[int]=None, wait: bool=True):
        """
        Move the pump to the given position in counts

        :param position_counts: Position to move pump to, in counts
        :param velocity_ml: [Optional] Velocity to use during move in mL / s
        :param velocity_counts: [Optional] Velocity to use during move in counts / s
        :param slope_code: [Optional] Slope code to use to calculate acceleration
        :param wait: [Optional] Wait for move to complete if True, defaults to True
        """
        velocity = self.ml_min_to_counts(velocity_ml, velocity_counts, self.counts_per_ml)

        try:
            if velocity is not None:
                self.command_request('V', [int(velocity * self.velocity_scale)], batch=True)
        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid velocity: {velocity_ml or velocity_counts}')
            else:
                raise err

        try:
            if slope_code is not None:
                self.command_request('L', [slope_code], batch=True)
        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid slope code: {slope_code}')
            else:
                raise err

        try:
            self.command_request('A', [int(position_counts)], batch=True)

            if velocity is not None:
                self.command_request('V', [int(self._velocity_counts * self.velocity_scale)], batch=True)

            if not self.batch_mode:
                self.execute(wait=wait)

        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid position: {position_counts}')
            else:
                raise err

        # resetting the slope
        try:
            if slope_code is not None:
                self.command_request('L', [self._slope_code], batch=True)
        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid slope code: {slope_code}')
            else:
                raise err

    def move_relative_counts(self, delta_counts: Union[float, int], velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None, slope_code: Optional[int]=None, wait: bool=True):
        """
        Move the pump relatively using the given delta in counts

        :param delta_counts: Distance to move pump, in counts
        :param velocity_ml: [Optional] Velocity to use during move in mL / s
        :param velocity_counts: [Optional] Velocity to use during move in counts / s
        :param slope_code: [Optional] Slope code to use to calculate acceleration
        :param wait: [Optional] Wait for move to complete if True, defaults to True
        """
        velocity = self.ml_min_to_counts(velocity_ml, velocity_counts, self.counts_per_ml)

        try:
            if velocity is not None:
                self.command_request('V', [int(velocity * self.velocity_scale)], batch=True)
        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid velocity: {velocity_ml or velocity_counts}')
            else:
                raise err

        try:
            if slope_code is not None:
                self.command_request('L', [slope_code], batch=True)
        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid slope code: {slope_code}')
            else:
                raise err

        try:
            command = 'P' if delta_counts > 0 else 'D'
            self.command_request(command, [abs(int(delta_counts))], batch=True, check_busy=True)

            if velocity is not None:
                self.command_request('V', [int(self._velocity_counts * self.velocity_scale)], batch=True)

            if not self.batch_mode:
                self.execute(wait=wait)

        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid position: {delta_counts}')
            else:
                raise err

        # resetting the slope
        try:
            if slope_code is not None:
                self.command_request('L', [self._slope_code], batch=True)
        except TecanCavroDeviceError as err:
            if err.status == TecanCavroDeviceError.INVALID_OPERAND:
                raise TecanCavroInvalidPositionError(f'Invalid slope code: {slope_code}')
            else:
                raise err

    def move_absolute_ml(self, position_ml: float, velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None, slope_code: Optional[int]=None, wait: bool=True):
        """
        Move the pump to the given position in mL

        :param position_ml: Position to move pump to, in mL
        :param velocity_ml: [Optional] Velocity to use during move in mL / s
        :param velocity_counts: [Optional] Velocity to use during move in counts / s
        :param slope_code: [Optional] Slope code to use to calculate acceleration
        :param wait: [Optional] Wait for move to complete if True, defaults to True
        """
        self.move_absolute_counts(ml_to_counts(position_ml, self.counts_per_ml), velocity_ml=velocity_ml, velocity_counts=velocity_counts, slope_code=slope_code, wait=wait)

    def move_relative_ml(self, delta_ml: float, velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None, slope_code: Optional[int]=None, wait: bool=True):
        """
        Move the pump relatively using the given delta in mL

        :param delta_ml: Distance to move pump, in mL
        :param velocity_ml: [Optional] Velocity to use during move in mL / s
        :param velocity_counts: [Optional] Velocity to use during move in counts / s
        :param slope_code: [Optional] Slope code to use to calculate acceleration
        :param wait: [Optional] Wait for move to complete if True, defaults to True
        """
        self.move_relative_counts(ml_to_counts(delta_ml, self.counts_per_ml), velocity_ml=velocity_ml, velocity_counts=velocity_counts, slope_code=slope_code, wait=wait)

    def switch_valve(self, position: int, wait: bool = True):
        """
        Switch the valve to the given position, starting at 1

        :param position: The valve position to switch to, numbered starting at 1
        :param wait: [Optional] Wait for switching to complete if True, defaults to True
        """
        if position < 1 or position > self.total_valve_positions:
            raise TecanCavroInvalidValvePositionError(f'Invalid valve position: {position}')

        if self.total_valve_positions == 2:
            if position == 1:
                command_string = 'I'
            elif position == 2:
                command_string = 'O'
        else:
            command_string = f'I{position}'

        self.command_request(command_string, batch=self.batch_mode, check_busy=True)

        if not self.batch_mode:
            time.sleep(0.1)  # wait 100ms before polling to avoid intermittent timing errors
            if wait:
                self.wait_for_ready()

    def dispense_ml(self, volume_ml: float, from_port: int, to_port: int, velocity_ml: Optional[float]=None,
                    velocity_counts: Optional[int]=None, dispense_velocity_ml: Optional[float]=None,
                    dispense_velocity_counts: Optional[int]=None, wait: bool=True, execute: bool=True):
        """
        Dispense the given volume from the ``from_port`` to the ``to_port`` on the valve, will build a loop with
        multiple cycles if needed

        :param volume_ml: The volume to dispense in mL
        :param from_port: The port to dispense from
        :param to_port: The port to dispense to
        :param velocity_ml: [Optional] Velocity to use when pulling from ``from_port``, in mL
        :param velocity_counts: [Optional] Velocity to use when pulling from ``from_port``, in counts
        :param dispense_velocity_ml: [Optional] Velocity to use when pushing to ``to_port``, in mL
        :param dispense_velocity_counts: [Optional] Velocity to use when pushing to ``to_port``, in counts
        :param wait: [Optional] Wait for dispense to complete if True, defaults to True
        :param execute: [Optional] Execute this command immediately if True, defaults to True
        """
        if dispense_velocity_ml is None:
            dispense_velocity_ml = velocity_ml
        if dispense_velocity_counts is None:
            dispense_velocity_counts = velocity_counts

        iterations = math.floor(volume_ml / self.syringe_volume_ml)
        remaining_volume = volume_ml % self.syringe_volume_ml
        batch_mode = self.batch_mode

        self.start_batch()

        # if we are dispensing more than the syringe volume create a loop of full syringe movements
        if iterations > 0:
            self.loop_start()
            self.switch_valve(from_port)
            self.move_absolute_ml(self.syringe_volume_ml, velocity_ml=velocity_ml, velocity_counts=velocity_counts)
            self.switch_valve(to_port)
            self.move_absolute_ml(0, velocity_ml=dispense_velocity_ml, velocity_counts=dispense_velocity_counts)
            self.loop_end(iterations)

        # dispense any remaining volume after looping
        self.switch_valve(from_port)
        self.move_absolute_ml(remaining_volume, velocity_ml=velocity_ml, velocity_counts=velocity_counts)
        self.switch_valve(to_port)
        self.move_absolute_ml(0, velocity_ml=dispense_velocity_ml, velocity_counts=dispense_velocity_counts)

        if not batch_mode:
            self.execute(wait=wait, execute=execute)

        self.batch_mode = batch_mode

    def pump_ml(self, volume_ml: float, port: int, velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None, wait: bool=True):
        """
        Pump the given volume from ``port``

        :param volume_ml: Volume to pump in mL
        :param port: Port to pump from
        :param velocity_ml: [Optional] Velocity to use during pump, in mL
        :param velocity_counts: [Optional] Velocity to use during pump, in counts
        """
        if abs(volume_ml) > self.syringe_volume_ml:
            raise TecanCavroInvalidPositionError(f'Error pumping {volume_ml} mL through {port} port, volume larger than syringe volume')

        self.switch_valve(port)
        self.move_relative_ml(volume_ml, wait=wait, velocity_ml=velocity_ml, velocity_counts=velocity_counts)

    def pump_to_ml(self, volume_ml: float, port: int, velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None, wait: bool=True):
        """
        Pump the given volume to ``port``

        :param volume_ml: Volume to pump in mL
        :param port: Port to pump to
        :param velocity_ml: [Optional] Velocity to use during pump, in mL
        :param velocity_counts: [Optional] Velocity to use during pump, in counts
        """
        self.pump_ml(-volume_ml, port, velocity_ml=velocity_ml, velocity_counts=velocity_counts)  # pumping "to" a port moves in a negative direction

    def pump_from_ml(self, volume_ml: float, port: int, velocity_ml: Optional[float]=None, velocity_counts: Optional[int]=None, wait: bool=True):
        """
        Pump the given volume from ``port``

        :param volume_ml: Volume to pump in mL
        :param port: Port to pump from
        :param velocity_ml: [Optional] Velocity to use during pump, in mL
        :param velocity_counts: [Optional] Velocity to use during pump, in counts
        """
        self.pump_ml(volume_ml, port, velocity_ml=velocity_ml, velocity_counts=velocity_counts)  # pumping "from" a port moves in a positive direction

#
# Errors
#

class TecanCavroError(Exception):
    pass


class TecanCavroInvalidSpeedError(TecanCavroError):
    pass


class TecanCavroDeviceBusyError(TecanCavroError):
    pass


class TecanCavroResponseError(TecanCavroError):
    pass


class TecanCavroResponseChecksumError(TecanCavroResponseError):
    pass


class TecanCavroInvalidResponseError(TecanCavroResponseError):
    pass


class TecanCavroReadyTimeout(TecanCavroError):
    pass


class TecanCavroInvalidPositionError(TecanCavroError):
    pass


class TecanCavroInvalidValvePositionError(TecanCavroError):
    pass


class TecanCavroCommandTooLongError(TecanCavroError):
    pass


class TecanCavroDeviceError(TecanCavroError):
    SUCCESS = 0
    INITIALIZATION = 1
    INVALID_COMMAND = 2
    INVALID_OPERAND = 3
    INVALID_COMMAND_SEQUENCE = 4
    EEPROM_FAILURE = 6
    DEVICE_NOT_INITIALIZED = 7
    PLUNGER_OVERLOAD = 9
    VALVE_OVERLOAD = 10
    PLUNGER_MOVE_NOT_ALLOWED = 11
    COMMAND_OVERFLOW = 15

    messages = {
        INITIALIZATION: 'Initialization',
        INVALID_COMMAND: 'Invalid Command',
        INVALID_OPERAND: 'Invalid Operand',
        INVALID_COMMAND_SEQUENCE: 'Invalid Command Sequence',
        EEPROM_FAILURE: 'EEPROM Failure',
        DEVICE_NOT_INITIALIZED: 'Pump Not Homed',
        PLUNGER_OVERLOAD: 'Plunger Overload',
        VALVE_OVERLOAD: 'Valve Overload',
        PLUNGER_MOVE_NOT_ALLOWED: 'Plunger Move Not Allowed',
        COMMAND_OVERFLOW: 'Command Overflow'
    }

    @classmethod
    def build_error(cls, status: int, data: bytes=b'') -> 'TecanCavroDeviceError':
        message = cls.messages.get(status, 'Unknown Error')
        return TecanCavroDeviceError(f'Tecan Cavro Error {status}: {message}', status, data)

    def __init__(self, message: str, status: int=-1, data: bytes=b''):
        Exception.__init__(self, message)
        self.status = status
        self.data = data
