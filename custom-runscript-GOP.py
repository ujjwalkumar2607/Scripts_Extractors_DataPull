#!/usr/bin/env python
'''
###########################################################################
#                         CubeX Runscript                                 #
#       Python script which executes the PCR or Isothermal protocols      #
#                                                                         #
#                                                                         #
###########################################################################
'''

#Runscript Metadata
__copyright__ = "Genomadix"
__license__ = "?"
__version__ = "1.0.0"
__maintainer__ = "?"
__email__ = "?"
__status__ = "Beta" #fix these up

# things not to forget to do
# cleaning up debugging code, any commented code that is not necessary, search ''', """, #
# comments that remain explain code
# need to think about pscs, codes, and texts output
# we need to figure out below 50 deg C temps
# figure out temp offset
# logic for handling below 50 temps or i guess below cal file lowest temp, likely 65C
# what limits do we input for ramp rate..... how can this be determined....
# add progress logic to change based on which steps like incubation or iso thermal or pcr are included

#Imports
import logging
from pickle import TRUE
import threading
import time
import csv
import queue
import typing as T
import traceback
import json
import numpy as np
import math
import re
from spylib3.thermal_predictors import ThermalPredictionException
from spylib3.exceptions import SleeveDeltaError
from concurrent.futures import TimeoutError

##AMPLIFICATION_METHOD##
##NUMBER_OF_CYCLES##
##CYCLIC_DENATURE_TEMP##
##CYCLIC_DENATURE_RAMP_RATE##
##CYCLIC_DENATURE_DWELL_TIME##
##CYCLIC_ANNEAL_TEMP##
##CYCLIC_ANNEAL_RAMP_RATE## 
##CYCLIC_ANNEAL_DWELL_TIME## 


##INCUBATION_1_TARGET_TEMP## 
##INCUBATION_1_RAMP_RATE##
##INCUBATION_1_DWELL_TIME##

##INCUBATION_2_TARGET_TEMP##
##INCUBATION_2_RAMP_RATE##
##INCUBATION_2_DWELL_TIME##

##INCUBATION_3_TARGET_TEMP##
##INCUBATION_3_RAMP_RATE##
##INCUBATION_3_DWELL_TIME##

##IMAGE_CAPTURE_PERIOD##
##ISOTHERMAL_AMPLIFICATION_TEMP##
##ISOTHERMAL_AMPLIFICATION_TIME##

##EXPERIMENT_ID##
##RUN_DATE##
##OPERATOR_NAME##
##NOTES##


# Start of Globally-used runscript constants
ROOT_MODULE_LOGGER = logging.getLogger("__runscript__")
ROOT_MODULE_LOGGER.propagate = False
ROOT_MODULE_LOGGER.setLevel(logging.NOTSET)
ROOT_MODULE_LOGGER.addHandler(logging.NullHandler())
NUMBER_OF_LOOPS = 0

# ?need to double check this works on a 1.4.31 fw prod cube
__PRODUCT_WHITELIST__ = r"\S*"
LED_YELLOW_FLASH = ('s', 20, 'y')
LED_GREEN_SOLID = ('h', 50, 'g')
LED_RED_SOLID = ('h', 50, 'r')
LED_YELLOW_SOLID = ('h', 50, 'y')

component_manifest = {
    "ManifestAPIVersion": ('', '3.0'),
    "ThermalInterface": ("RampInterface", "4.3"),
    "OpticalInterface": ("OpticalInterface", "4.0"),
    "AnalysisManager": ("AnalysisManager", "3.0"),
}
# End of Globally-used runscript constants

'''?RH create/call, a function or something that puts
a text file into tarball. filename....
{EXPERIMENT_ID}_metadata.log
it will contain, EXPERIMENT_ID,
RUN_DATE, OPERATOR, NOTES in
JSON format key value pairing'''

def generate_log(exp_id, run_date, operator, notes):
    json_string = {
        'EXPERIMENT ID': exp_id,
        'RUN DATE': run_date,
        'OPERATOR': operator,
        'NOTES': notes
    }
    pass

def calc_number_loops(run_type):
    global NUMBER_OF_LOOPS
    if run_type.lower() == "pcr":
        NUMBER_OF_LOOPS = NUMBER_OF_CYCLES
    elif run_type.lower() == "isothermal":
        if ((ISOTHERMAL_AMPLIFICATION_TIME % IMAGE_CAPTURE_PERIOD) == 0):
            NUMBER_OF_LOOPS = int(ISOTHERMAL_AMPLIFICATION_TIME/IMAGE_CAPTURE_PERIOD)
        else:
            NUMBER_OF_LOOPS = math.floor(ISOTHERMAL_AMPLIFICATION_TIME/IMAGE_CAPTURE_PERIOD)
    else:
        raise

class TimeoutManager:
    """Handles timeouts.
    Parameters
    ----------
    psc_callback : Callable
        Callback to trigger PSC. Arguments are (msg: str, code: int,
        payload: dict, led_override)
    cancel_callback : Callable[[], None]
        Callback to trigger a cancel. This will shutdown the runscript.
    strict : bool
        Default False. If True, will raise exceptions if registering an existing
        timeout.
    logger : Optional[RunLogger]
        Optional logging interface.

    Attributes
    ----------
    timeouts : dict
    """

    def __init__(self, psc_callback, cancel_callback, strict=False, logger=None):
        self.psc_callback = psc_callback
        self.cancel_callback = cancel_callback
        self.strict = strict
        self.logger = logger

        self.timeouts = {}

    def register(self, name: str, length: int, error_message: str, error_code: int, led_state=None, auto_arm=False):
        """ Register a timeout.
        Parameters
        ----------
        name : str
            Name of the callback. Used to identify callback.
        length : int
            Length of timeout in seconds.
        error_message : str
            The error message to pass into psc_callback
        error_code : int
            The error code to pass into psc_callback
        led_state
            The led state to use. Defaults to LED_RED_SOLID
        auto_arm : bool
            Default False. If True calls arm on timer right away.
        """
        if name in self.timeouts:
            if self.strict:
                raise ValueError(f"Timeout {name} already exists")
            else:
                return

        timer = threading.Timer(length, self.fire, args=(
            error_message, error_code, led_state))
        timer.daemon = True
        self.timeouts[name] = timer

        if auto_arm:
            self.arm(name)

    def arm(self, name):
        # Arm a timeout. Starts the timer
        if name not in self.timeouts:
            if self.strict:
                raise ValueError(f"Timeout {name} does not exist")
            else:
                return
        if self.logger is not None:
            try:
                self.logger.file_append(
                    "timeouts.log", f"arming {name} at {time.time()}\n")
            except:
                pass
        self.timeouts[name].start()

    def disarm(self, name):
        # Disarm a timeout. Cancels the timer
        if name not in self.timeouts:
            if self.strict:
                raise ValueError(f"Timeout {name} does not exist")
            else:
                return
        if self.logger is not None:
            try:
                self.logger.file_append(
                    "timeouts.log", f"disarming {name} at {time.time()}\n")
            except:
                pass
        self.timeouts[name].cancel()

    def fire(self, error_message, error_code, led_state):
        try:
            self.cancel_callback()
        except:
            pass

        if led_state is None:
            led_state = LED_RED_SOLID

        try:
            self.logger.file_append("timeout")
        except:
            pass

        self.psc_callback(error_message, error_code,
                          {}, led_override=led_state)


# Back up logger when no logger is passed
class BackupRunLogger:
    def _backup_callable(self, *args, **kwargs):
        pass

    def __getattr__(self, name):
        return self._backup_callable


'''
Parameter issue: something wrong with the parameters (3021)
KeyboardInterrupt: why does this even exist we can keep it i suppose but no PSC relating
ThermalPredictionException: we can keep, max oops related
TimeoutError: which timeout error is this watchdog  or the ones by timeout manager
SleeveDeltaError: we can keep
Exception as ex: catchall maybe we need a better text output like "Unexpected Error" which is what a catchall is right
def thermal_timeout(self, code=3055)
def timeout(self)
("Configuration Error", 3051, LED_RED_SOLID)
HT PSC fail
what other PSCs can we have what triggered system errors what if something fails on application FW side like spyscript etc, where does that trigger
'''
def error_log(string):
    with open("runscript_error.log", "a") as f:
        f.write(string + "\n")

def light_tarball_def(fn):
    if fn.endswith('.png'):
        return False
    return True

class StateMachine:
    def __init__(
        self,
        run_logger=None
    ):

        if run_logger is None:
            self.run_logger = BackupRunLogger()
        else:
            self.run_logger = run_logger

        self.data: T.List[T.Tuple[float]] = []
        self.psc_log_msgs: T.List[str] = []
        self.done_data_collection = False

        calc_number_loops(AMPLIFICATION_METHOD)
        self.number_of_loops = NUMBER_OF_LOOPS

    def collect_data(self, data):
        self.data.append(data)
        return self._logged_outcome()

    def _logged_outcome(self):
        resp = self._outcome()
        if resp is not None:
            try:
                self.run_logger.json_dump("data.log", resp.to_dict())
            except:
                pass
        return resp

    def _outcome(self):
        if len(self.data) == self.number_of_loops and not self.done_data_collection:
            self.output_to_csv()
            self.done_data_collection = True
            return {'msg': 'Done', 'payload': 'done', 'code': 0}
        # ? fix this what needs fixing

    def get_fl_data(self):
        self.data = np.array(self.data)
        ROW_PARAMS = (
            ("Red-1", 0),
            ("Green-1", 3),
            ("Red-2", 1),
            ("Green-2", 4),
            ("Red-3", 2),
            ("Green-3", 5),)
        ret = []
        for (channel, col_idx) in ROW_PARAMS:
            if self.data is None:
                fl = [None] * self.number_of_loops
            else:
                self.run_logger.file_append(
                            "column", col_idx)
                self.run_logger.file_append(
                            "output", self.data)
                fl = [float(x) for x in self.data[:, col_idx]] #error here
                if len(fl) < self.number_of_loops:
                    fl += [""] * (self.number_of_loops - len(fl))
            ret.append([channel, *fl])

        return ret

    def output_to_csv(self):
        """Write fluorescence data to run directory."""

        class SOpen:
            def __init__(self, filename):
                self.f = open(filename, 'w', newline='')
                self.f.seek(0)

            def __enter__(self): return self.f

            def __exit__(self, exc_type, exc_value, traceback):
                try:
                    self.f.truncate()
                except:
                    pass
                self.f.close()

        with SOpen("fl-report.csv") as fl_ofile:
            fl_csv = csv.writer(fl_ofile)

            # writing out headers
            FL_HEADER = ["channel"] + \
                [f"RFU_{i+1}" for i in range(0, self.number_of_loops)]
            fl_csv.writerow(FL_HEADER)

            try:
                fl_csv.writerows(self.get_fl_data())
            except:
                raise Exception(
                    f"Error while outputting to csv") #?

class _AnalysisLogger:
    pass

class _BasePscCallThread:
    pass

try:
    from rawhide_mid.assay_primitives.base_analysis import AnalysisLogger, BasePscCallThread
except:
    AnalysisLogger = _AnalysisLogger
    BasePscCallThread = _BasePscCallThread

class AnalysisLogger(AnalysisLogger):
    def file_append(self, fn, content):
        try:
            with open(fn, "a") as f:
                if isinstance(content, (list, tuple)):
                    f.writelines(content)
                else:
                    f.write("{}\n".format(content))
        except:
            pass

    def json_dump(self, fn, dd, *args, **kwargs):
        if "indent" not in kwargs:
            kwargs['indent'] = 1
        if "sort_keys" not in kwargs:
            kwargs['sort_keys'] = True
        if "default" not in kwargs:
            kwargs["default"] = repr
        try:
            with open(fn, "w") as f:
                json.dump(dd, f, *args, **kwargs)
        except:
            traceback.print_exc()


class OpenPscThread(BasePscCallThread):
    def __init__(self, input_queue, psc_callback, calling_callback, run_logger):
        self.run_logger = run_logger

        # default calling state machine
        self.sm = StateMachine(self.run_logger)

        BasePscCallThread.__init__(
            self, input_queue, psc_callback, calling_callback, run_logger)

    def stop(self):
        print("PSC Kill")
        self.looping = False

    def run(self):
        self.looping = True
        while self.looping:
            try:
                row = self.input_queue.get(block=False)
                self.input_queue.task_done()
                resp = None
                if row.label == "data":
                    resp = self.sm.collect_data(row.data)

                if resp is not None:
                    led_override = None
                    try:
                        if resp['code'] == 0:
                            led_override = LED_GREEN_SOLID
                        else:
                            led_override = LED_YELLOW_SOLID

                        if resp['code'] != 0:
                            led_override = LED_YELLOW_FLASH

                        ret_payload = {}
                        """? fix, need to return payload to include code and also the text of the psc how is this normally handled or pass along by runscript to GUI
                        I think at the very least we can add a psc.log file of sorts within tarball that gets generated... which i think it does
                        otherwise i dont see any reason for payload"""
                        # for k, v in resp.payload.items():
                        #     if v is None:
                        #         pass
                        #     else:
                        #         ret_payload[k] = v

                        self.calling_callback(
                            resp['msg'], resp['code'], ret_payload, led_override=led_override)
                        self.looping = False
                        continue
                    except:
                        self.run_logger.file_append(
                            "psc-error", traceback.format_exc())
                        # ?what is a config error if no barcode
                        self.calling_callback(
                            "Configuration Error", 3051, LED_RED_SOLID)

            except KeyboardInterrupt:
                # ? this has a psc?
                print("Keyboard Interrupt during Data Collection")
                self.stop()
                raise
            #No data to add this loop, wait for half a second before looping again
            except queue.Empty:
                time.sleep(0.5)
                pass
            except:  # no psc?
                self.run_logger.file_append(
                    "psc-error", traceback.format_exc())
                self.calling_callback(
                            "Runscript Exception", 3051, LED_RED_SOLID)
                raise


class runscript(object):

    def __init__(self, callback_manager=None,
                 thermal=None,
                 optics=None,
                 analysis=None,
                 timeout_manager=None):

        self.logger = ROOT_MODULE_LOGGER
        self.callback_manager = callback_manager
        self.thermal = thermal
        self.optics = optics
        self.analysis = analysis
        self.analysis.run_logger = AnalysisLogger()
        self.analysis.setup(psccall_thread_class=OpenPscThread)

        self.current_progress = 0

        self.hw = self.thermal.hardware_wrapper

        fh_handler = logging.FileHandler("runscript_log.log")
        fh_handler.setFormatter(logging.Formatter(
            "%(asctime)s::%(funcName)s::%(levelname)s::%(message)s"))
        fh_handler.setLevel(logging.WARNING)
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(logging.Formatter(
            "%(asctime)s::%(funcName)s::%(levelname)s::%(message)s"))
        stream_handler.setLevel(logging.DEBUG)
        self.logger.addHandler(fh_handler)
        self.logger.addHandler(stream_handler)

        self.thermal.spinup(3)

        self.watchdog_timer = threading.Timer(60 * 240, self.timeout) # 60*270 means run for 240mins, which is 4 hours.
        self.watchdog_timer.daemon = True

        if timeout_manager is None:
            self.timeout_manager = TimeoutManager(self.proxy_psc_callback,
                                                  self.cancel, logger=self.analysis.run_logger)
        else:
            self.timeout_manager = timeout_manager

        r = re.compile("INCUBATION_[0-9]*_TARGET_TEMP")
        # count how many incubation target temps are present in the runscript
        incubation_list = list(filter(r.match, globals().keys())) 
        # Gather the step number (i.e. 1, 2, 3) of all of the temp variables from the above step that have non zero values
        incubation_list = [re.findall(r'\d+',i)[0] for i in incubation_list if globals()[i] != 0]
        # Use the step numbers of incubations to create a list of lists, where each sublist consists of the inc information corresponding to that number
        self.incu_information = list([globals()[f"INCUBATION_{num}_TARGET_TEMP"], globals()[
                                 f"INCUBATION_{num}_RAMP_RATE"], globals()[f"INCUBATION_{num}_DWELL_TIME"]] for num in incubation_list)
        ''' RH - make this a class for better readability - this is too much'''
    
        self.prev_temp=0
        
        self.timeout_manager.register(
                 f"heatblock", 250, f"Initial Heatblock Timeout", 3091)

        calc_number_loops(AMPLIFICATION_METHOD)
        self.number_of_loops = NUMBER_OF_LOOPS

        if AMPLIFICATION_METHOD.lower() == "pcr":
            #Assign appropriate variables
            self.denature_temp = CYCLIC_DENATURE_TEMP
            self.denature_ramp_rate = CYCLIC_DENATURE_RAMP_RATE
            self.denature_dwell_time = CYCLIC_DENATURE_DWELL_TIME

            self.anneal_temp = CYCLIC_ANNEAL_TEMP
            self.anneal_ramp_rate = CYCLIC_ANNEAL_RAMP_RATE
            self.anneal_dwell_time = CYCLIC_ANNEAL_DWELL_TIME

            self.heating_ramp_time = abs(self.denature_temp - self.anneal_temp)/self.denature_ramp_rate
            self.cooling_ramp_time = abs(self.anneal_temp - self.denature_temp)/self.anneal_ramp_rate

            #set heatblock to minimum temperature used in the runscript
            # self.init_heatblock_temp = min([item[0] for item in self.incu_information] + [self.denature_temp]+[50])

            #NEED TO CHECK FOR LAST INCU -> FIRST PCR TEMP TOO

            cycling_time_estimate = self.number_of_loops*(self.denature_dwell_time + self.anneal_dwell_time)
                                                          
            if self.thermal.check_cal_data(self.denature_temp, self.anneal_temp, self.cooling_ramp_time):
                cycling_time_estimate += self.number_of_loops*self.cooling_ramp_time
            else:
                cycling_time_estimate += self.number_of_loops*0.4*abs(self.anneal_temp-self.denature_temp)

            if self.thermal.check_cal_data(self.anneal_temp, self.denature_temp, self.heating_ramp_time):
                cycling_time_estimate += self.number_of_loops*self.heating_ramp_time
            else:
                cycling_time_estimate += self.number_of_loops*0.4*abs(self.anneal_temp-self.denature_temp)

            self.time_estimate = cycling_time_estimate + 15
            for i in range(1, self.number_of_loops + 1):
                self.timeout_manager.register(
                    f"pcr_cycle_{i}", cycling_time_estimate/self.number_of_loops, f"Cycling {i} Timeout", 3091)

        elif AMPLIFICATION_METHOD.lower() == "isothermal":
            #Assign appropriate variables
            self.isothermal_amplification_time = self.number_of_loops * IMAGE_CAPTURE_PERIOD
            self.isothermal_amplification_temp = ISOTHERMAL_AMPLIFICATION_TEMP

            #set heatblock to minimum temperature used in the runscript
            # self.init_heatblock_temp =  min([item[0] for item in self.incu_information] + [self.isothermal_amplification_temp]+[50])

            iso_time_estimate = self.isothermal_amplification_time * 1.1 + 60 #is there a better number than 1.5?
            self.time_estimate = iso_time_estimate + 15

            self.timeout_manager.register(
                f"isothermal", iso_time_estimate, f"Isothermal Timeout", 3091)

        else:
            self.callback_manager.error_callback(
                "Runscript Parameter Error", 3021)  # ? psc number?
            
        #Estimate how much time each incubation will need
        
        for i, incu in enumerate(self.incu_information):
            prev_temp = 0
            target_temp = incu[0]
            ramp_rate = incu[1]
            dwell_time = incu[2]
            calc_time = abs(target_temp-prev_temp)/ramp_rate

            #Calculate time estimate for incubation period(s)
            if self.thermal.check_cal_data(prev_temp, target_temp, time):
                est_time = calc_time + dwell_time + 3 #3 seconds added as an error margin
            else:
                est_time = dwell_time + 0.5*abs(target_temp-prev_temp) #will this be ok for both directions??

            incu.append(est_time)
            self.time_estimate += est_time
            self.timeout_manager.register(
                 f"incubation_cycle_{i+1}", est_time, f"Incubation {i+1} Timeout", 3091)  # ? "XYZ" change #
         
        self.final_wait = threading.Event()

    def proxy_psc_callback(self, *args, **kwargs):
        self.callback_manager.psc_callback(*args, **kwargs)

    def timeout(self):
        self.cancel()
        self.callback_manager.psc_callback("Timeout Error", 3051, {})

    def error_log(string):
        with open("runscript_error.log", "a") as f:
            f.write(string + "\n")

    def run(self,config_payload=None):

        self.watchdog_timer.start()
        try:
            
            self.current_progress = 1
            self.progress(self.current_progress)  # ? QQQ
            self.thermal.set_sleeve(setpoint=37, tolerance=2.0, dwell=0.1, 
                    label="Initial Set Sleeve",
                    params={'Kp' : 40.0, 'Ki' : 15.0, 'Kd' : 1.0},
                    offset=0, reset_state=True)
            self.timeout_manager.arm("heatblock")
            self.thermal.sysfan_openloop(
                plant_value=-500, label="openloop-sys-fan")

            self.thermal.set_heatblock(
                setpoint=53,
                tolerance=2.0,
                label="set-heatblock",
                dwell=10.0,
                params={'Kp': 120, "Ki": 1.0, "Kd": 0.0}
            )
            self.timeout_manager.disarm("heatblock")
            self.current_progress = self.current_progress + 10
            self.progress(self.current_progress)

            # set initial prev temp to what the heatblock is set to
            self.prev_temp = 37

            
            for i, values in enumerate(self.incu_information):
                # self.timeout_manager.arm(f"incubation_cycle_{i+1}")

                target_temp = values[0]
                ramp_rate = values[1]
                dwell_time = values[2]
                calc_time = abs(self.prev_temp - target_temp)/ramp_rate

                #attempt to first use the provided ramp rate from the user
                try:
                    self.logger.error("trying ramp")
                    self.thermal.tec_ramp(
                        self.prev_temp,
                        target_temp,
                        calc_time,
                        label=f"Incubation (Tec Ramp) {i+1} Ramp",
                        keep_alive=True
                    )
                    self.logger.error("done ramp")
                    est_time_elapsed = values[3]

                #If the constant ramp rate fails, default back to openloop
                except Exception as e:
                    self.logger.error(f"why tec ramp failed: {e}")
                    sign = np.sign(target_temp-self.prev_temp)
                    self.thermal.tec_openloop(plant_value=sign*300, temperature_trigger=target_temp-(sign*5), #fix this
                    label=f"Incubation Ramping (Openloop) Cycle {i+1}", max_time=20.0)
                    
                #if annealing, use offset as -2
                if np.sign(target_temp-self.prev_temp) == -1:
                    self.thermal.set_sleeve(setpoint=target_temp, tolerance=2.0, dwell=dwell_time, 
                    label=f"Incubation {i+1} Dwell",
                    params={'Kp' : 40.0, 'Ki' : 15.0, 'Kd' : 1.0},
                    offset=-2, reset_state=True)
                    est_time_elapsed = dwell_time + 20
                #otherwise, use offset 200
                else:
                    self.thermal.set_sleeve(setpoint=target_temp, tolerance=2.0, dwell=dwell_time, 
                    label=f"Incubation {i+1} Dwell",
                    params={'Kp' : 40.0, 'Ki' : 15.0, 'Kd' : 1.0},
                    offset=200, reset_state=True)
                    est_time_elapsed = dwell_time + 20

                # self.timeout_manager.disarm(f"incubation_cycle_{i+1}")
                self.prev_temp = target_temp
                self.current_progress = self.current_progress + est_time_elapsed
                self.progress(self.current_progress)

            if AMPLIFICATION_METHOD.lower() == "pcr":
                self.run_pcr()
            elif AMPLIFICATION_METHOD.lower() == "isothermal":
                self.run_isothermal()
            else:
                pass  # add in error handling here

            try:
                self.thermal.tec_openloop(
                    plant_value=0,
                    max_time=0.1,
                    label="TEC Shutoff"
                )

                self.thermal.set_heatblock(
                    setpoint=30.0,
                    tolerance=10.0,
                    label="cooldown-heatblock",
                    dwell=1.0,
                    max_time=40.0,
                    params={'Kp': 120, "Ki": 1.0, "Kd": 0.0}
                )

                self.thermal.set_sleeve(
                    setpoint=20.0,
                    tolerance=20.0,
                    dwell=1.0,
                    params={'Kp': 20.0, 'Ki': 15.0, 'Kd': 5.0},
                    max_time=40.0,
                    label="cooldown-sleeve"
                )

            except:
                error_log("Error during cool down")  # make into psc?
                error_log(traceback.format_exc())

            self.thermal.kill_all()
            try:
                self.watchdog_timer.cancel()
            except:
                pass
            
            self.progress(self.time_estimate-10)
            self.final_wait.wait(timeout=10.0) #? get rid of?
            self.progress(self.time_estimate)
            self.callback_manager.calling_callback("Ending run",led_override=LED_GREEN_SOLID)

        except KeyboardInterrupt:
            self.cancel()
            error_log("Got KeyboardInterrupt")
            self.callback_manager.error_callback(
                "Keyboard Interrupted the Run", 3001)
        except ThermalPredictionException:
            '''
            why is it listed as camera error? and why is a timeouterror also a camera error?
            what produces a timeouterror... is it watchdog? is it the timeout manager?
            '''
            #probably a copy paste error?
            self.cancel()
            error_log(traceback.format_exc())
            self.callback_manager.psc_callback(
                "Thermal Error", 3091, {}, led_override=LED_RED_SOLID) #? number
        except TimeoutError:
            self.cancel()
            error_log(traceback.format_exc())
            self.callback_manager.psc_callback(
                "Camera Error", 3091, {}, led_override=LED_RED_SOLID)
        except SleeveDeltaError:  # ? look into, let us make the text in callback better
            self.cancel()
            error_log(traceback.format_exc())
            self.callback_manager.psc_callback(
                "Thermal Sleeve Delta Error", 3092, {}, led_override=LED_RED_SOLID)
        except Exception as ex:
            self.cancel()
            self.logger.error("Unexpected Runscript Error", exc_info=True)
            error_log(traceback.format_exc())
            self.callback_manager.psc_callback(
                "Unexpected Runscript Error", 3051, {}, led_override=LED_RED_SOLID)

    def run_pcr(self):
        for i in range(1, self.number_of_loops + 1):
            est_time_elapsed = self.denature_dwell_time + self.anneal_dwell_time
            sign = np.sign(self.denature_temp - self.prev_temp)
            # self.timeout_manager.arm(f"pcr_cycle_{i}")


            try:
                 self.thermal.tec_ramp(self.prev_temp, self.denature_temp-2, self.heating_ramp_time, label="Boost-Denature (Tec Ramp)", cycle=i, keep_alive = True)
                 est_time_elapsed += 0.4*abs(self.anneal_temp-self.denature_temp)
            except:

                self.thermal.tec_openloop(plant_value=sign*800, temperature_trigger=self.denature_temp-(sign*5), #fix this
                    label=f"Boost Denature (Openloop) Cycle {i}", cycle=i, max_time=20.0)
                est_time_elapsed += self.anneal_ramp_rate/abs(self.anneal_temp - self.denature_temp)

            self.thermal.set_sleeve(setpoint=self.denature_temp, tolerance=2.0, dwell=self.denature_dwell_time,
                    params={'Kp' : 40.0, 'Ki' : 15.0, 'Kd' : 1.0},
                    label="Denature", offset=205, cycle=i, reset_state=True)
            self.prev_temp = self.denature_temp
            sign = np.sign(self.anneal_temp - self.prev_temp)
            
            try:
                self.thermal.tec_ramp(self.denature_temp, self.anneal_temp+2, self.cooling_ramp_time,
                    label="Boost-Cool (Tec Ramp)", cycle=i, keep_alive = True)
                est_time_elapsed += 0.4*abs(self.anneal_temp-self.denature_temp)
            except:
                self.thermal.tec_openloop(plant_value=sign*800, temperature_trigger=self.anneal_temp-(sign*5), #fix this
                    label=f"Boost Cool (Openloop) (Anneal) Cycle {i}", cycle=i, max_time=20.0)
                est_time_elapsed += self.denature_ramp_rate/abs(self.anneal_temp - self.denature_temp)
                    
            self.thermal.set_sleeve(setpoint=self.anneal_temp, dwell=self.anneal_dwell_time, tolerance=2.0,
                        offset=-2, 
                        params={'Kp' : 40.0, 'Ki' : 15.0, 'Kd' : 1.0},
                        label="Anneal", cycle=i, reset_state=True)
            self.prev_temp = self.anneal_temp   

            # self.timeout_manager.disarm(f"pcr_cycle_{i}")
            self.optics.camera_click(
                "data", i, led_count=30, unblock_save=True)
            self.prev_temp = self.anneal_temp

            #Update progress
            self.current_progress = self.current_progress + est_time_elapsed
            self.progress(self.current_progress)

    def run_isothermal(self):
        # self.timeout_manager.arm("isothermal")
        sign = np.sign(self.isothermal_amplification_temp - self.prev_temp)

        self.logger.error("beginning of openloop in run-iso")
        self.thermal.tec_openloop(plant_value=sign*300, temperature_trigger=self.isothermal_amplification_temp-2, #fix this
                    label=f"Iso Openloop", max_time=50.0)
        
        self.logger.error("beginning of set sleeve in run-iso")

        if sign == 1:
            self.thermal.set_sleeve(setpoint=self.isothermal_amplification_temp, dwell=IMAGE_CAPTURE_PERIOD, tolerance=2.0,
                    offset=200, 
                    params={'Kp' : 40.0, 'Ki' : 15.0, 'Kd' : 1.0},
                    label="Iso Set Sleeve", cycle=0, reset_state=True)
        else:
            self.thermal.set_sleeve(setpoint=self.isothermal_amplification_temp, dwell=IMAGE_CAPTURE_PERIOD, tolerance=2.0,
                        offset=-2, 
                        params={'Kp' : 40.0, 'Ki' : 15.0, 'Kd' : 1.0},
                        label="Iso Set Sleeve", cycle=0, reset_state=True)
        
        self.optics.camera_click("data", 1, led_count=30, unblock_save=True)
        self.logger.error("took pic 1")
             
        self.logger.error("beginning of for loop in run-iso")
        self.logger.error(self.number_of_loops)
        for i in range(1, self.number_of_loops+1):
            start_time = time.time()

            if i != 1:
                self.optics.camera_click("data", i, led_count=30, unblock_save=True)
                self.logger.error(f"took pic {i}")

            while (time.time()-start_time < IMAGE_CAPTURE_PERIOD):
                pass

            self.current_progress = self.current_progress + IMAGE_CAPTURE_PERIOD
            self.progress(self.current_progress)
        # self.timeout_manager.disarm("isothermal")

    def progress(self, progress):
        time_remaining = max(int(self.time_estimate - progress), 60)
        second_progress = (progress/self.time_estimate)*100
        self.callback_manager.progress_callback(second_progress, time_remaining)

    def cancel(self):
        self.thermal.kill_all()
        self.analysis.kill_all()

    def preheat(self):
        time.sleep(0.5)
        self.callback_manager.preheat_callback("Done preheat")