from pymeasure.instruments.keithley import Keithley2450
import pyvisa
import time
import math
from random import choice
#from w1thermsensor import W1ThermSensor
import wx  # Import wx module to use wx.CallAfter
import logging
import csv
import wx.grid
import threading
import pandas as pd
import time
import datetime
import os, sys
import datetime 
import asyncio
from thread import *
import logging


# Set up logging configuration
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('battery_test_log.txt', mode='w')
    ]
)

mA2A = lambda mA: float(mA) / 1000

smu_list = { 'keithley': "05E6" }
manufacturer_keys = { 'Keithley': "keithley" }


def format_time(seconds):
    # Calculate hours, minutes, and seconds
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)

    # Return formatted string
    return f"{hours:02}:{minutes:02}:{seconds:02}"

class BatteryTest:
    def __init__(self, resource, data_file, output_file, termination_percentage=65, termination_current=0.01, virtual_battery=False, update_callback=None, results_update_callback = None, time_callback = None, battery_bad_callback = None, stop_callback = None, display_termination_message_callback = None, clear_termination_message_callback = None):
        self.rm = pyvisa.ResourceManager("C:\\Windows\\System32\\visa32.dll")
        try:
            self.smu = self.rm.open_resource('USB::0x05E6::0x2450::04387874::INSTR')#, timeout=30000) 
        except Exception as e:
            print(f"Error initializing SMU device: {e}")
            self.smu = None
        #self.smu.timeout = 100000
        #self.smu.baud_rate = 57600
        #self.smu.data_bits = 8 
        #self.smu.stop_bits = pyvisa.constants.StopBits.one
        #self.smu.parity = pyvisa.constants.Parity.none 
        #self.smu.write_termination = '\n'
        #self.smu.read_termination = '\n'    
        self.state = 'charging' 
        self.is_virtual_battery = virtual_battery
        self.terminate_test = False
        self.resource = resource 
        self.discharge = 0
        self.temperature = 0
        self.elapsed_time = 0
        self.termination_percentage = termination_percentage
        self.time_callback = time_callback
        self.battery_bad_callback = battery_bad_callback
        self.percentage = 0
        self.phases = 0 
        self.current_decrease_step = 200
        self.end_of_life = False
        self.discharge_rate = 0.02
        self.listen_for_termination = None
        self.eol_ratio = 100
        self.update_callback = update_callback  # Callback to update the frontend
        self.results_update_callback = results_update_callback  # Callback to update the frontend        
        self.max_temperature = 25.0  # Maximum temperature in °C        
        self.logger = logging.getLogger(__name__)
        self.cumulative_power = 0
        self.keithley = None
        self.phases_list = ['A','B','C','D']
        self.stop_callback = stop_callback
        self.termination_power = 0
        self.display_termination_message_callback = display_termination_message_callback
        self.clear_termination_message_callback = clear_termination_message_callback
        
        try:
            self.keithley = self.smu
        except Exception as e:
            print(f"Error initializing Keithley device: {e}")
            self.keithley = None

        try:
            self.data = pd.read_csv(data_file)
        except Exception as e:
            pass
        self.output_file = output_file

        # Testing parameters
        self.revolution_number = 1
        self.termination_percentage = termination_percentage
        self.termination_current = termination_current
        self.first_revolution_discharge_power = False
        self.results = []


    def log(self, level, message):
        # You could define your custom logging logic here
        if level == 'info':
            self.logger.info(message)
        elif level == 'error':
            self.logger.error(message)
        else:
            self.logger.debug(message)

    def log_data(self, revolution, state, elapsed_time, voltage, current, power, cumulative_power, phase='A', discharge=1, temperature=25, percentage = 0, is_discharge = False):
        """Log data and update GUI."""
        temperature = self.simulate_temperature(current, state)

        if not self.revolution_number is 1:  # Log every 10 minutes
            print("Cumulative power at revolution "+str(self.revolution_number)+":", self.cumulative_power)
            print("Discharge power at revolution "+str(self.revolution_number)+":", self.first_revolution_discharge_power)
            try:
               self.percentage = min(100, ((self.cumulative_power / self.first_revolution_discharge_power) * 100))
               print("Percentage at revolution "+str(self.revolution_number)+":", (self.cumulative_power / self.first_revolution_discharge_power) * 100)
               self.eol_ratio = max(self.termination_percentage, self.percentage)
               self.inputs["eol_ratio"].SetValue(f"{self.eol_ratio:,.2f}")
            except Exception as e:
               self.logger.exception(e)
               print("Percentage at revolution "+str(self.revolution_number)+":", 0)
               self.percentage = 0
               self.eol_ratio = max(self.termination_percentage, self.percentage)
               self.inputs["eol_ratio"].SetValue(f"{self.eol_ratio:,.2f}")
            
        is_about_ten_minutes = (self.elapsed_time % 60 < 1)
        
        print("Elapsed time  in seconds:", self.elapsed_time)

        if state == 'Discharge':
            state_message = state + ' - Period ' + self.phases_list[self.phases] 
        else:
            state_message = state 

        if ((self.elapsed_time != 0) and (is_about_ten_minutes)) or (is_discharge == True):  # Log every 10 minutes
            try:
                results = [{
                    "Elapsed Time (s)": format_time(elapsed_time),
                    "Revolution Number": revolution,
                    "State": state_message,
                    "Voltage (V)": round(float(voltage), 2),
                    "Current (mA)": round(float(current), 2),  # Convert A to mA
                    "Power (mW)": round(float(power), 2),
                    "Cumulative Energy (mW-Hrs)": max(0, round(float(cumulative_power / 1000),2)),
                    "Percentage of Cumulative Energy Compared With The First Revolution": round(float(self.percentage), 2),
                    "Temperature (°C)": round(float(temperature), 2)
                }]
            except Exception as e:
                try:
                    results = [{
                        "Elapsed Time (s)": format_time(elapsed_time),
                        "Revolution Number": revolution,
                        "State": state_message,
                        "Voltage (V)": round(float(voltage), 2),
                        "Current (mA)": round(float(current), 2),  # Convert A to mA
                        "Power (mW)": round(float(power), 2),
                        "Cumulative Power (mW-Hrs)": '-',
                        "Percentage of Cumulative Energy Compared With The First Revolution": self.percentage,
                        "Temperature (°C)": round(float(temperature), 2)
                    }]
                except Exception as e:
                    results = [{
                        "Elapsed Time (s)": format_time(elapsed_time),
                        "Revolution Number": revolution,
                        "State": state_message,
                        "Voltage (V)": round(float(voltage), 2),
                        "Current (mA)": round(float(current), 2),  # Convert A to mA
                        "Power (mW)": '-',
                        "Cumulative Power (mW-Hrs)": '-',
                        "Percentage of Cumulative Energy Compared With The First Revolution": self.percentage,
                        "Temperature (°C)": round(float(temperature), 2)
                    }]
                    
            print("Updating table")    
            # After logging data, update the frontend table via the callback
            wx.CallAfter(self.update_callback, results)

        try:
            cumulative_power = float(cumulative_power)
        except Exception as e:
            pass

        try:
            if power != '':
                self.resultsb = [{
                    "time_since_start": format_time(self.elapsed_time),
                    "cycle_number": revolution,
                    #"revolution_number": revolution,
                    "state": state,
                    "voltage": round(float(voltage), 2),
                    "current": round(float(current), 2),  # Convert A to mA
                    "power": round(float(power), 2),
                    "percentage": self.percentage,
                    "phase": phase,
                    "discharge": discharge,
                    #"cumulative_mwh": max(0, cumulative_power),
                    "most_recently_completed_revolution_number": revolution,
                    "most_recently_completed_cumulative_mw_hrs": round(float(cumulative_power), 2),
                    "temperature": round(float(temperature), 2)
                }]
            else:
                self.resultsb = [{
                    "time_since_start": format_time(self.elapsed_time),
                    "cycle_number": revolution,
                    #"revolution_number": revolution,
                    "state": state,
                    "voltage": round(float(voltage), 2),
                    "current": round(float(current), 2),  # Convert A to mA
                    "power": power,
                    "percentage": self.percentage,
                    "phase": phase,
                    "discharge": discharge,
                    #"cumulative_mwh": max(0, cumulative_power),
                    "most_recently_completed_revolution_number": revolution,
                    "most_recently_completed_cumulative_mw_hrs": round(float(cumulative_power), 2),
                    "temperature": round(float(temperature), 2)
                }]
        except Exception as e:
            if power != '':
                self.resultsb = [{
                    "time_since_start": format_time(self.elapsed_time),
                    "cycle_number": revolution,
                    #"revolution_number": revolution,
                    "state": state,
                    "voltage": round(float(voltage), 2),
                    "current": round(float(current), 2),  # Convert A to mA
                    "power": round(float(power), 2),
                    "percentage": self.percentage,
                    "phase": phase,
                    "discharge": discharge,
                    #"cumulative_mwh": cumulative_power,
                    "most_recently_completed_revolution_number": revolution,
                    "most_recently_completed_cumulative_mw_hrs": round(float(cumulative_power), 2),
                    "temperature": round(float(temperature), 2)
                }]
            else:
                if cumulative_power != '':
                    self.resultsb = [{
                        "time_since_start": format_time(self.elapsed_time),
                        "cycle_number": revolution,
                        #"revolution_number": revolution,
                        "state": state,
                        "voltage": round(float(voltage), 2),
                        "current": round(float(current), 2),  # Convert A to mA
                        "power": power,
                        "percentage": self.percentage,
                        "phase": phase,
                        "discharge": discharge,
                        #"cumulative_mwh": cumulative_power,
                        "most_recently_completed_revolution_number": revolution,
                        "most_recently_completed_cumulative_mw_hrs": round(float(cumulative_power), 2),
                        "temperature": round(float(temperature), 2)
                    }]
                else:
                    if self.cumulative_power == '':
                        self.resultsb = [{
                            "time_since_start": format_time(self.elapsed_time),
                            "cycle_number": revolution,
                            #"revolution_number": revolution,
                            "state": state,
                            "voltage": round(float(voltage), 2),
                            "current": round(float(current), 2),  # Convert A to mA
                            "power": power,
                            "percentage": self.percentage,
                            "phase": phase,
                            "discharge": discharge,
                            #"cumulative_mwh": cumulative_power,
                            "most_recently_completed_revolution_number": revolution,
                            "most_recently_completed_cumulative_mw_hrs": round(float(cumulative_power), 2),
                            "temperature": round(float(temperature), 2)
                        }]
                    else:
                        self.resultsb = [{
                            "time_since_start": format_time(self.elapsed_time),
                            "cycle_number": revolution,
                            #"revolution_number": revolution,
                            "state": state,
                            "voltage": round(float(voltage), 2),
                            "current": round(float(current), 2),  # Convert A to mA
                            "power": power,
                            "percentage": self.percentage,
                            "phase": phase,
                            "discharge": discharge,
                            #"cumulative_mwh": cumulative_power,
                            "most_recently_completed_revolution_number": revolution,
                            "most_recently_completed_cumulative_mw_hrs": cumulative_power,
                            "temperature": round(float(temperature), 2)
                        }]
                        
        #print("Updating output with:", self.resultsb)         
               
        wx.CallAfter(self.results_update_callback, self.resultsb)

        if (self.percentage != 0) and self.percentage < self.termination_percentage:
            self.terminate_test = True
            self.end_of_life = True
            print("Reaching End Of Life")
            wx.CallAfter(self.display_termination_message_callback, ["      Check Battery: Battery has reached end of life during overcurrent."] )
            wx.CallAfter(self.stop_callback)

    def simulate_temperature(self, current, state):
        """Simulate temperature based on the state and current."""
        cooling_rate = -0.1  # Cooling rate (°C per second)
        heating_factor = 0.2  # Heating factor (°C per mA)

        if state == "Charge":
            # Increase temperature during charging
            self.temperature += heating_factor * abs(current)
        elif state == "Discharge":
            # Increase temperature during discharging
            self.temperature += heating_factor * abs(current)
        else:
            # Cool down during idle/rest phases
            self.temperature += cooling_rate

        # Clamp temperature to maximum and minimum values
        self.temperature = max(0, min(self.temperature, self.max_temperature))

        return self.temperature

    def get_temperature(self):        
        sensor = W1ThermSensor()
        # Read the temperature in Celsius
        return sensor.get_temperature()


    def adjust_current_based_on_temp(self, current, temperature):
        """Adjust the current to avoid excessive temperature."""
        if temperature >= 25:
            print(f"Temperature {temperature}°C exceeds maximum threshold. Reducing current.")
            # Decrease current to reduce heat dissipation
            new_current = max(0.0, current - self.current_decrease_step)
            
            # Ensure cooling system is active
            #self.cooling_system.activate()
        
        elif temperature >= 20:  # If temperature is approaching 35°C
            print(f"Temperature {temperature}°C approaching limit. Reducing current.")
            # Reduce the current more gently
            current = max(0.0, current - (self.current_decrease_step / 2))
            
            # Ensure cooling system is active if not already
            #if not self.cooling_system.is_active():
            #    self.cooling_system.activate()

    def run_charge_phase(self, inputs):
        """Perform the charge phase."""
        if self.terminate_test is False:
            # Set source mode to voltage
            self.capacity = 0
            self.smu.write(":SOUR:FUNC VOLT")  # Set source function to voltage
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                if 'No error' in resp:
                    break
            # Set source voltage
            self.smu.write(f":SOUR:VOLT {inputs['charge_voltage'].GetValue()}")  # Set voltage level
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                if 'No error' in resp:
                    break
            # Set compliance current (current limit)
            self.smu.write(f":SOUR:VOLT:ILIM {str(mA2A(inputs['max_charge_current'].GetValue()))}")  # Set max charge current
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break
            # Enable source output
            #self.smu.write(":OUTP:STAT 1")
            #self.smu.write(":SYST:ERR?")
            #print("Current error: "+self.smu.read())
            # Set the capacity (assuming you want to store it for further use)

            start_time = time.time()
            last_log_time = start_time  # For logging every minute

            i = 0
            temperature = 26
       
            current = float(self.inputs['max_charge_current'].GetValue())
            self.smu.write(':SOUR:CURR:LEV '+str(mA2A(current)))
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break

            self.smu.write(':MEASURE:VOLT?')    
            #self.logger.exception(e)
            v = self.smu.read()
            print("V", v)  
            self.voltage = max(min(float(self.inputs['charge_voltage'].GetValue()), float(v.split(',')[0].replace('E', 'e')) ), float(self.inputs['discharge_voltage_limit'].GetValue()))
            self.smu.write(':SOUR:VOLT '+str(self.voltage))
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break
            while self.terminate_test is False:
                # Log data at regular intervals (e.g., every 10 seconds)
                self.smu.write(':MEASURE:VOLT?')  # Measure the voltage
                logged_voltage = float(self.smu.read().replace('E','e'))
                print("Measured voltage at this point:", logged_voltage)  # Read the result
                self.smu.write(':MEASURE:CURR?')  # Measure the voltage
                logged_current = float(self.smu.read().replace('E','e')) * 1000
                print("Measured current at this point:", logged_current)  # Read the result
                #current = self.adjust_current_based_on_temp(current, temperature)    
                self.capacity += (logged_voltage*current) * (self.elapsed_time / 3600)  # in mAh        
                normalized_capacity = (self.capacity / self.rated_capacity) * 100
                decay_factor = math.exp(-normalized_capacity / 100)
                current *= decay_factor

                current = min(float(self.inputs['max_charge_current'].GetValue()), current)
                print("Current during charge", current)

                self.capacity = min(self.capacity, self.rated_capacity)

                #try:  # Log every 10 seconds
                #    self.smu.write(':SOUR:CURR:LEV '+str(mA2A(current)))
                #    self.smu.write(":SYST:ERR?")
                #    print("Current error: "+self.smu.read())
                #    self.smu.write(':SOUR:VOLT:LEV:IMM '+str(self.voltage)+'V')
                #    self.smu.write(":SYST:ERR?")
                #    print("Current error: "+self.smu.read())
                #    print("Passed:", time.time() - self.elapsed_time)
                #except Exception as e:  # Log every 10 seconds
                #    pass
                while 1:
                    self.smu.write(":SYST:ERR?")
                    resp = self.smu.read()
                    #print("Errors:", resp)
                    if 'No error' in resp:
                        break                   

                self.termination_current = abs(float(inputs['terminal_stop_current'].GetValue()))

                self.voltage = logged_voltage

                power = (logged_voltage * current)  # Power in mW

                if self.revolution_number > 1:
                    # Clamp cumulative power to ensure it doesn't exceed the first revolution discharge power
                    self.power = min(self.power, self.first_revolution_discharge_power * 3600 / self.elapsed_time)

                #self.cumulative_power += (power * self.elapsed_time) / 3600

                #if self.revolution_number > 1:
                #    # Clamp cumulative power to ensure it doesn't exceed the first revolution discharge power
                #    self.cumulative_power = min(self.cumulative_power, self.first_revolution_discharge_power)

                if 1 >= self.revolution_number:
                    # Log data at regular intervals (e.g., every 10 seconds)
                    self.log_data(self.revolution_number, "Charge", self.elapsed_time, logged_voltage, current, power, 0)
                elif self.elapsed_time >= 600:
                    # Log data at regular intervals (e.g., every 10 seconds)
                    self.log_data(self.revolution_number, "Charge", self.elapsed_time, logged_voltage, current, power, self.cumulative_power)
                    break
                else:
                    # Log data at regular intervals (e.g., every 10 seconds)
                    self.log_data(self.revolution_number, "Charge", self.elapsed_time, logged_voltage, current, power, self.cumulative_power)
                    break

                #print(f"[Charging] Time: {self.elapsed_time:.1f}s, Termination current: {self.termination_current:.3f} V, Current: {current:.3f} A")
                #print(f"Cumulative Power:", self.cumulative_power)

                # Termination conditions
                if current <= self.termination_current:
                    self.state = 'discharging'
                    self.terminate_test = True

                if self.terminate_test is True:
                    self.log_data(self.revolution_number, "Charge", self.elapsed_time, logged_voltage, current, power, self.cumulative_power)    
                    try:
                        self.stop_test()
                    except Exception as e:
                        pass
                    wx.CallAfter(self.display_termination_message_callback, ["      Overcurrent: Battery Bad: Termination current has been reached."])
                    break

                current_date = datetime.datetime.now().strftime("%y-%m-%d")
                if current_date != self.inputs['start_date']:
                    self.time_callback()

                self.current = current
                self.power = power

                time.sleep(1)
                self.elapsed_time += 1
                i += 1

    def run_rest_phase(self, rest_time=None, voltage_stabilization_threshold=0.01, temperature_stabilization_threshold=0.1, inputs = None):
        """
        Introduce a rest phase between charge and discharge.
        The phase can end based on time, voltage stabilization, or temperature stabilization.
        """
        self.phases = 0
        if self.did_shutdown is True:
            self.terminate_test = True
            return
        if self.terminate_test is False:
            last_voltage = None
            voltage_stabilized = False
            temperature_stabilized = False
            start_time = self.elapsed_time
            rest_time = int(inputs['rest_time'].GetValue())
            while self.elapsed_time - start_time <= rest_time:

                # Check for temperature stabilization
                if self.did_shutdown is True:
                    self.terminate_test = True
                    return

                current_temperature = self.simulate_temperature(0, "Rest")
                print("Elapsed time during rest phase:", self.elapsed_time - start_time)

                try:
                    #self.smu.write(":MEASURE:VOLT?")
                    #print("Current V: "+self.smu.read())
                    #self.smu.write(":SOUR:FUNC CURR")
                    #print("Current C: "+self.smu.read())
                    self.smu.write(':SOUR:CURR:LEV 0')
                    while 1:
                        self.smu.write(":SYST:ERR?")
                        resp = self.smu.read()
                        #print("Errors:", resp)
                        if 'No error' in resp:
                            break
                    self.smu.write(':SOUR:VOLT '+self.inputs['discharge_voltage_limit'].GetValue())  # Ensure safety margin for voltage
                    while 1:
                        self.smu.write(":SYST:ERR?")
                        resp = self.smu.read()
                        #print("Errors:", resp)
                        if 'No error' in resp:
                            break
                    #self.smu.write(":OUTP:STAT 1")
                except Exception as e:
                    #self.logger.exception(e)
                    self.voltage = float(self.inputs['discharge_voltage_limit'].GetValue())
                # Check for voltage stabilization
                if last_voltage is not None:
                    voltage_change = abs(self.voltage - last_voltage)
                    if voltage_change <= voltage_stabilization_threshold:
                        voltage_stabilized = True
                last_voltage = self.voltage

                # Check for temperature stabilization
                if abs(current_temperature - self.temperature) <= temperature_stabilization_threshold:
                    temperature_stabilized = True

                self.smu.write(':MEASURE:VOLT?')  # Measure the voltage
                logged_voltage = float(self.smu.read().replace('E','e'))
                print("Measured voltage at this point:", logged_voltage)  # Read the result

                # Log rest phase data
                self.log_data(self.revolution_number, "Rest", self.elapsed_time, logged_voltage, 0, '', '', '', '')

                time.sleep(1)
                self.elapsed_time += 1
            
    def run_discharge_phase(self, inputs):
        """Perform the discharge phase."""
        #print(f"Revolution {self.revolution_number}: Discharging...")
        
        print("Discharging")
        voltage_time = 0
        phase_shift = False
        if self.revolution_number != 1:
            self.cumulative_power = 0
        self.phases = 0
        self.discharge += 1        
        while self.phases != 4:

            self.phase_currents = [inputs['phase_1_current'].GetValue(), inputs['phase_2_current'].GetValue(), inputs['phase_3_current'].GetValue(), inputs['phase_4_current'].GetValue() ]
            self.phase_times = [inputs['phase_1_time'].GetValue(), inputs['phase_2_time'].GetValue(), inputs['phase_3_time'].GetValue(), inputs['phase_4_time'].GetValue() ]
            #self.terminate_test = False
            phase = self.phases        
            try:
                current = float(self.phase_currents[self.phases])
            except Exception as e:
                self.phases = 0
                current = float(self.phase_currents[self.phases])
            print("Period discharge current:", current)
            input_current = current
            duration = self.phase_times[self.phases]
            self.smu.write(":SOUR:FUNC CURR")
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break
            self.smu.write(':SOUR:VOLT:ILIM '+str(mA2A(self.inputs['max_discharge_current'].GetValue())))
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break

            #self.smu.write(":OUTP:STAT 1")
            #self.smu.write(":SYST:ERR?")
            #print("Current error: "+self.smu.read())
            print("Terminate test during discharge:", self.terminate_test)
            self.smu.write(':SOUR:CURR '+str(mA2A(input_current)))
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break
            self.smu.write(':SOUR:VOLT '+str(self.voltage))
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break
            start_time = time.time()
            self.smu.write(':MEASURE:VOLT?')
            self.voltage = float(self.smu.read().split(',')[0].replace('E', 'e') )
            print("Voltage during discharge", self.voltage)
            self.smu.write(':MEASURE:CURR?')
            logged_current = float(self.smu.read().split(',')[0].replace('E', 'e') ) * 1000
            print("Current during discharge", logged_current)
            self.current = logged_current            
            #Stopping condition
            if self.elapsed_time >= float(duration) + start_time:
                self.phases += 1
                phase_shift = True
                self.phase = self.phases_list[self.phases]               
                self.log_data(self.revolution_number, f"Discharge", self.elapsed_time, self.voltage, logged_current, self.power, self.cumulative_power, self.phase, self.discharge, temperature)
                start_time = time.time()
            if self.elapsed_time >= float(self.inputs['discharge_cycle_time'].GetValue()) + start_time:
                if phase_shift is False:
                    self.phases += 1    
                start_time = time.time()                    
                print("Terminating after discharge cycle")
                return                    
            self.smu.write(':MEASURE:VOLT?')  # Measure the voltage
            print("Measured voltage at this point in volts:", float(self.smu.read()))  # Read the result
            self.smu.write(':MEASURE:CURR?')  # Measure the voltage
            print("Measured current at this point:", self.smu.read() * 1000)  # Read the result
            self.power = (self.voltage * abs(self.current))  # Power in mW
            #if self.revolution_number > 1:
            #    # Clamp cumulative power to ensure it doesn't exceed the first revolution discharge power
            #    self.power = min(self.power, self.first_revolution_discharge_power * 3600 / self.elapsed_time)

            if self.revolution_number > 1:
                # Clamp cumulative power to ensure it doesn't exceed the first revolution discharge power
                self.cumulative_power -= (self.power * self.elapsed_time) / 3600
            else:
                # Clamp cumulative power to ensure it doesn't exceed the first revolution discharge power
                self.cumulative_power += (self.power * self.elapsed_time) / 3600  # Add power in mW-Hours
                
            print("Power during discharge:", self.power)    
            print("Cumulative power during discharge:", self.cumulative_power)  

            temperature = 26  # Placeholder for temperature; integrate actual measurement if available
            #print("Current Period", phase)
            # Calculate percentage of cumulative power of the current revolution compared to the first revolution
            self.phase = self.phases_list[self.phases]

            if self.revolution_number == 1:
                self.first_revolution_discharge_power = self.cumulative_power
                self.percentage = 100

            #print("Revolution number:", self.revolution_number)    
            print("Cumulative Power during discharge:", self.cumulative_power)
            print("First Revolution Discharge Power during discharge:", self.first_revolution_discharge_power)

            if self.voltage < float(self.inputs['discharge_voltage_limit'].GetValue()):
                if self.cumulative_power < self.termination_power:
                    #Battery has reached end of life 
                    self.phase = self.phases_list[self.phases]
                    logged_voltage = self.voltage
                    self.log_data(self.revolution_number, f"Discharge", self.elapsed_time, logged_voltage, logged_current, self.power, self.cumulative_power, self.phase, self.discharge, temperature)
                    wx.CallAfter(self.display_termination_message_callback, ["      Check Battery: Termination current has been reached."])
                    self.terminate_test = True
                    return False  # Stop testing
                else:
                    self.revolution_number += 1
                    print("New Revolution", self.revolution_number)
                    self.discharge_voltage = self.voltage
                    print("New Revolution", self.revolution_number)
                    self.discharge_voltage = self.voltage

            logged_voltage = self.voltage
            print("Measured voltage at this point:", logged_voltage)  # Read the result
            self.log_data(self.revolution_number, f"Discharge", self.elapsed_time, logged_voltage, logged_current, self.power, self.cumulative_power, self.phase, self.discharge, temperature)

            if self.terminate_test is True or self.end_of_life is True:
                return False

            time.sleep(1)
            self.elapsed_time += 1
            voltage_time += 1
            print("Voltage time during discharge:", voltage_time)
            print("Duration:", duration)
            self.phase = self.phases_list[self.phases]

            if float(voltage_time) == float(duration):
                self.log_data(self.revolution_number, f"Discharge", self.elapsed_time, logged_voltage, logged_current, self.power, self.cumulative_power, self.phase, self.discharge, temperature, is_discharge = True)
                    
            if voltage_time >= float(duration):
                if phase_shift is False:
                    self.phases += 1
                    voltage_time = 0
                    continue

            # Stopping condition for this phase
            if self.terminate_test:
                self.state = 'charging'
                break

            current_date = datetime.datetime.now().strftime("%y-%m-%d")
            if current_date != self.inputs['start_date']:
                self.time_callback()

            #self.smu.write(":OUTP:STAT 0")
            #self.smu.write(":SYST:ERR?")
            #print("Current error: "+self.smu.read())
            if self.revolution_number == 1:
                self.first_revolution_discharge_power = self.cumulative_power
            self.termination_power = (self.first_revolution_discharge_power * (self.percentage / 100)) / 1000  #multiply for measurements
            #print("Termination power:", self.termination_power)
            #print("Cumulative power:", self.cumulative_power)
            #print(f"Initial discharge capacity: {self.first_revolution_discharge_power:.3f} mW-Hrs")

        print("Current Revolution", self.revolution_number)
        return True  # Continue testing

    def configure_keithley(self):
        """
        Configures the Keithley 2450 Source Measure Unit (SMU) for battery testing.
        This includes setting up the source and measurement parameters for the test.
        """
        try:
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break
            # Set the SMU to operate in DC mode
            self.smu.write("*RST")  # Reset to default settings
            while 1:
                self.smu.write(":SYST:ERR?")
                resp = self.smu.read()
                #print("Errors:", resp)
                if 'No error' in resp:
                    break
            # Configure source as a voltage source
            self.logger.info("Keithley 2450 configured successfully for battery testing.")
        except Exception as e:
            error_msg = f"Error configuring Keithley 2450: {e}"
            self.logger.error(error_msg)

    def run_test(self, is_virtual_battery, inputs):
        """Run the full test cycle."""
        self.phase = 'A'
        self.inputs = inputs
        self.capacity = abs(float(inputs['rated_capacity'].GetValue()))
        self.rated_capacity = abs(float(inputs['rated_capacity'].GetValue()))
        self.terminate_test = False       
        try:
            self.configure_keithley()
            discharge_increase = float(self.inputs['discharge_cycle_time'].GetValue())
            while (self.terminate_test is False) and not (28800 < self.elapsed_time):
                self.run_charge_phase(inputs)
                self.terminate_test = False    
                current_phase = float(inputs['phase_'+str(min(3,self.phases+1))+'_time'].GetValue())
                if self.did_shutdown is True:
                    self.stop_callback()
                    return
                if self.end_of_life is True:
                    self.stop_callback()
                    return
                self.run_rest_phase(rest_time=self.elapsed_time+current_phase, inputs=inputs)  # Rest for 1 minute
                #print("Test termination:", self.terminate_test)
                if self.did_shutdown is True:
                    self.stop_callback()
                    return
                if self.end_of_life is True:
                    self.stop_callback()
                    return                
                if self.run_discharge_phase(inputs) is False:
                    logging.info("Termination at discharge")
                    self.terminate_test = True
                    wx.CallAfter(self.display_termination_message_callback, ["Battery has reached end of life during overcurrent.", "End Of Life:"])
                    self.smu.write(':MEASURE:VOLT?')
                    logged_voltage = float(self.smu.read().split(',')[0].replace('E', 'e') )
                    print("Voltage during discharge", self.voltage)
                    self.smu.write(':MEASURE:CURR?')
                    logged_current = float(self.smu.read().split(',')[0].replace('E', 'e') ) * 1000
                    self.log_data(self.revolution_number, f"Discharge", self.elapsed_time, logged_voltage, logged_current, self.power, self.cumulative_power, self.phase, self.discharge, 0)
                    wx.CallAfter(self.stop_callback)
                    return
                if self.did_shutdown is True:
                    self.stop_callback()
                    return
                if self.end_of_life is True:
                    self.stop_callback()
                    return  
                #self.inputs['discharge_cycle_time'].SetValue(str(float(self.inputs['discharge_cycle_time'].GetValue()) + discharge_increase))  
                #print("Test termination:", self.terminate_test)
            self.stop_callback()
            #print(f"Test completed. Results saved to {self.output_file}.")
        except Exception as e:
            print(f"An error occurred during the test: {logging.exception(e)}")
        finally:
            try:
                self.stop_test()
            except Exception as e:
                pass

    def shutdown(self):
        """Shutdown the Keithley device and release resources."""
        try:
            self.terminate_test = True
            self.smu.close()
            self.log('info', "Keithley 2450 shutdown successfully.")
            #print("Keithley connection closed.")
        except Exception as e:
            self.log('error', e)

class BatteryTestApp(wx.Frame):
    def __init__(self, parent, title, is_virtual_battery = False, update_callback=None):
        super().__init__(parent, title=title, size=(1200, 800))
        logging.info("BatteryTestApp initialized.")
        self.data_file = 'test_template.csv'  # or appropriate default value
        self.output_file = 'revolutions_test_template.csv'  # or appropriate default value
        self.is_virtual_battery = is_virtual_battery  # Flag to toggle virtual battery mode
        #print("Is virtual battery:", is_virtual_battery)
        self.resource = ''  # or appropriate initialization here
           
        self.virtual_battery_data = {
            'voltage': 3.7,
            'current': 500,
            'capacity': 2000,
            'discharge_rate': 0.1
        }

        self.output_data = []
        self.phases_times = ['phase_1_time', 'phase_2_time', 'phase_3_time', 'phase_4_time']
        self.message_displayed = False
        self.test_thread = None
        self.test_active = False
        self.test_running = False
        self.start_time = None
        self.end_time = None
        self.init_ui()
        self.Centre()
        self.Show()
        
    def init_ui(self):
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create a scrolled window to allow scrolling
        self.scrolled_window = wx.ScrolledWindow(self, style=wx.VSCROLL | wx.HSCROLL)
        self.scrolled_window.SetScrollRate(10, 10)

        # Create input and output sections
        input_panel = self.create_input_section(self.scrolled_window)
        output_panel = self.create_output_section(self.scrolled_window)

        # Add both panels to the main sizer
        self.main_sizer.Add(input_panel, 0, wx.EXPAND | wx.ALL, 10)  # Allow flexible expansion
        self.main_sizer.Add(output_panel, 0, wx.EXPAND | wx.ALL, 10)

        # Initialize the grid here, after the parent is set up
        self.grid = wx.grid.Grid(self.scrolled_window)
        self.grid.SetDefaultCellBackgroundColour(wx.Colour(211, 211, 211))
        self.grid.CreateGrid(0, 30)  # 0 rows and 29 columns initially
        self.grid.SetColLabelValue(0, "Operator")
        self.grid.SetColLabelValue(1, "Template")
        self.grid.SetColLabelValue(2, "Manufacturer")
        self.grid.SetColLabelValue(3, "Model")
        self.grid.SetColLabelValue(4, "Nominal Voltage")
        self.grid.SetColLabelValue(5, "Rated Capacity")
        self.grid.SetColLabelValue(6, "Charge Voltage")
        self.grid.SetColLabelValue(7, "Max Charge Current")
        self.grid.SetColLabelValue(8, "Max Discharge Current")
        self.grid.SetColLabelValue(9, "Discharge Voltage Limit")
        self.grid.SetColLabelValue(10, "Terminal Stop Current")
        self.grid.SetColLabelValue(11, "Rest Time")
        self.grid.SetColLabelValue(12, "Phase A Current")
        self.grid.SetColLabelValue(13, "Phase A Time")
        self.grid.SetColLabelValue(14, "Phase B Current")
        self.grid.SetColLabelValue(15, "Phase B Time")
        self.grid.SetColLabelValue(16, "Phase C Current")
        self.grid.SetColLabelValue(17, "Phase C Time")
        self.grid.SetColLabelValue(18, "Phase D Current")
        self.grid.SetColLabelValue(19, "Phase D Time")
        self.grid.SetColLabelValue(20, "End Of Life (%)")
        # Set column output labels
        self.grid.SetColLabelValue(21, "Cumulative Time (Hrs:Min:Sec)")
        self.grid.SetColLabelValue(22, "Revolution Number")
        self.grid.SetColLabelValue(23, "State (Charge/Rest/Phase) - Period (A,B,C,D)")
        self.grid.SetColLabelValue(24, "Voltage (Volts)")
        self.grid.SetColLabelValue(25, "Current (mA)")
        self.grid.SetColLabelValue(26, "Power (mW)")
        self.grid.SetColLabelValue(27, "Cumulative Energy (W*Hrs)")
        self.grid.SetColLabelValue(28, "Percentage of Cumulative Energy Compared With The First Revolution")
        self.grid.SetColLabelValue(29, "Temperature (°C)")
        # Fit the scrollable area to the content inside
        #self.adjust_grid_column_sizes()
        self.grid.AutoSizeColumns()  # Ensure headers align
        self.main_sizer.Add(self.grid, 1, wx.EXPAND | wx.ALL, 10)  # Add grid to sizer
        #total_width = sum([
        #    160, 160, 140, 220, 220, 160, 160, 140,
        #    220, 220, 160, 160, 140, 220, 220, 160, 160, 140, 220, 220, 
        #    140, 190, 200, 160, 160, 140, 220, 220,160, 160
        #])
        #self.scrolled_window.SetMinSize((total_width, 400))
        #self.grid.DeleteCols(0)
        self.export_button = wx.Button(self.scrolled_window, label="Export to CSV")
        self.export_button.Bind(wx.EVT_BUTTON, self.on_export_csv)
        self.main_sizer.Add(self.export_button, 0, wx.LEFT, 310)
        self.export_button.SetPosition(wx.Point(100, self.main_sizer.GetPosition().y))
        # Assign the main sizer to the scrolled window
        self.scrolled_window.SetSizer(self.main_sizer)
        self.scrolled_window.FitInside()   

        # Allow layout updates and proper resizing
        self.Layout()
        self.grid.ForceRefresh()

    def define_tester(self, serial_number, model, smu):
        #self.resource = "USB0::0x"+smu_list[manufacturer_keys[smu.GetValue()]]+"::0x"+model.GetValue()+"::"+serial_number.GetValue()+"::INSTR"
        self.resource = "USB0::0x"+smu_list[manufacturer_keys['Keithley']]+"::0x"+'2450'+"::"+serial_number.GetValue()+"::INSTR"
        self.data_file = "test_data.csv"
        self.output_file = "BatteryTestResults.csv"
        try:
            term = float(self.inputs["eol_ratio"])
            self.tester = BatteryTest(self.resource, self.data_file, self.output_file, term, 0.01, self.is_virtual_battery, self.update_table, self.update_test_results, self.update_start_time, None, self.stop_test, self.display_termination_message, self.clear_termination_message)
        except Exception as e:
            self.tester = BatteryTest(self.resource, self.data_file, self.output_file, 65, 0.01, self.is_virtual_battery, self.update_table, self.update_test_results, self.update_start_time, None, self.stop_test, self.display_termination_message, self.clear_termination_message)

    def update_grid(self, row, col, value):
        """
        Safely update a specific cell in the grid from a background thread.

        Args:
        - row (int): The row index of the grid to update.
        - col (int): The column index of the grid to update.
        - value (str): The value to set for the grid cell.
        """
        self.grid.SetCellBackgroundColour(row, col, wx.Colour(211, 211, 211))
        self.grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))        
        if isinstance(value, (float, complex)):
            self.grid.SetCellValue(row, col, f"{value:.2f}")	
        else:
            try:
                self.grid.SetCellValue(row, col, f"{float(str(value)):.2f}")
            except Exception as e:
                self.grid.SetCellValue(row, col, str(value))

    def toggle_virtual_battery(self, event):
        """Toggle virtual battery mode on/off."""
        self.is_virtual_battery = not self.is_virtual_battery
        if self.is_virtual_battery:
            self.virtual_battery_data['voltage'] = float(self.inputs['nominal_voltage'].GetValue())
            self.virtual_battery_data['current'] = float(self.inputs['max_charge_current'].GetValue())
            self.virtual_battery_data['capacity'] = float(self.inputs['rated_capacity'].GetValue())
        wx.MessageBox(f"Virtual Battery Mode {'Activated' if self.is_virtual_battery else 'Deactivated'}", "Mode Changed", wx.OK | wx.ICON_INFORMATION)

    def on_load_save_csv(self, event):
        """Handles loading or saving CSV when button is clicked."""
        file_path = self.inputs['template'].GetValue()
        if self.is_data_filled() == True:
            self.save_csv(file_path)
        else:  # Otherwise, load the CSV to fill in data
            # Open file dialog to select and load a CSV file if any inputs are empty
            with wx.FileDialog(self, "Open CSV file", wildcard="CSV files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
                if file_dialog.ShowModal() == wx.ID_CANCEL:
                    return  # The user canceled the dialog

                # Get the selected file path
                file_path = file_dialog.GetPath()
                self.load_csv(file_path)


    def load_csv(self, file_path):
        """Load data from CSV and update input fields."""
        try:
            #print("Inputs:", self.inputs)
            with open(file_path, mode='r') as f:
                reader = csv.DictReader(f)
                #print("Reader:", reader)
                for row in reader:
                    for key, ctrl in self.inputs.items():
                        if key in row:
                            ctrl.SetValue(row[key])  # Set field value from CSV
        except Exception as e:
            wx.MessageBox(f"Error loading CSV: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def save_csv(self, file_path):
        """Save current input fields data to a CSV."""
        try:
            fieldnames = ['discharge_voltage_limit', 'phase_1_time', 'serial_number', 'rated_capacity', 'nominal_voltage', 'max_discharge_current', 'phase_4_current', 'template', 'manufacturer', 'max_charge_current', 'phase_3_time', 'start_date', 'start_time', 'model', 'operator', 'discharge_cycle_time', 'phase_4_time', 'phase_2_time', 'phase_1_current', 'phase_3_current', 'terminal_stop_current', 'charge_voltage', 'phase_2_current', 'eol_ratio', 'rest_time']
            data = {key: ctrl.GetValue() for key, ctrl in self.inputs.items()}

            with open(file_path+'.csv', mode='w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()  # Write header
                writer.writerow(data)  # Write current data from inputs

            wx.MessageBox("CSV saved successfully!", "Success", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Error saving CSV: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def compute_discharge_cycle_time(self):
        phase_times = [
            float(self.inputs['phase_1_time'].GetValue()),
            float(self.inputs['phase_2_time'].GetValue()),
            float(self.inputs['phase_3_time'].GetValue()),
            float(self.inputs['phase_4_time'].GetValue())
        ]
        self.inputs['discharge_cycle_time'].SetValue(str(sum(phase_times)))

    def is_data_filled(self):
        """Check if any input fields are filled to decide whether to load or save."""
        print([ctrl.GetValue() != '' for ctrl in self.inputs.values()])
        return all(ctrl.GetValue() != '' for ctrl in self.inputs.values())

    def show_popup(self, title, message):
        """Display a popup message."""
        wx.MessageBox(message, title, wx.OK | wx.ICON_WARNING)

    def create_input_section(self, parent):
        """Create the test input section."""
        #logging.debug("Creating input section.")
        input_panel = wx.Panel(parent)  # Create a separate panel for input
        input_panel.SetBackgroundColour(wx.Colour(211, 211, 211))  # Light blue
        input_sizer = wx.BoxSizer(wx.VERTICAL)  # Local sizer for the input panel

        # Create input section title
        label = wx.StaticText(input_panel, label="Battery Test Input Parameters")
        label.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        label.SetForegroundColour(wx.Colour(0, 0, 128))  # Dark blue text
        input_sizer.Add(label, flag=wx.LEFT, border=310)
        label.SetPosition(wx.Point(100, label.GetPosition().y))
        # Create fields in the input section (like operator, template, etc.)
        grid_sizer = wx.FlexGridSizer(cols=2, vgap=10, hgap=0)  # Set horizontal gap to 0 to reduce space usage
        grid_sizer.AddGrowableCol(1, 1)  # Allow input fields to grow and fill space

        fields = [
            ("Operator", "operator"),
            ("Template", "template"),
            ("Battery Manufacturer", "manufacturer"),
            ("Battery Model Number", "model"),
            ("Battery Serial Number", "serial_number"),
            ("Nominal Voltage", "nominal_voltage"),
            ("Rated Capacity", "rated_capacity"),
            ("Maximum Charge Voltage (charge source voltage)", "charge_voltage"),
            ("Maximum Charge Current (mA)", "max_charge_current"),
            ("Start Date", "start_date"), #automatically loaded by computer
            ("Start Time", "start_time"),
            ("Battery Low Discharge Voltage Limit (V)", "discharge_voltage_limit"),
            ("Maximum Discharge Current (mA)", "max_discharge_current"),
            ("Terminal Stop Current (mA)", "terminal_stop_current"),
            ("Rest phase time (seconds)", "rest_time"),
            ("Discharge Period A Current (mA)", "phase_1_current"),
            ("Discharge Period A Time (seconds)", "phase_1_time"),
            ("Discharge Period B Current (mA)", "phase_2_current"),
            ("Discharge Period B Time (seconds)", "phase_2_time"),
            ("Discharge Period C Current (mA)", "phase_3_current"),
            ("Discharge Period C Time (seconds)", "phase_3_time"),
            ("Discharge Period D Current (mA)", "phase_4_current"),
            ("Discharge Period D Time (seconds)", "phase_4_time"),
            ("Discharge Cycle Time (seconds)", "discharge_cycle_time"), #sum of phase times
            ("End Of Life (%)", "eol_ratio") 
        ]
        
        self.fields = fields

        self.inputs = {}
        for label, key in fields:
            field_label = wx.StaticText(input_panel, label=f"{label}:")
            field_label.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            field_label.SetForegroundColour(wx.Colour(0, 0, 128))  # Dark blue text
            input_field = wx.TextCtrl(input_panel)
            # Make "Template" field longer
            if key == "template":
                input_field.SetMinSize((240, -1))  # Set minimum width for the "Template" field  
                input_field.Bind(wx.EVT_TEXT, self.validate_template)        
            else:
                input_field.SetMinSize((120, -1))  # Set minimum width for the input field
            self.inputs[key] = input_field
            grid_sizer.Add(field_label, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_LEFT)  # Align label to the left
            grid_sizer.Add(input_field, flag=wx.EXPAND)  # Allow the field to expand horizontally

        input_sizer.Add(grid_sizer)

        # Start/Terminate button
        self.start_button = wx.Button(input_panel, label="Start/Terminate", size=(200, 50))
        self.start_button.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        self.start_button.SetBackgroundColour(wx.WHITE)
        self.start_button.SetForegroundColour(wx.RED)
        self.start_button.Bind(wx.EVT_BUTTON, self.toggle_test)
        input_sizer.Add(self.start_button, flag=wx.LEFT, border=330)
        self.start_button.SetPosition(wx.Point(100, self.start_button.GetPosition().y))
        # Load/Save CSV Button
        self.load_save_btn = wx.Button(input_panel, label="Load/Save CSV", pos=(10, 10))
        self.load_save_btn.Bind(wx.EVT_BUTTON, self.on_load_save_csv)
        input_sizer.Add(self.load_save_btn, flag=wx.LEFT, border=330)
        self.load_save_btn.SetPosition(wx.Point(100, input_sizer.GetPosition().y))
        input_panel.SetSizer(input_sizer)  # Set the sizer for the input panel
        #logging.debug("Input section created.")
        return input_panel

    def create_output_section(self, parent):
        """Create the output section."""
        #logging.debug("Creating output section.")
        output_panel = wx.Panel(parent)  # Create a separate panel for output
        output_panel.SetBackgroundColour(wx.Colour(211, 211, 211))
        output_sizer = wx.BoxSizer(wx.VERTICAL)  # Local sizer for output panel

        # Output section title
        label = wx.StaticText(output_panel, label="Test Results")
        label.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        label.SetForegroundColour(wx.Colour(0, 0, 128))  # Dark blue text
        output_sizer.Add(label, flag=wx.LEFT, border=310)
        label.SetPosition(wx.Point(100, label.GetPosition().y))
        # Most Recently Completed Revolution label
        #revolution_label = wx.StaticText(output_panel, label="Most Recently Completed Revolution")
        #revolution_label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        #revolution_label.SetForegroundColour(wx.Colour(0, 0, 128))
        #output_sizer.Add(revolution_label, flag=wx.EXPAND | wx.ALL, border=5)

        revolution_fields = [
            ("Revolution Number", "most_recently_completed_revolution_number"),
            ("Cumulative Power for this Revolution (mW-Hrs)", "most_recently_completed_cumulative_mw_hrs"),
            ("Percentage of Cumulative Power vs First Revolution", "percentage")
        ]

        field_sizer = wx.FlexGridSizer(cols=2, gap=(5, 5))
        self.revolution_outputs = {}
        for label, key in revolution_fields:
            field_label = wx.StaticText(output_panel, label=label + ":")
            field_label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            field_label.SetForegroundColour(wx.Colour(0, 0, 128))
            field_value = wx.StaticText(output_panel, label="0")
            field_value.SetForegroundColour(wx.Colour(0, 0, 0))
            field_value.SetMinSize((120, -1))  # Adjust width to make narrower inputs
            self.revolution_outputs[key] = field_value
            field_sizer.Add(field_label, flag=wx.ALIGN_CENTER_VERTICAL)
            field_sizer.Add(field_value, flag=wx.ALIGN_CENTER_VERTICAL)
        
        self.revolution_outputs['eol_ratio'] = 100

        # Current Revolution label
        current_revolution_label = wx.StaticText(output_panel, label="Current Revolution")
        current_revolution_label.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))  # Larger font
        current_revolution_label.SetForegroundColour(wx.Colour(0, 0, 128))
        output_sizer.Add(current_revolution_label, flag=wx.EXPAND | wx.ALL, border=5)

        current_revolution_fields = [
            ("Charge/Rest/Discharge", "state"),
            ("Discharge Cycle Number", "cycle_number"),
            ("Power (mW)", "power"),
            ("Discharge Period", "phase"),
            ("Time since test started", "time_since_start"),
            #("Revolution Number", "revolution_number"),
            ("Voltage", "voltage"),
            ("Current (mA)", "current"),
            #("Cumulative mW-Hours for this revolution", "cumulative_mwh")
        ]

        current_revolution_field_sizer = wx.FlexGridSizer(cols=2, gap=(10, 10))
        for label, key in current_revolution_fields:
            field_label = wx.StaticText(output_panel, label=label + ":")
            field_label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            field_label.SetForegroundColour(wx.Colour(0, 0, 128))
            field_value = wx.StaticText(output_panel, label="0")
            field_value.SetForegroundColour(wx.Colour(0, 0, 0))
            field_value.SetMinSize((120, -1))  # Adjust width
            self.revolution_outputs[key] = field_value
            current_revolution_field_sizer.Add(field_label, flag=wx.ALIGN_CENTER_VERTICAL)
            current_revolution_field_sizer.Add(field_value, flag=wx.ALIGN_CENTER_VERTICAL)

        # Add the grids to the main sizer
        output_sizer.Add(field_sizer, flag=wx.EXPAND | wx.ALL, border=10)
        output_sizer.Add(current_revolution_field_sizer, flag=wx.EXPAND | wx.ALL, border=10)

        # Add message panel to output sizer (instead of replacing the main sizer)
        self.message_panel = wx.Panel(output_panel)  # Create message panel
        self.message_label = wx.StaticText(self.message_panel)
        #self.message_label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL))
        self.message_label.SetForegroundColour(wx.RED)

        # Optionally, add an attention icon (e.g., using a red exclamation mark)
        self.icon = wx.ArtProvider.GetBitmap(wx.ART_ERROR, size=(16, 16))  # Error icon
        self.icon_bitmap = wx.StaticBitmap(self.message_panel)
        self.message_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add icon and message to sizer
        self.message_sizer.Add(self.icon_bitmap, 0, flag=wx.ALL, border=5)
        self.message_sizer.Add(self.message_label, 0, flag=wx.ALL, border=5)
        self.icon_bitmap.Hide()  # Hide icon by default

        # Add message section to the output sizer
        output_sizer.Add(self.message_panel, 0, flag=wx.EXPAND | wx.ALL, border=10)

        output_panel.SetSizer(output_sizer)  # Set the sizer for the output panel
        return output_panel


    def update_test_results(self, data):
        """Update the test results in the output section with current data."""
        #print("Revolution outputs fields", self.revolution_outputs.keys())
        for key, value in data[0].items():
            #print("Key:", key, "Value:", value)
            if key in self.revolution_outputs.keys():
                try:
                    if not key in ['revolution_number', 'state', 'temperature', 'phase', 'revolution', "most_recently_completed_revolution_number", 'cycle_number', "time_since_start"]:
                        try:
                            self.revolution_outputs[key].SetLabel(f"{value:,.2f}")
                        except Exception as e:
                            logger.exception(e)
                            self.revolution_outputs[key].SetLabel(f"{value}")
                    else:
                        self.revolution_outputs[key].SetLabel(str(value))
                except Exception as e:
                    continue
                self.revolution_outputs[key].GetParent().Refresh()  # Refresh the parent panel of the widgets 
                self.revolution_outputs[key].GetParent().Update() 
            elif key in self.inputs.keys():
                self.inputs[key].SetLabel(str(value))
                self.inputs[key].GetParent().Refresh()  # Refresh the parent panel of the widgets 
                self.inputs[key].GetParent().Update() 

    def increase_phase(self):
        # Check if phase time limit is exceeded
        if self.current_phase["elapsed_time"] >= phase_time_limit:
            self.current_phase["phase_number"] += 1  # Move to the next phase
            self.current_phase["elapsed_time"] = 0.0  # Reset elapsed time for the new phase
            self.current_phase["cumulative_energy"] = 0.0  # Reset energy tracking for the new phase
            logging.info(f"Transitioned to Phase {self.current_phase['phase_number']} ({phase_name})")

    def validate_template(self, event):
        if self.inputs['template'].GetValue().isalnum():
            self.inputs['template'].SetBackgroundColour(wx.NullColour)  # Reset background
        else:
            self.inputs['template'].SetValue("Template must be alphanumeric!")
            self.inputs['template'].SetBackgroundColour(wx.Colour(255, 200, 200))  # Light red
        self.inputs['template'].Refresh()  # Force visual update    

    def load_template(self, event):
        """Load test parameters from a CSV file."""
        with wx.FileDialog(self, "Open Template", wildcard="CSV files (*.csv)|*.csv", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = file_dialog.GetPath()
            # Read CSV and populate input fields
            with open(path, 'r') as file:
                reader = csv.reader(file)
                for row in reader:
                    # Assume first row contains field names, so skip
                    if reader.line_num == 1:
                        continue
                    for idx, (key, (_, limits)) in enumerate(self.inputs.items()):
                        self.inputs[key][0].SetValue(row[idx])

    def save_template(self, event):
        """Save current input parameters to a CSV file."""
        with wx.FileDialog(self, "Save Template", wildcard="CSV files (*.csv)|*.csv", style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = file_dialog.GetPath()
            with open(path, 'w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([key for key, (_, _) in self.inputs.items()])
                for key, (field, _) in self.inputs.items():
                    writer.writerow([field.GetValue()])

    def on_export_csv(self, event):
        """Export the grid data to a CSV file."""
        with wx.FileDialog(self, "Save CSV File", wildcard="CSV files (*.csv)|*.csv", style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return  # User cancelled the save operation

            # Get the selected file path
            path = file_dialog.GetPath()

            # Open the CSV file for writing
            with open(path, 'w', newline='') as file:
                writer = csv.writer(file)

                # Write column headers
                headers = [self.grid.GetColLabelValue(col) for col in range(self.grid.GetNumberCols())]
                writer.writerow(headers)

                # Write grid data (cell values)
                for row in range(self.grid.GetNumberRows()):
                    row_data = [self.grid.GetCellValue(row, col) for col in range(self.grid.GetNumberCols())]
                    writer.writerow(row_data)

            wx.MessageBox(f"Data exported successfully to {path}", "Export Successful", wx.OK | wx.ICON_INFORMATION)

    def update_start_time(self):
        current_time = time.strftime("%H:%M:%S")
        current_date = datetime.datetime.now().strftime("%y-%m-%d")
        if current_date != self.inputs['start_date']:
            self.inputs['start_date'].SetValue(current_date)
        self.inputs['start_time'].SetValue(current_time)

    def toggle_test(self, event):
        """Start or terminate the battery test."""
        # Validate inputs
        # Manufacturer, Model, Lot Code, Serial Number - non-empty check
        for key in ["manufacturer", "model", "serial_number"]:
            if not self.inputs[key].GetValue().strip():
                wx.MessageBox(f"Invalid input: {key} cannot be empty.", "Input Error", wx.OK | wx.ICON_ERROR)
                return
        # Nominal Voltage (1.2 to 13.6 volts)
        nominal_voltage = float(self.inputs["nominal_voltage"].GetValue())
        if not (1.2 <= nominal_voltage <= 13.6):
            wx.MessageBox(f"Invalid input: Nominal Voltage must be between 1.2 and 13.6 volts.", "Input Error", wx.OK | wx.ICON_ERROR)
            return                    
        if float(self.inputs["max_discharge_current"].GetValue()) < float(self.inputs["phase_1_current"].GetValue()) or float(self.inputs["max_discharge_current"].GetValue()) < float(self.inputs["phase_2_current"].GetValue()) or float(self.inputs["max_discharge_current"].GetValue()) < float(self.inputs["phase_3_current"].GetValue()) or float(self.inputs["max_discharge_current"].GetValue()) < float(self.inputs["phase_4_current"].GetValue()):
            wx.MessageBox(f"Invalid input: Discharge current for any phase must be lower than the maximum discharge current.", "Input Error", wx.OK | wx.ICON_ERROR)
            return
        # Rated Capacity (0.2 to 30,000 mW-Hrs)
        rated_capacity = float(self.inputs["rated_capacity"].GetValue())
        if not (0.2 <= rated_capacity <= 30000):
            wx.MessageBox(f"Invalid input: Rated Capacity must be between 0.2 and 30,000 mW-Hrs.", "Input Error", wx.OK | wx.ICON_ERROR)
            return
        # Charge Voltage (1.3 to 16V) and must be greater than the discharge voltage limit
        charge_voltage = float(self.inputs["charge_voltage"].GetValue())
        discharge_voltage_limit = float(self.inputs["discharge_voltage_limit"].GetValue())
        if not (1.3 <= charge_voltage <= 16):
            wx.MessageBox(f"Invalid input: Charge Voltage must be between 1.3 and 16 volts.", "Input Error", wx.OK | wx.ICON_ERROR)
            return
        if charge_voltage <= discharge_voltage_limit:
            wx.MessageBox(f"Invalid input: Charge Voltage must be greater than Discharge Voltage Limit.", "Input Error", wx.OK | wx.ICON_ERROR)
            return
        # Maximum Charge and Discharge Current (10 to 3,000 mA)
        max_charge_current = float(self.inputs["max_charge_current"].GetValue())
        max_discharge_current = float(self.inputs["max_discharge_current"].GetValue())
        if not (10 <= max_charge_current <= 3000):
            wx.MessageBox(f"Invalid input: Maximum Charge Current must be between 10 and 3,000 mA.", "Input Error", wx.OK | wx.ICON_ERROR)
            return
        if not (10 <= max_discharge_current <= 3000):
            wx.MessageBox(f"Invalid input: Maximum Discharge Current must be between 10 and 3,000 mA.", "Input Error", wx.OK | wx.ICON_ERROR)
            return
        # Discharge Low Voltage Limit (between 1 and charge voltage)
        if not (1 <= discharge_voltage_limit <= charge_voltage):
            wx.MessageBox(f"Invalid input: Discharge Low Voltage Limit must be between 1 and Charge Voltage.", "Input Error", wx.OK | wx.ICON_ERROR)
            return  # Stop further execution if validation fails

        # If validation passes, start the test
        if not self.test_running is True:
            self.test_running = True
            self.define_tester(self.inputs['serial_number'], self.inputs['model'], self.inputs['manufacturer'] )       
            self.start_test()
        else:
            self.stop_test()

    def display_termination_message(self, msg):
        # If message is already displayed, do nothing
        if self.message_displayed:
            return

        self.icon_bitmap.Show()             
        self.message_label.SetLabel(str('. ').join(msg))   

        self.message_panel.Layout()
        self.Layout()
        
        # Set flag to indicate message is being displayed
        self.message_displayed = True
        wx.CallLater(3000, self.clear_termination_message, True)
        return self.message_displayed

    def clear_termination_message(self, msg_displayed):
        if msg_displayed == True:
            self.message_displayed = False
            self.message_label.SetLabel("")
            self.icon_bitmap.Hide() 
            self.message_panel.Layout()
            self.Layout()

    async def display_msg(self):
            wx.MessageBox(
                "Error initializing Keithley device: No device found.",
                "Device Initialization Error",
                wx.OK | wx.ICON_ERROR
            )

    def clear_table(self):
        """Clear all rows from the 4th row downwards in the grid."""
        start_row = 1  # Row index to start clearing (4th row)
        num_rows = self.grid.GetNumberRows()
        # Clear content from rows starting at `start_row`
        for row in range(start_row, num_rows):
            for col in range(self.grid.GetNumberCols()):
                self.grid.SetCellValue(row, col, "")  # Clear cell content
        # Delete rows below `start_row`
        if num_rows > start_row:
            self.grid.DeleteRows(pos=start_row, numRows=num_rows - start_row)

    def start_test(self):
        """Start the test in a new thread."""
        # Check if Keithley device was initialized 
        try:
            self.clear_table()
        except Exception as e:
            print("No rows yet, just starting")
        self.compute_discharge_cycle_time()
        if not self.tester.keithley:
            asyncio.run(self.display_msg())
        self.test_running = True
        self.test_active = False
        self.tester.test_active = False
        self.tester.did_shutdown = False
        self.tester.end_of_life = False
        self.tester.revolution_number = 1
        self.tester.elapsed_time = 0
        self.tester.phases = 1
        self.tester.discharge = 0
        self.start_button.SetLabel("Terminate Test")
        self.test_thread = threading.Thread(target=self.run_test, daemon=True)
        self.test_thread.start()
        logging.info("Test started.")
        
    def stop_test(self):
        """Stop the test."""
        #This state structures the program's closure
        self.tester.did_shutdown = True        
        #This performs a test closure
        self.test_running = False
        self.start_button.SetLabel("Start Test")
        print("Stopping the test")
        self.tester.shutdown()
        self.test_thread = None
        logging.info("Test terminated.")

    def run_test(self):
        """Run the test, either real or virtual."""
        print("Test active:", self.test_active)
        if self.test_active == False:
            self.update_start_time()        	
            self.test_active = True
            test_data = self.tester.run_test(self.is_virtual_battery, self.inputs)
        else:
            self.stop_test()  # Stop the test in case of an error

    def simulate_virtual_battery(self):
        """Simulate the behavior of a virtual battery."""
        while self.test_running:
            try:
                # Simulate battery discharge
                #print("Virtual battery data:", self.virtual_battery_data)

                # Ensure voltage doesn't go below a safe threshold
                if self.virtual_battery_data['voltage'] <= float(self.inputs["discharge_voltage_limit"].GetValue()):
                    logging.info("Virtual battery test terminated: Voltage limit reached.")
                    self.stop_test()
                    break
                # Mock test data structure
                test_data = {
                    "most_recent_rev": {
                        "revolution_number": 1,
                        "cumulative_mw_hrs": self.virtual_battery_data['capacity'],
                        "percentage": (self.virtual_battery_data['capacity'] / 100) * 100
                    },
                    "current_revolution": [
                        self.virtual_battery_data['voltage'],
                        self.virtual_battery_data['current'],
                        self.virtual_battery_data['capacity']
                    ]
                }
                # Update the UI with simulated data
                # Simulate time delay
                time.sleep(1)  # 1-second interval for updates
            except Exception as e:
                logging.exception(f"Error during virtual battery simulation: {e}")
                self.stop_test()
                break

    def adjust_grid_column_sizes(self):
        """Adjust the column sizes of the grid."""
        #col_widths = [190, 200, 160, 160, 140, 220, 220]  # Adjust column widths as needed
        col_widths = [
            160, 160, 140, 220, 220, 160, 160, 140,
            220, 220, 160, 160, 140, 220, 220, 160, 160, 140, 220, 220, 
            140, 190, 200, 160, 160, 140, 220, 220,160, 160
        ]
        for col, width in enumerate(col_widths):
            self.grid.SetColSize(col, width)

    def update_output(self, data):
        """Update the output fields and table."""
        #print("Data:", data)
        try:
            most_recent_rev = data.get("most_recent_rev", {})
            for key, value in most_recent_rev.items():
                if key in self.revolution_outputs:
                    if key == "most_recently_completed_revolution_number":
                        self.revolution_outputs[key].SetLabel(f"{int(value)}")
                    elif key == "cycle_number":
                        self.revolution_outputs[key].SetLabel(f"{int(value)}")
                    else:
                        self.revolution_outputs[key].SetLabel(f"{value:,.2f}" if isinstance(value, (float, int)) else str(value))
            # Update grid/table data
            if "current_revolution" in data:
                self.update_table(data["current_revolution"])
        except Exception as e:
            #logging.exception(f"Error updating output: {e}")
            pass

    def update_table(self, revolution_data):
        """Update the grid with the current revolution data."""
        if not revolution_data:
            logging.warning("No data to update the grid.")
            return  # Skip if there's no data to populate the grid

        # Append a new row
        self.grid.AppendRows(1)

        # Update the last row
        last_row_index = self.grid.GetNumberRows() - 1  # Always get the last row dynamically
        last_row = revolution_data[-1]

        for j, value in enumerate(last_row.values()):
            try:
                if key == "cycle_number":
                    wx.CallAfter(self.grid.SetCellValue, last_row_index, j+21, str(f"{int(value)}"))
                elif key == "most_recently_completed_revolution_number":
                    wx.CallAfter(self.grid.SetCellValue, last_row_index, j+21, str(f"{int(value)}"))
                else:
                    wx.CallAfter(self.grid.SetCellValue, last_row_index, j+21, str(f"{value:,.2f}"))
                wx.CallAfter(self.grid.SetCellBackgroundColour, last_row_index, j+21, wx.Colour(211, 211, 211))
                wx.CallAfter(self.grid.SetCellTextColour, last_row_index, j+21, wx.Colour(0, 0, 0))
            except Exception as e:
                wx.CallAfter(self.grid.SetCellValue, last_row_index, j+21, str(value))
                wx.CallAfter(self.grid.SetCellBackgroundColour, last_row_index, j+21, wx.Colour(211, 211, 211))
                wx.CallAfter(self.grid.SetCellTextColour, last_row_index, j+21, wx.Colour(0, 0, 0))
        # Update input values row (second row)
        input_headers = [
            "operator",
            "template",
            "manufacturer",
            "model",
            "nominal_voltage",
            "rated_capacity",
            "charge_voltage",
            "max_charge_current",
            "max_discharge_current",
            "discharge_voltage_limit",
            "terminal_stop_current",
            "rest_time",
            "phase_1_current",
            "phase_1_time",
            "phase_2_current",
            "phase_2_time",
            "phase_3_current",
            "phase_3_time",
            "phase_4_current",
            "phase_4_time",
            "eol_ratio"
        ]
        for col, header in enumerate(input_headers):
            value = self.inputs[header].GetValue()  # Retrieve the current user setting for the header
            print("Value", value)
            self.grid.SetCellValue(last_row_index, col, str(value))  # Update the second row (input values row

        # Refresh layout
        self.scrolled_window.Layout()
        self.scrolled_window.FitInside()

app = wx.App(False)
if 'True' in sys.argv:
    BatteryTestApp(None, title="Battery Test Interface", is_virtual_battery=True)
else:
    BatteryTestApp(None, title="Battery Test Interface", is_virtual_battery=False)
app.MainLoop()
