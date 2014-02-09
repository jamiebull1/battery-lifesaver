'''
Created on 9 Feb 2014

@author: Jamie
'''
import logging
import wmi
import platform

VERSION_NUMBER = '0.0.6-beta'

class BatteryMonitor:
    ''' Class containing methods for testing power supply and battery
        charge levels, and suggesting action to be taken to extend battery life'''
    
    def __init__(self):
        logging.info('\r\r')
        logging.info('Starting laptop battery monitor application')
        logging.info('Initialising laptop battery monitor')
        self.record_system_info()
        logging.info('Initialising wmi.WMI()')        
        self.c = wmi.WMI()
        logging.info('Initialising wmi.WMI(moniker = "//./root/wmi)')
        self.t = wmi.WMI(moniker = "//./root/wmi")
        logging.info('Enabling alerts')
        self.unplug_alert_enabled = True # Initialise to True
        self.plugin_alert_enabled = True # Initialise to True
        self.fully_charged_alert_enabled = True # Initialise to True
        self.PLUGIN_LEVEL = 0.3
        self.UNPLUG_LEVEL = 0.8
        self.reset_time_remaining_queue()
    
    def record_system_info(self):
        logging.info('Battery Lifesaver version: %s' % VERSION_NUMBER)
        logging.info('System details: %s' % str(platform.uname()))
    
    @property
    def is_plugged_in(self):
        ''' Returns True if laptop is connected to a power supply '''
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        plugged_in = False
        for _i, b in enumerate(batts):
            if b.PowerOnline:
                plugged_in = True

        if plugged_in: logging.debug('Power is connected')
        if not plugged_in: logging.debug('Power is not connected')
        
        return plugged_in
        
    @property
    def is_fully_charged(self):
        ''' Returns True if laptop is fully charged '''        
        logging.debug('Battery %i%% charged' % (self.percentage_charge_remaining * 100))
        if self.percentage_charge_remaining >= 1.0:
            return True
        else:
            return False            

    @property
    def battery_statuses(self):
        ''' Returns a list with percentage remaining or "Not present" for each battery '''
        statuses = []
        batts_details = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        batts_charge = self.t.ExecQuery('Select * from BatteryFullChargedCapacity')
        for i, b in enumerate(batts_details):
            if b.RemainingCapacity:
                perc_charge = (b.RemainingCapacity or 0)/float(batts_charge[i].FullChargedCapacity)
                statuses.append('Battery #%i: %i%% available' % (i+1, (perc_charge * 100)))
            else:
                statuses.append('Battery #%i: Not present' % i+1)
        return statuses
            
    @property
    def full_charge_capacity(self):
        ''' Returns capacity of the battery or batteries when fully charged '''
        capacity = 0
        batts = self.t.ExecQuery('Select * from BatteryFullChargedCapacity')
        for _i, b in enumerate(batts):
            capacity += (b.FullChargedCapacity or 0)
        logging.debug('Full charge capacity: %s' % capacity)
        return capacity 

    @property
    def remaining_capacity(self):
        ''' Returns the remaining capacity of the battery or batteries '''
        capacity = 0
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        for _i, b in enumerate(batts):
            capacity += (b.RemainingCapacity or 0)            
        logging.debug('Remaining capacity: %s' % capacity)
        return capacity 
    
    @property
    def percentage_charge_remaining(self):    
        ''' Returns proportion of charge remaining as a float between 0.0 and 1.0 '''
        charge = float(self.remaining_capacity) / float(self.full_charge_capacity)
        logging.debug('Percentage charge remaining: %i%%' % (min(charge, 1.0) * 100))
        return min(charge, 1.0)        
        
    @property
    def time_remaining(self):
        ''' Returns time remaining, calculated as for the Windows Battery Meter. It finds
        a value for remaining battery life by dividing the remaining battery capacity 
        by the current battery draining rate as described in the ACPI specification 
        (chapter 3.9.3 'Battery Gas Gauge'). This is then averaged over a number of periods '''
        time_left = 0
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        for _i, b in enumerate(batts):
            time_left += float(b.RemainingCapacity) / float(b.DischargeRate)
            
        self.time_remaining_queue += [time_left]
        self.time_remaining_queue = self.time_remaining_queue[1:]
        
        if not float('-inf') in self.time_remaining_queue:
            average_time_remaining = sum(self.time_remaining_queue)/len(self.time_remaining_queue)
            hours = int(average_time_remaining)
            mins = 60 * (average_time_remaining % 1.0)
            return '%i hr %i min' % (hours, mins)

    def reset_time_remaining_queue(self):
        self.time_remaining_queue = [float('-inf')] * 20

    def should_unplug(self):
        ''' Tests whether conditions are met for unplugging the laptop '''
        unplug = (self.percentage_charge_remaining > self.UNPLUG_LEVEL and
                  self.is_plugged_in and
                  self.unplug_alert_enabled and
                  self.fully_charged_alert_enabled)
        if unplug: logging.info('Alerting to unplug')
        return unplug
        
    def should_plug_in(self):
        ''' Tests whether conditions are met for plugging in the laptop '''
        plugin = (self.percentage_charge_remaining < self.PLUGIN_LEVEL and
                  not self.is_plugged_in and
                  self.plugin_alert_enabled)
        if plugin: logging.info('Alerting to plug in')
        return plugin
