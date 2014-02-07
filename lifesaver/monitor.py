#!/usr/bin/env python
# coding=utf-8
'''
Created on 4 Feb 2014

@author: Jamie Bull - Young Bull Industries
'''

import webbrowser
import wmi
import winsound
from math import floor
import platform

import wx

import icons

# Logging setup
import logging
VERSION_NUMBER = '0.0.5-beta'

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)
debug_handler = logging.FileHandler('bl_%s.debug.log' % VERSION_NUMBER)
debug_handler.setLevel(logging.DEBUG)
info_handler = logging.FileHandler('bl_%s.info.log' % VERSION_NUMBER)
info_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
debug_handler.setFormatter(formatter)
info_handler.setFormatter(formatter)
logger.addHandler(debug_handler)
logger.addHandler(info_handler)


class BatteryMonitor:
    ''' Class containing methods for testing power supply and battery
        charge levels, and suggesting action to be taken to extend battery life'''
    
    def __init__(self):
        logger.info('\r\r')
        logger.info('Starting laptop battery monitor application')
        logger.info('Initialising laptop battery monitor')
        self.record_system_info()
        logger.info('Initialising wmi.WMI()')        
        self.c = wmi.WMI()
        logger.info('Initialising wmi.WMI(moniker = "//./root/wmi)')
        self.t = wmi.WMI(moniker = "//./root/wmi")
        logger.info('Enabling alerts')
        self.unplug_alert_enabled = True # Initialise to True
        self.plugin_alert_enabled = True # Initialise to True
        self.fully_charged_alert_enabled = True # Initialise to True
        self.PLUGIN_LEVEL = 0.3
        self.UNPLUG_LEVEL = 0.8
        self.reset_time_remaining_queue()
    
    def record_system_info(self):
        logger.info('Battery Lifesaver version: %s' % VERSION_NUMBER)
        logger.info('System details: %s' % str(platform.uname()))
    
    @property
    def is_plugged_in(self):
        ''' Returns True if laptop is connected to a power supply '''
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        plugged_in = False
        for _i, b in enumerate(batts):
            if b.PowerOnline:
                plugged_in = True

        if plugged_in: logger.debug('Power is connected')
        if not plugged_in: logger.debug('Power is not connected')
        
        return plugged_in
        
    
    @property
    def is_fully_charged(self):
        ''' Returns True if laptop is fully charged '''        
        logger.debug('Battery %i%% charged' % (self.percentage_charge_remaining * 100))
        if self.percentage_charge_remaining >= 1.0:
            return True
        else:
            return False            
    
    @property
    def full_charge_capacity(self):
        ''' Returns capacity of the battery or batteries when fully charged '''
        capacity = 0
        batts = self.t.ExecQuery('Select * from BatteryFullChargedCapacity')
        for _i, b in enumerate(batts):
            capacity += (b.FullChargedCapacity or 0)
        logger.debug('Full charge capacity: %s' % capacity)
        return capacity 

    @property
    def remaining_capacity(self):
        ''' Returns the remaining capacity of the battery or batteries '''
        capacity = 0
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        for _i, b in enumerate(batts):
            capacity += (b.RemainingCapacity or 0)            
        logger.debug('Remaining capacity: %s' % capacity)
        return capacity 
    
    @property
    def percentage_charge_remaining(self):    
        ''' Returns proportion of charge remaining as a float between 0.0 and 1.0 '''
        charge = float(self.remaining_capacity) / float(self.full_charge_capacity)
        logger.debug('Percentage charge remaining: %i%%' % (min(charge, 1.0) * 100))
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
            print self.time_remaining_queue
            return '%i hr %i min' % (hours, mins)

    def reset_time_remaining_queue(self):
        self.time_remaining_queue = [float('-inf')] * 20

    def should_unplug(self):
        ''' Tests whether conditions are met for unplugging the laptop '''
        unplug = (self.percentage_charge_remaining > self.UNPLUG_LEVEL and
                  self.is_plugged_in and
                  self.unplug_alert_enabled and
                  self.fully_charged_alert_enabled)
        if unplug: logger.info('Alerting to unplug')
        return unplug
        
    def should_plug_in(self):
        ''' Tests whether conditions are met for plugging in the laptop '''
        plugin = (self.percentage_charge_remaining < self.PLUGIN_LEVEL and
                  not self.is_plugged_in and
                  self.plugin_alert_enabled)
        if plugin: logger.info('Alerting to plug in')
        return plugin

ID_SILENCE_UNPLUG_ALERT = wx.NewId()
ID_SILENCE_PLUGIN_ALERT = wx.NewId()
ID_SILENCE_FULLY_CHARGED_ALERT = wx.NewId()

class BatteryTaskBarIcon(wx.TaskBarIcon):
    ''' Notification area (system tray) icon for output to user about their
        battery status '''
    
    def __init__(self,
                 laptop_batt
                 ):
        wx.TaskBarIcon.__init__(self)
        self.monitor_frequency = 2 # how often to check levels (secs)
        self.full_charge_reminder_frequency = 300 # how often to remind that battery is full (secs)
        self.laptop_batt = laptop_batt
        self.icon = self.BatteryIcon.GetIcon()
        self.SetIcon(self.icon, self.Tooltip)
        self.CreateMenu()
        self.Update()
    
    @property
    def Tooltip(self):
        ''' Generates a tooltip which replicates the Windows Battery Monitor '''
        charge = self.laptop_batt.percentage_charge_remaining * 100
        if self.laptop_batt.is_plugged_in:
            if self.laptop_batt.is_fully_charged:
                tooltip = "Fully charged (100%)"
            else:
                tooltip = "%i%% available (plugged in, charging)" % (charge)
        elif not self.laptop_batt.is_plugged_in:
            time_remaining = self.laptop_batt.time_remaining
            if not time_remaining is None:
                tooltip = "%s (%i%%) remaining" % (time_remaining, charge)
            else:
                tooltip = "%i%% remaining" % (charge)
        logging.debug("Tooltip is %s" % tooltip)
        return tooltip
    
    @property
    def BatteryIcon(self):
        ''' Returns the appropriate icon for the current charge level and whether
            the laptop is connected to a power supply '''
        charge = self.laptop_batt.percentage_charge_remaining * 100
        charge = int(floor(charge/20)*20) # round down to nearest multiple of 20
        if self.laptop_batt.is_plugged_in:
            ico = icons.icons["%s%03d" % ("battery_charging_", charge)]
        else:
            ico = icons.icons["%s%03d" % ("battery_discharging_", charge)]
        logging.debug("Icon is %s" % ico)
        return ico
           
    def Update(self):
        logger.info('')        
        logger.info('Updating')        
        self.RefreshIcon()
        self.ResetAlertsBasedOnPowerStatus()
        self.CheckAlertBalloons()
        wx.CallLater(self.full_charge_reminder_frequency * 1000,
                     self.CheckFullyChargedBalloon)
        wx.CallLater(self.monitor_frequency * 1000,
                     self.Update)
    
    def SilenceFullyChargedAlert(self, e):
        ''' Silences the full charge alert, for use when not ready to leave charging point '''
        self.laptop_batt.fully_charged_alert_enabled = False        
        self.menu.Enable(id=ID_SILENCE_FULLY_CHARGED_ALERT, enable=False) 
        logging.info("Silencing fully charged alert")

    def SilenceUplugAlert(self, e):
        ''' Silences the unplug alert, for use when a full charge is desired '''
        self.laptop_batt.unplug_alert_enabled = False
        self.menu.Enable(id=ID_SILENCE_UNPLUG_ALERT, enable=False) 
        logging.info("Silencing unplug alert")
    
    def SilencePluginAlert(self, e):
        ''' Silences the plugin alert, for use when away from a charging point '''
        self.laptop_batt.plugin_alert_enabled = False
        self.menu.Enable(id=ID_SILENCE_PLUGIN_ALERT, enable=False) 
        logging.info("Silencing plug in alert")
    
    def CheckFullyChargedBalloon(self):
        ''' Tests if fully charged and fires alert if required '''
        if (self.laptop_batt.is_fully_charged and
            self.laptop_batt.is_plugged_in and
            self.laptop_batt.fully_charged_alert_enabled):
            logging.info("Showing fully charged balloon notification")
            self.ShowBalloon("Fully charged",
                             "Your battery is now charged to 100%.")
    
    def ResetAlertsBasedOnPowerStatus(self):
        ''' Tests if plugged in and resets alerts if required'''
        if self.laptop_batt.is_plugged_in:
            logger.info('Plugged in. Resetting stored battery time-remaining values')
            self.laptop_batt.reset_time_remaining_queue
            if not self.laptop_batt.plugin_alert_enabled:
                logger.info('Plugged in. Resetting plugin alert')
                self.laptop_batt.plugin_alert_enabled = True
        elif not self.laptop_batt.is_plugged_in:
            if not self.laptop_batt.unplug_alert_enabled:            
                logger.info('Not plugged in. Resetting unplug alert')            
                self.laptop_batt.unplug_alert_enabled = True
            if not self.laptop_batt.fully_charged_alert_enabled:
                logger.info('Not plugged in. Resetting fully charged alert')            
                self.laptop_batt.fully_charged_alert_enabled = True
            
        #=======================================================================
        # logger.debug('\tPlugged in: %s' % self.laptop_batt.is_plugged_in)
        # logger.info('\tFully charged alert enabled: %s' % self.laptop_batt.fully_charged_alert_enabled)
        # logger.info('\tPlug in alert enabled: %s' % self.laptop_batt.plugin_alert_enabled)
        # logger.info('\tUnplug alert enabled: %s' % self.laptop_batt.unplug_alert_enabled)
        #=======================================================================

        self.menu.Enable(id=ID_SILENCE_FULLY_CHARGED_ALERT,
                         enable=(self.laptop_batt.fully_charged_alert_enabled and
                                 self.laptop_batt.is_plugged_in)) 
        self.menu.Enable(id=ID_SILENCE_PLUGIN_ALERT,
                         enable=(self.laptop_batt.plugin_alert_enabled and
                                 not self.laptop_batt.is_plugged_in))
        self.menu.Enable(id=ID_SILENCE_UNPLUG_ALERT,
                         enable=(self.laptop_batt.unplug_alert_enabled and
                                 self.laptop_batt.is_plugged_in)) 
    
    def CheckAlertBalloons(self):
        charge = self.laptop_batt.percentage_charge_remaining
        if self.laptop_batt.should_unplug():
            self.ShowBalloon("Unplug charger",
                             "Battery charge is at %i%%. Unplug your charger now to maintain battery life." % (charge * 100))
            logging.info("Showing unplug balloon notification")
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        if self.laptop_batt.should_plug_in():
            self.ShowBalloon("Plug in charger",
                             "Battery charge is at %i%%. Plug in your charger now to maintain battery life." % (charge * 100))
            logging.info("Showing plug in balloon notification")
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
    
    def RefreshIcon(self):
        ''' Sets the appropriate icon depending on power state '''
        logger.debug('Refreshing icon')
        self.icon = self.BatteryIcon.GetIcon()
        self.SetIcon(self.icon, self.Tooltip)        
        
    def CreateMenu(self):
        ''' Generates a context-aware menu. The user is only offered the relevant option
            depending on whether their laptop is plugged in '''
        logging.info("Creating icon menu")
        self.Bind(wx.EVT_TASKBAR_RIGHT_UP, self.OnPopup)        
        
        self.menu = wx.Menu()
        self.menu.Append(ID_SILENCE_FULLY_CHARGED_ALERT, 'Silence &Full Charge Alert', 'Continue not showing alerts')
        self.Bind(wx.EVT_MENU, self.SilenceFullyChargedAlert, id=ID_SILENCE_FULLY_CHARGED_ALERT)
        self.menu.Append(ID_SILENCE_PLUGIN_ALERT, 'Silence &Plugin Alert', 'Wait until next plugged in before resuming alerts')
        self.Bind(wx.EVT_MENU, self.SilencePluginAlert, id=ID_SILENCE_PLUGIN_ALERT)
        self.menu.Append(ID_SILENCE_UNPLUG_ALERT, 'Silence &Unplug Alert', 'Fully charge without showing alerts')
        self.Bind(wx.EVT_MENU, self.SilenceUplugAlert, id=ID_SILENCE_UNPLUG_ALERT)
        self.menu.AppendSeparator()
        self.menu.Append(wx.ID_ABOUT, '&Website', 'About this program')
        self.Bind(wx.EVT_MENU, self.OnAbout, id=wx.ID_ABOUT)
        self.menu.Append(wx.ID_EXIT, 'E&xit', 'Remove icon and quit application')
        self.Bind(wx.EVT_MENU, self.OnExit, id=wx.ID_EXIT)
        
    def OnPopup(self, event):
        ''' Generates the right click menu '''
        self.PopupMenu(self.menu)

    def OnExit(self, e):
        ''' Removes the icon from the notification area and closes the program '''
        logging.info("Closing application")
        self.Destroy()
        
    def OnAbout(self, e):
        ''' Launches the Battery Lifesaver webpage '''
        webbrowser.open("http://youngbullindustries.wordpress.com/about-battery-lifesaver/")
        

def main():
    batt = BatteryMonitor()
    
    app = wx.App(False)
    BatteryTaskBarIcon(batt)
    
    app.MainLoop()

if __name__ == '__main__':
    main()
