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

import wx

import icons
from balloon_task_bar import BalloonTaskBarIcon

class BatteryMonitor:
    ''' Class containing methods for testing power supply and battery
        charge levels, and suggesting action to be taken to extend battery life'''
    
    def __init__(self):
        self.c = wmi.WMI()
        self.t = wmi.WMI(moniker = "//./root/wmi")
        self.unplug_alert_enabled = True # Initialise to True
        self.plugin_alert_enabled = True # Initialise to True
        self.fully_charged_alert_enabled  = True # Initialise to True
        self.PLUGIN_LEVEL = 0.4 # Could be set to 0.3
        self.UNPLUG_LEVEL = 0.8
        self.rolling_power_level = []
    
    @property
    def is_plugged_in(self):
        ''' Returns True if laptop is connected to a power supply '''
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        for _i, b in enumerate(batts):
            if b.PowerOnline:
                return True
    
    @property
    def is_fully_charged(self):
        ''' Returns True if laptop is fully charged '''     
        return 1.0   
        return self.percentage_charge_remaining >= 1.0    
    
    @property
    def full_charge_capacity(self):
        ''' Returns capacity of the battery or batteries when fully charged '''
        capacity = 0
        batts = self.t.ExecQuery('Select * from BatteryFullChargedCapacity')
        for _i, b in enumerate(batts):
            capacity += (b.FullChargedCapacity or 0)
        return capacity 

    @property
    def remaining_capacity(self):
        ''' Returns the remaining capacity of the battery or batteries '''
        capacity = 0
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        for _i, b in enumerate(batts):
            capacity += (b.RemainingCapacity or 0)            
        return capacity 
    
    @property
    def percentage_charge_remaining(self):    
        ''' Returns proportion of charge remaining as a float between 0.0 and 1.0 '''
        charge = float(self.remaining_capacity) / float(self.full_charge_capacity)
        return min(charge, 1.0)
        
    @property
    def time_remaining(self):
        ''' Returns time remaining, calculated as for the Windows Battery Meter. It finds
        a value for remaining battery life by dividing the remaining battery capacity 
        by the current battery draining rate as described in the ACPI specification 
        (chapter 3.9.3 'Battery Gas Gauge'). '''
        time_left = 0
        batts = self.t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
        for _i, b in enumerate(batts):
            time_left += float(b.RemainingCapacity) / float(b.DischargeRate)
        hours = int(time_left)
        mins = 60 * (time_left % 1.0)
        return '%i hr %i min' % (hours, mins)

    def should_unplug(self):
        ''' Tests whether conditions are met for unplugging the laptop '''
        return (self.percentage_charge_remaining > self.UNPLUG_LEVEL and
                self.is_plugged_in and
                self.unplug_alert_enabled and
                self.fully_charged_alert_enabled)
        
    def should_plug_in(self):
        ''' Tests whether conditions are met for plugging in the laptop '''
        return (self.percentage_charge_remaining < self.PLUGIN_LEVEL and
                not self.is_plugged_in and
                self.plugin_alert_enabled)
        

ID_SILENCE_FULLY_CHARGED_ALERT = wx.NewId()
ID_SILENCE_UNPLUG_ALERT = wx.NewId()
ID_SILENCE_PLUGIN_ALERT = wx.NewId()

class BatteryTaskBarIcon(BalloonTaskBarIcon):
    ''' Notification area (system tray) icon for output to user about their
        battery status '''
    
    def __init__(self,
                 laptop_batt
                 ):
        wx.TaskBarIcon.__init__(self)
        self.monitor_frequency = 2 # how often to check levels (in seconds)
        self.laptop_batt = laptop_batt
        self.icon = self.BatteryIcon.GetIcon()
        self.SetIcon(self.icon, self.Tooltip)
        
        self.Update()
    
    @property
    def Tooltip(self):
        ''' Generates a tooltip which replicates the Windows Battery Monitor '''
        charge = self.laptop_batt.percentage_charge_remaining * 100
        if self.laptop_batt.is_plugged_in:
            return "%i%% available (plugged in, charging)" % (charge)
        else:
            time = self.laptop_batt.time_remaining
            return "%s (%i%%) remaining" % (time, charge)
        
    @property
    def BatteryIcon(self):
        ''' Returns the appropriate icon for the current charge level and whether
            the laptop is connected to a power supply '''
        charge = self.laptop_batt.percentage_charge_remaining * 100
        charge = int(floor(charge/20)*20) # round down to nearest multiple of 20
        if self.laptop_batt.is_plugged_in:
            return icons.icons["%s%03d" % ("battery_charging_", charge)]
        else:
            return icons.icons["%s%03d" % ("battery_discharging_", charge)]
                    
    def Update(self):
        ''' Checks the fully charged status only once a minute rather than every
            two seconds for all other checks'''
        self.CreateMenu()
        self.ResetAlertsFromPowerStatus()
        self.RefreshIcon()
        self.CheckAlertBalloons()
        self.CheckFullyChargedBalloon()
        wx.CallLater(self.monitor_frequency * 1000, self.Update)
    
    def SilenceUplugAlert(self, e):
        ''' Silences the unplug alert, for use when a full charge is desired '''
        self.laptop_batt.unplug_alert_enabled = False
    
    def SilencePluginAlert(self, e):
        ''' Silences the plugin alert, for use when away from a charging point '''
        self.laptop_batt.plugin_alert_enabled = False        
    
    def SilenceFullyChargedAlert(self, e):
        ''' Silences the plugin alert, for use when away from a charging point '''
        self.laptop_batt.fully_charged_alert_enabled = False        
    
    def CheckFullyChargedBalloon(self):
        ''' Tests if fully charged and fires alert if required '''
        if (self.laptop_batt.is_fully_charged and
            self.laptop_batt.is_plugged_in and
            self.laptop_batt.fully_charged_alert_enabled):
            self.ShowBalloon("Fully charged",
                             "Your battery is now charged to 100%.")
    
    def ResetAlertsFromPowerStatus(self):
        ''' Tests if plugged in enables relevant alerts'''
        if self.laptop_batt.is_plugged_in:
            self.laptop_batt.plugin_alert_enabled = True
        elif not self.laptop_batt.is_plugged_in:
            self.laptop_batt.fully_charged_alert_enabled = True
    
    def CheckAlertBalloons(self):
        charge = self.laptop_batt.percentage_charge_remaining
        if self.laptop_batt.should_unplug():
            self.ShowBalloon("Unplug charger",
                             "Battery charge is at %i%%. Unplug your charger now to maintain battery life." % (charge * 100))
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        if self.laptop_batt.should_plug_in():
            self.ShowBalloon("Plug in charger",
                             "Battery charge is at %i%%. Plug in your charger now to maintain battery life." % (charge * 100))
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
    
    def RefreshIcon(self):
        ''' Sets the appropriate icon depending on power state '''
        self.icon = self.BatteryIcon.GetIcon()
        self.SetIcon(self.icon, self.Tooltip)        
        
    def CreateMenu(self):
        ''' Generates a context-aware menu. The user is only offered the relevant option
            depending on whether their laptop is plugged in '''
        self.Bind(wx.EVT_TASKBAR_RIGHT_UP, self.OnPopup)        
        
        self.menu = wx.Menu()
        if self.laptop_batt.is_plugged_in:
            self.menu.Append(ID_SILENCE_UNPLUG_ALERT, '&Charge to full', 'Fully charge without showing alerts')
            self.Bind(wx.EVT_MENU, self.SilenceUplugAlert, id=ID_SILENCE_UNPLUG_ALERT)
            self.menu.Append(ID_SILENCE_FULLY_CHARGED_ALERT, '&Silence fully-charged alert', 'Wait until next unplugged before resuming alerts')
            self.Bind(wx.EVT_MENU, self.SilenceFullyChargedAlert, id=ID_SILENCE_UNPLUG_ALERT)            
        elif not self.laptop_batt.is_plugged_in:
            self.menu.Append(ID_SILENCE_PLUGIN_ALERT, '&Silence plugin alert', 'Wait until next plugged in before resuming alerts')
            self.Bind(wx.EVT_MENU, self.SilencePluginAlert, id=ID_SILENCE_PLUGIN_ALERT)
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
