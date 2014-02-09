#!/usr/bin/env python
# coding=utf-8
'''
Created on 4 Feb 2014

@author: Jamie Bull - Young Bull Industries
'''

import os
import re
import webbrowser
import wmi
import winsound
from math import floor
import platform

import wx

import icons

# Logging setup
import logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)
debug_handler = logging.FileHandler('bl.debug.log')
debug_handler.setLevel(logging.DEBUG)
info_handler = logging.FileHandler('bl.info.log')
info_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
debug_handler.setFormatter(formatter)
info_handler.setFormatter(formatter)
logger.addHandler(debug_handler)
logger.addHandler(info_handler)


VERSION_NUMBER = '0.0.6-beta'

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
ID_SCREEN_BRIGHTNESS = wx.NewId()
ID_POWER_OPTIONS = wx.NewId()
ID_MOBILITY_CENTER = wx.NewId()
ID_NOTIFICATION_ICONS = wx.NewId()

class BatteryTaskBarIcon(wx.TaskBarIcon):
    ''' Notification area (system tray) icon for output to user about their
        battery status '''
    
    def __init__(self):
        wx.TaskBarIcon.__init__(self)
        self.monitor_frequency = 2 # how often to check levels (secs)
        self.full_charge_reminder_frequency = 300 # how often to remind that battery is full (secs)
        self.batteries = BatteryMonitor()
        self.options_window = PowerStatusAndPlansWindow(self)
        self.icon = self.BatteryIcon.GetIcon()
        self.SetIcon(self.icon, self.Tooltip)
        self.CreateMenu()
        self.Update()
    
    @property
    def Tooltip(self):
        ''' Generates a tooltip which replicates the Windows Battery Monitor '''
        charge = self.batteries.percentage_charge_remaining * 100
        if self.batteries.is_plugged_in:
            if self.batteries.is_fully_charged:
                tooltip = "Fully charged (100%)"
            else:
                tooltip = "%i%% available (plugged in, charging)" % (charge)
        elif not self.batteries.is_plugged_in:
            time_remaining = self.batteries.time_remaining
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
        charge = self.batteries.percentage_charge_remaining * 100
        charge = int(floor(charge/20)*20) # round down to nearest multiple of 20
        if self.batteries.is_plugged_in:
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
    
    def CheckFullyChargedBalloon(self):
        ''' Tests if fully charged and fires alert if required '''
        if (self.batteries.is_fully_charged and
            self.batteries.is_plugged_in and
            self.batteries.fully_charged_alert_enabled):
            logging.info("Showing fully charged balloon notification")
            self.ShowBalloon("Fully charged",
                             "Your battery is now charged to 100%.")
    
    def ResetAlertsBasedOnPowerStatus(self):
        ''' Tests if plugged in and resets alerts if required'''
        if self.batteries.is_plugged_in:
            logger.info('Plugged in. Resetting stored battery time-remaining values')
            self.batteries.reset_time_remaining_queue
            if not self.batteries.plugin_alert_enabled:
                logger.info('Plugged in. Resetting plugin alert')
                self.batteries.plugin_alert_enabled = True
        elif not self.batteries.is_plugged_in:
            if not self.batteries.unplug_alert_enabled:            
                logger.info('Not plugged in. Resetting unplug alert')            
                self.batteries.unplug_alert_enabled = True
            if not self.batteries.fully_charged_alert_enabled:
                logger.info('Not plugged in. Resetting fully charged alert')            
                self.batteries.fully_charged_alert_enabled = True
            
        self.menu.Enable(id=ID_SILENCE_FULLY_CHARGED_ALERT,
                         enable=(self.batteries.fully_charged_alert_enabled and
                                 self.batteries.is_plugged_in)) 
        self.menu.Enable(id=ID_SILENCE_PLUGIN_ALERT,
                         enable=(self.batteries.plugin_alert_enabled and
                                 not self.batteries.is_plugged_in))
        self.menu.Enable(id=ID_SILENCE_UNPLUG_ALERT,
                         enable=(self.batteries.unplug_alert_enabled and
                                 self.batteries.is_plugged_in)) 
    
    def CheckAlertBalloons(self):
        charge = self.batteries.percentage_charge_remaining
        if self.batteries.should_unplug():
            self.ShowBalloon("Unplug charger",
                             "Battery charge is at %i%%. Unplug your charger now to maintain battery life." % (charge * 100))
            logging.info("Showing unplug balloon notification")
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        if self.batteries.should_plug_in():
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
        # Alert buttons
        self.menu.Append(ID_SILENCE_FULLY_CHARGED_ALERT, 'Silence &Full Charge Alert', 'Continue not showing alerts')
        self.Bind(wx.EVT_MENU, self.SilenceFullyChargedAlert, id=ID_SILENCE_FULLY_CHARGED_ALERT)
        self.menu.Append(ID_SILENCE_PLUGIN_ALERT, 'Silence &Plugin Alert', 'Wait until next plugged in before resuming alerts')
        self.Bind(wx.EVT_MENU, self.SilencePluginAlert, id=ID_SILENCE_PLUGIN_ALERT)
        self.menu.Append(ID_SILENCE_UNPLUG_ALERT, 'Silence &Unplug Alert', 'Fully charge without showing alerts')
        self.Bind(wx.EVT_MENU, self.SilenceUplugAlert, id=ID_SILENCE_UNPLUG_ALERT)
        
        # Replicated Windows Power Widget options here
        self.menu.AppendSeparator()
        self.menu.Append(ID_SCREEN_BRIGHTNESS, 'Adjust screen brightness', 'Launch Power Options dialogue')
        self.Bind(wx.EVT_MENU, self.LaunchPowerOptions, id=ID_SCREEN_BRIGHTNESS)
        self.menu.Append(ID_POWER_OPTIONS, 'Power Options', 'Launch Power Options dialogue')
        self.Bind(wx.EVT_MENU, self.LaunchPowerOptions, id=ID_POWER_OPTIONS)
        self.menu.Append(ID_MOBILITY_CENTER, 'Windows Mobility Center', 'Launch Windows Mobility Center dialogue')
        self.Bind(wx.EVT_MENU, self.LaunchMobilityCenter, id=ID_MOBILITY_CENTER)
        self.menu.AppendSeparator()
        self.menu.Append(ID_NOTIFICATION_ICONS, 'Turn system icons on or off', 'Launch Windows Notification Area Icons Options dialogue')
        self.Bind(wx.EVT_MENU, self.LaunchNotificationAreaIconsOptions, id=ID_NOTIFICATION_ICONS)
        # About and Exit options
        self.menu.AppendSeparator()
        self.menu.Append(wx.ID_ABOUT, '&Website', 'About this program')
        self.Bind(wx.EVT_MENU, self.OnAbout, id=wx.ID_ABOUT)
        self.menu.Append(wx.ID_EXIT, 'E&xit', 'Remove icon and quit application')
        self.Bind(wx.EVT_MENU, self.OnExit, id=wx.ID_EXIT)
    
    def SilenceFullyChargedAlert(self, e):
        ''' Silences the full charge alert, for use when not ready to leave charging point '''
        self.batteries.fully_charged_alert_enabled = False        
        self.menu.Enable(id=ID_SILENCE_FULLY_CHARGED_ALERT, enable=False) 
        logging.info("Silencing fully charged alert")

    def SilenceUplugAlert(self, e):
        ''' Silences the unplug alert, for use when a full charge is desired '''
        self.batteries.unplug_alert_enabled = False
        self.menu.Enable(id=ID_SILENCE_UNPLUG_ALERT, enable=False) 
        logging.info("Silencing unplug alert")
    
    def SilencePluginAlert(self, e):
        ''' Silences the plugin alert, for use when away from a charging point '''
        self.batteries.plugin_alert_enabled = False
        self.menu.Enable(id=ID_SILENCE_PLUGIN_ALERT, enable=False) 
        logging.info("Silencing plug in alert")
    
    def LaunchPowerOptions(self, e):
        ''' Opens the Control Panel Power Options dialogue '''
        try:
            os.system('control.exe /name Microsoft.PowerOptions')
        except:
            os.system('control.exe main.cpl power')
            
    def LaunchNotificationAreaIconsOptions(self, e):
        ''' Opens the Control Panel Notification Area Icons Options dialogue '''
        os.system('control.exe /name Microsoft.NotificationAreaIcons')

    def LaunchMobilityCenter(self, e):
        ''' Opens the Windows Mobility Center dialogue '''
        windir = os.environ['WINDIR']
        os.system(os.path.join(windir, 'Sysnative\mblctr.exe'))
        
    def OnPopup(self, event):
        ''' Generates the right click menu '''
        self.PopupMenu(self.menu)

    def OnExit(self, e):
        ''' Removes the icon from the notification area and closes the program '''
        logger.info("Closing application")
        self.options_window.Destroy()
        self.Destroy()
        self.GetParent.Destroy()
        
    def OnAbout(self, e):
        ''' Launches the Battery Lifesaver webpage '''
        webbrowser.open("http://youngbullindustries.wordpress.com/about-battery-lifesaver/")

        
class PowerStatusAndPlansWindow(wx.Frame):
    
    def __init__(self, tbicon):
        frame = wx.Frame(None)
        super(PowerStatusAndPlansWindow, self).__init__(frame, style=wx.FRAME_NO_TASKBAR|wx.CAPTION)
        self.tbicon = tbicon
        self.tbicon.Bind(wx.EVT_TASKBAR_LEFT_UP, self.ToggleVisibility)  
        self.is_visible = False
        self.SetSize(wx.Size(270,350))
        self.AlignToBottomCentre()
        self.panel = wx.Panel(self, wx.ID_ANY) 
        self.panel.SetBackgroundColour('white')
        self.tbicon.Bind(wx.EVT_KILL_FOCUS, lambda e: self.ToggleVisibility("Bound to window"))      
        self.Bind(wx.EVT_KILL_FOCUS, lambda e: self.ToggleVisibility("Bound to window"))      
        self.panel.Bind(wx.EVT_KILL_FOCUS, lambda e: self.ToggleVisibility("Bound to panel"))      
        wx.EVT_KILL_FOCUS(self.panel, lambda e: self.ToggleVisibility("Bound to panel (old way)")) 
        self.InitUI()

    def AlignToBottomCentre(self):
        dw, dh = wx.GetMousePosition()
        w, h = self.GetSize()
        x = dw - w/2
        y = dh - h - 25
        self.SetPosition((x, y))

    def ToggleVisibility(self, e):
#        import pdb; pdb.set_trace()
        self.is_visible = not self.is_visible
        if self.is_visible == True:
            self.AlignToBottomCentre()
            self.Show()
            self.Raise()
        else:
            self.Hide()
              
    def InitUI(self):
        self.GetDataForUI()
        self.PopulateUI()
#        self.Show()
     
    def PopulateUI(self):
        self.main_vbox = wx.BoxSizer(wx.VERTICAL)
        self.top_hbox = wx.BoxSizer(wx.HORIZONTAL)
         
        self.top_hbox.Add((20, -1))
        self.top_hbox.Add(self.icon_vbox, flag=wx.LEFT)
        self.top_hbox.Add((10, -1))
        self.top_hbox.Add(self.summary_vbox, flag=wx.RIGHT)
 
        self.main_vbox.Add(self.top_hbox)
        self.main_vbox.Add((-1, 10))
        self.main_vbox.Add(self.statuses_hbox)
         
        self.main_vbox.Add((-1, 10))
         
        self.main_vbox.Add(self.plans_vbox, flag=wx.LEFT, border=20)    
        self.main_vbox.Add(self.links_vbox, flag=wx.TOP|wx.CENTER, border=20)
        self.panel.SetSizer(self.main_vbox)
     
    def GetDataForUI(self):
        self.BatteryIcon(self.tbicon.BatteryIcon)
        self.SetPowerStatusText(self.tbicon.Tooltip)
        self.SetBatteryStatuses(self.tbicon.batteries.battery_statuses)
        self.SetPowerPlanOptions()
        self.SetLinks()
     
    def BatteryIcon(self, icon):
        self.icon_vbox = wx.BoxSizer(wx.VERTICAL)
        self.pic = wx.StaticBitmap(self.panel)
        self.pic.SetBitmap(icon.GetBitmap()) 
        self.icon_vbox.Add(self.pic, flag=wx.LEFT|wx.TOP, border=10)
     
    def SetPowerStatusText(self, text):
        self.summary_vbox = wx.BoxSizer(wx.VERTICAL)
        self.power_status_txt = wx.StaticText(self.panel, wx.ID_ANY, text)
        self.summary_vbox.Add(self.power_status_txt, flag=wx.RIGHT|wx.TOP, border=10)
     
    def SetBatteryStatuses(self, statuses):
        self.statuses_hbox = wx.BoxSizer(wx.HORIZONTAL)
        for status in statuses:
            self.battery_statuses_txt = wx.StaticText(self.panel, wx.ID_ANY, status)
            self.statuses_hbox.Add(self.battery_statuses_txt, flag=wx.TOP|wx.LEFT, border=10)
     
    def SetPowerPlanOptions(self):
        self.plans_vbox = wx.BoxSizer(wx.VERTICAL)
        txt = wx.StaticText(self.panel, wx.ID_ANY, 'Select a power plan:')
        txt.SetForegroundColour('gray')
        self.plans_vbox.Add(txt, flag=wx.LEFT)
        plans = wmi.WMI(moniker = "//./root/cimv2/power").Win32_PowerPlan()
        self.names = [p.ElementName for p in plans]
        self.guids = [re.findall('\{(.*?)\}', p.InstanceID)[0] for p in plans]
        self.plan_is_active = [-1 if p.IsActive else 1 for p in plans]
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.radios = []
        for i, name in enumerate(self.names):
            if i == 0:
                radio = wx.RadioButton(self.panel,
                                       label=name,
                                       style = wx.RB_GROUP)
            else:
                radio = wx.RadioButton(self.panel,
                                   label=name)
            if self.plan_is_active[i] == -1:
                radio.SetValue(True)
            self.radios += [radio]
            sizer.Add(self.radios[i], border=5) 
        self.Bind(wx.EVT_RADIOBUTTON, self.ActivatePowerPlan)
        self.plans_vbox.Add(sizer, flag=wx.TOP|wx.LEFT)
                 
    def ActivatePowerPlan(self, e):
        name = e.EventObject.GetLabel()
        guid = self.guids[self.names.index(name)]
        os.system('powercfg -setactive %s' % guid)   
     
    def SetLinks(self):
        self.links_vbox = wx.BoxSizer(wx.VERTICAL)
        link1 = wx.HyperlinkCtrl(self.panel, wx.ID_ANY, 'Adjust screen brightness')
        link1.Bind(wx.EVT_HYPERLINK, self.tbicon.LaunchPowerOptions)
        self.links_vbox.Add(link1, flag=wx.CENTER|wx.TOP, border=5)
        
        link2 = wx.HyperlinkCtrl(self.panel, wx.ID_ANY, 'More power options')
        link2.Bind(wx.EVT_HYPERLINK, self.tbicon.LaunchPowerOptions)
        self.links_vbox.Add(link2, flag=wx.CENTER|wx.TOP, border=5)


def main():
    
    app = wx.App(False)
    BatteryTaskBarIcon()
    app.MainLoop()

if __name__ == '__main__':
    main()
