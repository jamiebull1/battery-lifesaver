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
import monitor

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
    def __init__(self, frame):
        wx.TaskBarIcon.__init__(self)
        self.frame = frame
        self.monitor_frequency = 2 # how often to check levels (secs)
        self.full_charge_reminder_frequency = 300 # how often to remind that battery is full (secs)
        self.batt_mon = monitor.BatteryMonitor()
        self.icon = self.ChooseIcon.GetIcon()
        self.SetIcon(self.icon, self.Tooltip)
        self.BindEvents()
        self.CreateMenu()
        self.Update()
    
    @property
    def Tooltip(self):
        ''' Generates a tooltip which replicates the Windows Battery Monitor '''
        charge = self.batt_mon.percentage_charge_remaining * 100
        if self.batt_mon.is_plugged_in:
            if self.batt_mon.is_fully_charged:
                tooltip = "Fully charged (100%)"
            else:
                tooltip = "%i%% available (plugged in, charging)" % (charge)
        elif not self.batt_mon.is_plugged_in:
            time_remaining = self.batt_mon.time_remaining
            if not time_remaining is None:
                tooltip = "%s (%i%%) remaining" % (time_remaining, charge)
            else:
                tooltip = "%i%% remaining" % (charge)
        logger.debug("Tooltip is %s" % tooltip)
        return tooltip
    
    @property
    def ChooseIcon(self):
        ''' Returns the appropriate icon for the current charge level and whether
            the laptop is connected to a power supply '''
        charge = self.batt_mon.percentage_charge_remaining * 100
        charge = int(floor(charge/20)*20) # round down to nearest multiple of 20
        if self.batt_mon.is_plugged_in:
            ico = icons.icons["%s%03d" % ("battery_charging_", charge)]
        else:
            ico = icons.icons["%s%03d" % ("battery_discharging_", charge)]
        logger.debug("Icon is %s" % ico)
        self.current_icon = ico
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
        if (self.batt_mon.is_fully_charged and
            self.batt_mon.is_plugged_in and
            self.batt_mon.fully_charged_alert_enabled):
            logger.info("Showing fully charged balloon notification")
            self.ShowBalloon("Fully charged",
                             "Your battery is now charged to 100%.")
    
    def ResetAlertsBasedOnPowerStatus(self):
        ''' Tests if plugged in and resets alerts if required'''
        if self.batt_mon.is_plugged_in:
            logger.info('Plugged in. Resetting stored battery time-remaining values')
            self.batt_mon.reset_time_remaining_queue
            if not self.batt_mon.plugin_alert_enabled:
                logger.info('Plugged in. Resetting plugin alert')
                self.batt_mon.plugin_alert_enabled = True
        elif not self.batt_mon.is_plugged_in:
            if not self.batt_mon.unplug_alert_enabled:            
                logger.info('Not plugged in. Resetting unplug alert')            
                self.batt_mon.unplug_alert_enabled = True
            if not self.batt_mon.fully_charged_alert_enabled:
                logger.info('Not plugged in. Resetting fully charged alert')            
                self.batt_mon.fully_charged_alert_enabled = True
            
        self.menu.Enable(id=ID_SILENCE_FULLY_CHARGED_ALERT,
                         enable=(self.batt_mon.fully_charged_alert_enabled and
                                 self.batt_mon.is_plugged_in)) 
        self.menu.Enable(id=ID_SILENCE_PLUGIN_ALERT,
                         enable=(self.batt_mon.plugin_alert_enabled and
                                 not self.batt_mon.is_plugged_in))
        self.menu.Enable(id=ID_SILENCE_UNPLUG_ALERT,
                         enable=(self.batt_mon.unplug_alert_enabled and
                                 self.batt_mon.is_plugged_in)) 
    
    def CheckAlertBalloons(self):
        charge = self.batt_mon.percentage_charge_remaining
        if self.batt_mon.should_unplug():
            self.ShowBalloon("Unplug charger",
                             "Battery charge is at %i%%. Unplug your charger now to maintain battery life." % (charge * 100))
            logger.info("Showing unplug balloon notification")
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        if self.batt_mon.should_plug_in():
            self.ShowBalloon("Plug in charger",
                             "Battery charge is at %i%%. Plug in your charger now to maintain battery life." % (charge * 100))
            logger.info("Showing plug in balloon notification")
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
    
    def RefreshIcon(self):
        ''' Sets the appropriate icon depending on power state '''
        logger.debug('Refreshing icon')
        self.icon = self.ChooseIcon.GetIcon()
        self.SetIcon(self.icon, self.Tooltip)        
    
    def BindEvents(self):
        ''' Binds the taskbar click events to their event handlers '''
        logger.info("Binding taskbar icon click events")
        self.Bind(wx.EVT_TASKBAR_RIGHT_UP, self.OnPopup)        
        self.Bind(wx.EVT_TASKBAR_LEFT_UP, self.OnLeftClick)  
    
    def CreateMenu(self):
        ''' Generates a context-aware menu. The user is only offered the relevant option
            depending on whether their laptop is plugged in '''
        logger.info("Creating icon menu")
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
        self.batt_mon.fully_charged_alert_enabled = False        
        self.menu.Enable(id=ID_SILENCE_FULLY_CHARGED_ALERT, enable=False) 
        logger.info("Silencing fully charged alert")

    def SilenceUplugAlert(self, e):
        ''' Silences the unplug alert, for use when a full charge is desired '''
        self.batt_mon.unplug_alert_enabled = False
        self.menu.Enable(id=ID_SILENCE_UNPLUG_ALERT, enable=False) 
        logger.info("Silencing unplug alert")
    
    def SilencePluginAlert(self, e):
        ''' Silences the plugin alert, for use when away from a charging point '''
        self.batt_mon.plugin_alert_enabled = False
        self.menu.Enable(id=ID_SILENCE_PLUGIN_ALERT, enable=False) 
        logger.info("Silencing plug in alert")
    
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

    def OnLeftClick(self, event):
        ''' Generates the left click ui '''
        print "fired"
        self.options_window = LeftClickFrame(self.GetParent())

    def OnExit(self, e):
        ''' Removes the icon from the notification area and closes the program '''
        logger.info("Closing application")
        self.frame.Destroy()
        self.Destroy()
        
    def OnAbout(self, e):
        ''' Launches the Battery Lifesaver webpage '''
        webbrowser.open("http://youngbullindustries.wordpress.com/about-battery-lifesaver/")

        
class LeftClickFrame(wx.Frame):
    
    def __init__(self, frame):
        super(LeftClickFrame, self).__init__(frame, style=wx.FRAME_NO_TASKBAR|wx.CAPTION)
        self.parent = frame
        self.tbicon = frame.tbicon
        self.tbicon.options_window = self
        self.WIDTH = 270
        self.InitUI()

    def AlignToBottomCentre(self):
        logger.debug("Aligning left click UI")
        dw, dh = wx.GetMousePosition()
        w, h = self.GetSize()
        x = dw - w/2
        y = dh - h - 25
        self.SetPosition((x, y))

    def InitUI(self):
        logger.info("Initialising left click UI")
            
        self.SetSize(wx.Size(self.WIDTH,350))
        self.AlignToBottomCentre()
        self.panel = wx.Panel(self, wx.ID_ANY) 
        self.panel.SetBackgroundColour('white')
        self.Bind(wx.EVT_ACTIVATE, self.Close)  #bind it to Frame
        self.PopulateUI()
        
        self.Show()
        self.Raise()
    
    def Close(self, e):
        if e.GetActive() != True:
            self.tbicon.options_window.Destroy()

    def PopulateUI(self):
        logger.info("Populating left click UI")
        self.main_vbox = wx.BoxSizer(wx.VERTICAL)
        
        top_hbox = self.GetTopHBox()
        self.main_vbox.Add(top_hbox)
        self.main_vbox.Add((-1, 10))
        
        statuses_hbox = self.GetStatusesVBox()
        self.main_vbox.Add(statuses_hbox)
        self.main_vbox.Add((-1, 10))
        
        plans_vbox = self.GetPowerPlansVBox()
        self.main_vbox.Add(plans_vbox, flag=wx.LEFT, border=20)

        links_vbox = self.GetLinksVBox()
        self.main_vbox.Add(links_vbox, flag=wx.TOP|wx.CENTER, border=20)
        self.main_vbox.Add((-1, 20))

        self.panel.SetSizer(self.main_vbox)
     
    def GetTopHBox(self):
        icon_vbox = self.RetrieveCurrentTaskbarIcon()
        summary_status_vbox = self.RetrievePowerStatusSummary()
        top_hbox = wx.BoxSizer(wx.HORIZONTAL)
        top_hbox.Add((20, -1))
        top_hbox.Add(icon_vbox, flag=wx.LEFT)
        top_hbox.Add((10, -1))
        top_hbox.Add(summary_status_vbox, flag=wx.RIGHT)
        return top_hbox

    def RetrieveCurrentTaskbarIcon(self):
        logger.debug("Setting up battery icon")
        icon = self.tbicon.current_icon
        icon_vbox = wx.BoxSizer(wx.VERTICAL)
        pic = wx.StaticBitmap(self.panel)
        pic.SetBitmap(icon.GetBitmap()) 
        icon_vbox.Add(pic, flag=wx.LEFT|wx.TOP, border=10)
        return icon_vbox
     
    def RetrievePowerStatusSummary(self):
        logger.debug("Setting up power status text")
        tooltip = self.tbicon.Tooltip
        summary_status_vbox = wx.BoxSizer(wx.VERTICAL)
        txt = wx.StaticText(self.panel, wx.ID_ANY, tooltip)
        txt.Wrap(self.WIDTH - 100)
        summary_status_vbox.Add(txt, flag=wx.RIGHT|wx.TOP, border=10)
        return summary_status_vbox
     
    def GetStatusesVBox(self):
        logger.debug("Setting up battery statuses")
        statuses = self.tbicon.batt_mon.battery_statuses
        
        statuses_vbox = wx.BoxSizer(wx.VERTICAL)
        for status in statuses:
            battery_statuses_txt = wx.StaticText(self.panel, wx.ID_ANY, status)
            statuses_vbox.Add(battery_statuses_txt, flag=wx.TOP|wx.LEFT, border=10)
        return statuses_vbox
     
    def GetLinksVBox(self):
        logger.debug("Setting links")
        links_vbox = wx.BoxSizer(wx.VERTICAL)
        link1 = wx.HyperlinkCtrl(self.panel, wx.ID_ANY, 'Adjust screen brightness')
        link1.Bind(wx.EVT_HYPERLINK, self.tbicon.LaunchPowerOptions)
        links_vbox.Add(link1, flag=wx.CENTER|wx.TOP, border=5)
        
        link2 = wx.HyperlinkCtrl(self.panel, wx.ID_ANY, 'More power options')
        link2.Bind(wx.EVT_HYPERLINK, self.tbicon.LaunchPowerOptions)
        links_vbox.Add(link2, flag=wx.CENTER|wx.TOP, border=5)
        return links_vbox

    def GetPowerPlansVBox(self):
        logger.debug("Setting up power plan radio buttons")
        plans_vbox = wx.BoxSizer(wx.VERTICAL)
        txt = wx.StaticText(self.panel, wx.ID_ANY, 'Select a power plan:')
        txt.SetForegroundColour('gray')
        plans_vbox.Add(txt, flag=wx.LEFT)
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
        plans_vbox.Add(sizer, flag=wx.TOP|wx.LEFT)
        return plans_vbox
                 
    def ActivatePowerPlan(self, e):
        name = e.EventObject.GetLabel()
        logger.info("Activating power plan %s" % name)
        guid = self.guids[self.names.index(name)]
        os.system('powercfg -setactive %s' % guid)   
     

class TaskBarFrame(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, style=wx.FRAME_NO_TASKBAR)
        self.tbicon = BatteryTaskBarIcon(self)
        wx.EVT_TASKBAR_LEFT_UP(self.tbicon, self.OnTaskBarLeftClick)

    def OnTaskBarLeftClick(self, evt):
        ''' Creates/destroys left click menu '''
        self.left_frame = LeftClickFrame(self)


def main():

    app = wx.App(False)
    TaskBarFrame(None, "TaskBarFrame")
    app.MainLoop()
    
if __name__ == '__main__':
    main()
