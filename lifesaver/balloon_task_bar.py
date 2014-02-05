#!/usr/bin/env python
# coding=utf-8
'''
Created on 4 Feb 2014

@author: Jamie Bull
'''

import wx

try:
    import win32gui #, win32con
    WIN32 = True
except:
    WIN32 = False

class BalloonTaskBarIcon(wx.TaskBarIcon):
    """
    Base Taskbar Icon Class
    """
    def __init__(self):
        wx.TaskBarIcon.__init__(self)

    def ShowBalloon(self, title, text, msec = 0, flags = 0):
        """
        Show Balloon tooltip
         @param title - Title for balloon tooltip
         @param msg   - Balloon tooltip text
         @param msec  - Timeout for balloon tooltip, in milliseconds
         @param flags -  one of wx.ICON_INFORMATION, wx.ICON_WARNING, wx.ICON_ERROR
        """
        if WIN32 and self.IsIconInstalled():
            try:
                self.__SetBalloonTip(self.icon.GetHandle(), title, text, msec, flags)
            except Exception:
                pass # print(e) Silent error

    def __SetBalloonTip(self, hicon, title, msg, msec, flags):

        # translate flags
        infoFlags = 0

        if flags & wx.ICON_INFORMATION:
            infoFlags |= win32gui.NIIF_INFO
        elif flags & wx.ICON_WARNING:
            infoFlags |= win32gui.NIIF_WARNING
        elif flags & wx.ICON_ERROR:
            infoFlags |= win32gui.NIIF_ERROR

        # Show balloon
        lpdata = (self.__GetIconHandle(),   # hWnd
                  99,                       # ID
                  win32gui.NIF_MESSAGE|win32gui.NIF_INFO|win32gui.NIF_ICON, # flags: Combination of NIF_* flags
                  0,                        # CallbackMessage: Message id to be pass to hWnd when processing messages
                  hicon,                    # hIcon: Handle to the icon to be displayed
                  '',                       # Tip: Tooltip text
                  msg,                      # Info: Balloon tooltip text
                  msec,                     # Timeout: Timeout for balloon tooltip, in milliseconds
                  title,                    # InfoTitle: Title for balloon tooltip
                  infoFlags                 # InfoFlags: Combination of NIIF_* flags
                  )
        win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, lpdata)

        self.SetIcon(self.icon, self.tooltip)   # Hack: because we have no access to the real CallbackMessage value

    def __GetIconHandle(self):
        """
        Find the icon window.
        This is ugly but for now there is no way to find this window directly from wx
        """
        if not hasattr(self, "_chwnd"):
            try:
                for handle in wx.GetTopLevelWindows():
                    if handle.GetWindowStyle():
                        continue
                    handle = handle.GetHandle()
                    if len(win32gui.GetWindowText(handle)) == 0:
                        self._chwnd = handle
                        break
                if not hasattr(self, "_chwnd"):
                    raise Exception
            except:
                raise Exception, "Icon window not found"
        return self._chwnd

    def SetIcon(self, icon, tooltip = ""):
        self.icon = icon
        self.tooltip = tooltip
        wx.TaskBarIcon.SetIcon(self, icon, tooltip)

    def RemoveIcon(self):
        self.icon = None
        self.tooltip = ""
        wx.TaskBarIcon.RemoveIcon(self)
