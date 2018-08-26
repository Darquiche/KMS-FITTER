from appJar import gui
import platform
import requests
import subprocess
import os
import ctypes

class Gui:
    def __init__(self):

        self.author = "Darqi"
        self.title = "KMS-FITTER"
        self.version = 1.0

        self.width = "480"
        self.height = "630"
        self.isResizable = False
        self.isAdmin = False

        self.GVLK = "https://gist.githubusercontent.com/Darquiche/ef366385a6c6c646a55c970afb46e6a0/raw/"
        self.SERVERS = "https://gist.githubusercontent.com/Darquiche/ba9bd846a4a0a6ccd6e2c6143b2c1330/raw/"

        self.WIN_KEYS = {}
        self.WIN_PRODUCTS = []

        self.OFF_KEYS = {}
        self.OFF_PRODUCTS = []

        self.KMS_HOSTS = []

    # Build the GUI
    def build(self, app):
        win_colspan = 3
        off_colspan = 6

        app.setTitle("{t} by {a} | v{v}".format(t=self.title, a=self.author, v=self.version))
        app.setSize(self.width + "x" + self.height)
        app.setResizable(canResize=self.isResizable)
        app.setStopFunction(self.onExit)
        app.addStatusbar(fields=1)
        app.setStatusbar("", 0)
        app.startTabbedFrame("TabbedFrame")
        app.startTab("Windows")
        # Windows tab
        app.addLabel("l_win_title", "Select a GVLK", colspan=win_colspan)
        app.addOptionBox("opt_win_prod", self.WIN_PRODUCTS, colspan=win_colspan)
        app.setOptionBoxChangeFunction("opt_win_prod", self.opt_changed)
        app.addEntry("e_win_key", colspan=win_colspan)
        app.setEntryDefault("e_win_key", "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")
        app.addWebLink("More Windows keys", "https://www.virtualease.fr/microsoft-officewindows-cles-dactivation-kms-gvlk/", colspan=win_colspan)
        
        app.addHorizontalSeparator(colour="black", colspan=win_colspan)
        
        app.addLabel("l_win_kms", "Select a KMS Host", colspan=win_colspan)
        app.addOptionBox("opt_win_kms", self.KMS_HOSTS, colspan=win_colspan)
        app.setOptionBoxChangeFunction("opt_win_kms", self.opt_changed)
        app.addEntry("e_win_kms", colspan=win_colspan)
        app.setEntryDefault("e_win_kms", "kms.address.here[:port]")
        app.addNamedButton("Activate Windows", "win_activate", self.press, row=9, column=0)
        app.addNamedButton("Get license status", "win_get_info", self.press, row=9, column=1)
        app.stopTab()

        app.startTab("Office")
        # Office tab
        app.addEntry("e_off_path", row=0, colspan=off_colspan-1)
        app.setEntryDefault("e_off_path", "C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS")
        app.addNamedButton(">", "off_get_path", self.press, row=0, column=5)
        app.addLabel("l_office_title", "Select a GVLK")
        app.addOptionBox("opt_off_prod", self.OFF_PRODUCTS, colspan=off_colspan)
        app.setOptionBoxChangeFunction("opt_off_prod", self.opt_changed)
        app.addEntry("e_off_key", colspan=off_colspan)
        app.setEntryDefault("e_off_key", "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX")
        app.addWebLink("More Office keys", "https://github.com/SystemRage/py-kms/wiki/Office-GVLK-Keys/", colspan=off_colspan)

        app.addHorizontalSeparator(colour="black", colspan=off_colspan)

        app.addLabel("l_off_kms", "Select a KMS Host")
        app.addOptionBox("opt_off_kms", self.KMS_HOSTS, colspan=off_colspan)
        app.setOptionBoxChangeFunction("opt_off_kms", self.opt_changed)
        app.addEntry("e_off_kms", colspan=off_colspan)
        app.setEntryDefault("e_off_kms", "kms.address.here[:port]")
        app.addNamedButton("Activate Office", "off_activate", self.press, row=10, column=0, colspan=3)
        app.addNamedButton("Get product infos", "off_get_info", self.press, row=10, column=3, colspan=3)

        app.stopTab()

        app.startTab("About")
        app.addMessage("mess_about", """This application was written by Darqi.
You can use it to activate Windows 8, Windows 8.1, Windows 10, Office 2016 and previous versions without warranty.
== Why doesn't Office accept a GVLK? ==
You'll have to install a volume license (VL) version of Office. Office versions downloaded from MSDN and/or Technet are non-VL.
== Where can I find other keys? ==
That is relatively simple. The GVLKs are published on Microsoft's Technet web site.
These lists only include products that Microsoft sells to corporations via volume license contracts. For Windows there are inofficial GVLKs that work with consumer-only versions of Windows.
== Where can I find other KMS servers ? ==
There are server lists on the internet. But there is no guarantee that these are up! The best option is to create your own emulate server on your NAS with 'vlmcsd'.""")
        app.setMessageAspect("mess_about", 140)
        app.addWebLink("Microsoft Technet - Windows GVLK", "http://technet.microsoft.com/en-us/library/jj612867.aspx")
        app.addWebLink("Microsoft Technet - Office 2010 GVLK", "http://technet.microsoft.com/en-us/library/ee624355(v=office.14).aspx#section2_3")
        app.addWebLink("Microsoft Technet - Office 2013 GVLK", "http://technet.microsoft.com/en-us/library/dn385360.aspx")
        app.addWebLink("SystemRage Github- Windows GVLK", "https://github.com/SystemRage/py-kms/wiki/Windows-GVLK-Keys")
        app.addWebLink("SystemRage Github- Office GVLK", "https://github.com/SystemRage/py-kms/wiki/Office-GVLK-Keys")
        app.addWebLink("Virtualease - Microsoft GVLK", "https://www.virtualease.fr/microsoft-officewindows-cles-dactivation-kms-gvlk/")
        app.addWebLink("CHEF-KOCH - KMS Gist", "https://gist.github.com/CHEF-KOCH/29cac70239eed583ad1c96dcb6de364b")
        app.addWebLink("10ks - KMS Server list", "https://textuploader.com/10ks/raw")
        app.addWebLink("Wind4 - vlmcsd", "https://github.com/Wind4/vlmcsd")
        app.addHorizontalSeparator(colour="black", colspan=off_colspan)
        app.addWebLink("http://darqi.fr", "http://darqi.fr")

        app.stopTab()
        
        app.stopTabbedFrame()

        return app

    # Build and Start the application
    def start(self):

        # Creates a UI
        app = gui(showIcon=False)

        self.getWinGVLK()
        self.getOffGVLK()
        self.getKMS()

        # Run the prebuild method that adds items to the UI
        app = self.build(app)

        # Check admin rights
        if ctypes.windll.shell32.IsUserAnAdmin():
            self.isAdmin = True
        else:
            app.setStatusbarBg("red", 0)
            app.setStatusbarFg("white", 0)
            app.setStatusbar("Need Admin rights!", 0)

        # Make the app class-accessible
        self.app = app

        # Start appJar
        app.go()

    # Callback execute before quitting the app
    def onExit(self):
        return True

    # handle button events
    def press(self, button):
        if button == "win_activate":
            self.activateWindows()

        elif button == "win_get_info":
            self.getWindowsStatus()

        elif button == "off_get_path":
            start_dir = os.environ['PROGRAMFILES']
            path = app.openBox(title="Open OSPP.VBS", dirName=start_dir, fileTypes=[('Script', '*.VBS'), ('Script', '*.vbs')], asFile=False, parent=None)
            
            if "OSPP" not in path:
                pass
                #error
            else:
                app.setEntry("e_off_path", path, callFunction=False)

        elif button == "off_activate":
            self.activateOffice()
        elif button == "off_get_info":
            self.getOfficeStatus()

    def activateWindows(self):
        app = self.app
        value = app.getEntry("e_win_key")
        addr = app.getEntry("e_win_kms")

        if not self.isAdmin:
            #error
            app.errorBox("Admin rights", "Need admin rights!")
            app.stop()

        if len(value) != 29:
            #error
            app.warningBox("GVLK", "Please enter a correct Global Volume License Key!")
            return
        
        if (addr == "") or (len(addr) < 2):
            #error
            app.warningBox("KMS", "Please enter a correct address!")
            return

        subprocess.call('cmd /c slmgr /ipk {v}'.format(v=value))
        subprocess.call('cmd /c slmgr /skms {v}'.format(v=addr))

        act = app.yesNoBox("Windows Activation", "Activate now?")
        if act:
            subprocess.call('cmd /c slmgr /ato')

    def getWindowsStatus(self):
        app = self.app

        if not self.isAdmin:
            #error
            app.errorBox("Admin rights", "Need admin rights!")
            app.stop()

        subprocess.call("cmd /c slmgr /dli")

    def activateOffice(self):
        app = self.app
        path = app.getEntry("e_off_path")
        value = app.getEntry("e_off_key")
        addr = app.getEntry("e_off_kms")

        if not self.isAdmin:
            #error
            app.errorBox("Admin rights", "Need admin rights!")
            app.stop()

        if "OSPP" not in path:
            #error
            app.warningBox("OSPP.VBS", "Please enter a correct path!")
            return

        if len(value) != 29:
            #error
            app.warningBox("GVLK", "Please enter a correct Global Volume License Key!")
            return
        
        if (addr == "") or (len(addr) < 2):
            #error
            app.warningBox("KMS", "Please enter a correct address!")
            return

        subprocess.call('cscript "{p}" /inpkey:{v}'.format(p=path,v=value))

        # Parse address
        if ":" in addr:
            address = addr.split(":")
            addr = address[0]
            port = address[1]
        else:                
            port = "1688"

        subprocess.call('cscript "{p}" /sethst:{v}'.format(p=path,v=addr))
        subprocess.call('cscript "{p}" /setprt:{v}'.format(p=path,v=port))

        act = app.yesNoBox("Office Activation", "Activate now?")
        if act:
            subprocess.call('cscript "{p}" /act'.format(p=path))
        
    def getOfficeStatus(self):
        app = self.app
        path = app.getEntry("e_off_path")

        if not self.isAdmin:
            #error
            app.errorBox("Admin rights", "Need admin rights!")
            app.stop()

        if "OSPP" not in path:
            #error
            app.warningBox("OSPP.VBS", "Please enter a correct path!")
            return

        subprocess.call('cscript "{p}" /dstatus'.format(p=path))        

    # retrieve all GVLK
    def getWinGVLK(self):
        response = requests.get(self.GVLK)
        data = response.text
        lines = data.split("\n")
        for l in lines:
            if "Windows" in l:
                line = l.split("|")
                self.WIN_KEYS[line[1]] = line[0]
                self.WIN_PRODUCTS.append(line[1])
    
    def getOffGVLK(self):
        response = requests.get(self.GVLK)
        data = response.text
        lines = data.split("\n")
        for l in lines:
            if (not "Windows" in l) and ("|" in l):
                line = l.split("|")
                self.OFF_KEYS[line[1]] = line[0]
                self.OFF_PRODUCTS.append(line[1])

    # retrieve all KMS Hosts
    def getKMS(self):
        response = requests.get(self.SERVERS)
        data = response.text
        lines = data.split("\n")
        for l in lines:
            self.KMS_HOSTS.append(l)

    # event function
    def opt_changed(self, opt):
        option = self.app.getOptionBox(opt)
        if opt == "opt_win_prod":
            self.app.setEntry("e_win_key", self.WIN_KEYS[option], callFunction=False)
        elif opt == "opt_win_kms":
            self.app.setEntry("e_win_kms", option, callFunction=False)
        elif opt == "opt_off_prod":
            self.app.setEntry("e_off_key", self.OFF_KEYS[option], callFunction=False)
        elif opt == "opt_off_kms":
            self.app.setEntry("e_off_kms", option, callFunction=False)


if __name__ == "__main__":
    app = Gui()
    app.start()