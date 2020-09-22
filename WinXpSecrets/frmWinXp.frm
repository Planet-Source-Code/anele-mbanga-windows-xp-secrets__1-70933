VERSION 5.00
Begin VB.Form frmWinXp 
   Caption         =   "Windows Xp Secrets"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWinXp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin WindowsXpSecrets.TreeList TreeList1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3836
      OuterAppearance =   1
      ListViewGridLines=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TreeViewHideSelection=   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
      End
      Begin VB.Menu xx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmWinXp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error Resume Next
    TreeList1.TreeViewSorted = True
    TreeList1.Headings = "Description"
    Dim treeItems(20) As MSComctlLib.Node
    Dim msc(50) As MSComctlLib.Node
    
    Set msc(1) = TreeList1.TreeViewAddPath("Security\Certificates")
    msc(1).Tag = "run-certmgr.msc"
    TreeList1.ListViewAddItem msc(1).FullPath, "cert", "See Certificates Help for an overview."
    
    Set msc(2) = TreeList1.TreeViewAddPath("Services\Indexing Service")
    msc(2).Tag = "run-ciadv.msc"
    TreeList1.ListViewAddItem msc(2).FullPath, "is", "See Indexing Services Help for an overview."
    
    Set msc(3) = TreeList1.TreeViewAddPath("Management\Computer Management")
    msc(3).Tag = "run-compmgmt.msc"
    TreeList1.ListViewAddItem msc(3).FullPath, "is", "See Computer Management Help for an overview."
    
    Set msc(4) = TreeList1.TreeViewAddPath("Hardware\Device Manager")
    msc(4).Tag = "run-devmgmt.msc"
    TreeList1.ListViewAddItem msc(4).FullPath, "is", "See Device Manager Help for an overview."
    
    Set msc(5) = TreeList1.TreeViewAddPath("Storage\Disk Defragmenter")
    msc(5).Tag = "run-dfrg.msc"
    TreeList1.ListViewAddItem msc(5).FullPath, "is", "See Disk Defragmenter Help for an overview."
    
    Set msc(6) = TreeList1.TreeViewAddPath("Storage\Disk Management")
    msc(6).Tag = "run-diskmgmt.msc"
    TreeList1.ListViewAddItem msc(6).FullPath, "is", "See Disk Management Help for an overview."
    
    Set msc(7) = TreeList1.TreeViewAddPath("Management\Event Viewer")
    msc(7).Tag = "run-eventvwr.msc"
    TreeList1.ListViewAddItem msc(7).FullPath, "is", "See Event Viewer Help for an overview."
    
    Set msc(8) = TreeList1.TreeViewAddPath("Storage\Shared Folders")
    msc(8).Tag = "run-fsmgmt.msc"
    TreeList1.ListViewAddItem msc(8).FullPath, "is", "See Shared Folders Help for an overview."
    
    Set msc(9) = TreeList1.TreeViewAddPath("Security\Local Users And Groups")
    msc(9).Tag = "run-lusrmgr.msc"
    TreeList1.ListViewAddItem msc(9).FullPath, "is", "See Local Users And Groups Help for an overview."
    
    Set msc(10) = TreeList1.TreeViewAddPath("Storage\Removable Storage")
    msc(10).Tag = "run-ntmsmgr.msc"
    TreeList1.ListViewAddItem msc(10).FullPath, "is", "See Removable Storage Help for an overview."
    
    Set msc(11) = TreeList1.TreeViewAddPath("Storage\Removable Storage Operator Requests")
    msc(11).Tag = "run-ntmsoprq.msc"
    TreeList1.ListViewAddItem msc(11).FullPath, "is", "See Removable Storage Operator Requests Help for an overview."
    
    Set msc(12) = TreeList1.TreeViewAddPath("Management\Performance")
    msc(12).Tag = "run-perfmon.msc"
    TreeList1.ListViewAddItem msc(12).FullPath, "is", "See Performance Help for an overview."
    
    Set msc(13) = TreeList1.TreeViewAddPath("Management\Resultant Set of Policy")
    msc(13).Tag = "run-rsop.msc"
    TreeList1.ListViewAddItem msc(13).FullPath, "is", "See Resultant Set of Policy Help for an overview."
    
    Set msc(14) = TreeList1.TreeViewAddPath("Security\Local Security Settings")
    msc(14).Tag = "run-secpol.msc"
    TreeList1.ListViewAddItem msc(14).FullPath, "is", "See Local Security Settings Help for an overview."
    
    Set msc(15) = TreeList1.TreeViewAddPath("Management\Group Policy")
    msc(15).Tag = "run-gpedit.msc"
    TreeList1.ListViewAddItem msc(15).FullPath, "is", "See Group Policy Help for an overview."
    
    Set msc(16) = TreeList1.TreeViewAddPath("Databases\SQL Server Configuation Manager")
    msc(16).Tag = "run-SQLServerManager.msc"
    TreeList1.ListViewAddItem msc(16).FullPath, "is", "See SQL Server Configuation Manager Help for an overview."
    
    Set msc(17) = TreeList1.TreeViewAddPath("Management\Windows Management Infrastructure")
    msc(17).Tag = "run-wmimgmt.msc"
    TreeList1.ListViewAddItem msc(17).FullPath, "is", "See Windows Management Infrastructure Help for an overview."
    
    Set msc(18) = TreeList1.TreeViewAddPath("Management\.Net Configuration")
    msc(18).Tag = "run-c:\windows\servicepackfiles\i386\mscorcfg.msc"
    TreeList1.ListViewAddItem msc(18).FullPath, "is", "See .Net Configuration Help for an overview."
    
    Set msc(19) = TreeList1.TreeViewAddPath("Services\Component Services")
    msc(19).Tag = "run-c:\windows\system32\com\comexp.msc"
    TreeList1.ListViewAddItem msc(19).FullPath, "is", "See Component Services Help for an overview."
    
    Set msc(20) = TreeList1.TreeViewAddPath("Services\Internet Information Services")
    msc(20).Tag = "run-c:\windows\system32\inetsrv\iis.msc"
    TreeList1.ListViewAddItem msc(20).FullPath, "is", "See Internet Information Services Help for an overview."
    
    ''Set msc(21) = TreeList1.TreeViewAddPath("Management\.Net Framework 1.1 Configuration")
    'msc(21).Tag = "run-mscorcfg.msc"
    'TreeList1.ListViewAddItem msc(21).FullPath, "is", "See .Net Framework 1.1 Configuration Help for an overview."
    
    'Set msc(22) = TreeList1.TreeViewAddPath("Applications\Frontpage Server Extensions")
    'msc(22).Tag = "run-fpmmc.msc"
    'TreeList1.ListViewAddItem msc(22).FullPath, "is", "See Frontpage Server Extensions Help for an overview."
    
    Set msc(23) = TreeList1.TreeViewAddPath("Databases\SQL Server Client Network Utility")
    msc(23).Tag = "run-cliconfg.exe"
    TreeList1.ListViewAddItem msc(23).FullPath, "is", "See SQL Server Client Network Utility Help for an overview."
    
    Set msc(24) = TreeList1.TreeViewAddPath("Applications\DDE Share")
    msc(24).Tag = "run-ddeshare.exe"
    TreeList1.ListViewAddItem msc(24).FullPath, "is", "See DDE Share Help for an overview."
    
    Set msc(25) = TreeList1.TreeViewAddPath("Applications\DirectX Diagnostic Tool")
    msc(25).Tag = "run-dxdiag.exe"
    TreeList1.ListViewAddItem msc(25).FullPath, "is", "See DirectX Diagnostic Tool Help for an overview."
    
    Set msc(26) = TreeList1.TreeViewAddPath("Storage\Disk Cleanup")
    msc(26).Tag = "run-cleanmgr"
    TreeList1.ListViewAddItem msc(26).FullPath, "is", "See Disk Cleanup Help for an overview."
    
    Set msc(27) = TreeList1.TreeViewAddPath("Applications\Bluetooth File Transfer")
    msc(27).Tag = "run-fsquirt.exe"
    TreeList1.ListViewAddItem msc(27).FullPath, "is", "See Bluetooth File Transfer Help for an overview."
    
    Set msc(28) = TreeList1.TreeViewAddPath("Applications\Fax Console")
    msc(28).Tag = "run-fxsclnt.exe"
    TreeList1.ListViewAddItem msc(28).FullPath, "is", "See Fax Console Help for an overview."
    
    Set msc(29) = TreeList1.TreeViewAddPath("Applications\Fax Send Wizard")
    msc(29).Tag = "run-fxssend.exe"
    TreeList1.ListViewAddItem msc(29).FullPath, "is", "See Fax Send Wizard Help for an overview."
    
    Set msc(30) = TreeList1.TreeViewAddPath("Applications\Java Application Cache Viewer")
    msc(30).Tag = "run-javaws.exe"
    TreeList1.ListViewAddItem msc(30).FullPath, "is", "See Java Application Cache Viewer Help for an overview."
    
    Set msc(31) = TreeList1.TreeViewAddPath("Applications\Remote Desktop Connection")
    msc(31).Tag = "run-mstsc.exe"
    TreeList1.ListViewAddItem msc(31).FullPath, "is", "See Remote Desktop Connection Help for an overview."
    
    Set msc(32) = TreeList1.TreeViewAddPath("Applications\Narrator")
    msc(32).Tag = "run-narrator.exe"
    TreeList1.ListViewAddItem msc(32).FullPath, "is", "See Narrator Help for an overview."
    
    Set msc(33) = TreeList1.TreeViewAddPath("Applications\Network Setup Wizard")
    msc(33).Tag = "run-netsetup.exe"
    TreeList1.ListViewAddItem msc(33).FullPath, "is", "See Network Setup Wizard Help for an overview."
    
    Set msc(34) = TreeList1.TreeViewAddPath("Storage\Backup Restore Wizard")
    msc(34).Tag = "run-ntbackup.exe"
    TreeList1.ListViewAddItem msc(34).FullPath, "is", "See Backup Restore Wizard Help for an overview."
    
    Set msc(35) = TreeList1.TreeViewAddPath("Databases\ODBC Data Source Administrator")
    msc(35).Tag = "run-odbcad32.exe"
    TreeList1.ListViewAddItem msc(35).FullPath, "is", "See ODBC Data Source Administrator Help for an overview."
    
    Set msc(36) = TreeList1.TreeViewAddPath("Applications\On Screen Keyboard")
    msc(36).Tag = "run-osk.exe"
    TreeList1.ListViewAddItem msc(36).FullPath, "is", "See On Screen Keyboard Help for an overview."
    
    Set msc(37) = TreeList1.TreeViewAddPath("Applications\Remote Automation Connection Manager")
    msc(37).Tag = "run-RACMGR32.EXE"
    TreeList1.ListViewAddItem msc(37).FullPath, "is", "See Remote Automation Connection Manager Help for an overview."
    
    Set msc(38) = TreeList1.TreeViewAddPath("Trace\Trace Settings")
    msc(38).Tag = "run-regtrace.exe"
    TreeList1.ListViewAddItem msc(38).FullPath, "is", "See Trace Settings Help for an overview."
    
    Set msc(39) = TreeList1.TreeViewAddPath("Applications\Phone Dialer")
    msc(39).Tag = "run-dialer"
    TreeList1.ListViewAddItem msc(39).FullPath, "is", "See Phone Dialer Help for an overview."
    
    Set msc(40) = TreeList1.TreeViewAddPath("Applications\Create A Shared Folder Wizard")
    msc(40).Tag = "run-shrpubw"
    TreeList1.ListViewAddItem msc(40).FullPath, "is", "See Create A Shared Folder Wizard Help for an overview."
    
    Set msc(41) = TreeList1.TreeViewAddPath("Applications\File Signature Verification")
    msc(41).Tag = "run-sigverif"
    TreeList1.ListViewAddItem msc(41).FullPath, "is", "See File Signature Verification Help for an overview."
    
    Set msc(42) = TreeList1.TreeViewAddPath("Applications\Address Book")
    msc(42).Tag = "run-wab"
    TreeList1.ListViewAddItem msc(42).FullPath, "is", "See Address Book Help for an overview."
    
    Set msc(43) = TreeList1.TreeViewAddPath("Security\Securing the Windows XP Database")
    msc(43).Tag = "run-syskey"
    TreeList1.ListViewAddItem msc(43).FullPath, "is", "See Securing the Windows XP Database Help for an overview."
    
    Set msc(44) = TreeList1.TreeViewAddPath("Storage\Driver Verifier Manager")
    msc(44).Tag = "run-verifier"
    TreeList1.ListViewAddItem msc(44).FullPath, "is", "See Driver Verifier Manager Help for an overview."
    
    Set msc(45) = TreeList1.TreeViewAddPath("Network\Chat")
    msc(45).Tag = "run-winchat"
    TreeList1.ListViewAddItem msc(45).FullPath, "is", "See Chat Help for an overview."
    
    Set msc(46) = TreeList1.TreeViewAddPath("Databases\SQL Server Client Configuration Utility")
    msc(46).Tag = "run-WINDBVER"
    TreeList1.ListViewAddItem msc(46).FullPath, "is", "See SQL Server Client Configuration Utility Help for an overview."
    
    Set msc(47) = TreeList1.TreeViewAddPath("Management\Windows Update")
    msc(47).Tag = "run-wupdmgr"
    TreeList1.ListViewAddItem msc(47).FullPath, "is", "See Windows Update Help for an overview."
    
    
    Set treeItems(1) = TreeList1.TreeViewAddPath("Tweaks\Group Policy")
    treeItems(1).Tag = "run-gpedit.msc"
    TreeList1.ListViewAddItem treeItems(1).FullPath, "gpedit.msc", "1. For example, if you hate CD autoplay like I do and want to permanently disable it, you can use this tool to do so."
    TreeList1.ListViewAddItem treeItems(1).FullPath, "gpedit.msc.1", "2. Just run gpedit.msc, then go to Computer Configuration -> Administrative Templates -> System."
    TreeList1.ListViewAddItem treeItems(1).FullPath, "gpedit.msc.2", "3. In here you can see the value 'Turn Off Autoplay'. Right-click on it and then click 'Properties'."
    
    Set treeItems(2) = TreeList1.TreeViewAddPath("Tweaks\System Configuration Utility")
    treeItems(2).Tag = "run-msconfig"
    TreeList1.ListViewAddItem treeItems(2).FullPath, "msconfig", "1. This displays all of the programs that will be started when Windows boots up."
    TreeList1.ListViewAddItem treeItems(2).FullPath, "msconfig.1", "2. If you uncheck some boxes, windows should start up faster and will take less resources by not running these programs in the background."
    
    Set treeItems(3) = TreeList1.TreeViewAddPath("Services\Services")
    treeItems(3).Tag = "run-services.msc"
    TreeList1.ListViewAddItem treeItems(3).FullPath, "services.msc", "1. This is a more detailed list of processes that are starting up with Windows."
    TreeList1.ListViewAddItem treeItems(3).FullPath, "services.msc.1", "2. All those items with 'Automatic' listed next to their names are booting with Windows."
    TreeList1.ListViewAddItem treeItems(3).FullPath, "services.msc.2", "3. Click on the items to find out just what they do."
    TreeList1.ListViewAddItem treeItems(3).FullPath, "services.msc.3", "4. If you decide you don't need a certain service, you can simply right-click on it and change it's properties from 'Automatic' to 'Manual'."
    
    Set treeItems(4) = TreeList1.TreeViewAddPath("Tweaks\Internet Explorer Boot Speed")
    TreeList1.ListViewAddItem treeItems(4).FullPath, "internet", "1. Simply right-click on a shortcut to Internet Explorer and add the parameter '-nohome' to the end of the command line."
    
    Set treeItems(5) = TreeList1.TreeViewAddPath("Tweaks\Menu Delays")
    treeItems(5).Tag = "reg-HKEY_CURRENT_USER|Control Panel\Desktop|MenuShowDelay|0"
    TreeList1.ListViewAddItem treeItems(5).FullPath, "menudelay", "1. Click Start -> Run, then type 'regedit' and press enter."
    TreeList1.ListViewAddItem treeItems(5).FullPath, "menudelay.1", "2. The key you need to change is located in HKEY_CURRENT_USER\Control Panel\Desktop"
    TreeList1.ListViewAddItem treeItems(5).FullPath, "menudelay.2", "3. Change MenuShowDelay to 0."
    TreeList1.ListViewAddItem treeItems(5).FullPath, "menudelay.3", "4. Reboot Windows."
    
    Set treeItems(6) = TreeList1.TreeViewAddPath("Tweaks\Add Remove Programs")
    treeItems(6).Tag = "run-c:\Windows\inf\sysoc.inf"
    TreeList1.ListViewAddItem treeItems(6).FullPath, "msn", "1. For example, you want to remove MSN Messenger or Windows Media Player, open c:\Windows\inf\sysoc.inf."
    TreeList1.ListViewAddItem treeItems(6).FullPath, "msn.1", "2. Locate 'msmsgs', the word 'hide' is the string which tells Windows not to display the component in the Add/Remove Programs list"
    TreeList1.ListViewAddItem treeItems(6).FullPath, "msn.2", "3. Remove the hide string so that the entry is 'msmsgs=msgrocm.dll,OcEntry,msmsgs.inf,,7', reboot."
    TreeList1.ListViewAddItem treeItems(6).FullPath, "msn.3", "4. This should now appear in your Add Remove Programs list."
    
    Set treeItems(7) = TreeList1.TreeViewAddPath("Tweaks\Windows File Protection - Disable")
    treeItems(7).Tag = "reg-HKEY_LOCAL_MACHINE|SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon|SFCDisable|0xFFFFFF9D"
    TreeList1.ListViewAddItem treeItems(7).FullPath, "wfp", "1. WARNING: Using this tweak means you will be able to delete vital Windows files, this is not recommended."
    TreeList1.ListViewAddItem treeItems(7).FullPath, "wfp.1", "2. To disable, simply find the key SFCDisable in HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon."
    TreeList1.ListViewAddItem treeItems(7).FullPath, "wfp.2", "3. Change the value to '0xFFFFFF9D'."
    
    Set msc(48) = TreeList1.TreeViewAddPath("Tweaks\Windows File Protection - Enable")
    msc(48).Tag = "reg-HKEY_LOCAL_MACHINE|SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon|SFCDisable|0"
    
    Set treeItems(8) = TreeList1.TreeViewAddPath("Tweaks\Automatically Kill Programs At Shutdown")
    treeItems(8).Tag = "reg-HKEY_CURRENT_USER|Control Panel\Desktop|AutoEndTasks|1"
    TreeList1.ListViewAddItem treeItems(8).FullPath, "akpas", "1. Simply navigate to the HKEY_CURRENT_USER\Control Panel\Desktop directory in the Registry, then alter the key AutoEndTasks to the value 1."
    
    Set treeItems(9) = TreeList1.TreeViewAddPath("Tweaks\Memory Performance")
    TreeList1.ListViewAddItem treeItems(9).FullPath, "mp", "1. Open the Registry and locate HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
    
    Set treeItems(10) = TreeList1.TreeViewAddPath("Tweaks\Memory Performance\Disable Paging Executive")
    treeItems(10).Tag = "reg-HKEY_LOCAL_MACHINE|SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management|DisablePagingExecutive|1"
    TreeList1.ListViewAddItem treeItems(10).FullPath, "dpe", "1. In normal usage, XP pages sections from RAM memory to the hard drive."
    TreeList1.ListViewAddItem treeItems(10).FullPath, "dpe.1", "2. We can stop this happening and keep the data in RAM, resulting in improved performance."
    TreeList1.ListViewAddItem treeItems(10).FullPath, "dpe.2", "3. Note that only users with a large amount of RAM (256MB+) should use this setting. "
    TreeList1.ListViewAddItem treeItems(10).FullPath, "dpe.3", "4. The setting we want to change to disable the 'Paging Executive', as it is called, is called DisablePagingExecutive. "
    TreeList1.ListViewAddItem treeItems(10).FullPath, "dpe.4", "5. Changing the value of this key from 0 to 1 will de-activate memory paging."
    
    Set treeItems(11) = TreeList1.TreeViewAddPath("Tweaks\Memory Performance\System Cache Boost")
    treeItems(11).Tag = "reg-HKEY_LOCAL_MACHINE|SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management|LargeSystemCache|1"
    TreeList1.ListViewAddItems treeItems(11).FullPath, "scb", "Changing the value of the key LargeSystemCache from 0 to 1 will tell Windows XP to allocate all but 4MB of system memory to the file system cache." & _
    "Basically meaning that the XP Kernel can run in memory, greatly improving it's speed.The 4MB of memory left is used for disk caching, but if for any reason more is needed, XP allocates more." & _
    "Generally, this tweak improves performance by a fair bit but can, in some intensive applications, degrade performance." & _
    "As with the above tweak, you should have at least 256MB of RAM before attempting to enable LargeSystemCache.", ".", True, True
    
    Set treeItems(12) = TreeList1.TreeViewAddPath("Tweaks\Memory Performance\Input & Output Performance")
    treeItems(12).Tag = "reg-HKEY_LOCAL_MACHINE|SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management|IoPageLockLimit|16777216"
    TreeList1.ListViewAddItems treeItems(12).FullPath, "iop", "This tweak is only really valuable to anyone running a server - it improves performace while a computer is performing large file transfer operations." & _
    "By default, the value does not appear in the registry, so you will have to create a REG_DWORD value called IOPageLockLimit." & _
    "The data for this value is in bytes, and defaults to 512KB on machines that have the value." & _
    "Most people using this tweak have found maximum performance in the 8 to 16 megabyte range.So you will have to play around with the value to find the best performance." & _
    "Remember that the value is measured in bytes, so if you want, say, 12MB allocated, it's 12 * 1024 * 1024, or 12582912." & _
    "As with all these memory tweaks, you should only use this if you have 256MB or more of RAM.", ".", True, True
    
    Set treeItems(13) = TreeList1.TreeViewAddPath("Tweaks\Speeding Up Share Viewing")
    treeItems(13).Tag = "regdel-HKEY_LOCAL_MACHINE|Software\Microsoft\Windows\CurrentVersion\Explorer\RemoteComputer\NameSpace|{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"
    TreeList1.ListViewAddItems treeItems(13).FullPath, "susv", "Basically, when you connect to another computer with Windows XP, it checks for any Scheduled tasks on that computer." & _
    "Whilst this is a fairly useless task, but one that can add up to 30 seconds of waiting on the other end - not good!." & _
    "Fortunately, it's fairly easy to disable this process.First, navigate to HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\Current Version\Explorer\RemoteComputer\NameSpace in the Registry." & _
    "Below that, there should be a key called {D6277990-4C6A-11CF-8D87-00AA0060F5BF}.Just delete this, and after a restart, Windows will no longer check for scheduled tasks.", ".", True, True
    
    Set treeItems(14) = TreeList1.TreeViewAddPath("Tweaks\Prioritizing Individual Processes")
    TreeList1.ListViewAddItems treeItems(14).FullPath, "pip", "If you press Control+Alt+Delete, then click on the 'Processes' tab, you should get a task bar." & _
    "You can see a list of all the processes running at the time.Now, if you are running a program that you want to dedicate more processing time to - eg, 3D Studio Max." & _
    "You can just right-click on the process, move your cursor down to 'Set Priority >', then select how high you want that program prioritized." & _
    "While checking email, you might want a Normal priority for Max, but if I leave my Computer, I can increass it to 'RealTime' to get the most rendering done.", ".", True, True
    
    Set treeItems(15) = TreeList1.TreeViewAddPath("Tweaks\Prioritizing IRQs")
    TreeList1.ListViewAddItems treeItems(15).FullPath, "pirq", _
    "The main components of your computer have an IRQ number assigned to them." & _
    "With this tweak we can increase the priority given to any IRQ number, thereby improving the performance of that component." & _
    "The most common component this tweak is used for is the System CMOS/real time clock, which improves performance across the board." & _
    "First of all, decide which component you want to give a performance boost to.Next, you have to discover which IRQ that piece of hardware is using." & _
    "To do this, simply go to Control Panel, then open the System panel. You can also press the shortcut of Windows+Break.Click the 'Hardware' tab, then on the 'Device Manager' button." & _
    "Now, right click on the component you want to discover the IRQ for and click 'Properties', then click on the 'Resources' tab." & _
    "You can plainly see which IRQ this device is using (if there is no IRQ number, select another device).Remember the number and close down all of the dialog boxes you have opened, then start up RegEdit." & _
    "Navigate to HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\PriorityControl in the registry." & _
    "Now, we have to create a new DWORD value - called IRQ#Priority (where '#' is the IRQ number), then set the data to 1." & _
    "For example, the IRQ of my System CMOS is 8, so I would create the key IRQ8Priority." & _
    "Now, after restarting, you should notice improved performance in the component you tweaked. I would strongly recommend the CMOS, as it improves performance around the board." & _
    "Also note that you can have multiple IRQ prioritized, but it is fairly inefficient and can cause instability. To remove this tweak, simply delete the value you created.", ".", True, True
    
    Set treeItems(16) = TreeList1.TreeViewAddPath("Tweaks\Scandisk")
    TreeList1.ListViewAddItems treeItems(16).FullPath, "scandisk", "It is now hidden.To use it Go to >my Computer>Right click on your hard drive Icon>Properties>Tools>Error-checking.It works just like the old scandisk.", ".", True, True
    
    Set treeItems(17) = TreeList1.TreeViewAddPath("Tweaks\Winxp Clear Page File On Shutdown")
    TreeList1.ListViewAddItems treeItems(17).FullPath, "clearpage", _
    "Go to Control panel> Administrative tools> local security policy.Go to local policies> security options.Then change the option for 'Shutdown: Clear Virtual Memory Pagefile'.", ".", True, True
    
    Set treeItems(18) = TreeList1.TreeViewAddPath("Tweaks\Remove Shortcut Arrow From Desktop Icons")
    TreeList1.ListViewAddItems treeItems(18).FullPath, "rsafdi", _
    "Start regedit.Navigate to HKEY_CLASSES_ROOT\lnkfile.Delete the IsShortcut registry value.You may need to restart Windows XP.", ".", True, True
    
    
    
    Err.Clear
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    TreeList1.Height = Me.ScaleHeight - 240
    TreeList1.Width = Me.ScaleWidth - 240
    Err.Clear
End Sub

Public Function ExecuteProgram(ByVal strProgramName As String) As Long
    On Error GoTo ErrHandler:
    Dim objWSHShell As IWshRuntimeLibrary.IWshShell
    Set objWSHShell = CreateObject("WScript.Shell")
    ExecuteProgram = objWSHShell.Run(Chr$(34) & strProgramName & Chr$(34), 2, False)
    Err.Clear
    Exit Function
    
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCr & vbCr & Err.Description, vbOKOnly + vbExclamation + vbSystemModal, strProgramName
    Err.Clear
End Function

Private Sub mnuExit_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub

Private Sub mnuRun_Click()
    On Error Resume Next
    Dim mNode As MSComctlLib.Node
    Dim mTag As String
    Dim mAction As String
    Dim mProgram As String
    Dim bRead As Boolean
    Dim regType As String
    Dim regPath As String
    Dim regKey As String
    Dim regValue As String
    Dim regReturn As String
    Dim resp As Long
    Dim pathTitle As String
    Dim regTypeE As REGTool5.REGToolRootTypes
    Set mNode = TreeList1.TreeViewSelectedItem
    If TypeName(mNode) = "Nothing" Then Exit Sub
    
    mTag = mNode.Tag
    If Len(mTag) = 0 Then Exit Sub
    
    mAction = Split(mTag, "-")(0)
    mProgram = Split(mTag, "-")(1)
    pathTitle = Replace$(mProgram, "|", "\")
    Select Case mAction
    Case "run"
        ExecuteProgram mProgram
    Case "reg"
TryTweak:
        regType = Split(mProgram, "|")(0)
        regPath = Split(mProgram, "|")(1)
        regKey = Split(mProgram, "|")(2)
        regValue = Split(mProgram, "|")(3)
        
        Select Case LCase$(regType)
        Case "hkey_current_user"
            regTypeE = HKEY_CURRENT_USER
        Case "hkey_local_machine"
            regTypeE = HKEY_LOCAL_MACHINE
        End Select
        bRead = Registry_Read(regTypeE, regPath, regKey, regReturn)
        If bRead = True Then
            If regReturn = regValue Then
                MsgBox pathTitle & vbCr & vbCr & "This registry value is already tweaked.", vbOKOnly + vbInformation + vbApplicationModal, "Tweak"
            Else
                resp = MsgBox(pathTitle & vbCr & vbCr & "This registry value needs to be updated to " & regValue & " from " & regReturn & ", continue?", vbYesNo + vbQuestion + vbApplicationModal, "Tweak")
                If resp = vbNo Then Exit Sub
                bRead = Registry_Save(regTypeE, regPath, regKey, regValue)
                If bRead = False Then
                    resp = MsgBox(pathTitle & vbCr & vbCr & "This registry value could not be tweaked.", vbRetryCancel + vbExclamation + vbApplicationModal, "Tweak")
                    If resp = vbCancel Then Exit Sub
                    GoTo TryTweak
                End If
            End If
        Else
            MsgBox pathTitle & vbCr & vbCr & "This registry value does not exist.", vbOKOnly + vbExclamation + vbApplicationModal, "Tweak"
        End If
    Case "regdel"
        regType = Split(mProgram, "|")(0)
        regPath = Split(mProgram, "|")(1)
        regKey = Split(mProgram, "|")(2)
        
        Select Case LCase$(regType)
        Case "hkey_current_user"
            regTypeE = HKEY_CURRENT_USER
        Case "hkey_local_machine"
            regTypeE = HKEY_LOCAL_MACHINE
        End Select
        
        bRead = Registry_Read(regTypeE, regPath, regKey, regReturn)
        If bRead = True Then
            bRead = Registry_Delete(regTypeE, regPath & "\" & regKey)
        Else
            MsgBox pathTitle & vbCr & vbCr & "This registry value does not exist.", vbOKOnly + vbExclamation + vbApplicationModal, "Delete"
        End If
    End Select
    Err.Clear
End Sub

Private Sub TreeList1_TreeViewDblClick()
    On Error Resume Next
    mnuRun_Click
    Err.Clear
End Sub
