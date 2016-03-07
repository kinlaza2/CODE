=============================================================
Osprey Video
ViewCast Corporation

Osprey MultiMedia Capture Driver
For the following Card(s):
Osprey-100   Osprey-101   Osprey-200
Osprey-210   Osprey-220   Osprey-230
Osprey-300   Osprey-440   Osprey-530
Osprey-540   Osprey-560

For Windows XP and Windows 2003, 32-bit version
Version 4.0.0
9/05/2006

Sections in this file:
- Introduction
- Troubleshooting
- Osprey-300 IEEE 1394 Ports
- Testing the Driver
- Manuals and Help
- Latest Drivers
- Known Issues in version 4.0.0
- Installation
- Important Notice
- Contact Information


Introduction:
=============================================================
=============================================================

This is the first "Series IV" AVStream video and audio capture
driver for all models of ViewCast's Osprey video capture cards 
listed above. This includes selected Legacy cards no longer in 
production. 

Like the earlier Series III drivers, these drivers
are built on Microsoft's AVStream / DirectShow driver model. 

This version of the driver is for Windows XP and 2003 32-bit 
O/S versions and is plug and play compatible. 

Please note that ViewCast no longer supports new development or 
bug fixes in older Video for Windows (VfW) drivers used in 
Microsoft O/S versions prior to Windows XP. Those legacy drivers 
may still be downloaded from the ViewCast web site. 

This driver has not been signed by WHQL but has successfully 
passed the WHQL developer test programs and can be installed
with confidence. WHQL version will be posted on the ViewCast 
web site when available.

This driver works only with genuine Osprey Video(R)
video capture cards designed and marketed by 
ViewCast Corporation of Plano, TX. USA. 


NEW FEATURES
------------

Series IV drivers are based on Osprey Series III drivers with 
a number of enhancements that can take advantage of the
increased performance of the latest generation of PCs. 

Significant new features include:

- Improved video post-capture processing, including an enhanced
de-interlacer, automatic telecine detection, and other enhancements
that generally improve video quality on current and many older
Osprey board-level products.

-Improved S-video picture sharpness on some current and older Osprey cards
equipped with S-video inputs.

- Substantially different implementation of Simulstream(TM) that allows
much greater flexibility for configuring each stream individually.
Streams replicated via SimulStream(TM) can now be individually configured
for sizing, scaling, cropping,and watermarking.

-SimulStream (TM) Test Drive. Users can now set an option on the 
Filter tab in the Driver Properties dialog to toggle this capability
on and off. While in evaluation mode a watermark will appear on the active
video region which denotes usage of an evaluation version of
Osprey SimulStream.

Note that many of the performance enhancements in version 4.x must
be enabled manually in the driver's property pages. This ensures 
that the driver initially operates with equivilent features and 
performance to Series 3 drivers to ensure out-of-the-box compatability 
with applications already been installed and set up on Series 
III or earlier Osprey drivers. 

Please refer to the Osprey AVStream 4.x User Guide included in this 
distribution for setup and operation information on these and other 
new features of the advanced Series IV drivers. SimulStream users are
strongly advised to become familiar with the extensive changes and new
options detailed in the User Guide. 

 
Troubleshooting
=============================================================
=============================================================
If the installation program for this driver appears to hang,
press Alt-Tab to ensure that the installation screen is not
hidden.

The 'Digital Signature Not Found' window may appear during the
installation. Select the Continue Anyway button to dismiss this
dialog.

Osprey-300 IEEE 1394 Ports
=============================================================
=============================================================
The Osprey-300 has the same features of an Osprey-230 with two
IEEE 1394 ports added to it.  The 1394 ports reside on a separate
Plug and Play device that is controlled by standard windows
drivers; they are separate from the Osprey analog Bt878A device
and the older Osprey Video for Windows analog drivers (Series II
and earlier). The 1394 ports support Digital Video devices such 
as camcorders, and also non-DV devices such as hard drives.

The 1394 ports are WDM rather than Video for Windows devices.
DirectShow applications such as Windows Media Encoder 9 will
work directly with them.  Video for Windows applications will
probably not work with 1394 devices connected to the Osprey 
300, including VidCap32 which is included with this driver 
distribution. Many DirectShow applications will capture DV 
video but not DV audio - for example, the AMCap test application
included in this release has this limitation.

The 1394 device works only on Windows 2000 and XP. Windows NT4.0
does not have built-in 1394 support.


Implications Related to 4.0.0 Patches
==========================================================
==========================================================
This release 4.0.0 completely replaces all previous Osprey 
drivers. All installations of previous Osprey drivers must 
be removed prior to installing this version.


Testing the Driver:
=============================================================
=============================================================
Click on "SwiftCap" application icon that is installed on the
desktop to test the driver.


Manuals and Help:
=============================================================
=============================================================

Refer to the Osprey User Guide for detailed information about 
the Osprey drivers.

Latest Drivers:
=============================================================
=============================================================
Before installing, check the ViewCast Corporation website
www.viewcast.com for the latest drivers:

If there isn't a newer driver at the time of
your installation, periodically check the website for newer
versions.


Known issues in version 4.0.0:
=============================================================
=============================================================
- Release 4.0.0 is not yet WHQL-certified.

- PCI bus numbers of 16 or higher: 
The drivers support multiple Osprey single-channel cards 
(up to 15) on a PCI bus. Due to PCI bus limitations, attempting 
to enumerate 16 or more devices will cause failures. This may 
affect users who require a large number of Osprey-230, 
Osprey 440, Osprey 300, or Osprey 560 cards in a single system.  
These cards have on-board PCI bridge chips that creates one 
additional PCI bus for each card.  As a result, adding multiples
of these cards to a system may quickly create more than 16 PCI 
devices on a system. The drivers will fail to function properly 
with cards on PCI bus number 16 or higher.

- Sndvol32.exe error message under Windows 2003 SP1
The Sndvol32.exe application will display an error message
under the Windows 2003 SP1. Select OK through the two message
dialogs and the sound panel applet will function normally.

- GraphEdit:
In the Filter Properties section of GraphEdit, the size and
crop settings may not work properly.

- Horizontal Delay:
When adjusting the Horizontal Delay on the RefSize tab, the
negative delay may not allow adjustment beyond negative 6.

Osprey - Windows XP Installation:
=============================================================
=============================================================
Depending on your system setup, you will have multiple 
options for the installation of the Osprey MultiMedia
drivers. By default, the installer will install the drivers
for all supported Osprey cards. Select the Custom radio
box during the install process to select installation of 
specific drivers.

Following are the different scenarios and their methods of
installation:

* INSTALL SCENARIOS

 There are three main situations that might apply to you:

 Scenario 1: Osprey Card(s) Physically Installed, but Osprey 
 Software not Installed * RECOMMENDED METHOD* 

 Scenario 2: Osprey Card(s) Physically Installed, and 
 Previous Osprey Software Installed.

 Scenario 3: Osprey Card(s) not Physically Installed in the 
 PC.

 You must uninstall all previous installations of the Osprey
 drivers prior to installing this version. You must also
 reboot your computer after uninstalling.

 In all cases, the most efficient and complete installation 
 method is to run the setup.exe program on the product CD or 
 in the web package that you downloaded after you have 
 installed the Osprey card(s). The setup program 
 automates the Plug and Play steps required to install the 
 drivers and ensures that they are performed correctly. It 
 also installs the bundled applets and Users' Guide. If you 
 have multiple Osprey capture cards in the system it sets 
 all of them up at once.

 It is possible to install the Osprey drivers using the
 Hardware Installation Wizard. Select Have disk and
 navigate to the card specific Drivers directory that is located
 on the installation disk to select the inf file. This is
 an advanced feature and will not be supported by
 additional documentation or Customer Support.
 Use this method at your own risk.

 The installer provides a Custom installation option that
 allows selected installation.

 Although the installer allows drivers to reside across mapped
 network drives, this method is not recommended because it
 will not allow a proper unistall.

-- SCENARIO 1: OSPREY CARD(S) PHYSICALLY INSTALLED, BUT
   OSPREY SOFTWARE NOT INSTALLED

 Run the Installation Program:
 When the OS is first started for the first time after
 the Osprey card is installed, the New Hardware Found wizard
 will appear one or more times. CANCEL OUT OF THESE WIZARDS.

 After the OS has finished starting, do the following:

 1. Double-click the setup.exe file. This will start the 
    installation program.

 2. If you choose to do a custom install:
    Select Destination Folders and Program Folders when
    prompted.

 3. When the installation is finished, there is no need to 
    restart the PC. Your Osprey card(s) can be used directly.

 If there are multiple Osprey cards in the system, this 
 installation method will set up all of them at once
 automatically.

-- SCENARIO 2: OSPREY CARD(S) PHYSICALLY INSTALLED, AND
   PREVIOUS OSPREY SOFTWARE INSTALLED

 This scenario is for the case when the Osprey card is
 physically installed in the PC and there is a previous
 installation of the Osprey drivers.

 It is necessary to uninstall the old driver
 before installing the new driver. You must also
 reboot your computer after uninstalling.

 After restarting, the New Hardware Found wizard
 will appear one or more times. CANCEL OUT OF THESE WIZARDS.

 Run the Installation Program

 1. Double click the setup.exe file. This will start the
    installation program.

 2. If you choose to do a custom install:
    Select Destination Folders and Program Folders when
    prompted.

 3. If you have one or more Osprey-200s in the system, you
 will need to restart the system before you can use the
 updated audio. If you have video-only cards, you do not
 need to restart the system - the Osprey card(s) can be
 used immediately.

-- SCENARIO 3: OSPREY CARD(S) NOT PHYSICALLY INSTALLED IN THE 
   PC

 This scenario  is called the "Preinstall Scenario". After
 the install is  run, as soon as an Osprey card is installed
 in the PC, it  is detected and its drivers are started
 automatically.

 1. Double-click the setup.exe file to start the 
    installation.

 2. If you choose to do a custom install:
    Select Destination Folders and Program Folders when
    prompted.

 3. You will then be prompted to preinstall the drivers.
    Select Yes to continue.
 
 4. The Osprey software is now fully installed. It will be ready
    for use after you install the Osprey card in your computer.
 
 5. When you are ready to install the card, shut down and
    install the Osprey card inside your computer. Then power
    up the computer.  The OS will detect the newly present Osprey 
    card, and begin to activate the pre-installed driver. The
    Osprey card will then be ready for use.


Important Notice
=============================================================
=============================================================
If you are using an Osprey-500 or Osprey-2000 card
in the same machine as your Osprey MultiMedia card, we 
strongly recommend that you upgrade the Osprey-500
or Osprey-2000 driver to the version 2.1.0 or later 
available on the ViewCast web site. ViewCast has not 
developed Series III or IV AVStream drivers for these models. 
Note that Windows XP still supports the older Video for Windows 
(VfW) driver architecture, but newer Microsoft O/S versions 
are expected to drop support for VfW-style drivers at any time. 


Support Contact information:
=============================================================
=============================================================
voice:  972-488-7156
email:  support@viewcast.com
web:    www.viewcast.com

