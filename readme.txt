VB.NET USB Sample
==================

This file contains:
1. An overview of the contents of the vb.net\usb_sample\ directory.
2. List of files.
3. Instructions for using the sample application.
4. Instructions for building the sample code.


1. Overview
   =========
   The vb.net\usb_sample\ directory contains a graphical (GUI) Visual Basic
   (VB) .NET USB sample diagnostics application (vb_usb_sample.exe) and the
   source code for this application.

   The code was written by Jungo Connectivity using WinDriver's USB API.

   The sample application supports handling of multiple USB devices and
   simultaneous execution of several tasks (on different devices or on the same
   device) - such as reading from one pipe while writing to the other, etc.

   The GUI reflects the insertion and removal of USB devices registered to work
   with WinDriver (via an *.inf file) and displays the configuration and
   resources (pipes) information for each connected device.
   The application also enables changing the active alternate setting for each
   connected device.

   You can easily switch between the connected devices from the GUI and test
   the communication with the device. You can issue requests on the control
   pipe, write and read (listen) from/to the pipes, or reset a pipe.

   The sample code uses the WinDriver .NET API DLL - wdapi_dotnet<version>.dll
   - which provides the required .NET interface for the WinDriver USB API and
   enables the use of the C WinDriver API DLL - wdapi<version>.dll (found under
   the WinDriver\redist\ directory).
   A copy of wdapi_dotnet<version>.dll is provided with the sample (see below),
   as well as under the WinDriver\lib\<CPU>\<.NET Version>\ directory
   (e.g. WinDriver\lib\x86\v1.1.4322 - for Windows x86 32-bit, .NET v1.1.4322).
   The source code of wdapi_dotnet<version>.dll is found in the
   WinDriver\src\wdapi.net\ directory.

   The sample code also uses the usb_lib_dotnet.dll WinDriver C# .NET USB
   library, which provides an interface between the WinDriver .NET API DLL
   (wdapi_dotnet<version>.dll - see above) and .NET applications, such as the
   vb_usb_sample.exe sample. For a description of the usb_lib_dotnet.dll
   library, refer to the WinDriver\csharp.net\usb_sample\readme.txt file.

   Note: This code sample is provided AS-IS and as a guiding sample only.

2. Files
   ======
   This section describes the sub-directories and files provided under the
   vb.net\usb_sample\ directory.

   - readme.txt:
         Describes the contents of the vb.net\usb_sample\ directory.

   - diag\ sub-directory:
     --------------------

     For all Windows platforms:
     --------------------------
     - UsbSample.vb:
           VB .NET source file.
           The sample's main form. Implements the sample's graphical
           user interface (GUI) and related code.

     - FormChangeSettings.vb:
           VB .NET source file.
           Input form for changing a device's active alternate setting.

     - FormTransfers.vb:
           VB .NET source file.
           Input form for performing pipe transfers.

     - DeviceTabPage.vb:
           VB .NET source file.
           Defines a new USB device tab page class (DeviceTabPage), which
           inherits from the graphical .NET System.Windows.Forms.TabPage class.


     - diag\Release\ sub-directory:

       - vb_usb_sample.exe:
             A pre-compiled sample executable.

       - usb_lib_dotnet.dll:
             A copy of the C# .NET USB library.

       - wdapi_dotnet<version>.dll:
            A copy of the WinDriver .NET API DLL.

     - Additional files required for building the sample code (resources,
       assembly, etc.)

     For x86 32-bit platforms:
     -------------------------
     - x86 subdirectory:
        - msdev_20xx subdirectory
           - vb_usb_sample.sln:
             Visual Studio 20xx VB .NET project file.

     For x64 64-bit platforms:
     -------------------------
     - x64 subdirectory:
        - msdev_20xx subdirectory
           - vb_usb_sample.sln:
             Visual Studio 20xx VB .NET project file.


3. Using the sample application
   =============================
   To use the sample vb_usb_sample.exe application, follow these steps:

   1) Install the .NET Framework.
   2) Build the vb_usb_sample.exe application by following the instructions in
      section #4 of this file, or use the pre-built version of the
      application from the Release\<.NET Version> sub-directory
      (see section #2 above.)
   3) Install an INF file for any USB device that you wish to view and control
      from the application, which registers the device(s) to work with
      WinDriver. You can use the DriverWizard (Start | WinDriver | DriverWizard)
      to create and install the required INF file(s).
   4) Run the vb_usb_sample.exe executable and use it to test the communication
      with your USB device(s).
      NOTE: In order to run the executable the wdapi_dotnet<version>.dll and
      usb_lib_dotnet.dll files must be found in the same directory as
      vb_usb_sample.exe.


4. Building the sample code
   =========================
   1) Install the .NET Framework and Visual Studio IDE (2012 and higher).
   2) Open the relevant library solution file (vb_usb_sample_msdev_xxx.sln).
   3) Build the solution (CTRL+SHIFT+B OR from the Build | Build Solution menu).

