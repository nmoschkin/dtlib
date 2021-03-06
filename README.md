# DataTools

**Memory and Hardware Interop Library**

**Version 3.2**

An extensive memory and hardware abstraction and wrapper library featuring a number of unique and useful classes and structures built in **Visual Basic .NET** and **[CIL](https://en.wikipedia.org/wiki/Common_Intermediate_Language)** for native Windows operating systems.

_**Note:** Run these projects with elevated permission, and native to your platform. If you are on a 64-bit machine, compile with the 64-bit configuration._

Visit [The Wiki](https://github.com/nmoschkin/dtlib/wiki).

_(See below for **important namespaces and classes**)_


_(Click Here For The [Unabridged Namespace Map](https://github.com/nmoschkin/dtlib/wiki/NamespaceMap))_

## Individual Libraries

### __DTCore__ 

DTCore contains the core unmanaged memory manipulation structures and classes that **DTInterop** is built on. 
It is heavily coded in **[CIL](https://en.wikipedia.org/wiki/Common_Intermediate_Language)** on the backend, with **VB.NET** on the front using **[ILSupport](https://github.com/ins0mniaque/ILSupport)**.

This library features *****VB.NET** language-specific*** indexed accessor properties for **[MemPtr](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_MemPtr)**, **[SafePtr](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_SafePtr)**, and **[Blob](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_Blob)**.  These tools can directly interface with unmanaged memory and pinned GCHandles to quickly manipulate data at the byte-level for just about any kind of blittable object.  

These classes **CAN** be used in **C#**, but since the **C#** language does not support properties with indexers, the functions will appear in **Intellisense** by their internal 'magic' names.  For example, a property called __ByteAt(index As Integer)__ will appear in **C#** as two separate functions: __byte get_ByteAt(int index)__ and __void set_ByteAt(int index, byte value)__.

These classes were written with **VB.NET** developers in mind, since the **VB.NET** language specification does not define any method by which to utilize pointers or unsafe calls.

These classes help **VB.NET** developers do with greater ease things that **C#** developers take for granted, often more conveniently so.  Some of the features of these classes may also be of interest to **C#** and **F#** developers because they make working with unmanaged memory (and memory buffers, in general) a bit more concise and monolithic.

### __DTInterop__ 

DTInterop is an OS and hardware abstraction and wrapper library for systems running on native Windows platforms.

This library was written entirely in in **VB.NET** using **[DTCore](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Memory)** as the backbone for easily manipulating unmanaged memory in **VB.NET**.

***This library is not specific to **VB.NET**, and works well in all .NET languages, including **C#** and **F#**.***

**[DevEnumPublic](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_DevEnumPublic)** is the jump-off point for the computer hardware information and manipulation portion of this library.

Other independent subsystems include **[Desktop and Shell](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Desktop)**, **[Virtual Media Control](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Disk_VirtualDisk)**, **[Networking](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Network)**, **[System Information](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_SystemInfo)**, and others.  See below, or see the [Wiki](https://github.com/nmoschkin/dtlib/wiki) for more details.

* Note, running the examples **will** require you to run Visual Studio as Administrator.

## Important Namespaces and Classes
**[Visit the Wiki](https://github.com/nmoschkin/dtlib/wiki)**

**Click Here For The [Unabridged Namespace Map](NamespaceMap)**

 - **[DataTools.Memory Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Memory)**
   - **[Blob Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_Blob)**
   - **[MemPtr Structure](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_MemPtr)**
   - **[SafePtr Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_SafePtr)**
 - **[DataTools.BitStream Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_BitStream)**
   - **[Crc32 Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_BitStream_Crc32)**
 - **[DataTools.Interop Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop)**
   - **[DevEnumPublic Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_DevEnumPublic)**
 - **[DataTools.Interop.Desktop Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Desktop)**
   - **[AllSystemFileTypes Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Desktop_AllSystemFileTypes)**
   - **[FontInfo Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Desktop_FontInfo)**
   - **[IconImage Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Desktop_IconImage)**
   - **[Resources Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Desktop_Resources)**
 - **[DataTools.Interop.Disk Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Disk)**
   - **[DiskDeviceInfo Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Disk_DiskDeviceInfo)**
   - **[FSMonitor Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Disk_FSMonitor)**
   - **[VirtualDisk Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Disk_VirtualDisk)**
 - **[DataTools.Interop.Display Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Display)**
   - **[MonitorInfo Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Display_MonitorInfo)**
 - **[DataTools.Interop.Network Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Network)**
   - **[AdaptersCollection Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Network_AdaptersCollection)**
 - **[DataTools.Interop.Printers Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Printers)**
   - **[PrinterDeviceInfo Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Printers_PrinterDeviceInfo)**
 - **[DataTools.Interop.Usb Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Interop_Usb)**
   - **[HidDeviceInfo Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Interop_Usb_HidDeviceInfo)**
 - **[DataTools.Strings Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_Strings)**
 - **[DataTools.SystemInfo Namespace](https://github.com/nmoschkin/dtlib/wiki/N_DataTools_SystemInformation)**
   - **[SysInfo Class](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_SystemInformation_SysInfo)**



## Other Libraries

### __DTMath__

DTMath is an extended miscellaneous mathemetics library, including working with polar coordinates and an extensive color space math calculation library. 

  
### __DTExtended__

DTExtended is a collection of advanced modules including __UnitConverter__ and __BinarySearcher__.
  
