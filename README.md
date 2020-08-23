# The DataTools Desktop Utility Library
## Version 4.2
### A general memory, hardware, math, text and utility library.

An extensive low-level and hardware utility library featuring a number of useful classes and structures built in Visual Basic .NET and MSIL.

### [Visit the Wiki](https://github.com/nmoschkin/dtlib/wiki)

## The Pieces

### __DTCore__ 

DTCore is heavily coded in IL on the backend, with VB.NET on the front using __ILSupport__.

This module features VB.NET language-specific indexed accessor properties for [__MemPtr__](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_MemPtr), [__SafePtr__](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_SafePtr)__, and [__Blob__](https://github.com/nmoschkin/dtlib/wiki/T_DataTools_Memory_Blob).  These tools can directly interface with unmanaged memory and pinned GCHandles to quickly manipulate data at the byte-level for just about any kind of blittable object.  

These classes **CAN** be used in C#, but since the C# language does not support properties with indexers, the functions will appear in Intellisense by their internal 'magic' names.  For example, a property called __ByteAt(index As Integer)__ will appear in C# as two separate functions: __byte get_ByteAt(int index)__ and __void set_ByteAt(int index, byte value)__.

These classes were written with VB.NET developers in mind, since the VB.NET language specification does not define any method by which to utilize pointers or unsafe calls.
These classes help VB.NET developers do with greater ease things that C# developers take for granted, often more conveniently so.  Some of the features of these classes may also be of interest to C# and F# developers because they make working with unmanaged memory (and memory buffers, in general) a bit more concise and monolithic.

### __DTInterop__ 

DTInterop is an extensive OS and hardware library for Windows Desktop, written in VB.NET using __DTCore__ as the backbone for its many interop functions.

This module is of more universal appeal to all .NET developers, regardless of the language used.
 
### __DTMath__

DTMath is an extended miscellaneous mathemetics library, including working with polar coordinates and an extensive color space math calculation library. 
This module is a bit outdated and some parts that were unfinished will probably never be finished or will be done a different way, so components may be added or removed
from this library, in the near future.
  
### __DTExtended__

DTExtended is a collection of advanced modules including __UnitConverter__ and __BinarySearcher__.  This project is actually do for an intense overhaul as many of its features have become obsoleted or outdated, as well.

  * Note, running the examples **will** require you to run Visual Studio as Administrator.



