**CopyMemory** is actually a common alias for the **RtlMoveMemory**. Many tutorials on the Internet demonstrate how to use it in 32-bit VBA, but there is not much material on 64-bit VBA. Thanks to Bytecomb I managed to get the basics and will share them here, one tip a time.

The reason to use **CopyMemory**, was that 20 years ago VBA was kind of slow when performing certain functionalities, e.g. appending Strings, so it would be very useful to squeeze as much power out of it as possible, especially when you need to perform expensive operations for a thousand times.

**Declaration**

This is the easiest part. Most of the tutorials are for 32-bit, so remember to add PtrSafe.

Instead of:
```VBA
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
```
Use:
```VBA
Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
```
BTW this line must be at the top, before any Sub or Functions.

**VarPtr**

VarPtr returns the address of the variable (could be anything in VBA)
