**CopyMemory** is actually a common alias for the **RtlMoveMemory**. Many tutorials on the Internet demonstrate how to use it in 32-bit VBA, but there is not much material on 64-bit VBA. Thanks to Bytecomb I managed to get the basics and will share them here, one tip a time.

The reason to use **CopyMemory**, was that 20 years ago VBA was kind of slow when performing certain functionalities, e.g. appending Strings, so it would be very useful to squeeze as much power out of it as possible, especially when you need to perform expensive operations for a thousand times.
