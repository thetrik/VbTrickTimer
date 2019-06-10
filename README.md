# VbTrickTimer

Hello everyone!

This class allows to use the timers in VB6/VBA. It contains the **Interval** propertiy which sets the interval between the **Tick** events. It also contains the **Tag** property which allows to hold any data associated with the timer instance.
The class uses the assembly thunks which check allow use the single class without any other dependencies. It also has the simple checking to reduce the crashes. It check the **Ebmode** function and if the code is stopped it automatically disables the timers and if the code is in the stepping mode it just bypasses the events generation until the code is running. This checking simplifies debugging but doesn't exclude the crashes (because if the timer wasn't disabled since the last debugging session it'll continue execution with the old invalid data.) 
This code is compatible with the 64 bit office as well.
If you want to add a method to the class you should update the **TIMERPROC_INDEX** constant according to the offset.

Thanks for your attention!

The trick,
2019.
