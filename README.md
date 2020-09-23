<div align="center">

## KeyBoardHook


</div>

### Description

This code will work for win 2000 and XP. It hooks the systems keyboard and runs a callback function to determine if the "delete" key has been hit and disables the delete (system wide not just application wide). This program could also be easily resigned as a keylogger if you feel the need to do so, also it may be used to disable any number of keys you wish. In the LowLevelKetBoardProc just use debug.print xpInfo.vkCode and you will see the value of the key and you could set that value to de disabled (as VK_DELETE is currently.) When I get a chance I will also try to make it for the other MS OS's

----

Warning

----

Since this program uses callbacks DO NOT HIT THE END BUTTON IN VB'S IDE .....YOU WILL CRASH.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-08-24 01:32:12
**By**             |[ebred](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ebred.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[KeyBoardHo1634428242003\.zip](https://github.com/Planet-Source-Code/ebred-keyboardhook__1-47941/archive/master.zip)








