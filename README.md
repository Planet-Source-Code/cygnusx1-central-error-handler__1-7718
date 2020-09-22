<div align="center">

## Central Error Handler


</div>

### Description

The following code provides the means of adding a Centralized Error Handler to an application. I keep the Public Function in a Module named ErrorHandling to keep things simple.
 
### More Info
 
The appropriate comments have been added to the code so that you know where to place each item. So long as you follow that the code will work great for anyone.

Note that the Public Function Error Handler can accept any additions concerning the expansion of error codes. All functions/procedures/events that use this error handler do not need to be updated.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[CygnusX1](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cygnusx1.md)
**Level**          |Beginner
**User Rating**    |3.9 (31 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cygnusx1-central-error-handler__1-7718/archive/master.zip)





### Source Code

```
Option Explicit
'The Following Code gets added to each Sub in which you
'Would like to trap errors in.
ErrHandler:
    Dim iErrorAction As Long
    iErrorAction = ErrorHandler(Err)
    Select Case iErrorAction
    Case 1
      Resume
    Case 2
      Resume Next
    Case 3
    'Case 3 is for Resume to a Line, otherwise left blank
    Case 4
      Exit Sub
    Case 5
      End
    End Select
'The code below remains in a Module where it can be expanded in one central location
Public Function ErrorHandler(iErrNum) As Long
Dim iAction As Long
  Select Case iErrNum
    Case -2147467259
    MsgBox "A database data entry violation has occurred. " & "Error Number = " & iErrNum
    iAction = 5
    Case 5
    'Invalid Procedure Call
    MsgBox Error(iErrNum) & " Contact Help Desk."
    iAction = 2
    Case 7
    'Out of memory
    MsgBox "Out of Memory. Close all unnecessary applications."
    iAction = 1
    Case 11
    'Divide by 0
    MsgBox "Zero is not a valid value."
    iAction = 1
    Case 48, 49, 51
    'Error in loading DLL
    MsgBox iErrNum & " Contact Help Desk"
    iAction = 5
    Case 57
    'Device I/O error
    MsgBox "Insert a disk into Drive A."
    iAction = 1
    Case 68
    'Device Unavailable
    MsgBox "Device is unavailable(the device may not exist or it is currently unavailable)."
    iAction = 4
    Case 482, 483
    'General Printer Error
    MsgBox "A general printer error has occurred. Your printer may be offline."
    iAction = 4
    Case Else
    MsgBox "Unrecoverable Error. Exiting Application. " & "Error Number = " & iErrNum
    iAction = 5
    End Select
    ErrorHandler = iAction
  End Function
```

