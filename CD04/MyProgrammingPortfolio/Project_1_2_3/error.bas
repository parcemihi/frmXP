Attribute VB_Name = "errNumber"
Option Explicit
Select Case Err.Number
Case 3 'Return without GoSub
Case 5 'Invalid procedure call
Case 6 'Overflow
Case 7 'Out of memory
Case 9 'Subscript out of range
Case 10 'This array is fixed or temporarily locked
Case 11 'Division by zero
Case 13 'Type mismatch
Case 14 'Out of string space
Case 16 'Expression too complex
Case 17 'Can't perform requested operation
Case 18 'User interrupt occurred
Case 20 'Resume without error
Case 28 'Out of stack space
Case 35 'Sub, Function, or Property not defined
Case 47 'Too many DLL application clients
Case 48 'Error in loading DLL
Case 49 'Bad DLL calling convention
Case 51 'Internal error
Case 52 'Bad file name or number
Case 53 'File not found
Case 54 'Bad file mode
Case 55 'File already open
Case 57 'Device I/O error
Case 58 'File already exists
Case 59 'Bad record length
Case 61 'Disk full
Case 62 'Input past end of file
Case 63 'Bad record number
Case 67 'Too many files
Case 68 'Device unavailable
Case 70 'Permission denied
Case 71 'Disk not ready
Case 74 'Can't rename with different drive
Case 75 'Path/File access error
Case 76 'Path not found
Case 91 'Object variable or With block variable not set
Case 92 'For loop not initialized
Case 93 'Invalid pattern string
Case 94 'Invalid use of Null
Case 97 'Can't call Friend procedure on an object that is not an instance of the defining class
Case 98 'A property or method call cannot include a reference to a private object, either as an argument or as a return value
Case 298 'System DLL could not be loaded
Case 320 'Can't use character device names in specified file names
Case 321 'Invalid file format
Case 322 'Can’t create necessary temporary file
Case 325 'Invalid format in resource file
Case 327 'Data value named not found
Case 328 'Illegal parameter; can't write arrays
Case 335 'Could not access system registry
Case 336 'Component not correctly registered
Case 337 'Component not found
Case 338 'Component did not run correctly
Case 360 'Object already loaded
Case 361 'Can't load or unload this object
Case 363 'Control specified not found
Case 364 'Object was unloaded
Case 365 'Unable to unload within this context
Case 368 'The specified file is out of date. This program requires a later version
Case 371 'The specified object can't be used as an owner form for Show
Case 380 'Invalid property value
Case 381 'Invalid property-array index
Case 382 'Property Set can't be executed at run time
Case 383 'Property Set can't be used with a read-only property
Case 385 'Need property-array index
Case 387 'Property Set not permitted
Case 393 'Property Get can't be executed at run time
Case 394 'Property Get can't be executed on write-only property
Case 400 'Form already displayed; can't show modally
Case 402 'Code must close topmost modal form first
Case 419 'Permission to use object denied
Case 422 'Property not found
Case 423 'Property or method not found
Case 424 'Object required
Case 425 'Invalid object use
Case 429 'Component can't create object or return reference to this object
Case 430 'Class doesn't support Automation
Case 432 'File name or class name not found during Automation operation
Case 438 'Object doesn't support this property or method
Case 440 'Automation error
Case 442 'Connection to type library or object library for remote process has been lost
Case 443 'Automation object doesn't have a default value
Case 445 'Object doesn't support this action
Case 446 'Object doesn't support named arguments
Case 447 'Object doesn't support current locale setting
Case 448 'Named argument not found
Case 449 'Argument not optional or invalid property assignment
Case 450 'Wrong number of arguments or invalid property assignment
Case 451 'Object not a collection
Case 452 'Invalid ordinal
Case 453 'Specified not found
Case 454 'Code resource not found
Case 455 'Code resource lock error
Case 457 'This key is already associated with an element of this collection
Case 458 'Variable uses a type not supported in Visual Basic
Case 459 'This component doesn't support the set of events
Case 460 'Invalid Clipboard format
Case 461 'Method or data member not found
Case 462 'The remote server machine does not exist or is unavailable
Case 463 'Class not registered on local machine
Case 480 'Can't create AutoRedraw image
Case 481 'Invalid picture
Case 482 'Printer error
Case 483 'Printer driver does not support specified property
Case 484 'Problem getting printer information from the system. Make sure the printer is set up correctly
Case 485 'Invalid picture type
Case 486 'Can't print form image to this type of printer
Case 520 'Can't empty Clipboard
Case 521 'Can't open Clipboard
Case 735 'Can't save file to TEMP directory
Case 744 'Search text not found
Case 746 'Replacements too long
Case 31001 'Out of memory
Case 31004 'No object
Case 31018 'Class is not set
Case 31027 'Unable to activate object
Case 31032 'Unable to create embedded object
Case 31036 'Error saving to file
Case 31037 'Error loading from file
End Select
