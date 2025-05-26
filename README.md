## VBA_Shell - concept of creating an analogue of the system console based on VBA.
In this project I will try to present some, perhaps minimal example of creating a console based on the VBA environment. It seems that VBA has everything for this, the console - immediate, the ability to create functions and classes, use winapi.. Why not VBA?
Let's imagine a console in which we enter
```vba
GetFile("File name")
```
then **"."** and we get a hint with a dozen actions that can be done with this file. Or
```vba
Processes.filter("*chrome*").kill forsed:=true
```
