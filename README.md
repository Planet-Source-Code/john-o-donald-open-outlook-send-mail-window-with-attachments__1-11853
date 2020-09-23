<div align="center">

## Open  Outlook send mail window with attachments


</div>

### Description

Open Outlook send mail window with attachments from a vb application. Also, change m.display to m.send if you want to just send the email and not preview it! Also works in vbscript! Copy code to a text file and save with a *.vbs extension and double click to activate. Super Cool.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John O'Donald](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-o-donald.md)
**Level**          |Intermediate
**User Rating**    |4.9 (108 globes from 22 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-o-donald-open-outlook-send-mail-window-with-attachments__1-11853/archive/master.zip)





### Source Code

```
Dim o
 Dim m
 Set o = CreateObject("Outlook.Application")
 Set m = o.CreateItem(0)
 m.To = "xxxx@yyyy.com"
 m.Subject = "This is the Subject"
 m.Body = "Hey, this is cool!"
 m.Attachments.Add "C:\Temp\FileToAttach.txt"
 'Repeat this line if there are more Attachments
 m.Display
 'm.Send 'If you want to just send it
```

