<div align="center">

## Determine if Connected to Internet \(even through network\)


</div>

### Description

Simple way to tell if a user is connected to the internet
 
### More Info
 
Assume MS Winsock control is added to app with name Winsock1

you can replace www.yahoo.com with any site you want, just make sure its on the net somewhere...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Dryden](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-dryden.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-dryden-determine-if-connected-to-internet-even-through-network__1-9184/archive/master.zip)





### Source Code

```
Public Sub CheckIfConnected()
 Winsock1.Close
 Winsock1.Connect "www.yahoo.com", 80
 While Winsock1.state <> sckConnected
  If Winsock1.state = sckError Then GoTo Offline
  DoEvents
 Wend
 MsgBox "Online"
 Winsock1.Close
 Exit Sub
Offline:
 MsgBox "Offline"
End Sub
```

