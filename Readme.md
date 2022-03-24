# **The notify connect component**

##### Author: Fabrice Sanga
<br/>
<br/>

The component augments functionalities to the network system icon. It helps notify of the status of the network connexion and of the network interface card (NIC). It is not limited to the binary value, either `connected` or `no internet`. 
The notification is a visual icon that changes form according to the following states:
- `Disabled`: The administrative status of the NIC is disabled
- `Disconnected`: The media connecting the NIC to the router is disconnected
- `NoInternet`: The internet access is interrupted
- `Connected`: The workstation has access to the internet.

An example of use:
```vbscript
With CreateObject("CustomUI.Shuffler")
    .NetAdapterID = "Ethernet"                              '(1)
    .ShortcutPath = "..\Start Menu\Programs\Ethernet.lnk"   '(2)
    .IconsDir = "\Path\to\ConnectionStatusIcons"            '(3)
    While True
        Wscript.Sleep 1000
        CurrentState = .NetStatus                           '(4)
        .ChangeShortcutIcon                                 '(5)
        While CurrentState = .NetStatus : Wend
    Wend
End With
```
**(1)** `NetAdapterID` is the name of the network connection ID of the adapter

**(2)** `ShortcutPath` is the path to the link to the notifier that is shortcut link

**(3)** `IconsDir` is the path to the directory containing the icons, each named by the state it represents.

**(4)** `NetStatus` returns the state of the connexion

**(5)** `ChangeShortcutIcon` change the shortcut icon
<br/>

The filesystem:
```
    ConnectionStatusIcons
        |--Connected.ico
        |--Disabled.ico
        |--Disabled.ico
        |--NoInternet.ico
```
