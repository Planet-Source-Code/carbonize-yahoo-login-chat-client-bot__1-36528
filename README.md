<div align="center">

## Yahoo Login / Chat Client / Bot


</div>

### Description

This is Just an updated version of Alfie Pughs bot/chat client source. It will retrieve and parse the cookies, log you onto the chat server, then join a room. I changed the cookie parsing as Yahoo altered what was required to log on to the servers. I fixed the think packet as well as yahoo had stopped it. I added the PM and Leave Room packets and added more incoming packet remarks. I also added the facility to log in space ID's.

Just to inform you that after I uploaded this Yahoo made some changes to YCHT. Emotes and thoughts work again after yahoo blocking them.

Yahoo stopped support for ncclogin so you need to change

"GET /config/ncclogin?.src=bl&login=" & USERID & "&passwd=" & PASS & "&n=1 HTTP/1.0"

to

"GET /config/login?.src=&login=" & USERID & "&passwd=" & PASS & "&n=1 HTTP/1.0"
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-05 02:23:08
**By**             |[Carbonize](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/carbonize.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Yahoo\_Logi102030742002\.zip](https://github.com/Planet-Source-Code/carbonize-yahoo-login-chat-client-bot__1-36528/archive/master.zip)








