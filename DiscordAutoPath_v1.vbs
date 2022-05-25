' DiscordAutoPath v1.0 -  Automatic retrieval of Discord path and launch silently in the background and minimize it to system tray
' Home URL: https://github.com/avimanyu786/DiscordAutoPath/ 

' Copyright (C) 2022-23 Avimanyu Bandyopadhyay

' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>. 

strUser = CreateObject("WScript.Network").UserName

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set DiscordFolder = objFSO.GetFolder("C:\Users\" + strUser + "\AppData\Local\Discord")
Set AllDiscordSubFolders = DiscordFolder.SubFolders

For Each DiscordSubFolder in AllDiscordSubFolders
	t = left(DiscordSubFolder.name,8)
	If t = "app-0.0." Then
		x = right(DiscordSubFolder.name,3)
		If right(DiscordSubFolder.name,3) > x Then
			x = right(DiscordSubFolder.name,3)
		Exit For
		End If
	End If
Next

Set DiscordPath = objFSO.GetFolder("C:\Users\" + strUser + "\AppData\Local\Discord\app-0.0." & x & "")

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.CurrentDirectory = DiscordPath

WshShell.Run "Discord.exe --start-minimized"

WScript.Quit
