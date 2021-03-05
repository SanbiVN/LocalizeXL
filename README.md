# LocalizeXL
Highlights the orientation of the selected cell in worksheet.

### Execute commands in an Excel application using functions:
Execute | Fx | argument
---|---|---
On |	=LocalizeOn()	|
Off | =LocalizeOff() |
Change Color | =LocalizeSetColor("#FCD220")	| <Color Number (or Hex)>
Change Opacity |	=LocalizeSetOpacity(40)	| <Number 20~255>
On/Off Fading |	=LocalizeSetFading(400) |	<miliseconds 0~4000>
Change All |	=LocalizeSet(16711680, 40, True)	| <Color, Opacity, Fading>
Change Add-in or Book |	=LocalizeSpin() |
Reset Add-in |	=LocalizeReset() |
Close Add-in |	=LocalizeQuit()	|
Uninstall Add-in |	=LocalizeUninstall()	|

# SETUP
In the Ribbon tab 'Deverloper' choose Excel Add-ins, Choose button 'Browse...' -> choose file download, Tick Add-in added and click 'OK'
(Enable 'Deverloper': right click on the Ribbon, choose 'Customize the Ribbon')


# EXAMPLE

![LocalizeXL](https://user-images.githubusercontent.com/58664571/110070199-082dd380-7dac-11eb-8b9e-06707ddad1b8.gif)

# SCAN VIRUS
https://www.virustotal.com/gui/
