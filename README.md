# LocalizeXL - Add-in Excel
Highlights the orientation of the selected cell in Excel Worksheet (Window  Only).

[Nhấn tải LocalizeXL v1.71](https://github.com/SanbiVN/LocalizeXL/releases/download/localize_xl/LocalizeXL-v1.71.zip)
[![Tổng tải xuống](https://img.shields.io/github/downloads/SanbiVN/LocalizeXL/total.svg)]()

# EXAMPLE

![LocalizeXL](https://user-images.githubusercontent.com/58664571/110070199-082dd380-7dac-11eb-8b9e-06707ddad1b8.gif)

![LocalizeXL](https://github.com/SanbiVN/LocalizeXL/blob/main/test/vba%20localize%20style.gif)

### Execute commands in an Excel application using functions:

Execute | Fx | argument
---|---|---
On |	=LocalizeOn()	|
Off | =LocalizeOff() |
| Change Color | =LocalizeSetColor("#FCD220") | Hexadecimal Number
|  | =LocalizeSetColor(255)	| Color Number
|  | =LocalizeSetColor("yellow")	| Name color (yellow/ye/yl, red/re, blue, green/gr, cyan/cy, magenta/ma, white/wh/wi, black/bl/bk, orange/or, pink, purple/pu, silver/si, violet/vi, Brown/br, Beige/be)
Change Opacity | =LocalizeSetOpacity(40)	| <Number 20~255>
On/Off Fading |	=LocalizeSetFading(400) |	<miliseconds 0~4000>
Change All |	=LocalizeSet(16711680, 40, True)	| <Color, Opacity, Fading>
Change Add-in or Book |	=LocalizeSpin() |
Reset Add-in |	=LocalizeReset() |
Close Add-in |	=LocalizeQuit()	|
Uninstall Add-in |	=LocalizeUninstall()	|

# SETUP

## Add developer tab to ribbon:
1. Right-click on the Ribbon, and click Customize the Ribbon.
2. In the Customize the Ribbon list, add a check mark to the Developer tab.

![developertabadd](https://user-images.githubusercontent.com/58664571/110081294-4d5b0100-7dbe-11eb-814b-946de593dc11.png)

3. Add developer tab to ribbon

## Install Add-in
1. Open Excel, and on the Ribbon, click the Developer tab
2. Click the Add-ins button.

![ribbontabmacros](https://user-images.githubusercontent.com/58664571/110081583-b773a600-7dbe-11eb-81f4-8958c2999e31.png)

3. In the Add-in dialog box, find the My Macros Custom Ribbon Tab add-in, and add a check mark to its name.
install the Add-in

![Add-ins](https://user-images.githubusercontent.com/58664571/110081743-f73a8d80-7dbe-11eb-89c0-fc136b9573eb.jpg)

4. Click OK, to close the Add-ins window.

# NOTE
If your project has real-time working Macros VBA, then turn off LocalizeXL, to avoid collisions, by typing in the empty cell function =LocalizeOff()

# SCAN VIRUS
https://www.virustotal.com/gui/

