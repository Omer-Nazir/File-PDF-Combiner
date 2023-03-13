For pyinstaller add custom tkinter library path in your pc. for example:
```bash
--add-data "c:\users\[USER NAME]\appdata\local\programs\python\python310\lib\site-packages/customtkinter;customtkinter/"
```
For nuikta add following arguments
```bash
--enable-plugin=tk-inter
