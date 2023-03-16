# File-PDF-Combiner
A Desktop App that can combine excel, word, presentations, images and pdf files into one single pdf file with bookmarks. Thing that makes it standout is that it can also merge password protected Powerpoints. Currently In Release v1.0.0, Modern UI does not support Drag and Drop for external files while Simple UI does.
1. [Features](#features)
2. [Requirements](#requirements)
2. [Screen Shots](#screen-shots)
3. [Download](#download)
4. [Languages and Packages](#languages-and-packages) 
5. [Bugs](#bugs)
6. [Licence](#licence) 

## Features
1. Can combine password proctected Read only PowerPoints
2. Many ppt to pdf converters/compilers tend to open ppts during compilation. this can cause quite a hinderance in users normal pc usage. This app avoids it, making it user friendly 
3. Drag and Drop Files into app (Simple UI Only)
4. Can move multiple selected files within table
5. Can move files with in table by dragging

## Requirements
64-bit Windows Operating System
MS Office (for compiling Excel, PowerPoints or .docx)

## Screen Shots
### Light Mode
![Light Mode ScreenShot](https://github.com/OM3R-Nazir/File-PDF-Combiner/blob/main/screenshots/light.png?raw=true)
### Dark Mode
![Dark Mode ScreenShot](https://github.com/OM3R-Nazir/File-PDF-Combiner/blob/main/screenshots/dark.png?raw=true)
### Simple
![Simple UI ScreenShot](https://github.com/OM3R-Nazir/File-PDF-Combiner/blob/main/screenshots/simple.png?raw=true)

## Download
### Note
Please note that anti-virus and windows might find these files harmful and give false positive beacuse this project is not from a famous publisher. You can continue without any worries.
If following message shows up

![Windows protected your pc](https://github.com/OM3R-Nazir/File-PDF-Combiner/blob/main/screenshots/winprot1.jpeg)

Just click on more info and you will see the following message

![Windows protected your pc](https://github.com/OM3R-Nazir/File-PDF-Combiner/blob/main/screenshots/winprot2.jpeg)

Clicking on runany way will run the file.
Also dont forget to add this file to exception list in windows defender. [Here is how to do it](https://support.microsoft.com/en-us/windows/add-an-exclusion-to-windows-security-811816c0-4dfd-af4a-47e4-c301afe13b26#:~:text=Go%20to%20Start%20%3E%20Settings%20%3E%20Update,%2C%20file%20types%2C%20or%20process.)

### Major Types
There are 2 UI distributions
1. [Modern UI](https://github.com/OM3R-Nazir/File-PDF-Combiner/releases/download/v0.1.0-mui/File-PDF.Combiner.exe) : (does not support drag and drop for external files and have some minor bugs(not really effeccting the user) given in [Bugs Section](#bugs))
2. [Simple UI](https://github.com/OM3R-Nazir/File-PDF-Combiner/releases/download/v0.1.0-sui/File-PDF.Combiner.exe)


Note that do try Single Directory build if standalone fails

Check in the Releases section in the side panel to download the files. Once you selected from MUI (Modern) or SUI (Simple), you can then download respective distribution type, either standalone or single directory build

## Languages and Packages
Python Language is used along with tkinter (Simple UI) and custom tkinter (Modern UI). Executable packaging are made through nuikta.

## Licence
[MIT](https://choosealicense.com/licenses/mit/)

## Bugs
### Modern UI
1. Changing Windows Theme mode does not update treeview style
2. Toplevel windows donot open in focus and their icons are changed to default custom tkinter icon

## TODO
- Enabling Drag and Drop in Modern UI, so that Simple UI can be made discontinued
- Make 32 bit release
