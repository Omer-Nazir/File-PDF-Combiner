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
(NOTE!!: Please note that anti-virus and windows might find these files harmful and give false positive. You can continue without any worries.)
If following message shows up
![Windows protected your pc](https://github.com/OM3R-Nazir/File-PDF-Combiner/blob/main/screenshots/winprot1.png?raw=true)
Just click on more info and you will see the following message
![Windows protected your pc](https://github.com/OM3R-Nazir/File-PDF-Combiner/blob/main/screenshots/winprot2.png?raw=true)
Clicking on runany way will run the file

There are 2 UI distributions
1. Modern UI: (does not support drag and drop for external files and have some minor bugs(not really effeccting the user) given in [Bugs Section](#bugs))
2. Simple UI
Each UI have 2 distributions
1. Standalone file build: consists of single Standalone exe file
2. Single Directory build: consists of a directory. Place directory somewhere in PC, then use shortcut to access the app. (can have better performance then standalone file build)

Check in the Releases in the side panel to download the files

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
