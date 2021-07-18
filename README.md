# Automation in CorelDraw
VBA and CorelDraw

## Summary
1. [About](#about)
2. [Function](#function)
3. [Command](#command)
4. [Screenshots](#prints)

<div id="about" />

## About
This code was developed for a clothing company, focused on the sports area. The program has the functionality to change texts inside the corel and export the vectors in JPG.

<div id="function" />

## Functions

- Change words (shapes).
- Export vector groups, organized in layers.

<div id="command" />

## Command
syntax: ```exportPNG([layerNameShort] As String, [playerName] As String, [numberID] As String, [sizeN] As String)```

|Comand|Description|Requery|
| ---------- | ---------- | ---------- |
| ``exportPNG(...)`` | Function for called comand | Yes |
|  ``layerNameShort`` | Abbreviation to layer name | Yes |
|  ``playerName`` | String for player name | No |
|  ``numberID`` | String for player number | No |
|  ``sizeN`` | String for size peace paper id  | No |

**Attenction**: *If items playerName, numberID and sizeN were empty ("") it will give the impression that the item does not exist in the corel vector.*

<div id="prints" />

## Screenshots
-01 - Windows Corel
![01 - Windows Corel](https://raw.githubusercontent.com/AlysonRM/corel-exportjpg/main/_screenshots/1-WindowCorel.png)
-02 - Select option "Macro Global" inside macros
![2-MacroGlobalOpen](https://raw.githubusercontent.com/AlysonRM/corel-exportjpg/main/_screenshots/2-MacroGlobalOpen.png)
-03 - Fill in the text in the form
![3-FormToExport](https://raw.githubusercontent.com/AlysonRM/corel-exportjpg/main/_screenshots/3-FormToExport.png)
-04 - Exported folders and vectors
![4-VectorExported](https://raw.githubusercontent.com/AlysonRM/corel-exportjpg/main/_screenshots/4-VectorExported.png)