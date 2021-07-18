# Automation in CorelDraw
VBA and CorelDraw

## Summary
1. [About](#about)
2. [Function](#function)

<div id="about" />

## About
This code was developed for a clothing company, focused on the sports area. The program has the functionality to change texts inside the corel and export the vectors in JPG.

<div id="function" />

## Functions

- Change words (shapes).
- Export vector groups, organized in layers.

## Comando
syntax: ```exportPNG([layerNameShort] As String, [playerName] As String, [numberID] As String, [sizeN] As String)```

|Comand|Description|Requery|
| ---------- | ---------- | ---------- |
| ``exportPNG(...)`` | Function for called comand | Yes |
|  ``layerNameShort`` | Abbreviation to layer name | Yes |
|  ``playerName`` | String for player name | No |
|  ``numberID`` | String for player number | No |
|  ``sizeN`` | String for size peace paper id  | No |

**Attenction**: *If items playerName, numberID and sizeN were empty ("") it will give the impression that the item does not exist in the corel vector.*


