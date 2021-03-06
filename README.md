# Excel Inventory with photos of fire bulkheads in a communications building
## by Alejandra Ramos de Vega
### This Excel was used by the company "Team für Technik" for the project in Germany "Allach Inventur der Brandschotts" ("Inventory of fire bulkheads in Allach") at the request of the company ISS.


## Executive Summary 
* [Introduction](#introduction) 
* [Methodology](#Methodology)
* [Results](#results)
* [Upgrates](#upgrates)


## Introduction
Fire barriers guarantee the safety of each individual room against fire, preventing the spread of smoke and fire and therefore keeping areas and escape routes clear. 
 
![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Example%20fire%20protection2.JPG)

Normally they are used for cables and pipes running through a building or construction.  
 
 ![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Example%20fire%20protection.jpg)

The principle of functioning of these protections, is the use of a material that expands under high temperatures and in this way seals any gap that exists in the wall, in this way the transmission of fire can be avoided for 30 or 90 minutes depending on the material used (these specifications are marked with F30 and F90).
All the bulkheads must come identified with a label, where the manufacturer, the date of revision, the type of bulkhead, among other information is specified. Without this label the bulkhead is considered "not suitable". 

 ![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Example%20fire%20protection3.jpg)



## Methodology 
For this project we created the following Excel file in which we present two sheets, the first one where we can find the main inventory and the second one where we can find dropdown lists for the use of the inventory, such as the status of the bulkhead: 

![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Status.JPG)

In the first sheet, a VBA code is used to insert the images of each bulkhead, which operates in column I and row 7. 

![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Code_II.JPG)

These pictures must be in a folder with the name "Photos" which is located in the same file as the code: 

![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/File.JPG)

In addition to this, the photos must have the same name as the name found in column C (Fire bulkheads ID):


![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/ID.JPG)

The second code found in the Excel sheet is to delete the photos, so that there are no repetitions, every time one wants to run the code to add photos (because there are new bulkheads to add to the inventory) one should first delete all the ones that are already currently in the inventory. This way the final file ends up being less heavy.


![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Code_I.JPG)

## Results
The resulting table with the inventory is presented below: 
![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Excel.png)
A brief explanation of each column:
* Column A: Here the status of the bulkhead is entered (Properly, Unproperly, Need for Intervention )
* Column B: The condition of the bulkhead (OK, Damaged, Open)
* Column C: The ID or name of the bulkhead (this must be unique for each one)
* Column D: The floor where the bulkhead is located (e.g.: Ground floor, 1st floor...)
* Column E: The room number or name of the room from which the cable or pipe runs through the bulkhead.
* Column F: The number or name of the room to which the pipe or cable through the bulkhead leads (all bulkheads in the list must have two sides, one side for each wall). 
* Column G: The function or use of the room in which the bulkhead is located (e.g.: Hallway, Office, Warehouse...)
* Column H: Where in the room the bulkhead is located (near the ceiling or the floor for example)
* Column I: The photo that is added using the mentioned code.
* Column J: The dimensions of the bulkhead (all measurements must be in cm)
* Column K: The measurements of the position of the bulkhead in the room (distance to the nearest wall or walls, to the ceiling and/or to the floor)
* Column L: The use of the bulkhead (e.g.: cable or pipe)
* Column M: The company that manufactured the bulkhead. 
* Column N: The material it is made of (e.g.: Cement bulkhead, Foam bulkhead...)
* Column O: Any comments or deficiencies that need to be mentioned. 
* Column P: The date on which the information was collected.

## Side note: all photos and information used in this excel sheet are for illustrative purposes only, they do not represent real life data. 


## Upgrades
Two new columns have been added (column J and column K) the first one (column J) is to insert the photos of the bulkheads after the due repairs and in the second one (column K) the company that performed such works, for this, the following codes were added to the excel sheet:

The deletion code works in the same way as in the case of column I: 

![Code_I_ColumnJ](https://user-images.githubusercontent.com/93596082/154290045-50f21aee-b351-4526-8f11-d3cfbce99b4c.JPG)

And the code to add images, too, regardless of the order in which the bulkheads are repaired (these are stored in a new folder called "new photos")
![Code_II_ColumnJ](https://user-images.githubusercontent.com/93596082/154291549-d8594599-2066-4835-a29c-461a1364db8b.JPG)

These are the new columns: 
![new columns](https://user-images.githubusercontent.com/93596082/154295399-1e6ecc76-a6e0-4f17-988c-2b7cfa2370a0.JPG)


A **new** brief explanation of each column:
* Column A: Here the status of the bulkhead is entered (Properly, Unproperly, Need for Intervention )
* Column B: The condition of the bulkhead (OK, Damaged, Open)
* Column C: The ID or name of the bulkhead (this must be unique for each one)
* Column D: The floor where the bulkhead is located (e.g.: Ground floor, 1st floor...)
* Column E: The room number or name of the room from which the cable or pipe runs through the bulkhead.
* Column F: The number or name of the room to which the pipe or cable through the bulkhead leads (all bulkheads in the list must have two sides, one side for each wall). 
* Column G: The function or use of the room in which the bulkhead is located (e.g.: Hallway, Office, Warehouse...)
* Column H: Where in the room the bulkhead is located (near the ceiling or the floor for example)
* Column I: The photo that is added using the mentioned code.
* **Column J: Photos after repair (also added with the code)**
* **Column K: Company that performed the repair**
* Column L: The dimensions of the bulkhead (all measurements must be in cm) 
* Column M: The measurements of the position of the bulkhead in the room (distance to the nearest wall or walls, to the ceiling and/or to the floor)
* Column N: The use of the bulkhead (e.g.: cable or pipe)
* Column O: The company that manufactured the bulkhead.  
* Column P: The material it is made of (e.g.: Cement bulkhead, Foam bulkhead...) 
* Column Q: Any comments or deficiencies that need to be mentioned.
* Column R:The date on which the information was collected.
