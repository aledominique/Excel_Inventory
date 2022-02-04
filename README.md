# Excel Inventory with photos of firewalls in a communications building
## by Alejandra Ramos de Vega
### This Excel was used by the company "Team für Technik" for the project in Germany "Allach Inventur der Brandschotts" ("Inventory of firewalls in Allach") at the request of the company ISS.


## Executive Summary 
* [Introduction](#introduction) 
* [Methodology](#Methodology)
* [Results](#results)


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

