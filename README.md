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
Para este proyecto se creo el siguiente [archivo de Excel](https://github.com/aledominique/Excel_Inventory/blob/main/Inventory%20of%20fire%20bulkheads.xlsm) En el cual se presentan dos hojas, la primera donde se encuentra el inventario principal y el segundo se encuentran dropdown listas para el uso del inventario, como por ejemplo el estado de la mampara: 

![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Status.JPG)

En la primera hoja, se utiliza un codigo de VBA, para insertar las imagenes de cada mampara, el cual actua en la columna I y en la fila 7. 

![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Code_II.JPG)

Estas fotos deben encontrarse en una carpeta con el nombre "Photos" la cual se encuentra localizada en la misma carpeta que el código: 

![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/File.JPG)

Ademas de esto, las fotos deben tener el mismo nombre que el nombre que se encuentra en la columna C (Fire bulkheads ID):


![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/ID.JPG)

El segundo codigo que se encuentra en la hoja de Excel es para borrar las fotos, para que no existan repeticiones, cada vez que se desee correr el codigo de añadir fotos (porque hay nuevas mamparas que agregar al inventario) se deberian borrar primero todas la que ya se encuentran actualmente en el inventario para que de esta forma no añadir las viejas fotos otra vez y así evitar que el archivo final sea extremandamente pesado. 

![grafik](https://github.com/aledominique/Excel_Inventory/blob/main/Photos/Code_I.JPG)

## Results

