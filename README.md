# BOM_OOP_Python
Graphical BOM Processor using Python OOP
# Description
The engineering design process consists of 5 stages: Need Assessment, Problem Formulation, Abstraction and Synthesis, Analysis, and Implementation. In the Abstraction and Synthesis phase, it is common and useful to decompose the product into its constituent parts in a hierarchical manner in a table called the Bill of Materials (BOM). A fictitious and unrealistic example of this table can be seen in the figure below and in the "MySampleData.xlsx" file in the repository.

![image](https://github.com/Sina-Taghizadeh/BOM_OOP_Python/assets/162900845/c169faca-bd27-4c23-b92d-66d749e404bb)

As can be seen in the table above, each component has a level that specifies its hierarchy. For example, the 'body' is at level 1, and 'Front body', 'Back body', and 'Polymer handle' are its first sub-assemblies and are at level 2. The 'Front body' itself has two sub-assemblies called 'Nozzle' and 'Power cable', which are at level 3, and the 'Power cable' itself has one more sub-assembly called 'The end of the cable', which is at level 4. The same can be extracted from the table above for other components.

In addition to the level, each component in the table above has other information such as id, unit, Quantity, and also Labor Cost (LC), Machine Cost (MC), Other Cost (OC), and Purchase Cost (PC). Therefore, to find the cost required for each component, one must use the different costs of itself and its sub-assemblies in the number that exists from them, and reach the desired results of these diagrams. Doing this manually or with the Excel file itself is very time-consuming and inefficient, so here a graphical program is written in Python in an Object-Oriented manner so that these tables can be easily used.

# How to use
## 1. Open the program:
Run the program to open the following window.

![image](https://github.com/Sina-Taghizadeh/BOM_OOP_Python/assets/162900845/ee13af1e-4ea6-4490-a514-d68ff45817bc)

## 2. Select the BOM file:
In this window, after selecting the desired Excel file and filling in the information above, click on "Confirm". If there is no error in entering the information, the "Confirm" button will change to the "Start" button.

![image](https://github.com/Sina-Taghizadeh/BOM_OOP_Python/assets/162900845/f6aded86-2369-48e9-a9a9-04560bf00a01)

## 3. Start the analysis:
Clicking on this button opens the following window where you can request and receive the desired information from the program.

![image](https://github.com/Sina-Taghizadeh/BOM_OOP_Python/assets/162900845/0f34de3b-3caf-4abb-b990-903fc9d3188f)

## 4. Select the component and information:
After selecting the desired component, you can open the second Combobox which will be as follows and will show the desired information by selecting each one.

![image](https://github.com/Sina-Taghizadeh/BOM_OOP_Python/assets/162900845/7797a723-ab8b-48d3-a4a9-45235b64700a)

## 5. Get the results:
For example, by selecting "body" and "Children", it shows us the first sub-assemblies of it along with the required number of each one as follows.

![image](https://github.com/Sina-Taghizadeh/BOM_OOP_Python/assets/162900845/28b3b388-43f1-44fa-9fc3-53218430692c)

# Further development:
Other features can also be added to this program and make it more complete. I would be happy if someone could develop this program.
sina.taghizadeh123@gmail.com

# Acknowledgements
This program was developed as a course project under the supervision of Dr. Mehrdad Kazerooni. I would like to thank my friend Mohammad Zolfaghari for his valuable contributions to this project.


