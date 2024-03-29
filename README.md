# VBS Scripts for Precast Concrete and Material Take-off
#### Video Demo:  <URL https://youtu.be/RkUI8aCbwws>
#### Description:
In this final project of the CS50X, I would like to present the project that try to solve the repetitive work which related to Structural Engineering field as following.

There 2 types of concrete work when we try to build some building. First the "Pre-Cast Concrete" which is the concrete element that have manufacturing at the factory then bring to combine at the construction site to create an building. In the other, there is concrete type called "Cast-in-place Concrete" which is casted at the construction site. When comparing these 2 types of concrete, the precast is quite better because this one is easy to control the quality in the factory. So, now the precast concrete is popular that the cast-in-place concrete. But the majority of the structural analysis still did not have the full workflow support for the precast concrete Design that we have to do the normal design process and repetitive work for applied this to the precast concrete workflow such as finding the critical element from many scenarios and category its into the group that match the precast concrete from the factory and check the quantity of material that will be used. So, this is project I try to solve this problem by create programing script that can check "Moment" and "Shear" from concrete element then category into the group that match the value from the user's factory and calculate the quantity of concrete which will be use in this project.

##### Requirement:
1. STAAD.Pro 2023 (Structural Analysis software from Bentley System)
2. STAAD.Pro 2023 Script Editor base on WinWrap Basic Language v10 (Build-in STAAD.Pro software)
3. OpenSTAAD Registry file

##### Pre-Cast Concrete grouping script
In this script my code will try to check every element in each loading situation (load cases and load combination) to find the absolute maximum value of Moment and Shear force for category which group this element should be. The detail of my coding is following
1. Start to declare the group name and deleting the group name in case which the script will be used in the second, third, .... to prevent that the result did not update after make some adjustment in the structural model
2. Count all of the member(structural element) in the file for declare the variable to make sure that the script can work in every size of the structural model
3. Count all of load case and load combination scenarios in the model and combine its into array
4. Finding the absolute maximum moment value of each member from every load scenarios
5. Finding the absolute maximum shear force value of each member from every load scenarios
6. Category each member into the group by considering between the maximum moment value and the precast value group
7. Category each member into the group by considering between the maximum shear force value and the precast value group
8. Make a decision to select the envelop case when compare between the moment and the shear group by remark on the member which group it should be
9. Check each member and select only the vertical member for do the CreateGroup operation
10. Do the Analysis of the structural model because the create group operation is in the modeling phase after done that have to do the analysis again.

##### Material Take off script
This script try to solve the problem that the staad pro basic software did not support the material take off function for the concrete material. The script will check all of the member in the structural model next calculate the properties and sum its into the same type(Section) then report to the user by creating the table which has the detail as following
1. Count all of the member in the structural model to use as variable and get the member list
2. Check and Get the properties of each member including member dimension, section name, material property and weight.
3. Check the reference member number for prevent the case that some member was deleted from the model that affect the reference number did not run in sequence
4. Collect the member which has the same property into the same variable to calculate the summation of each type
5. Convert KN to Kg unit and roundup the decimal
6. Create the empty table and fill the result into the table by loop to the column of table and then show the result to user.
