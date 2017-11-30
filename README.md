# Workbook Map

This macro creates a map of the activeworkbook in the activesheet. This map is
made of one box for each tab. Arrows between boxes indicate relationships
between tabs (eg: for each worksheet, the width of the arrow from any other
worksheet depends on the number of cells of the first worksheet making use or
calling the second one). Example:

![Alt text](https://github.com/jabellcu/workbook_map/blob/master/Workbook_map_EXAMPLE.png)

# Instructions

1. Import "Workbook_map.bas" to your workbook
2. Activate the tab where you want the workbook map to be created. Add a new
   tab if necesary.
3. Run "create_wb_map"

# Notes

1. "add_dependency_arrows_to_boxes" macro can take several minutes, depending
   on the workbook's complexity.
2. Additional auxiliary tools are included:
    - "AUX_clean_shapes": deletes all shapes
    - "AUX_clean_dependecy_arrows": deletes all connectors
    - "AUX_change_dependecy_arrows": modify existing connectors

# Disclaimer

Use at yout own risk! The author doesn't accept any responisbility over and is
not liable for any damage caused by execution of this code and/or any modified
version of it.
