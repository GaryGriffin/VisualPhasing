# VisualPhasing



Visual Phasing of genealogy DNA data identifies the grandparents for specific DNA segments. This includes a LibreOffice Calc spreadsheet to perform Visual Phasing from GEDmatch graphs or any other source of segment images. The process for using this spreadsheet is

SETUP

Create shared segment images. This should be done when you want to add new sibling or cousin graphs.

Step 1: run GEDmatch for each pair of siblings. Save the Web Page for each (graphs) locally. This will create a subdirectory of <name>_files which contains each 26 graphs, 2 for each chromosome. They are named <kit1>_<kit2>_chromosome_runID_image.gif. The runID is a unique string for the webpage. The image is either 1 or 2. 

Step 2: Clean up any existing GEDmatchImages data. The command is: make clobber

Step 3: Copy graphs to subdirectory GEDmatchImagesUse the Makefile to create the folder GEDmatchImages and copy all of the graphs to that folder. The command is: make GEDmatchImages

Step 4: Rename the graphs for use with the spreadsheet. This uses the rename.awk script to reformat the names to be kit1_kit2_chromosome_image and creates the linked file <kit2>_<kit1>_chromosome_image.gif. The command is make rename

To add additional siblings or cousins, you need to run step 1 to create the new images and then run the remaining steps to populate the new GEDmatchImages subdirectory.

EXECUTION

Run LibreOffice with the VirtualPhasingTemplate.ods . You will have to Enable Macros to run this spreadsheet.


Sibling Data

The Sibling data is the information and GEDmatch kit number. 

Cousin Data

The Cousing data is the information and the GEDmatch kit number

FIR Data

The FIR data is optional to assist in generating the Mbp value for the recombination point. For this, you run GEDmatch for the sibling pairs with Position Only and copy/paste the table into a sheet named FIR. 

Starting

The first step is to enter all of the Sibling data so that the spreadsheet can import the graphs. 

Once the Sibling Data is entered, you can start the visual phasing process. The first step is to create the individual Chromosome sheets. Click the Reset Sheets button on the Main sheet. This will delete any existing chromosome sheets and create new ones named based on the list in the Sibling Table. If you add or remove a sibling, you should click the Reset Sheets button to regenerate all of the chromosome sheets. 

Once the sheets have been created, you can click the Load Images button to load all of the GEDmatch images into the chromosome sheets. 

To Add the Cousin graphs, click the Show Cousins button. Doing this will rename the button to Clear Cousins. To remove the Cousin graphs, click the Clear Cousins button. You can filter on which cousins are populated in the chromosome sheet by putting any character in the Skip column of the Cousin Table. This will ignore this cousin when you press the Show Cousins button.

If you want to create a segment list to export from the visual phasing, click the Generate Segments button. This will create a sheet named Segments and populate it with all of the segments that the visual phasing has identified – for each sibling. 

If you change the color for the format labeling in the Grandparent Table, you should click Legend Format Update for that change to propagate to the chromosome sheets. 

If you have loaded the FIR Data into the FIR sheet, you can copy each chromosome FIR data to the appropriate sheet by clicking the Load FIR Data button. 

SAMPLE DATA

The SampleData.zip file is a complete set of sample data with a spreadsheet set up for using this data. This can be used for learning how to use the program.

