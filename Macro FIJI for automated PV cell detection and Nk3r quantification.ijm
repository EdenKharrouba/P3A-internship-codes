// This code is used to get automatically the results related to a batch of images saved within the same folder

// Make sure the size of your images is the right one --> Image --> Properties --> pixel width and height in micon
// run("Properties...", "channels=1 slices=11 frames=1 pixel_width=0.3125 pixel_height=0.3125 voxel_depth=1.0000000");

// Select a directory(folder) of images to open
inputFolder = getDirectory("input folder where images are stored");
// Open a dialogue to select the location where the results will be stored
outputFolder = getDirectory("output folder to save the results");

// Get the list of the files, images in the input folder
list = getFileList(inputFolder);
print("Nombre de fichiers détectés dans le dossier d'entrée : " + list.length);
setBatchMode(true);
// Désactiver l'affichage des fenêtres intermédiaires
//setOption("Show Intermediate Results", false);

//run("Read and Write Excel", "file=[C:/Users/uqekharr/OneDrive - The University of Queensland/Documents/2. Immunohistochemistry/Resultstest/Results.xlsx]");

// Create a loop to apply a macro to all the files
for (i=0; i<list.length; i++){  // list.length is equal to the number of images to analyse
    
    nomFeuille = "Image " + (i + 1);  
    
    open(inputFolder + list[i]);  // open the image i
    
    // Insert below the macro you want to apply to all the files
    Original=list[i];
	Original=getTitle();
	run("Properties...", "pixel_width=0.3125 pixel_height=0.3125");
	print(Original);
	run("Clear Results");
	run("Set Measurements...", "mean redirect=None decimal=3");
	selectWindow("Log");
	print("\\Clear");
	print(Original);
	selectWindow("Log");
	filename = getInfo("log");
	//String.copy(filename);
	print("\\Clear");
	// waitForUser("Filename is copied, please paste into spreadsheet.\nPress OK to continue.");
	
	selectWindow(Original);
	setBatchMode("show");
	//waitForUser("Adjust Contrast to see PV neurons for future ROI quality control.\nPress OK to continue.");
	
	run("Duplicate...", "duplicate");
	run("Z Project...", "projection=[Max Intensity]");
	imageTitle=getTitle();
	// Split channels of the duplicated image
	run("Split Channels");

    selectWindow("C2-"+imageTitle);
    run("Duplicate...", " ");
    run("Despeckle");
    run("Remove Outliers...", "radius=2.5 threshold=50 which=Bright");
    Nk3r=getTitle();

    selectWindow("C4-"+imageTitle);
    run("Duplicate...", " ");
    run("Despeckle");
    PV=getTitle();
    run("Duplicate...", " ");

    //Percentage threshold code
    percentage = 98.00; 
    nBins = 256; 
    resetMinAndMax(); 
    getHistogram(values, counts, nBins);  
    // find culmulative sum 
    nPixels = 0; 
    for (j = 0; j<counts.length; j++) 
        nPixels += counts[j]; 
        nBelowThreshold = nPixels * percentage / 100; 
    sum = 0; 
    for (j = 0; j<counts.length; j++) { 
        sum = sum + counts[j]; 
        if (sum >= nBelowThreshold) { 
            setThreshold(values[0], values[j]); 
            print(values[0]+"-"+values[j]+": "+sum/nPixels*100+"%"); 
            j = 99999999;//break 
        } 
    } 
    run("Convert to Mask");
    run("Remove Outliers...", "radius=5 threshold=50 which=Dark");
    run("Invert");
    run("Analyze Particles...", "size=80-Infinity add");
    setBatchMode("show");

    selectWindow(Nk3r);
    run("Duplicate...", " ");
    roiManager("Show All without labels");
    setBatchMode("show");
    // waitForUser("Mask quality control.\nPress OK to continue.");
    roiManager("Measure");
    roiManager("Show None");
    FinalNk3r=getTitle();
    selectWindow("Results");
    
    // Save NK3R expression results
    
    run("Read and Write Excel", "no_count_column file=[" + outputFolder + "Results98criteria.xlsx] dataset_label=[Nk3r expression]");
    //waitForUser("NK3R expression in whole PV+ cells has been copied.\nPress OK to continue.");
    run("Clear Results");

    selectWindow(PV);
    roiManager("Show All without labels");
    roiManager("Measure");
    roiManager("Show None");
    selectWindow("Results");
    
    // Save PV expression results
    run("Read and Write Excel", "no_count_column file=[" + outputFolder + "Results98criteria.xlsx] dataset_label=[PV expression]");
    //waitForUser("PV expression in whole PV+ cells has been copied.\nPress OK to continue.");
    run("Clear Results");


	run("Clear Results");
    run("Close All");
    
    // Clear all the ROI from memory
    roiManager("Reset");

}
