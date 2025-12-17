// Unified Olympus OIR Batch Processing Macro
// Automatically detects Z-stack vs Channel-only and applies appropriate workflow

inputDir = getDirectory("Choose input directory with .oir files");
if (inputDir == "") exit("No input folder chosen.");
outputDir = getDirectory("Choose output directory");
if (outputDir == "") exit("No output folder chosen.");

// Create output subdirectories
File.makeDirectory(outputDir + "Individual_Channels/");
File.makeDirectory(outputDir + "Composite_Images/");

// Get .oir files
fileList = getFileList(inputDir);
oirFiles = newArray();
for (i = 0; i < fileList.length; i++) {
    if (endsWith(fileList[i], ".oir")) {
        oirFiles[oirFiles.length] = fileList[i];
    }
}

if (oirFiles.length == 0) {
    showMessage("No .oir files found!");
    exit();
}

setBatchMode(true);

for (f = 0; f < oirFiles.length; f++) {
    fileName = oirFiles[f];
    fullPath = inputDir + fileName;
    
    // Extract base name
    baseName = fileName;
    if (indexOf(baseName, ".") > 0) {
        baseName = substring(baseName, 0, indexOf(baseName, "."));
    }
    
    print("Processing (" + (f+1) + "/" + oirFiles.length + "): " + fileName);
    
    // Close all windows first
    run("Close All");
    
    // Open file to check dimensions
    run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Composite view=Hyperstack stack_order=XYCZT");
    originalTitle = getTitle();
    
    // Check dimensions to determine workflow
    getDimensions(width, height, channels, slices, frames);
    
    print("  Dimensions: " + width + "x" + height + ", Channels: " + channels + ", Z-slices: " + slices + ", Frames: " + frames);
    
    if (slices > 1) {
        // Z-STACK WORKFLOW
        print("  -> Z-stack detected (" + slices + " slices) - Using Z-stack workflow");
        processZStack(originalTitle, baseName, fullPath, outputDir);
        
    } else {
        // CHANNEL-ONLY WORKFLOW  
        print("  -> Channel-only stack detected - Using channel workflow");
        processChannelStack(originalTitle, baseName, fullPath, outputDir);
    }
    
    // Clean up
    run("Close All");
}

setBatchMode(false);
print("Processing complete! Files processed: " + oirFiles.length);

// Function for Z-stack processing
function processZStack(originalTitle, baseName, fullPath, outputDir) {
    // Step 1: Scale first
    selectWindow(originalTitle);
    run("Scale...", "x=10 y=10 z=1.0 width=5120 height=5120 depth=4 interpolation=Bilinear average create");
    scaledTitle = getTitle();
    close(originalTitle);
    
    // Step 2: Z Project
    selectWindow(scaledTitle);
    run("Z Project...", "projection=[Max Intensity]");
    projTitle = getTitle();
    close(scaledTitle);
    
    // Step 3: Stack to Images
    selectWindow(projTitle);
    run("Stack to Images");
    close(projTitle);
    
    // Step 4: Convert each image to RGB and save
    allImages = getList("image.titles");
    
    for (i = 0; i < allImages.length; i++) {
        selectWindow(allImages[i]);
        run("RGB Color");
        
        // Save with MAX_ prefix for Z-stack projections
        saveAs("Tiff", outputDir + "Individual_Channels/MAX_" + baseName + "-1-" + IJ.pad(i+1, 4) + ".tif");
        close();
    }
    
    // Step 5: Create composite for Z-stack
    run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Composite view=Hyperstack stack_order=XYCZT");
    run("Scale...", "x=10 y=10 z=1.0 width=5120 height=5120 depth=4 interpolation=Bilinear average create");
    run("Z Project...", "projection=[Max Intensity]");
    
    // Set active channels (exclude last channel if it's TD/brightfield)
    getDimensions(w, h, ch, sl, fr);
    if (ch > 1) {
        activePattern = "";
        for (c = 1; c <= ch; c++) {
            if (c < ch) activePattern += "1";
            else activePattern += "0";
        }
        
        Stack.setActiveChannels(activePattern);
        run("Stack to RGB");
        
        saveAs("Tiff", outputDir + "Composite_Images/MAX_" + baseName + ".oir (RGB)-1.tif");
    } else {
        // Single channel Z-stack
        run("RGB Color");
        saveAs("Tiff", outputDir + "Composite_Images/MAX_" + baseName + "_Single_Channel.tif");
    }
}

// Function for channel-only processing
function processChannelStack(originalTitle, baseName, fullPath, outputDir) {
    // Step 1: Scale
    selectWindow(originalTitle);
    run("Scale...", "x=10 y=10 z=1.0 width=5120 height=5120 depth=4 interpolation=Bilinear average create");
    scaledTitle = getTitle();
    close(originalTitle);
    
    // Step 2: Set scale
    selectWindow(scaledTitle);
    run("Set Scale...", "distance=300 known=1 unit=inch");
    
    // Step 3: Stack to Images
    run("Stack to Images");
    
    // Step 4: Process each channel image
    allImages = getList("image.titles");
    
    // Set scale and convert to RGB for each channel
    for (i = 0; i < allImages.length; i++) {
        selectWindow(allImages[i]);
        run("Set Scale...", "distance=300 known=1 unit=inch");
        run("RGB Color");
        
        // Save individual channel (no MAX_ prefix for non-Z-stack)
        saveAs("Tiff", outputDir + "Individual_Channels/" + baseName + "-1-" + IJ.pad(i+1, 4) + ".tif");
        close();
    }
    
    // Step 5: Create composite for channel stack
    run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Composite view=Hyperstack stack_order=XYCZT");
    run("Scale...", "x=10 y=10 z=1.0 width=5120 height=5120 depth=4 interpolation=Bilinear average create");
    run("Set Scale...", "distance=300 known=1 unit=inch");
    
    getDimensions(w, h, ch, sl, fr);
    
    if (ch > 1) {
        Stack.setDisplayMode("composite");
        
        // Set active channels (exclude last channel if it's TD/brightfield)
        activePattern = "";
        for (c = 1; c <= ch; c++) {
            if (c < ch) activePattern += "1";
            else activePattern += "0";
        }
        
        Stack.setActiveChannels(activePattern);
        run("Stack to RGB");
        
        saveAs("Tiff", outputDir + "Composite_Images/" + baseName + " (RGB).tif");
    } else {
        // Single channel
        run("RGB Color");
        saveAs("Tiff", outputDir + "Composite_Images/" + baseName + "_Single_Channel.tif");
    }
}