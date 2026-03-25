// Unified Olympus OIR Batch Processing Macro
// Automatically detects Z-stack vs Channel-only and applies appropriate workflow

inputDir = getDirectory("Choose input directory with .oir files");
if (inputDir == "") exit("No input folder chosen.");
outputDir = getDirectory("Choose output directory");
if (outputDir == "") exit("No output folder chosen.");

// Create output subdirectories
File.makeDirectory(outputDir + "Individual_Channels/");
File.makeDirectory(outputDir + "Composite_Images/");

// Iteratively collect all .oir files from inputDir and every subfolder at any depth.
// Uses an explicit directory queue to avoid recursion (ImageJ macro variables are global,
// so recursive functions corrupt shared loop variables).
// Stores results directly into arrays to avoid split() tab-collapsing issues.
dirQueue = newArray(inputDir);
relQueue = newArray("");
queueSize = 1;

oirFiles = newArray(0);
oirPaths = newArray(0);
oirRelDirs = newArray(0);

for (qi = 0; qi < queueSize; qi++) {
    currentDir = dirQueue[qi];
    currentRel = relQueue[qi];
    fileList = getFileList(currentDir);
    for (fli = 0; fli < fileList.length; fli++) {
        fileItem = fileList[fli];
        if (endsWith(fileItem, "/")) {
            dirQueue = Array.concat(dirQueue, currentDir + fileItem);
            relQueue = Array.concat(relQueue, currentRel + fileItem);
            queueSize++;
        } else if (endsWith(toLowerCase(fileItem), ".oir")) {
            oirFiles   = Array.concat(oirFiles,   fileItem);
            oirPaths   = Array.concat(oirPaths,   currentDir + fileItem);
            oirRelDirs = Array.concat(oirRelDirs,  currentRel);
        }
    }
}

if (oirFiles.length == 0) {
    showMessage("No .oir files found!");
    exit();
}

setBatchMode(true);

for (f = 0; f < oirFiles.length; f++) {
    fileName = oirFiles[f];
    fullPath = oirPaths[f];
    relSubDir = oirRelDirs[f];

    // Mirror the input subfolder structure inside each output folder (multi-level safe)
    makeNestedDirs(outputDir + "Individual_Channels/" + relSubDir);
    makeNestedDirs(outputDir + "Composite_Images/" + relSubDir);

    // Extract base name
    baseName = fileName;
    if (indexOf(baseName, ".") > 0) {
        baseName = substring(baseName, 0, indexOf(baseName, "."));
    }

    print("Processing (" + (f+1) + "/" + oirFiles.length + "): " + fileName);

    // Close all windows first
    run("Close All");

    // Open file to check dimensions
    run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Composite view=Hyperstack stack_order=XYCZT windowless=true");
    originalTitle = getTitle();

    // Check dimensions to determine workflow
    getDimensions(width, height, channels, slices, frames);

    print("  Dimensions: " + width + "x" + height + ", Channels: " + channels + ", Z-slices: " + slices + ", Frames: " + frames);

    if (slices > 1) {
        // Z-STACK WORKFLOW
        print("  -> Z-stack detected (" + slices + " slices) - Using Z-stack workflow");
        processZStack(originalTitle, baseName, fullPath, outputDir, relSubDir);

    } else {
        // CHANNEL-ONLY WORKFLOW
        print("  -> Channel-only stack detected - Using channel workflow");
        processChannelStack(originalTitle, baseName, fullPath, outputDir, relSubDir);
    }

    // Clean up
    run("Close All");
}

setBatchMode(false);
print("Processing complete! Files processed: " + oirFiles.length);

// Function for Z-stack processing
function processZStack(originalTitle, baseName, fullPath, outputDir, relSubDir) {
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
        saveAs("Tiff", outputDir + "Individual_Channels/" + relSubDir + "MAX_" + baseName + "-1-" + IJ.pad(i+1, 4) + ".tif");
        close();
    }

    // Step 5: Create composite for Z-stack
    run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Composite view=Hyperstack stack_order=XYCZT windowless=true");
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

        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir + "MAX_" + baseName + ".oir (RGB)-1.tif");
    } else {
        // Single channel Z-stack
        run("RGB Color");
        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir + "MAX_" + baseName + "_Single_Channel.tif");
    }
}

// Helper: creates a directory path level by level (handles multi-level paths)
function makeNestedDirs(fullDirPath) {
    // Normalize slashes to forward slash
    p = replace(fullDirPath, "\\\\", "/");
    p = replace(p, "\\", "/");
    parts = split(p, "/");
    built = "";
    for (pi = 0; pi < parts.length; pi++) {
        if (lengthOf(parts[pi]) == 0) continue;
        // Rebuild path segment by segment
        if (lengthOf(built) == 0) {
            // Handle Windows drive letter (e.g. "E:")
            built = parts[pi];
        } else {
            built = built + "/" + parts[pi];
        }
        if (!File.exists(built + "/")) {
            File.makeDirectory(built + "/");
        }
    }
}

// Function for channel-only processing
function processChannelStack(originalTitle, baseName, fullPath, outputDir, relSubDir) {
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
        saveAs("Tiff", outputDir + "Individual_Channels/" + relSubDir + baseName + "-1-" + IJ.pad(i+1, 4) + ".tif");
        close();
    }

    // Step 5: Create composite for channel stack
    run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Composite view=Hyperstack stack_order=XYCZT windowless=true");
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

        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir + baseName + " (RGB).tif");
    } else {
        // Single channel
        run("RGB Color");
        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir + baseName + "_Single_Channel.tif");
    }
}
