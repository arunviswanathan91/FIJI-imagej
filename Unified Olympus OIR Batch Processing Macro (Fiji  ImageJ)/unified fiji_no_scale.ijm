// Unified Olympus OIR Batch Processing Macro
// Automatically detects Z-stack vs Channel-only and applies appropriate workflow
//
// Changelog (quality fixes):
//   - Removed 10x upscaling (was primary cause of pixelation/blurring)
//   - Removed hardcoded Set Scale in inches (Bio-Formats already sets correct µm calibration)
//   - Added per-channel Enhance Contrast before 16→8-bit RGB conversion
//   - Fixed individual channel saves to use per-channel duplication (preserves LUT color)
//   - Composite now uses Flatten instead of Stack to RGB for correct LUT rendering

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
dirQueue  = newArray(inputDir);
relQueue  = newArray("");
queueSize = 1;

oirFiles   = newArray(0);
oirPaths   = newArray(0);
oirRelDirs = newArray(0);

for (qi = 0; qi < queueSize; qi++) {
    currentDir = dirQueue[qi];
    currentRel = relQueue[qi];
    fileList   = getFileList(currentDir);
    for (fli = 0; fli < fileList.length; fli++) {
        fileItem = fileList[fli];
        if (endsWith(fileItem, "/")) {
            dirQueue  = Array.concat(dirQueue,  currentDir + fileItem);
            relQueue  = Array.concat(relQueue,  currentRel + fileItem);
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
    fileName  = oirFiles[f];
    fullPath  = oirPaths[f];
    relSubDir = oirRelDirs[f];

    // Mirror the input subfolder structure inside each output folder (multi-level safe)
    makeNestedDirs(outputDir + "Individual_Channels/" + relSubDir);
    makeNestedDirs(outputDir + "Composite_Images/"    + relSubDir);

    // Extract base name (strip extension)
    baseName = fileName;
    if (indexOf(baseName, ".") > 0) {
        baseName = substring(baseName, 0, indexOf(baseName, "."));
    }

    print("Processing (" + (f+1) + "/" + oirFiles.length + "): " + fileName);

    run("Close All");

    // Open at native resolution — Bio-Formats reads the correct µm/pixel calibration
    // from the OIR metadata automatically; no Set Scale needed.
    run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Composite view=Hyperstack stack_order=XYCZT windowless=true");
    originalTitle = getTitle();

    getDimensions(width, height, channels, slices, frames);
    print("  Dimensions: " + width + "x" + height + ", Channels: " + channels + ", Z-slices: " + slices + ", Frames: " + frames);

    if (slices > 1) {
        print("  -> Z-stack detected (" + slices + " slices) - Using Z-stack workflow");
        processZStack(originalTitle, baseName, fullPath, outputDir, relSubDir);
    } else {
        print("  -> Channel-only stack detected - Using channel workflow");
        processChannelStack(originalTitle, baseName, fullPath, outputDir, relSubDir);
    }

    run("Close All");
}

setBatchMode(false);
print("Processing complete! Files processed: " + oirFiles.length);

// =============================================================================
// Z-STACK WORKFLOW
// =============================================================================
function processZStack(originalTitle, baseName, fullPath, outputDir, relSubDir) {

    // --- Step 1: Max Intensity Z-Projection (operates on the original hyperstack) ---
    selectWindow(originalTitle);
    run("Z Project...", "projection=[Max Intensity]");
    projTitle = getTitle();
    close(originalTitle);

    // --- Step 2: Save each channel individually, preserving LUT color ---
    selectWindow(projTitle);
    getDimensions(w, h, ch, sl, fr);

    for (c = 1; c <= ch; c++) {
        selectWindow(projTitle);
        Stack.setChannel(c);

        // Normalize contrast for 16→8-bit conversion so intensity range is preserved
        run("Enhance Contrast", "saturated=0.35");

        // Duplicate this single channel (retains its LUT/color assignment)
        run("Duplicate...", "title=ch" + c + "_save duplicate channels=" + c);
        run("RGB Color");

        saveAs("Tiff", outputDir + "Individual_Channels/" + relSubDir
               + "MAX_" + baseName + "-1-" + IJ.pad(c, 4) + ".tif");
        close();
    }

    // --- Step 3: Build composite (exclude last channel = TD/brightfield) ---
    selectWindow(projTitle);
    getDimensions(w, h, ch, sl, fr);

    if (ch > 1) {
        // Normalize each fluorescence channel before flattening
        for (c = 1; c < ch; c++) {
            Stack.setChannel(c);
            run("Enhance Contrast", "saturated=0.35");
        }

        // Build active-channel pattern: all fluorescence ON, last (TD) OFF
        activePattern = "";
        for (c = 1; c <= ch; c++) {
            if (c < ch) activePattern += "1";
            else        activePattern += "0";
        }
        Stack.setActiveChannels(activePattern);
        Stack.setDisplayMode("composite");

        // Flatten renders the composite exactly as displayed (respects LUTs + active channels)
        run("Flatten");
        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir
               + "MAX_" + baseName + "_composite.tif");
        close();
    } else {
        // Single-channel Z-stack
        Stack.setChannel(1);
        run("Enhance Contrast", "saturated=0.35");
        run("Flatten");
        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir
               + "MAX_" + baseName + "_single_channel.tif");
        close();
    }

    close(projTitle);
}

// =============================================================================
// CHANNEL-ONLY WORKFLOW
// =============================================================================
function processChannelStack(originalTitle, baseName, fullPath, outputDir, relSubDir) {

    selectWindow(originalTitle);
    getDimensions(w, h, ch, sl, fr);

    // --- Step 1: Save each channel individually, preserving LUT color ---
    for (c = 1; c <= ch; c++) {
        selectWindow(originalTitle);
        Stack.setChannel(c);

        // Normalize contrast for 16→8-bit conversion so intensity range is preserved
        run("Enhance Contrast", "saturated=0.35");

        // Duplicate this single channel (retains its LUT/color assignment)
        run("Duplicate...", "title=ch" + c + "_save duplicate channels=" + c);
        run("RGB Color");

        saveAs("Tiff", outputDir + "Individual_Channels/" + relSubDir
               + baseName + "-1-" + IJ.pad(c, 4) + ".tif");
        close();
    }

    // --- Step 2: Build composite (exclude last channel = TD/brightfield) ---
    selectWindow(originalTitle);

    if (ch > 1) {
        // Normalize each fluorescence channel before flattening
        for (c = 1; c < ch; c++) {
            Stack.setChannel(c);
            run("Enhance Contrast", "saturated=0.35");
        }

        // Build active-channel pattern: all fluorescence ON, last (TD) OFF
        activePattern = "";
        for (c = 1; c <= ch; c++) {
            if (c < ch) activePattern += "1";
            else        activePattern += "0";
        }
        Stack.setActiveChannels(activePattern);
        Stack.setDisplayMode("composite");

        // Flatten renders the composite exactly as displayed (respects LUTs + active channels)
        run("Flatten");
        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir
               + baseName + "_composite.tif");
        close();
    } else {
        // Single channel
        Stack.setChannel(1);
        run("Enhance Contrast", "saturated=0.35");
        run("Flatten");
        saveAs("Tiff", outputDir + "Composite_Images/" + relSubDir
               + baseName + "_single_channel.tif");
        close();
    }

    close(originalTitle);
}

// =============================================================================
// HELPER: creates a directory path level by level (handles multi-level paths)
// =============================================================================
function makeNestedDirs(fullDirPath) {
    p     = replace(fullDirPath, "\\\\", "/");
    p     = replace(p, "\\", "/");
    parts = split(p, "/");
    built = "";
    for (pi = 0; pi < parts.length; pi++) {
        if (lengthOf(parts[pi]) == 0) continue;
        if (lengthOf(built) == 0) {
            built = parts[pi];          // Windows drive letter (e.g. "E:")
        } else {
            built = built + "/" + parts[pi];
        }
        if (!File.exists(built + "/")) {
            File.makeDirectory(built + "/");
        }
    }
}
