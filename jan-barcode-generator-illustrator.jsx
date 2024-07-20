#target illustrator

// Define constants
var BATCH_SIZE = 10;
var BREAK_POINT = 300;
var SLEEP_TIME_SHORT = 100;
var SLEEP_TIME_LONG = 2000;

/**
 * Creates a UI panel and gets user input
 * @returns {Object} User input data
 */
function createUIPanel() {
    var dialog = new Window('dialog', 'Barcode Generator');
    
    dialog.alignChildren = 'left';
    
    // JAN code input
    dialog.add('statictext', undefined, 'Enter JAN codes (one per line):');
    var janInput = dialog.add('edittext', undefined, '', {multiline: true});
    janInput.size = [300, 200];
    
    // OK button
    var okButton = dialog.add('button', undefined, 'OK');
    okButton.onClick = function() {
        dialog.close();
    };
    
    dialog.show();
    
    // Process JAN codes
    var janCodes = janInput.text.split('\n');
    var cleanedJanCodes = [];
    for (var i = 0; i < janCodes.length; i++) {
        var trimmedCode = janCodes[i].replace(/^\s+|\s+$/g, '');
        if (trimmedCode.length > 0) {
            cleanedJanCodes.push(trimmedCode);
        }
    }
    
    return {
        janList: cleanedJanCodes
    };
}

/**
 * Creates a progress bar
 * @param {string} title Title of the progress bar
 * @param {number} maxValue Maximum value
 * @returns {Window} Progress bar window
 */
function createProgressBar(title, maxValue) {
    var win = new Window("palette", title, undefined, {closeButton: false});
    win.progressBar = win.add("progressbar", undefined, 0, maxValue);
    win.progressBar.preferredSize.width = 300;
    win.status = win.add("statictext", undefined, "Processing: 0/" + maxValue);
    win.cancelButton = win.add("button", undefined, "Cancel");
    win.cancelButton.onClick = function() { win.close(); };
    win.show();
    return win;
}

/**
 * Main function to generate barcodes
 * @param {Array} janList List of JAN codes
 * @param {string} outputFolder Path to output folder
 */
function createBarcodes(janList, outputFolder) {
    var progressWin = createProgressBar("Generating Barcodes", janList.length);
    var isCancelled = false;
    var skippedCount = 0;

    for (var i = 0; i < janList.length && !isCancelled; i += BATCH_SIZE) {
        var batch = janList.slice(i, Math.min(i + BATCH_SIZE, janList.length));
        skippedCount += processBatchAndCountSkipped(batch, outputFolder);
        
        // Update progress bar
        progressWin.progressBar.value = i + batch.length;
        progressWin.status.text = "Processing: " + (i + batch.length) + "/" + janList.length + " (Skipped: " + skippedCount + ")";
        
        // Update UI
        progressWin.update();
        
        // Check for cancellation
        if (progressWin.cancelButton.pressed) {
            isCancelled = true;
        }
        
        $.sleep(SLEEP_TIME_SHORT);

        if ((i + batch.length) % BREAK_POINT === 0 && i + batch.length < janList.length) {
            progressWin.status.text = "Taking a short break...";
            progressWin.update();
            $.sleep(SLEEP_TIME_LONG);
        }
    }

    progressWin.close();
    if (isCancelled) {
        alert("Process was cancelled.");
    } else {
        alert("All processing completed.\nProcessed JAN codes: " + (janList.length - skippedCount) + "\nSkipped JAN codes: " + skippedCount);
    }
}

/**
 * Processes a batch and counts skipped files
 * @param {Array} batch Batch of JAN codes to process
 * @param {string} outputFolder Path to output folder
 * @returns {number} Number of skipped files
 */
function processBatchAndCountSkipped(batch, outputFolder) {
    var skippedCount = 0;
    for (var i = 0; i < batch.length; i++) {
        var jan = Jan(batch[i]);
        if (jan) {
            var svgFile = new File(outputFolder + "/" + batch[i] + ".svg");
            var pngFile = new File(outputFolder + "/" + batch[i] + ".png");
            
            if (svgFile.exists && pngFile.exists) {
                skippedCount++;
                continue;
            }
            
            var doc = createDocument();
            createBarcodeText(doc, jan);
            
            if (!svgFile.exists) {
                exportSVG(doc, batch[i], outputFolder);
            }
            
            if (!pngFile.exists) {
                exportPNG(doc, batch[i], outputFolder);
            }
            
            doc.close(SaveOptions.DONOTSAVECHANGES);
        } else {
            alert("Invalid JAN code: " + batch[i]);
        }
    }
    $.gc();
    return skippedCount;
}

/**
 * Creates a new document
 * @returns {Document} Created document
 */
function createDocument() {
    return app.documents.add(DocumentColorSpace.RGB, 200, 90);
}

/**
 * Creates barcode text and centers it on the canvas
 * @param {Document} doc Illustrator document
 * @param {string} jan JAN code
 */
function createBarcodeText(doc, jan) {
    var textFrame = doc.textFrames.add();
    textFrame.contents = jan;
    var font = app.textFonts.getByName("JANCODE-nicWabun");
    textFrame.textRange.characterAttributes.textFont = font;
    textFrame.textRange.characterAttributes.size = 72;

    // Get document dimensions
    var docWidth = doc.width;
    var docHeight = doc.height;

    // Get text frame dimensions
    var textWidth = textFrame.width;
    var textHeight = textFrame.height;

    // Calculate center position
    var centerX = (docWidth - textWidth) / 2;
    var centerY = (docHeight + textHeight) / 2; // Illustrator's coordinate system goes from bottom to top

    // Position the barcode at the center
    textFrame.position = [centerX+5, centerY];

}

/**
 * Exports as SVG
 * @param {Document} doc Document
 * @param {string} jan JAN code
 * @param {string} outputFolder Path to output folder
 */
function exportSVG(doc, jan, outputFolder) {
    var file = new File(outputFolder + "/" + jan + ".svg");
    var exportOptions = new ExportOptionsSVG();
    exportOptions.embedRasterImages = true;
    doc.exportFile(file, ExportType.SVG, exportOptions);
}

/**
 * Exports as PNG
 * @param {Document} doc Document
 * @param {string} jan JAN code
 * @param {string} outputFolder Path to output folder
 */
function exportPNG(doc, jan, outputFolder) {
    var file = new File(outputFolder + "/" + jan + ".png");
    var exportOptions = new ExportOptionsPNG24();
    exportOptions.antiAliasing = true;
    exportOptions.transparency = true;
    exportOptions.artBoardClipping = true;
    doc.exportFile(file, ExportType.PNG24, exportOptions);
}

/**
 * Gets the path to the JAN folder on the desktop
 * @returns {string} Path to the JAN folder
 */
function getDesktopJANFolder() {
    var desktopPath = Folder.desktop.fsName;
    var janFolder = new Folder(desktopPath + "/JAN");
    if (!janFolder.exists) {
        janFolder.create();
    }
    return janFolder.fsName;
}

/**
 * Calculates the check digit
 * @param {string} strJancode JAN code
 * @returns {number} Check digit
 */
function CD(strJancode) {
    var g = 0, k = 0, h = 0;
    for (var i = 12; i > 0; i -= 2) {
        g += parseInt(strJancode.charAt(i - 1));
        k += parseInt(strJancode.charAt(i - 2));
    }
    h = (g * 3 + k) % 10;
    return (h === 0) ? 0 : 10 - h;
}

/**
 * Gets the character array for barcodes
 * @returns {Array} Character array for barcodes
 */
function getBar() {
    return [
        ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"],
        ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"],
        ["L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    ];
}

/**
 * Processes 8-digit JAN code
 * @param {string} n 8-digit JAN code
 * @returns {string} Processed JAN code
 */
function Eight(n) {
    var strJanfont = "Y";
    var BAR = getBar();
    for (var i = 0; i < 4; i++) {
        strJanfont += n.charAt(i);
    }
    strJanfont += "K";
    for (var i = 4; i < 7; i++) {
        strJanfont += BAR[2][parseInt(n.charAt(i))];
    }
    strJanfont += BAR[2][CD("00000" + n)];
    strJanfont += "Z";
    return strJanfont;
}

/**
 * Processes 13-digit JAN code
 * @param {string} n 13-digit JAN code
 * @returns {string|null} Processed JAN code or null if check digit is incorrect
 */
function Thirteen(n) {
    if (parseInt(n.charAt(12)) !== CD(n.slice(0, 12))) {
        return null;
    }

    var Initial = [
        "000000", "001011", "001101", "001110",
        "010011", "011001", "011100", "010101",
        "010110", "011010"
    ];
    var BAR = getBar();
    var strJanfont = getStartCode(parseInt(n.charAt(0)));
    for (var i = 1; i <= 6; i++) {
        strJanfont += BAR[parseInt(Initial[parseInt(n.charAt(0))].charAt(i - 1))][parseInt(n.charAt(i))];
    }
    strJanfont += "K";
    for (var i = 7; i <= 12; i++) {
        strJanfont += BAR[2][parseInt(n.charAt(i))];
    }
    strJanfont += "Z";
    return strJanfont;
}

/**
 * Gets the start code
 * @param {number} n Index of start code
 * @returns {string} Start code
 */
function getStartCode(n) {
    var Startbar = ["a", "b", "W", "d", "X", "f", "g", "h", "i", "j"];
    return Startbar[n];
}

/**
 * Processes JAN code
 * @param {string} JANCODE JAN code
 * @returns {string|null} Processed JAN code or null
 */
function Jan(JANCODE) {
    if (JANCODE === "") return null;
    if (!/^\d+$/.test(JANCODE)) return null;
    switch (JANCODE.length) {
        case 7:
            return Eight("0" + JANCODE);
        case 8:
            return Eight(JANCODE);
        case 12:
            return Thirteen("0" + JANCODE);
        case 13:
            return Thirteen(JANCODE);
        default:
            return null;
    }
}

// Script execution
try {
    var uiInput = createUIPanel();
    var outputFolder = getDesktopJANFolder();
    if (uiInput.janList.length > 0 && outputFolder) {
        createBarcodes(uiInput.janList, outputFolder);
    } else {
        alert("Invalid input. Please enter JAN codes and ensure the output folder is correctly set.");
    }
} catch (e) {
    alert("An error occurred: " + e.message);
}
