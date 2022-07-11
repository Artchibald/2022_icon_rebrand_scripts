// #target Illustrator
/************************************************
Script to automate creating variations and exporting files for WTW icons
Starting with an open AI file with a single icon on a single 256 x 256 artboard
– Creates a new artboard at 16x16
- Creates a new artboard at 24x24
- Creates a new artboard at 1400x128
(if these artboards already exist, optionally clears and rebuilds these artboards instead)

- Adds resized copies of the icon to the artboards
- Asks for the name of the icon and adds text to the masthead icon

- Creates exports of the icon:
- RGB EPS
- RGB inverse EPS
- RGB inactive EPS
- PNGs at 1024, 256, 128, 64, 48, 32
- RGB masthead
- CMYK EPS
- CMYK inverse EPS
************************************************/
/*********************************
VARIABLES YOU MIGHT NEED TO CHANGE
**********************************/
var sourceDoc = app.activeDocument;
var RGBColorElements = [
    [127, 53, 178],
    [191, 191, 191],
    [201, 0, 172],
    [50, 127, 239],
    [58, 220, 201],
    [255, 255, 255],
    [128, 128, 128], // Dark grey (unused)
];
// New CMYK values dont math rgb exatcly in new branding 2022 so we stopped the exact comparison part of the script.
// Intent is different colors in print for optimum pop of colors
var CMYKColorElements = [
    [65, 91, 0, 0],
    [0, 0, 0, 25],
    [16, 96, 0, 0],
    [78, 47, 0, 0],
    [53, 0, 34, 0],
    [0, 0, 0, 0],
    [0, 0, 0, 50], // Dark grey (unused)
];
// old not needed, no longer match exactly
// let CMYKColorElements = [
//    [29, 70, 0, 30], //ultraviolet purple
//    [0, 0, 0, 25], //Gray matter light grey
//    [0, 100, 14, 21], // Fireworks magenta
//    [79, 47, 0, 6], //Stratosphere blue
//    [74, 0, 9, 14], // Inifinity turquoise  
//    [0, 0, 0, 0], // white
// ];
var desiredFont = "Graphik-Regular";
var exportSizes = [1024, 512, 256, 128, 64, 48, 32, 24, 16]; //sizes to export
var violetIndex = 0; //these are for converting to inverse and inactive versions
var grayIndex = 1;
var whiteIndex = 5;
//loop default 
var i;
// folder creations
var coreName = "Core";
var expressiveName = "Expressive";
var inverseName = "Inverse";
var inactiveName = "Inactive";
// Masthead
var mastheadName = "Masthead";
// Colors
var rgbName = "RGB";
var cmykName = "CMYK";
var onWhiteName = "onFFF";
//Folder creations
var pngName = "png";
var jpgName = "jpg";
var svgName = "svg";
var epsName = "eps";
var iconFilename = sourceDoc.name.split(".")[0];
var rebuild = true;
// let gutter = 32;
// hide guides
var guideLayer = sourceDoc.layers["Guidelines"];
var name = sourceDoc.name.split(".")[0];
var destFolder = Folder(sourceDoc.path + "/" + name);
var CSTasks = (function () {
    var tasks = {};
    /********************
       POSITION AND MOVEMENT
       ********************/
    //takes an artboard
    //returns its left top corner as an array [x,y]
    tasks.getArtboardCorner = function (artboard) {
        var corner = [artboard.artboardRect[0], artboard.artboardRect[1]];
        return corner;
    };
    //takes an array [x,y] for an item's position and an array [x,y] for the position of a reference point
    //returns an array [x,y] for the offset between the two points
    tasks.getOffset = function (itemPos, referencePos) {
        var offset = [itemPos[0] - referencePos[0], itemPos[1] - referencePos[1]];
        return offset;
    };
    //takes an object (e.g. group) and a destination array [x,y]
    //moves the group to the specified destination
    tasks.translateObjectTo = function (object, destination) {
        var offset = tasks.getOffset(object.position, destination);
        object.translate(-offset[0], -offset[1]);
    };
    //takes a document and index of an artboard
    //deletes everything on that artboard
    tasks.clearArtboard = function (doc, index) {
        //clears an artboard at the given index
        doc.selection = null;
        doc.artboards.setActiveArtboardIndex(index);
        doc.selectObjectsOnActiveArtboard();
        var sel = doc.selection; // get selection
        for (i = 0; i < sel.length; i++) {
            sel[i].remove();
        }
    };
    /*********************************
       SELECTING, GROUPING AND UNGROUPING
       **********************************/
    //takes a document and the index of an artboard in that document's artboards array
    //returns a selection of all the objects on that artboard
    tasks.selectContentsOnArtboard = function (doc, i) {
        doc.selection = null;
        doc.artboards.setActiveArtboardIndex(i);
        doc.selectObjectsOnActiveArtboard();
        return doc.selection;
    };
    //takes a document and a collection of objects (e.g. selection)
    //returns a group made from that collection
    tasks.createGroup = function (doc, collection) {
        var newGroup = doc.groupItems.add();
        for (i = 0; i < collection.length; i++) {
            collection[i].moveToBeginning(newGroup);
        }
        return newGroup;
    };
    //takes a group
    //ungroups that group at the top layer (no recursion for nested groups)
    tasks.ungroupOnce = function (group) {
        for (i = group.pageItems.length - 1; i >= 0; i--) {
            group.pageItems[i].move(group.pageItems[i].layer, 
            /*@ts-ignore*/
            ElementPlacement.PLACEATEND);
        }
    };
    /****************************
       CREATING AND SAVING DOCUMENTS
       *****************************/
    //take a source document and a colorspace (e.g. DocumentColorSpace.RGB)
    //opens and returns a new document with the source document's units and the specified colorspace
    tasks.newDocument = function (sourceDoc, colorSpace) {
        var preset = new DocumentPreset();
        /*@ts-ignore*/
        preset.colorMode = colorSpace;
        /*@ts-ignore*/
        preset.units = sourceDoc.rulerUnits;
        /*@ts-ignore*/
        var newDoc = app.documents.addDocument(colorSpace, preset);
        newDoc.pageOrigin = sourceDoc.pageOrigin;
        newDoc.rulerOrigin = sourceDoc.rulerOrigin;
        return newDoc;
    };
    //take a source document, artboard index, and a colorspace (e.g. DocumentColorSpace.RGB)
    //opens and returns a new document with the source document's units and specified artboard, the specified colorspace
    tasks.duplicateArtboardInNewDoc = function (sourceDoc, artboardIndex, colorspace) {
        var rectToCopy = sourceDoc.artboards[artboardIndex].artboardRect;
        var newDoc = tasks.newDocument(sourceDoc, colorspace);
        newDoc.artboards.add(rectToCopy);
        newDoc.artboards.remove(0);
        return newDoc;
    };
    //takes a document, destination file, starting width and desired width
    //scales the document proportionally to the desired width and exports as a PNG
    tasks.scaleAndExportPNG = function (doc, destFile, startWidth, desiredWidth) {
        var scaling = (100.0 * desiredWidth) / startWidth;
        var options = new ExportOptionsPNG24();
        /*@ts-ignore*/
        options.antiAliasing = true;
        /*@ts-ignore*/
        options.transparency = true;
        /*@ts-ignore*/
        options.artBoardClipping = true;
        /*@ts-ignore*/
        options.horizontalScale = scaling;
        /*@ts-ignore*/
        options.verticalScale = scaling;
        doc.exportFile(destFile, ExportType.PNG24, options);
    };
    tasks.scaleAndExportNonTransparentPNG = function (doc, destFile, startWidth, desiredWidth) {
        var scaling = (100.0 * desiredWidth) / startWidth;
        var options = new ExportOptionsPNG24();
        /*@ts-ignore*/
        options.antiAliasing = true;
        /*@ts-ignore*/
        options.transparency = false;
        /*@ts-ignore*/
        options.artBoardClipping = true;
        /*@ts-ignore*/
        options.horizontalScale = scaling;
        /*@ts-ignore*/
        options.verticalScale = scaling;
        doc.exportFile(destFile, ExportType.PNG24, options);
    };
    //takes a document, destination file, starting width and desired width
    //scales the document proportionally to the desired width and exports as a SVG
    tasks.scaleAndExportSVG = function (doc, destFile, startWidth, desiredWidth) {
        var scaling = (100.0 * desiredWidth) / startWidth;
        var options = new ExportOptionsSVG();
        /*@ts-ignore*/
        options.horizontalScale = scaling;
        /*@ts-ignore*/
        options.verticalScale = scaling;
        // /*@ts-ignore*/
        // options.transparency = true;
        /*@ts-ignore*/
        // options.compressed = false; 
        // /*@ts-ignore*/
        // options.saveMultipleArtboards = true;
        // /*@ts-ignore*/
        // options.artboardRange = ""
        // options.cssProperties.STYLEATTRIBUTES = false;
        // /*@ts-ignore*/
        // options.cssProperties.PRESENTATIONATTRIBUTES = false;
        // /*@ts-ignore*/
        // options.cssProperties.STYLEELEMENTS = false;
        // /*@ts-ignore*/
        // options.artBoardClipping = true;
        doc.exportFile(destFile, ExportType.SVG, options);
    };
    //takes a document, destination file, starting width and desired width
    //scales the document proportionally to the desired width and exports as a SVG
    tasks.scaleAndExportJPEG = function (doc, destFile, startWidth, desiredWidth) {
        var scaling = (100.0 * desiredWidth) / startWidth;
        var options = new ExportOptionsJPEG();
        /*@ts-ignore*/
        options.antiAliasing = true;
        /*@ts-ignore*/
        options.artBoardClipping = true;
        /*@ts-ignore*/
        options.horizontalScale = scaling;
        /*@ts-ignore*/
        options.verticalScale = scaling;
        doc.exportFile(destFile, ExportType.JPEG, options);
    };
    //takes left x, top y, width, and height
    //returns a Rect that can be used to create an artboard
    tasks.newRect = function (x, y, width, height) {
        var rect = [];
        rect[0] = x;
        rect[1] = -y;
        rect[2] = width + x;
        rect[3] = -(height + y);
        return rect;
    };
    /***
    TEXT
    ****/
    //takes a text frame and a string with the desired font name
    //sets the text frame to the desired font or alerts if not found
    tasks.setFont = function (textRef, desiredFont) {
        var foundFont = false;
        /*@ts-ignore*/
        for (var i_1 = 0; i_1 < textFonts.length; i_1++) {
            /*@ts-ignore*/
            if (textFonts[i_1].name == desiredFont) {
                /*@ts-ignore*/
                textRef.textRange.characterAttributes.textFont = textFonts[i_1];
                foundFont = true;
                break;
            }
        }
        if (!foundFont)
            alert("Didn't find the font. Please check if the font is installed or check the script to make sure the font name is right.");
    };
    //takes a document, message string, position array and font size
    //creates a text frame with the message
    tasks.createTextFrame = function (doc, message, pos, size) {
        var textRef = doc.textFrames.add();
        textRef.contents = message;
        textRef.left = pos[0];
        textRef.top = pos[1];
        textRef.textRange.characterAttributes.size = size;
    };
    /***************
       COLOR CONVERSION
       ****************/
    //takes two equal-length arrays of corresponding colors [[R,G,B], [R2,G2,B2],...] and [[C,M,Y,K],[C2,M2,Y2,K2],...] (fairly human readable)
    //returns an array of ColorElements [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...] (usable by the script for fill colors etc.)
    tasks.initializeColors = function (RGBArray, CMYKArray) {
        var colors = new Array(RGBArray.length);
        for (var i_2 = 0; i_2 < RGBArray.length; i_2++) {
            var rgb = new RGBColor();
            rgb.red = RGBArray[i_2][0];
            rgb.green = RGBArray[i_2][1];
            rgb.blue = RGBArray[i_2][2];
            var cmyk = new CMYKColor();
            cmyk.cyan = CMYKArray[i_2][0];
            cmyk.magenta = CMYKArray[i_2][1];
            cmyk.yellow = CMYKArray[i_2][2];
            cmyk.black = CMYKArray[i_2][3];
            colors[i_2] = [rgb, cmyk];
        }
        return colors;
    };
    //take a single RGBColor and an array of corresponding RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns the index in the array if it finds a match, otherwise returns -1
    tasks.matchRGB = function (color, matchArray) {
        //compares a single color RGB color against RGB colors in [[RGB],[CMYK]] array
        for (var i_3 = 0; i_3 < matchArray.length; i_3++) {
            if (Math.abs(color.red - matchArray[i_3][0].red) < 1 &&
                Math.abs(color.green - matchArray[i_3][0].green) < 1 &&
                Math.abs(color.blue - matchArray[i_3][0].blue) < 1) {
                //can't do equality because it adds very small decimals
                return i_3;
            }
        }
        return -1;
    };
    //take a single RGBColor and an array of corresponding RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns the index in the array if it finds a match, otherwise returns -1
    tasks.matchColorsRGB = function (color1, color2) {
        //compares two colors to see if they match
        if (Math.abs(color1.red - color2.red) < 1 &&
            Math.abs(color1.green - color2.green) < 1 &&
            Math.abs(color1.blue - color2.blue) < 1) {
            //can't do equality because it adds very small decimals
            return true;
        }
        return false;
    };
    //takes a pathItems array, startColor and endColor and converts all pathItems with startColor into endColor
    tasks.convertColorCMYK = function (pathItems, startColor, endColor) {
        for (i = 0; i < pathItems.length; i++) {
            if (tasks.matchColorsCMYK(pathItems[i].fillColor, startColor))
                pathItems[i].fillColor = endColor;
        }
    };
    //take a single CMYKColor and an array of corresponding RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns the index in the array if it finds a match, otherwise returns -1
    tasks.matchColorsCMYK = function (color1, color2) {
        //compares two colors to see if they match
        if (Math.abs(color1.cyan - color2.cyan) < 1 &&
            Math.abs(color1.magenta - color2.magenta) < 1 &&
            Math.abs(color1.yellow - color2.yellow) < 1 &&
            Math.abs(color1.black - color2.black) < 1) {
            //can't do equality because it adds very small decimals
            return true;
        }
        return false;
    };
    //takes a pathItems array, startColor and endColor and converts all pathItems with startColor into endColor
    tasks.convertColorRGB = function (pathItems, startColor, endColor) {
        for (i = 0; i < pathItems.length; i++) {
            if (tasks.matchColorsRGB(pathItems[i].fillColor, startColor))
                pathItems[i].fillColor = endColor;
        }
    };
    //takes a pathItems array, endColor and opacity and converts all pathItems into endColor at the specified opacity
    tasks.convertAll = function (pathItems, endColor, opcty) {
        for (i = 0; i < pathItems.length; i++) {
            pathItems[i].fillColor = endColor;
            pathItems[i].opacity = opcty;
        }
    };
    //takes a collection of pathItems and an array of specified RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns an array with an index to the RGB color if it is in the array
    tasks.indexRGBColors = function (pathItems, matchArray) {
        var colorIndex = new Array(pathItems.length);
        for (i = 0; i < pathItems.length; i++) {
            var itemColor = pathItems[i].fillColor;
            colorIndex[i] = tasks.matchRGB(itemColor, matchArray);
        }
        return colorIndex;
    };
    //takes a doc, collection of pathItems, an array of specified colors and an array of colorIndices
    //converts the fill colors to the indexed CMYK colors and adds a text box with the unmatched colors
    //Note that this only makes sense if you've previously indexed the same path items and haven't shifted their positions in the pathItems array
    tasks.convertToCMYK = function (doc, pathItems, colorArray, colorIndex) {
        var unmatchedColors = [];
        for (i = 0; i < pathItems.length; i++) {
            if (colorIndex[i] >= 0 && colorIndex[i] < colorArray.length)
                pathItems[i].fillColor = colorArray[colorIndex[i]][1];
            else {
                var unmatchedColor = "(" +
                    pathItems[i].fillColor.red +
                    ", " +
                    pathItems[i].fillColor.green +
                    ", " +
                    pathItems[i].fillColor.blue +
                    ")";
                unmatchedColors.push(unmatchedColor);
            }
        }
        if (unmatchedColors.length > 0) {
            // NOTE: Don't perform the Artboard Creation Work if there are unmatched colors due to new palettes CMYK and RGB no longer matching.
            return;
            alert("One or more colors don't match the brand palette and weren't converted.");
            unmatchedColors = tasks.unique(unmatchedColors);
            var unmatchedString = "Unconverted colors:";
            for (var i_4 = 0; i_4 < unmatchedColors.length; i_4++) {
                unmatchedString = unmatchedString + "\n" + unmatchedColors[i_4];
            }
            var errorMsgPos = [Infinity, Infinity]; //gets the bottom left of all the artboards
            for (var i_5 = 0; i_5 < doc.artboards.length; i_5++) {
                var rect = doc.artboards[i_5].artboardRect;
                if (rect[0] < errorMsgPos[0])
                    errorMsgPos[0] = rect[0];
                if (rect[3] < errorMsgPos[1])
                    errorMsgPos[1] = rect[3];
            }
            errorMsgPos[1] = errorMsgPos[1] - 20;
            tasks.createTextFrame(doc, unmatchedString, errorMsgPos, 18);
        }
    };
    //takes an array
    //returns a sorted array with only unique elements
    tasks.unique = function (a) {
        var sorted;
        var uniq;
        if (a.length > 0) {
            sorted = a.sort();
            uniq = [sorted[0]];
            for (var i_6 = 1; i_6 < sorted.length; i_6++) {
                if (sorted[i_6] != sorted[i_6 - 1])
                    uniq.push(sorted[i_6]);
            }
            return uniq;
        }
        return [];
    };
    return tasks;
})();
/****************
 * CORE SCRIPT START
 ***************/
function mainCore() {
    /**********************************
   ** HIDE / SHOW SOME LAYERS NEEDED
   ***********************************/
    try {
        guideLayer.visible = false;
    }
    catch (e) {
        alert("Issue with layer hiding the Guidelines layer (do not change name from exactly Guidelines).", e.message);
    }
    /*****************************
    create export folder if needed
    ******************************/
    try {
        // Core folder
        new Folder("".concat(sourceDoc.path, "/").concat(coreName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(jpgName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(pngName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(svgName)).create();
        // Expressive folder
        new Folder("".concat(sourceDoc.path, "/").concat(expressiveName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(jpgName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(pngName)).create();
        new Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(svgName)).create();
    }
    catch (e) {
        alert("Issues with creating setup folders.", e.message);
    }
    /******************
    set up artboards
    ******************/
    //if there are two artboards at 256x256, create the new third masthead artboard
    if (sourceDoc.artboards.length == 2 &&
        sourceDoc.artboards[0].artboardRect[2] -
            sourceDoc.artboards[0].artboardRect[0] ==
            256 &&
        sourceDoc.artboards[0].artboardRect[1] -
            sourceDoc.artboards[0].artboardRect[3] ==
            256) {
        // alert("More than 2 artboards detected!");
        var firstRect = sourceDoc.artboards[0].artboardRect;
        sourceDoc.artboards.add(
        // CSTasks.newRect(firstRect[1] * 2.5 + gutter, firstRect[2], 2400, 256)
        // CSTasks.newRect(firstRect[1] * 0.5, firstRect[2], 2400, 256)
        CSTasks.newRect(firstRect[1], firstRect[2] + 128, 2400, 256));
    }
    //if the masthead artboard is present, check if rebuilding or just exporting
    else if (sourceDoc.artboards.length == 3 &&
        sourceDoc.artboards[1].artboardRect[1] -
            sourceDoc.artboards[1].artboardRect[3] ==
            256) {
        rebuild = confirm("It looks like your artwork already exists. This script will rebuild the masthead and export various EPS and PNG versions. Do you want to proceed?");
        if (rebuild)
            CSTasks.clearArtboard(sourceDoc, 1);
        else
            return;
    }
    //otherwise abort
    else {
        alert("Please try again with 2 artboards that are 256x256px.");
        return;
    }
    //select the contents on artboard 0
    var sel = CSTasks.selectContentsOnArtboard(sourceDoc, 0);
    // make sure all colors are RGB, equivalent of Edit > Colors > Convert to RGB
    app.executeMenuCommand('Colors9');
    if (sel.length == 0) {
        //if nothing is in the artboard
        alert("Please try again with artwork on the main 256x256 artboard.");
        return;
    }
    var colors = CSTasks.initializeColors(RGBColorElements, CMYKColorElements); //initialize the colors from the brand palette
    var iconGroup = CSTasks.createGroup(sourceDoc, sel); //group the selection (easier to work with)
    var iconOffset = CSTasks.getOffset(iconGroup.position, CSTasks.getArtboardCorner(sourceDoc.artboards[0]));
    /********************************
    Create new artboard with masthead
    *********************************/
    //place icon on masthead
    /*@ts-ignore*/
    var mast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
    var mastPos = [
        sourceDoc.artboards[2].artboardRect[0] + iconOffset[0],
        sourceDoc.artboards[2].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(mast, mastPos);
    //request a name for the icon, and place that as text on the masthead artboard
    var appName = prompt("What name do you want to put in the masthead?");
    var textRef = sourceDoc.textFrames.add();
    textRef.contents = appName;
    textRef.textRange.characterAttributes.size = 178;
    CSTasks.setFont(textRef, desiredFont);
    //vertically align the baseline to be 64 px above the botom of the artboard
    var bottomEdge = sourceDoc.artboards[2].artboardRect[3] +
        0.25 * sourceDoc.artboards[0].artboardRect[2] -
        sourceDoc.artboards[0].artboardRect[0]; //64px (0.25*256px) above the bottom edge of the artboard
    var vOffset = CSTasks.getOffset(textRef.anchor, [0, bottomEdge]);
    textRef.translate(0, -vOffset[1]);
    //create an outline of the text
    var textGroup = textRef.createOutline();
    //horizontally align the left edge of the text to be 96px to the right of the edge
    var rightEdge = mast.position[0] +
        mast.width +
        0.375 * sourceDoc.artboards[0].artboardRect[2] -
        sourceDoc.artboards[0].artboardRect[0]; //96px (0.375*256px) right of the icon
    var hOffset = CSTasks.getOffset(textGroup.position, [rightEdge, 0]);
    textGroup.translate(-hOffset[0], 0);
    //resize the artboard to be only a little wider than the text
    var leftMargin = mast.position[0] - sourceDoc.artboards[2].artboardRect[0];
    var newWidth = textGroup.position[0] +
        textGroup.width -
        sourceDoc.artboards[2].artboardRect[0] +
        leftMargin;
    var resizedRect = CSTasks.newRect(sourceDoc.artboards[2].artboardRect[0], -sourceDoc.artboards[2].artboardRect[1], newWidth, 256);
    sourceDoc.artboards[2].artboardRect = resizedRect;
    //get the text offset for exporting
    var mastTextOffset = CSTasks.getOffset(textGroup.position, CSTasks.getArtboardCorner(sourceDoc.artboards[2]));
    /*********************************************************************
    RGB export (EPS, PNGs at multiple sizes, inactive EPS and inverse EPS)
    **********************************************************************/
    var rgbDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 0, DocumentColorSpace.RGB);
    rgbDoc.swatches.removeAll();
    var rgbGroup = iconGroup.duplicate(rgbDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var rgbLoc = [
        rgbDoc.artboards[0].artboardRect[0] + iconOffset[0],
        rgbDoc.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(rgbGroup, rgbLoc);
    CSTasks.ungroupOnce(rgbGroup);
    //save a master PNG
    var masterStartWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_7 = 0; i_7 < exportSizes.length; i_7++) {
        var filename_1 = "/".concat(iconFilename, ".png");
        var destFile_1 = new File(Folder("".concat(sourceDoc.path)) + filename_1);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_1, masterStartWidth, exportSizes[2]);
    }
    //save a master SVG 
    var svgMasterCoreStartWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_8 = 0; i_8 < exportSizes.length; i_8++) {
        var filename_2 = "/".concat(iconFilename, ".svg");
        var destFile_2 = new File(Folder("".concat(sourceDoc.path)) + filename_2);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_2, svgMasterCoreStartWidth, exportSizes[2]);
    }
    //save all sizes of PNG into the export folder
    // let startWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    // non transparent png exports
    // let startWidthonFFF =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${onWhiteName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile, startWidthonFFF, exportSizes[i]);
    // }
    //save all sizes of SVG into the export folder
    // let svgCoreStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgCoreStartWidth, exportSizes[i]);
    // }
    //save all sizes of JPEG into the export folder
    // let jpegStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.jpg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${jpgName}`) + filename);
    //    CSTasks.scaleAndExportJPEG(rgbDoc, destFile, jpegStartWidth, exportSizes[i]);
    // }
    //save EPS into the export folder
    var filename = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, ".eps");
    var destFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + filename);
    var rgbSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    rgbSaveOpts.cmykPostScript = false;
    rgbDoc.saveAs(destFile, rgbSaveOpts);
    //index the RGB colors for conversion to CMYK. An inelegant location.
    var colorIndex = CSTasks.indexRGBColors(rgbDoc.pathItems, colors);
    //convert violet to white and save as EPS
    CSTasks.convertColorRGB(rgbDoc.pathItems, colors[violetIndex][0], colors[whiteIndex][0]);
    var inverseFilename = "/".concat(iconFilename, "_").concat(inverseName, "_").concat(rgbName, ".eps");
    var inverseFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + inverseFilename);
    rgbDoc.saveAs(inverseFile, rgbSaveOpts);
    //save inverse file in all the PNG sizes
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${inverseName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //save inverse file in all the SVG sizes
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${inverseName}_${rgbName}_${exportSizes[i]}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //convert to inactive color (WTW Icon grey at 100% opacity) and save as EPS
    CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 100);
    var inactiveFilename = "/".concat(iconFilename, "_").concat(inactiveName, "_").concat(rgbName, ".eps");
    var inactiveFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + inactiveFilename);
    rgbDoc.saveAs(inactiveFile, rgbSaveOpts);
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${inactiveName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${inactiveName}_${rgbName}_${exportSizes[i]}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //close and clean up
    rgbDoc.close(SaveOptions.DONOTSAVECHANGES);
    rgbDoc = null;
    /****************
    CMYK export (EPS) (Inverse?)
    ****************/
    //open a new document with CMYK colorspace, and duplicate the icon to the new document
    var cmykDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 0, DocumentColorSpace.CMYK);
    cmykDoc.swatches.removeAll();
    //need to reverse the order of copying the group to get the right color ordering
    var cmykGroup = iconGroup.duplicate(cmykDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATBEGINNING);
    var cmykLoc = [
        cmykDoc.artboards[0].artboardRect[0] + iconOffset[0],
        cmykDoc.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(cmykGroup, cmykLoc);
    CSTasks.ungroupOnce(cmykGroup);
    CSTasks.convertToCMYK(cmykDoc, cmykDoc.pathItems, colors, colorIndex);
    //save EPS into the export folder
    var cmykFilename = "/".concat(iconFilename, "_").concat(coreName, "_").concat(cmykName, ".eps");
    var cmykDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + cmykFilename);
    var cmykSaveOpts = new EPSSaveOptions();
    cmykDoc.saveAs(cmykDestFile, cmykSaveOpts);
    //convert violet to white and save as EPS
    CSTasks.convertColorCMYK(cmykDoc.pathItems, colors[violetIndex][1], colors[whiteIndex][1]);
    var cmykInverseFilename = "/".concat(iconFilename, "_").concat(inverseName, "_").concat(cmykName, ".eps");
    var cmykInverseFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + cmykInverseFilename);
    cmykDoc.saveAs(cmykInverseFile, rgbSaveOpts);
    //close and clean up
    cmykDoc.close(SaveOptions.DONOTSAVECHANGES);
    cmykDoc = null;
    /********************
    Masthead export core CMYK (EPS)
    ********************/
    //open a new doc and copy and position the icon and the masthead text
    var mastCMYKDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 1, DocumentColorSpace.CMYK);
    mastCMYKDoc.swatches.removeAll();
    var mastCMYKGroup = iconGroup.duplicate(mastCMYKDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var mastCMYKLoc = [
        mastCMYKDoc.artboards[0].artboardRect[0] + iconOffset[0],
        mastCMYKDoc.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(mastCMYKGroup, mastCMYKLoc);
    CSTasks.ungroupOnce(mastCMYKGroup);
    var mastCMYKText = textGroup.duplicate(mastCMYKDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var mastCMYKTextLoc = [
        mastCMYKDoc.artboards[0].artboardRect[0] + mastTextOffset[0],
        mastCMYKDoc.artboards[0].artboardRect[1] + mastTextOffset[1],
    ];
    CSTasks.translateObjectTo(mastCMYKText, mastCMYKTextLoc);
    //save CMYK EPS into the export folder
    var mastCMYKFilename = "/".concat(iconFilename, "_").concat(mastheadName, "_").concat(cmykName, ".eps");
    var mastCMYKDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + mastCMYKFilename);
    var mastCMYKSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    mastCMYKSaveOpts.cmykPostScript = false;
    mastCMYKDoc.saveAs(mastCMYKDestFile, mastCMYKSaveOpts);
    //close and clean up
    mastCMYKDoc.close(SaveOptions.DONOTSAVECHANGES);
    mastCMYKDoc = null;
    /********************
    Masthead export core RGB (EPS)
    ********************/
    //open a new doc and copy and position the icon and the masthead text
    var mastDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 1, DocumentColorSpace.RGB);
    mastDoc.swatches.removeAll();
    var mastGroup = iconGroup.duplicate(mastDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var mastLoc = [
        mastDoc.artboards[0].artboardRect[0] + iconOffset[0],
        mastDoc.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(mastGroup, mastLoc);
    CSTasks.ungroupOnce(mastGroup);
    var mastText = textGroup.duplicate(mastDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var mastTextLoc = [
        mastDoc.artboards[0].artboardRect[0] + mastTextOffset[0],
        mastDoc.artboards[0].artboardRect[1] + mastTextOffset[1],
    ];
    CSTasks.translateObjectTo(mastText, mastTextLoc);
    //save RGB EPS into the export folder
    var mastFilename = "/".concat(iconFilename, "_").concat(mastheadName, "_").concat(rgbName, ".eps");
    var mastDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + mastFilename);
    var mastSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    mastSaveOpts.cmykPostScript = false;
    mastDoc.saveAs(mastDestFile, mastSaveOpts);
    //close and clean up
    mastDoc.close(SaveOptions.DONOTSAVECHANGES);
    mastDoc = null;
    /************
    Final cleanup
    ************/
    CSTasks.ungroupOnce(iconGroup);
    CSTasks.ungroupOnce(mast);
    sourceDoc.selection = null;
}
mainCore();
/****************
 * Expressive
 ***************/
function mainExpressive() {
    /******************
    Set up purple artboard 3
    ******************/
    //if there are three artboards, create the new third masthead artboard
    if (sourceDoc.artboards.length == 3 &&
        sourceDoc.artboards[1].artboardRect[2] -
            sourceDoc.artboards[1].artboardRect[0] ==
            256 &&
        sourceDoc.artboards[1].artboardRect[1] -
            sourceDoc.artboards[1].artboardRect[3] ==
            256) {
        // IF there are already  3 artboards. Add a 4th one.
        var firstRect = sourceDoc.artboards[1].artboardRect;
        sourceDoc.artboards.add(
        // this fires but then gets replaced further down
        CSTasks.newRect(firstRect[1], firstRect[2] + 128, 1024, 512));
    }
    //if the masthead artboard is present, check if rebuilding or just exporting
    else if (sourceDoc.artboards.length == 3 &&
        sourceDoc.artboards[1].artboardRect[1] -
            sourceDoc.artboards[1].artboardRect[3] ==
            512) {
        rebuild = confirm("It looks like your artwork already exists. This script will rebuild the masthead and export various EPS and PNG versions. Do you want to proceed?");
        if (rebuild)
            CSTasks.clearArtboard(sourceDoc, 3);
        else
            return;
    }
    //otherwise abort
    else {
        alert("Please try again with 2 artboards that are 256x256px. Had trouble building artboard 3 (artb. 2 in js)");
        return;
    }
    /* try {
       sourceDoc.artboards.setActiveArtboardIndex(3);//change which artboard you want to crop
       sourceDoc.artboards[3].artboardRect = new Rect(0, 0, 1024, 512);
    } catch (error) {
       alert(error);
    } */
    //select the contents on artboard 1
    var sel = CSTasks.selectContentsOnArtboard(sourceDoc, 1);
    // make sure all colors are RGB, equivalent of Edit > Colors > Convert to RGB
    app.executeMenuCommand('Colors9');
    if (sel.length == 0) {
        //if nothing is in the artboard
        alert("Please try again with artwork on the main second 256x256 artboard.");
        return;
    }
    var colors = CSTasks.initializeColors(RGBColorElements, CMYKColorElements); //initialize the colors from the brand palette
    var iconGroup = CSTasks.createGroup(sourceDoc, sel); //group the selection (easier to work with)
    var iconOffset = CSTasks.getOffset(iconGroup.position, CSTasks.getArtboardCorner(sourceDoc.artboards[1]));
    /********************************
    Create new artboard with masthead
    *********************************/
    //place icon on masthead
    /*@ts-ignore*/
    var mast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
    var mastPos = [
        sourceDoc.artboards[3].artboardRect[0] + iconOffset[0] * 24,
        sourceDoc.artboards[3].artboardRect[1] + iconOffset[1] * 2.3,
    ];
    CSTasks.translateObjectTo(mast, mastPos);
    mast.width = 460;
    mast.height = 460;
    // new purple bg
    // Add new layer above Guidelines and fill white
    var myMainArtworkLayer = sourceDoc.layers.getByName('Art');
    var myMainPurpleBgLayer = sourceDoc.layers.add();
    myMainPurpleBgLayer.name = "Main_Purple_BG_layer";
    var GetMyMainPurpleBgLayer = sourceDoc.layers.getByName('Main_Purple_BG_layer');
    // mastDoc.activeLayer = GetMyMainPurpleBgLayer;
    // mastDoc.activeLayer.hasSelectedArtwork = true;
    var mainRect = GetMyMainPurpleBgLayer.pathItems.rectangle(-784, 0, 1024, 512);
    var setMainVioletBgColor = new RGBColor();
    setMainVioletBgColor.red = 72;
    setMainVioletBgColor.green = 8;
    setMainVioletBgColor.blue = 111;
    mainRect.filled = true;
    mainRect.fillColor = setMainVioletBgColor;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayer.move(myMainArtworkLayer, ElementPlacement.PLACEATEND);
    var rectRef = sourceDoc.pathItems.rectangle(-850, -800, 400, 300);
    var setTextBoxBgColor = new RGBColor();
    setTextBoxBgColor.red = 141;
    setTextBoxBgColor.green = 141;
    setTextBoxBgColor.blue = 141;
    rectRef.filled = true;
    rectRef.fillColor = setTextBoxBgColor;
    // svg wtw logo for new purple masthead
    var imagePlacedItem = myMainArtworkLayer.placedItems.add();
    var svgFile = File("".concat(sourceDoc.path, "/../images/wtw_logo.ai"));
    imagePlacedItem.file = svgFile;
    imagePlacedItem.top = -1188;
    imagePlacedItem.left = 62;
    /*@ts-ignore*/
    // svgFile.embed();
    //request a name for the icon, and place that as text on the masthead artboard
    var appName = prompt("What name do you want to put in second the masthead?");
    var textRef = sourceDoc.textFrames.add();
    //use the areaText method to create the text frame
    /*@ts-ignore*/
    textRef = sourceDoc.textFrames.areaText(rectRef);
    textRef.contents = appName;
    textRef.textRange.characterAttributes.size = 62;
    // textRef.textRange.characterAttributes.horizontalScale = 2299;
    textRef.textRange.characterAttributes.fillColor = colors[whiteIndex][0];
    CSTasks.setFont(textRef, desiredFont);
    //vertically align the baseline to be 64 px above the botom of the artboard
    // let bottomEdge =
    //    sourceDoc.artboards[3].artboardRect[3] +
    //    1.58 * sourceDoc.artboards[0].artboardRect[2] -
    //    sourceDoc.artboards[0].artboardRect[0]; //64px (0.25*256px) above the bottom edge of the artboard
    // let vOffset = CSTasks.getOffset(textRef.anchor, [0, bottomEdge]);
    // textRef.translate(0, -vOffset[1]);
    //create an outline of the text
    var textGroup = textRef.createOutline();
    //horizontally align the left edge of the text to be 96px to the right of the edge
    var rightEdge = 64;
    var hOffset = CSTasks.getOffset(textGroup.position, [rightEdge, 0]);
    textGroup.translate(-hOffset[0], 0);
    //resize the artboard to be only a little wider than the text
    var leftMargin = mast.position[0] - sourceDoc.artboards[3].artboardRect[0];
    var newWidth = textGroup.position[0] +
        textGroup.width -
        sourceDoc.artboards[3].artboardRect[0] +
        leftMargin;
    var resizedRect = CSTasks.newRect(sourceDoc.artboards[3].artboardRect[0], -sourceDoc.artboards[3].artboardRect[1], 1024, 512);
    sourceDoc.artboards[3].artboardRect = resizedRect;
    /******************
    Set up purple artboard 4 in file
    ******************/
    //if there are 3 artboards, create the new fourth masthead artboard
    // creat last artboard in file
    var firstRect2 = sourceDoc.artboards[1].artboardRect;
    sourceDoc.artboards.add(
    // this fires but then gets replaced further down
    CSTasks.newRect(firstRect2[1], firstRect2[2] + 772, 800, 400));
    /* try {
       sourceDoc.artboards.setActiveArtboardIndex(3);//change which artboard you want to crop
       sourceDoc.artboards[3].artboardRect = new Rect(0, 0, 1024, 512);
    } catch (error) {
       alert(error);
    } */
    //select the contents on artboard 1
    var selBanner2 = CSTasks.selectContentsOnArtboard(sourceDoc, 1);
    // make sure all colors are RGB, equivalent of Edit > Colors > Convert to RGB
    app.executeMenuCommand('Colors9');
    if (selBanner2.length == 0) {
        //if nothing is in the artboard
        alert("Please try again with artwork on the main second 256x256 artboard.");
        return;
    }
    /********************************
    Create new artboard with masthead
    *********************************/
    //place icon on masthead
    /*@ts-ignore*/
    var mast2 = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
    var mastPos2 = [
        sourceDoc.artboards[4].artboardRect[0] + iconOffset[0] * 18.5,
        sourceDoc.artboards[4].artboardRect[1] + iconOffset[1] * 1.6,
    ];
    CSTasks.translateObjectTo(mast2, mastPos2);
    mast2.width = 360;
    mast2.height = 360;
    // new purple bg
    // Add new layer above Guidelines and fill white
    var myMainArtworkLayer2 = sourceDoc.layers.getByName('Art');
    var myMainPurpleBgLayer2 = sourceDoc.layers.add();
    myMainPurpleBgLayer2.name = "Main_Purple_BG_layer_two";
    var GetMyMainPurpleBgLayer2 = sourceDoc.layers.getByName('Main_Purple_BG_layer_two');
    // mastDoc.activeLayer = GetMyMainPurpleBgLayer2;
    // mastDoc.activeLayer.hasSelectedArtwork = true;
    var mainRect2 = GetMyMainPurpleBgLayer2.pathItems.rectangle(-1428, 0, 800, 400);
    var setMainVioletBgColor2 = new RGBColor();
    setMainVioletBgColor2.red = 72;
    setMainVioletBgColor2.green = 8;
    setMainVioletBgColor2.blue = 111;
    mainRect2.filled = true;
    mainRect2.fillColor = setMainVioletBgColor2;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayer2.move(myMainArtworkLayer2, ElementPlacement.PLACEATEND);
    /*@ts-ignore*/
    // svgFile.embed();
    var resizedRect2 = CSTasks.newRect(sourceDoc.artboards[4].artboardRect[0], -sourceDoc.artboards[4].artboardRect[1], 800, 400);
    sourceDoc.artboards[4].artboardRect = resizedRect2;
    /*********************************************************************
    RGB export (EPS, PNGs at multiple sizes, inactive EPS and inverse EPS)
    **********************************************************************/
    var rgbDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 1, DocumentColorSpace.RGB);
    rgbDoc.swatches.removeAll();
    var rgbGroup = iconGroup.duplicate(rgbDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var rgbLoc = [
        rgbDoc.artboards[0].artboardRect[0] + iconOffset[0],
        rgbDoc.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(rgbGroup, rgbLoc);
    CSTasks.ungroupOnce(rgbGroup);
    //save a master PNG
    // let masterStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, masterStartWidth, exportSizes[2]);
    // }
    //save a master SVG 
    // let svgMasterCoreStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgMasterCoreStartWidth, exportSizes[2]);
    // }
    //save all sizes of PNG into the export folder
    // let startWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    // non transparent png exports
    // let startWidthonFFF =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${onWhiteName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile, startWidthonFFF, exportSizes[i]);
    // }
    //save all sizes of SVG into the export folder
    // let svgCoreStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[1];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${exportSizes[i]}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgCoreStartWidth, exportSizes[i]);
    // }
    //save all sizes of JPEG into the export folder
    // let jpegStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[1];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${exportSizes[i]}.jpg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${jpgName}`) + filename);
    //    CSTasks.scaleAndExportJPEG(rgbDoc, destFile, jpegStartWidth, exportSizes[i]);
    // }
    //save EPS into the export folder
    var filename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, ".eps");
    var destFile = new File(Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)) + filename);
    var rgbSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    rgbSaveOpts.cmykPostScript = false;
    rgbDoc.saveAs(destFile, rgbSaveOpts);
    //index the RGB colors for conversion to CMYK. An inelegant location.
    var colorIndex = CSTasks.indexRGBColors(rgbDoc.pathItems, colors);
    //convert violet to white and save as EPS
    CSTasks.convertColorRGB(rgbDoc.pathItems, colors[violetIndex][0], colors[whiteIndex][0]);
    var inverseFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inverseName, "_").concat(rgbName, ".eps");
    var inverseFile = new File(Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)) + inverseFilename);
    rgbDoc.saveAs(inverseFile, rgbSaveOpts);
    //save inverse file in all the PNG sizes
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${inverseName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //save inverse file in all the SVG sizes
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${inverseName}_${rgbName}_${exportSizes[i]}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //convert to inactive color (WTW Icon grey at 100% opacity) and save as EPS
    CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 100);
    var inactiveFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inactiveName, "_").concat(rgbName, ".eps");
    var inactiveFile = new File(Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)) + inactiveFilename);
    rgbDoc.saveAs(inactiveFile, rgbSaveOpts);
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${inactiveName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${inactiveName}_${rgbName}_${exportSizes[i]}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //close and clean up
    rgbDoc.close(SaveOptions.DONOTSAVECHANGES);
    rgbDoc = null;
    /****************
    CMYK export (EPS)
    ****************/
    //open a new document with CMYK colorspace, and duplicate the icon to the new document
    var cmykDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 0, DocumentColorSpace.CMYK);
    cmykDoc.swatches.removeAll();
    //need to reverse the order of copying the group to get the right color ordering
    var cmykGroup = iconGroup.duplicate(cmykDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATBEGINNING);
    var cmykLoc = [
        cmykDoc.artboards[0].artboardRect[0] + iconOffset[0],
        cmykDoc.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(cmykGroup, cmykLoc);
    CSTasks.ungroupOnce(cmykGroup);
    CSTasks.convertToCMYK(cmykDoc, cmykDoc.pathItems, colors, colorIndex);
    //save EPS into the export folder
    var cmykFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(cmykName, ".eps");
    var cmykDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)) + cmykFilename);
    var cmykSaveOpts = new EPSSaveOptions();
    cmykDoc.saveAs(cmykDestFile, cmykSaveOpts);
    //convert violet to white and save as EPS
    CSTasks.convertColorCMYK(cmykDoc.pathItems, colors[violetIndex][1], colors[whiteIndex][1]);
    var cmykInverseFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inverseName, "_").concat(cmykName, ".eps");
    var cmykInverseFile = new File(Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)) + cmykInverseFilename);
    cmykDoc.saveAs(cmykInverseFile, rgbSaveOpts);
    //close and clean up
    cmykDoc.close(SaveOptions.DONOTSAVECHANGES);
    cmykDoc = null;
    /********************
    Purple Masthead with text export export
    ********************/
    //open a new doc and copy and position the icon and the masthead text
    // duplication did not work as expected here. I have used a less elegant solution whereby I recreated the purple banner instead of copying it.
    var mastDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 3, DocumentColorSpace.RGB);
    mastDoc.swatches.removeAll();
    var mastGroup = iconGroup.duplicate(mastDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    // new icon width in rebrand
    mastGroup.width = 460;
    mastGroup.height = 460;
    // new icon position
    var mastLoc = [
        mastDoc.artboards[0].artboardRect[0] + 576,
        mastDoc.artboards[0].artboardRect[1] - 62,
    ];
    CSTasks.translateObjectTo(mastGroup, mastLoc);
    CSTasks.ungroupOnce(mastGroup);
    var mastText = textGroup.duplicate(mastDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    // text position
    var mastTextLoc = [
        mastDoc.artboards[0].artboardRect[0] + 62,
        mastDoc.artboards[0].artboardRect[1] - 62,
    ];
    CSTasks.translateObjectTo(mastText, mastTextLoc);
    // add new style purple banner elements
    var myMainArtworkLayerMastDoc = mastDoc.layers.getByName('Layer 1');
    var myMainPurpleBgLayerMastDoc = mastDoc.layers.add();
    myMainPurpleBgLayerMastDoc.name = "Main_Purple_BG_layer";
    var GetMyMainPurpleBgLayerMastDoc = mastDoc.layers.getByName('Main_Purple_BG_layer');
    // mastDoc.activeLayer = GetMyMainPurpleBgLayerMastDoc;
    // mastDoc.activeLayer.hasSelectedArtwork = true;
    var mainRectMastDoc = GetMyMainPurpleBgLayerMastDoc.pathItems.rectangle(-781, 0, 1024, 512);
    var setMainVioletBgColorMastDoc = new RGBColor();
    setMainVioletBgColorMastDoc.red = 72;
    setMainVioletBgColorMastDoc.green = 8;
    setMainVioletBgColorMastDoc.blue = 111;
    mainRectMastDoc.filled = true;
    mainRectMastDoc.fillColor = setMainVioletBgColorMastDoc;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayerMastDoc.move(myMainArtworkLayerMastDoc, ElementPlacement.PLACEATEND);
    // svg wtw logo for new purple masthead
    var imagePlacedItemMastDoc = myMainArtworkLayerMastDoc.placedItems.add();
    var svgFileMastDoc = File("".concat(sourceDoc.path, "/../images/wtw_logo.ai"));
    imagePlacedItemMastDoc.file = svgFileMastDoc;
    imagePlacedItemMastDoc.top = -1181;
    imagePlacedItemMastDoc.left = 62;
    // we need to make artboard clipping mask here for the artboard to crop expressive icons correctly.
    var myCroppingLayerMastDoc = mastDoc.layers.add();
    myCroppingLayerMastDoc.name = "crop";
    var GetMyCroppingLayerMastDoc = mastDoc.layers.getByName('crop');
    mastDoc.activeLayer = GetMyCroppingLayerMastDoc;
    mastDoc.activeLayer.hasSelectedArtwork = true;
    // insert clipping rect here
    var mainClipRectMastDoc = GetMyCroppingLayerMastDoc.pathItems.rectangle(-781, 0, 1024, 512);
    var setClipBgColorMastDoc = new RGBColor();
    setClipBgColorMastDoc.red = 0;
    setClipBgColorMastDoc.green = 255;
    setClipBgColorMastDoc.blue = 255;
    mainClipRectMastDoc.filled = true;
    mainClipRectMastDoc.fillColor = setClipBgColorMastDoc;
    // select all for clipping here
    sourceDoc.selectObjectsOnActiveArtboard();
    // clip!
    app.executeMenuCommand('makeMask');
    //save a banner PNG
    var masterStartWidthMastDoc = 1024;
    for (var i_9 = 0; i_9 < exportSizes.length; i_9++) {
        var filename_3 = "/".concat(iconFilename, "_banner.png");
        var destFile_3 = new File(Folder("".concat(sourceDoc.path)) + filename_3);
        CSTasks.scaleAndExportPNG(mastDoc, destFile_3, masterStartWidthMastDoc, exportSizes[0]);
    }
    //save RGB EPS into the export folder
    var mastFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(mastheadName, "_").concat(rgbName, ".eps");
    var mastDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)) + mastFilename);
    var mastSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    mastSaveOpts.cmykPostScript = false;
    /*@ts-ignore*/
    mastSaveOpts.embedLinkedFiles = true;
    mastDoc.saveAs(mastDestFile, mastSaveOpts);
    //close and clean up
    mastDoc.close(SaveOptions.DONOTSAVECHANGES);
    mastDoc = null;
    /********************
      Purple Masthead with text export export
      ********************/
    //open a new doc and copy and position the icon and the masthead text
    // duplication did not work as expected here. I have used a less elegant solution whereby I recreated the purple banner instead of copying it.
    var mastDoc2 = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 3, DocumentColorSpace.RGB);
    mastDoc2.swatches.removeAll();
    var mastGroup2 = iconGroup.duplicate(mastDoc2.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    // new icon width in rebrand
    mastGroup2.width = 460;
    mastGroup2.height = 460;
    // new icon position
    var mastLoc2 = [
        mastDoc2.artboards[0].artboardRect[0] + 576,
        mastDoc2.artboards[0].artboardRect[1] - 62,
    ];
    CSTasks.translateObjectTo(mastGroup2, mastLoc2);
    CSTasks.ungroupOnce(mastGroup2);
    var mastText2 = textGroup.duplicate(mastDoc2.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    // text position
    var mastTextLoc2 = [
        mastDoc2.artboards[0].artboardRect[0] + 62,
        mastDoc2.artboards[0].artboardRect[1] - 62,
    ];
    CSTasks.translateObjectTo(mastText2, mastTextLoc2);
    // add new style purple banner elements
    var myMainArtworkLayerMastDoc2 = mastDoc2.layers.getByName('Layer 1');
    var myMainPurpleBgLayerMastDoc2 = mastDoc2.layers.add();
    myMainPurpleBgLayerMastDoc2.name = "Main_Purple_BG_layer";
    var GetMyMainPurpleBgLayerMastDoc2 = mastDoc2.layers.getByName('Main_Purple_BG_layer');
    // mastDoc.activeLayer = GetMyMainPurpleBgLayerMastDoc2;
    // mastDoc.activeLayer.hasSelectedArtwork = true;
    var mainRectMastDoc2 = GetMyMainPurpleBgLayerMastDoc2.pathItems.rectangle(-781, 0, 1024, 512);
    var setMainVioletBgColorMastDoc1 = new RGBColor();
    setMainVioletBgColorMastDoc1.red = 72;
    setMainVioletBgColorMastDoc1.green = 8;
    setMainVioletBgColorMastDoc1.blue = 111;
    mainRectMastDoc2.filled = true;
    mainRectMastDoc2.fillColor = setMainVioletBgColorMastDoc1;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayerMastDoc2.move(myMainArtworkLayerMastDoc2, ElementPlacement.PLACEATEND);
    // svg wtw logo for new purple masthead
    var imagePlacedItemMastDoc2 = myMainArtworkLayerMastDoc2.placedItems.add();
    var svgFileMastDoc2 = File("".concat(sourceDoc.path, "/../images/wtw_logo.ai"));
    imagePlacedItemMastDoc2.file = svgFileMastDoc2;
    imagePlacedItemMastDoc2.top = -1181;
    imagePlacedItemMastDoc2.left = 62;
    // we need to make artboard clipping mask here for the artboard to crop expressive icons correctly.
    var myCroppingLayerMastDoc2 = mastDoc2.layers.add();
    myCroppingLayerMastDoc2.name = "crop";
    var GetMyCroppingLayerMastDoc2 = mastDoc2.layers.getByName('crop');
    mastDoc2.activeLayer = GetMyCroppingLayerMastDoc2;
    mastDoc2.activeLayer.hasSelectedArtwork = true;
    // insert clipping rect here
    var mainClipRectMastDoc2 = GetMyCroppingLayerMastDoc2.pathItems.rectangle(-781, 0, 1024, 512);
    var setClipBgColorMastDoc2 = new RGBColor();
    setClipBgColorMastDoc2.red = 0;
    setClipBgColorMastDoc2.green = 255;
    setClipBgColorMastDoc2.blue = 255;
    mainClipRectMastDoc2.filled = true;
    mainClipRectMastDoc2.fillColor = setClipBgColorMastDoc2;
    // select all for clipping here
    sourceDoc.selectObjectsOnActiveArtboard();
    // clip!
    app.executeMenuCommand('makeMask');
    //save a banner PNG
    var masterStartWidthMastDoc2 = 1024;
    for (var i_10 = 0; i_10 < exportSizes.length; i_10++) {
        var filename_4 = "/".concat(iconFilename, "______LASTBANNER.png");
        var destFile_4 = new File(Folder("".concat(sourceDoc.path)) + filename_4);
        CSTasks.scaleAndExportPNG(mastDoc2, destFile_4, masterStartWidthMastDoc2, exportSizes[0]);
    }
    //save RGB EPS into the export folder
    var mastFilename2 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(mastheadName, "_").concat(rgbName, "______LASTBANNER.eps");
    var mastDestFile2 = new File(Folder("".concat(sourceDoc.path, "/").concat(expressiveName, "/").concat(epsName)) + mastFilename2);
    var mastSaveOpts2 = new EPSSaveOptions();
    /*@ts-ignore*/
    mastSaveOpts2.cmykPostScript = false;
    /*@ts-ignore*/
    mastSaveOpts2.embedLinkedFiles = true;
    mastDoc2.saveAs(mastDestFile2, mastSaveOpts2);
    //close and clean up
    // mastDoc2.close(SaveOptions.DONOTSAVECHANGES);
    // mastDoc2 = null;
    /************
    Final cleanup
    ************/
    // CSTasks.ungroupOnce(iconGroup);
    // CSTasks.ungroupOnce(mast);
}
mainExpressive();
