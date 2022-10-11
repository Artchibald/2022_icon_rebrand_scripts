// #target Illustrator 
/************************************************
 * ** README https://github.com/Artchibald/2022_icon_rebrand_scripts
 
Script to automate creating variations and exporting files for WTW icons
Starting with an open AI file with a single icon on a single 256 x 256 artboard
– Creates a new artboard at 16x16
- Creates a new artboard at 24x24
- Creates a new artboard at 1400x128
(if these artboards already exist, optionally clears and rebuilds these artboards instead)

- Adds resized copies of the icon to the artboards
- Asks for the name of the icon and adds text to the lockup icon

- Creates exports of the icon:
- RGB EPS
- RGB inverse EPS
- RGB inactive EPS
- PNGs at 1024, 256, 128, 64, 48, 32
- RGB lockup
- CMYK EPS
- CMYK inverse EPS
************************************************/
alert(" \n\nThis script only works locally not on a server. \n\nDon't forget to change .txt to .js on the script. \n\nFULL README: https://github.com/Artchibald/2022_icon_rebrand_scripts   \n\nVideo set up tutorial available here: https://youtu.be/yxtrt7nkOjA. \n\nOpen your own.ai template or the provided ones in folders called test. \n\nGo to file > Scripts > Other Scripts > Import our new script. \n\nYou must have the folder called images in the parent folder, this is where wtw_logo.ai is saved so it can be imported into the big purple banner and exported as assets.Otherwise you will get an error that says error = svgFile. \n\nIllustrator says(not responding) on PC but it will respond, give Bill Gates some time XD!). \n\nIf you run the script again, you should probably delete the previous assets created.They get intermixed and overwritten. \n\nBoth artboard sizes must be exactly 256px x 256px. \n\nGuides must be on a layer called exactly 'Guidelines'. \n\nIcons must be on a layer called exactly 'Art'. \n\nMake sure all layers are unlocked to avoid bugs. \n\nExported assets will be saved where the.ai file is saved. \n\nPlease try to use underscore instead of spaces to avoid bugs in filenames. \n\nMake sure you are using the correct swatches / colours. \n\nIllustrator check advanced colour mode is correct: Edit > Assign profile > Must match sRGB IEC61966 - 2.1. \n\nSelect each individual color shape and under Window > Colours make sure each shape colour is set to rgb in tiny top right burger menu if bugs encountered. \n\nIf it does not save exports as intended, check the file permissions of where the.ai file is saved(right click folder > Properties > Visibility > Read and write access ? Also you can try apply permissions to sub folders too if you find that option) \n\nAny issues: archie ATsymbol archibaldbutler.com.");
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
var sourceDocName = sourceDoc.name.slice(0, -3);
// Lockups
var lockupName = "Lockup";
var lockup1 = "Lockup1";
var lockup2 = "Lockup2";
var eightByFour = "800x400";
var tenByFive = "1024x512";
// Colors
var rgbName = "RGB";
var cmykName = "CMYK";
var onWhiteName = "onFFF";
// type
var croppedName = "Cropped";
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
    //takes a document, destination file, starting width and desired width
    //scales the document proportionally to the desired width and exports as a NON TRANSPARENT PNG
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
    //scales the document proportionally to the desired width and exports as a JPG
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
        /*@ts-ignore*/
        options.qualitySetting = 100;
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
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName)).create();
    // Core folder
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(epsName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(jpgName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(pngName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(svgName)).create();
    // Expressive folder 
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(epsName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(jpgName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)).create();
    new Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)).create();
}
catch (e) {
    alert("Issues with creating setup folders. Check your file permission properties.", e.message);
}
/****************
 _________
\_   ___ \  ___________   ____
/    \  \/ /  _ \_  __ \_/ __ \
\     \___(  <_> )  | \/\  ___/
 \______  /\____/|__|    \___  >
        \/                   \/
 ***************/
/**********************************
 Core
***********************************/
function mainCore() {
    /******************
    set up artboards
    ******************/
    //if there are two artboards at 256x256, create the new third lockup artboard
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
    //if the lockup artboard is present, check if rebuilding or just exporting
    else if (sourceDoc.artboards.length == 3 &&
        sourceDoc.artboards[1].artboardRect[1] -
            sourceDoc.artboards[1].artboardRect[3] ==
            256) {
        rebuild = confirm("It looks like your artwork already exists. This script will rebuild the lockup and export various EPS and PNG versions. Do you want to proceed?");
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
    Create new artboard with lockup
    *********************************/
    //place icon on lockup
    /*@ts-ignore*/
    var mast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
    var mastPos = [
        sourceDoc.artboards[2].artboardRect[0] + iconOffset[0],
        sourceDoc.artboards[2].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(mast, mastPos);
    //request a name for the icon, and place that as text on the lockup artboard
    var appName = prompt("What name do you want to put in the lockup?");
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
        var filename_1 = "/".concat(iconFilename, "_").concat(coreName, ".png");
        var destFile_1 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName)) + filename_1);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_1, masterStartWidth, exportSizes[0]);
    }
    //save a master SVG 
    var svgMasterCoreStartWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_8 = 0; i_8 < exportSizes.length; i_8++) {
        var filename_2 = "/".concat(iconFilename, "_").concat(coreName, ".svg");
        var destFile_2 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName)) + filename_2);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_2, svgMasterCoreStartWidth, exportSizes[0]);
    }
    //save a master SVG in Core SVG
    var svgMasterCoreStartWidth2 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_9 = 0; i_9 < exportSizes.length; i_9++) {
        var filename_3 = "/".concat(iconFilename, "_").concat(coreName, ".svg");
        var destFile_3 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(svgName)) + filename_3);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_3, svgMasterCoreStartWidth2, exportSizes[0]);
    }
    //save all sizes of PNG into the export folder
    var startWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    // non transparent png exports
    // let startWidthonFFF =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${onWhiteName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile, startWidthonFFF, exportSizes[i]);
    // }
    //save all sizes of JPEG into the export folder
    // let jpegStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.jpg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${jpgName}`) + filename);
    //    CSTasks.scaleAndExportJPEG(rgbDoc, destFile, jpegStartWidth, exportSizes[i]);
    // }
    //save EPS into the export folder
    var filename = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, ".eps");
    var destFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(epsName)) + filename);
    var rgbSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    rgbSaveOpts.cmykPostScript = false;
    rgbDoc.saveAs(destFile, rgbSaveOpts);
    //index the RGB colors for conversion to CMYK. An inelegant location.
    var colorIndex = CSTasks.indexRGBColors(rgbDoc.pathItems, colors);
    //convert violet to white and save as EPS
    CSTasks.convertColorRGB(rgbDoc.pathItems, colors[violetIndex][0], colors[whiteIndex][0]);
    // let inverseFilename = `/${iconFilename}_${inverseName}_${rgbName}.eps`;
    // let inverseFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${epsName}`) + inverseFilename);
    // rgbDoc.saveAs(inverseFile, rgbSaveOpts);
    //save inverse file in all the PNG sizes
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${inverseName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //convert to inactive color (WTW Icon grey at 100% opacity) and save as EPS
    CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 100);
    // let inactiveFilename = `/${iconFilename}_${inactiveName}_${rgbName}.eps`;
    // let inactiveFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${epsName}`) + inactiveFilename);
    // rgbDoc.saveAs(inactiveFile, rgbSaveOpts);
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${inactiveName}_${rgbName}_${exportSizes[i]}.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
    // }
    //close and clean up
    rgbDoc.close(SaveOptions.DONOTSAVECHANGES);
    rgbDoc = null;
    /*********************************************************************
    Export an SVG on a white background
    **********************************************************************/
    var rgbDocOnFFF = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 0, DocumentColorSpace.RGB);
    rgbDocOnFFF.swatches.removeAll();
    var rgbGroupOnFFF = iconGroup.duplicate(rgbDocOnFFF.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATBEGINNING);
    var rgbLocOnFFF = [
        rgbDocOnFFF.artboards[0].artboardRect[0] + iconOffset[0],
        rgbDocOnFFF.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(rgbGroupOnFFF, rgbLocOnFFF);
    CSTasks.ungroupOnce(rgbGroupOnFFF);
    // Add a white square bg here to save svg on white
    var getArtLayer = rgbDocOnFFF.layers.getByName('Layer 1');
    var myMainPurpleBgLayerOnFFF = rgbDocOnFFF.layers.add();
    myMainPurpleBgLayerOnFFF.name = "Main_Purple_BG_layer";
    var GetMyMainPurpleBgLayerOnFFF = rgbDocOnFFF.layers.getByName('Main_Purple_BG_layer');
    var landingZoneSquareOnFFF = GetMyMainPurpleBgLayerOnFFF.pathItems.rectangle(0, 0, 256, 256);
    var setLandingZoneSquareOnFFFColor = new RGBColor();
    setLandingZoneSquareOnFFFColor.red = 255;
    setLandingZoneSquareOnFFFColor.green = 255;
    setLandingZoneSquareOnFFFColor.blue = 255;
    landingZoneSquareOnFFF.fillColor = setLandingZoneSquareOnFFFColor;
    landingZoneSquareOnFFF.name = "bgSquareForSvgOnFF";
    landingZoneSquareOnFFF.filled = true;
    landingZoneSquareOnFFF.stroked = false;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayerOnFFF.move(getArtLayer, ElementPlacement.PLACEATEND);
    //close and clean up
    rgbDocOnFFF.close(SaveOptions.DONOTSAVECHANGES);
    rgbDocOnFFF = null;
    /*********************************************************************
    RGB cropped export (JPG, PNGs at 16 and 24 sizes), squares, cropped to artwork
    **********************************************************************/
    var rgbDocCroppedVersion = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 0, DocumentColorSpace.RGB);
    rgbDocCroppedVersion.swatches.removeAll();
    var rgbGroupCropped = iconGroup.duplicate(rgbDocCroppedVersion.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var rgbLocCropped = [
        rgbDocCroppedVersion.artboards[0].artboardRect[0] + iconOffset[0],
        rgbDocCroppedVersion.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(rgbGroupCropped, rgbLocCropped);
    // remove padding here befor exporting
    function placeIconLockup1Correctly(rgbGroupCropped, maxSize) {
        var W = rgbGroupCropped.width, H = rgbGroupCropped.height, MW = maxSize.W, MH = maxSize.H, factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
        rgbGroupCropped.resize(factor, factor);
    }
    placeIconLockup1Correctly(rgbGroupCropped, { W: 256, H: 256 });
    CSTasks.ungroupOnce(rgbGroupCropped);
    // below we export croped only versions
    // // non transparent png exports
    var startWidthCroppedOnFFF = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    //24x24
    // let filenameCropped24PngOnFFF = `/${iconFilename}_${coreName}_${rgbName}_${onWhiteName}_${exportSizes[7]}.png`;
    // let destFileCropped24PngOnFFF = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${pngName}`) + filenameCropped24PngOnFFF);
    // CSTasks.scaleAndExportNonTransparentPNG(rgbDocCroppedVersion, destFileCropped24PngOnFFF, startWidthCroppedOnFFF, exportSizes[7]);
    //16x16
    // let filenameCropped16PngOnFFF = `/${iconFilename}_${coreName}_${rgbName}_${onWhiteName}_${exportSizes[8]}.png`;
    // let destFileCropped16PngOnFFF = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${coreName}/${pngName}`) + filenameCropped16PngOnFFF);
    // CSTasks.scaleAndExportNonTransparentPNG(rgbDocCroppedVersion, destFileCropped16PngOnFFF, startWidthCroppedOnFFF, exportSizes[8]);
    // save cropped 16 and 24 sizes of PNG into the export folder
    var startWidthCropped = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    //24x24
    var filenameCropped24Png = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[7], ".png");
    var destFileCropped24Png = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(pngName)) + filenameCropped24Png);
    CSTasks.scaleAndExportPNG(rgbDocCroppedVersion, destFileCropped24Png, startWidthCropped, exportSizes[7]);
    //16x16
    var filenameCropped16Png = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[8], ".png");
    var destFileCropped16Png = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(pngName)) + filenameCropped16Png);
    CSTasks.scaleAndExportPNG(rgbDocCroppedVersion, destFileCropped16Png, startWidthCropped, exportSizes[8]);
    // save cropped 16 and 24 sizes of JPEG into the export folder
    var jpegStartWidthCropped = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    //24x24
    var filenameCropped24Jpg = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[7], ".jpg");
    var destFileCropped24Jpg = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(jpgName)) + filenameCropped24Jpg);
    CSTasks.scaleAndExportJPEG(rgbDocCroppedVersion, destFileCropped24Jpg, jpegStartWidthCropped, exportSizes[7]);
    //16x16
    var filenameCropped16Jpg = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[8], ".jpg");
    var destFileCropped16Jpg = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(jpgName)) + filenameCropped16Jpg);
    CSTasks.scaleAndExportJPEG(rgbDocCroppedVersion, destFileCropped16Jpg, jpegStartWidthCropped, exportSizes[8]);
    // Save a cropped SVG 
    var svgMasterCoreStartWidthCroppedSvg = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    var filenameCroppedSvg = "/".concat(iconFilename, "_").concat(coreName, "_").concat(croppedName, ".svg");
    var destFileCroppedSvg = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(svgName)) + filenameCroppedSvg);
    CSTasks.scaleAndExportSVG(rgbDocCroppedVersion, destFileCroppedSvg, svgMasterCoreStartWidthCroppedSvg, exportSizes[0]);
    //convert color to white
    CSTasks.convertColorRGB(rgbDocCroppedVersion.pathItems, colors[violetIndex][0], colors[whiteIndex][0]);
    // save cropped 16 and 24 sizes of PNG into the export folder
    var startWidthCroppedInversed = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    //24x24  
    var filenameCropped24PngInversed = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inverseName, "_").concat(rgbName, "_").concat(exportSizes[7], ".png");
    var destFileCropped24PngInversed = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(pngName)) + filenameCropped24PngInversed);
    CSTasks.scaleAndExportPNG(rgbDocCroppedVersion, destFileCropped24PngInversed, startWidthCroppedInversed, exportSizes[7]);
    //16x16
    var filenameCropped16PngInversed = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inverseName, "_").concat(rgbName, "_").concat(exportSizes[8], ".png");
    var destFileCropped16PngInversed = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(pngName)) + filenameCropped16PngInversed);
    CSTasks.scaleAndExportPNG(rgbDocCroppedVersion, destFileCropped16PngInversed, startWidthCropped, exportSizes[8]);
    //convert to inactive color
    CSTasks.convertAll(rgbDocCroppedVersion.pathItems, colors[grayIndex][0], 100);
    // save cropped 16 and 24 sizes of PNG into the export folder
    var startWidthCroppedInactive = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    //24x24  
    var filenameCropped24PngInactive = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inactiveName, "_").concat(rgbName, "_").concat(exportSizes[7], ".png");
    var destFileCropped24PngInactive = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(pngName)) + filenameCropped24PngInactive);
    CSTasks.scaleAndExportPNG(rgbDocCroppedVersion, destFileCropped24PngInactive, startWidthCroppedInactive, exportSizes[7]);
    //16x16
    var filenameCropped16PngInactive = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inactiveName, "_").concat(rgbName, "_").concat(exportSizes[8], ".png");
    var destFileCropped16PngInactive = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(pngName)) + filenameCropped16PngInactive);
    CSTasks.scaleAndExportPNG(rgbDocCroppedVersion, destFileCropped16PngInactive, startWidthCropped, exportSizes[8]);
    //close and clean up
    rgbDocCroppedVersion.close(SaveOptions.DONOTSAVECHANGES);
    rgbDocCroppedVersion = null;
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
    var cmykDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(epsName)) + cmykFilename);
    var cmykSaveOpts = new EPSSaveOptions();
    cmykDoc.saveAs(cmykDestFile, cmykSaveOpts);
    //convert violet to white and save as EPS
    CSTasks.convertColorCMYK(cmykDoc.pathItems, colors[violetIndex][1], colors[whiteIndex][1]);
    //close and clean up
    cmykDoc.close(SaveOptions.DONOTSAVECHANGES);
    cmykDoc = null;
    /********************
    Lockup export core CMYK (EPS)
    ********************/
    //open a new doc and copy and position the icon and the lockup text
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
    var mastCMYKFilename = "/".concat(iconFilename, "_").concat(lockupName, "_").concat(cmykName, ".eps");
    var mastCMYKDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(epsName)) + mastCMYKFilename);
    var mastCMYKSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    mastCMYKSaveOpts.cmykPostScript = false;
    mastCMYKDoc.saveAs(mastCMYKDestFile, mastCMYKSaveOpts);
    //close and clean up
    mastCMYKDoc.close(SaveOptions.DONOTSAVECHANGES);
    mastCMYKDoc = null;
    /********************
    Lockup export core RGB (EPS)
    ********************/
    //open a new doc and copy and position the icon and the lockup text
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
    var mastFilename = "/".concat(iconFilename, "_").concat(lockupName, "_").concat(rgbName, ".eps");
    var mastDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(coreName, "/").concat(epsName)) + mastFilename);
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
___________                                          .__
\_   _____/__  ________________   ____   ______ _____|__|__  __ ____
 |    __)_\  \/  /\____ \_  __ \_/ __ \ /  ___//  ___/  \  \/ // __ \
 |        \>    < |  |_> >  | \/\  ___/ \___ \ \___ \|  |\   /\  ___/
/_______  /__/\_ \|   __/|__|    \___  >____  >____  >__| \_/  \___  >
        \/      \/|__|               \/     \/     \/              \/
 ***************/
/****************
 * Expressive
 ***************/
function mainExpressive() {
    /******************
    Set up purple artboard 3
    ******************/
    //if there are three artboards, create the new fourth(3rd in array) lockup artboard
    if (sourceDoc.artboards.length == 3 &&
        sourceDoc.artboards[1].artboardRect[2] -
            sourceDoc.artboards[1].artboardRect[0] ==
            256 &&
        sourceDoc.artboards[1].artboardRect[1] -
            sourceDoc.artboards[1].artboardRect[3] ==
            256) {
        // If there are already  3 artboards. Add a 4th one.
        var firstRect = sourceDoc.artboards[1].artboardRect;
        sourceDoc.artboards.add(
        // this fires but then gets replaced further down
        CSTasks.newRect(firstRect[1], firstRect[2] + 128, 1024, 512));
    }
    //if the lockup artboard is present, check if rebuilding or just exporting
    else if (sourceDoc.artboards.length == 3 &&
        sourceDoc.artboards[1].artboardRect[1] -
            sourceDoc.artboards[1].artboardRect[3] ==
            512) {
        rebuild = confirm("It looks like your artwork already exists. This script will rebuild the lockup and export various EPS and PNG versions. Do you want to proceed?");
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
    Create new expressive artboard with lockup and text
    *********************************/
    /*@ts-ignore*/
    var mastBannerIconOnText = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
    var mastBannerIconOnTextPos = [
        sourceDoc.artboards[3].artboardRect[0],
        sourceDoc.artboards[3].artboardRect[1],
    ];
    CSTasks.translateObjectTo(mastBannerIconOnText, mastBannerIconOnTextPos);
    /********************************
    Custom function to create a landing square to place the icon correctly
    Some icons have width or height less than 256 so it needed special centering geometrically
    you can see the landing zone square by changing fill to true and uncommenting color
    *********************************/
    // create a landing zone square to place icon inside
    //moved it outside the function itself so we can delete it after so it doesn't get exported
    var getArtLayer = sourceDoc.layers.getByName('Art');
    var landingZoneSquare = getArtLayer.pathItems.rectangle(-844, 571, 460, 460);
    function placeIconLockup1Correctly(mastBannerIconOnText, maxSize) {
        // let setLandingZoneSquareColor = new RGBColor();
        // setLandingZoneSquareColor.red = 12;
        // setLandingZoneSquareColor.green = 28;
        // setLandingZoneSquareColor.blue = 151;
        // landingZoneSquare.fillColor = setLandingZoneSquareColor;
        landingZoneSquare.name = "LandingZone";
        landingZoneSquare.filled = false;
        /*@ts-ignore*/
        landingZoneSquare.move(getArtLayer, ElementPlacement.PLACEATEND);
        // start moving expressive icon into our new square landing zone
        var placedmastBannerIconOnText = mastBannerIconOnText;
        var landingZone = sourceDoc.pathItems.getByName("LandingZone");
        var preferredWidth = (460);
        var preferredHeight = (460);
        // do the width
        var widthRatio = (preferredWidth / placedmastBannerIconOnText.width) * 100;
        if (placedmastBannerIconOnText.width != preferredWidth) {
            placedmastBannerIconOnText.resize(widthRatio, widthRatio);
        }
        // now do the height
        var heightRatio = (preferredHeight / placedmastBannerIconOnText.height) * 100;
        if (placedmastBannerIconOnText.height != preferredHeight) {
            placedmastBannerIconOnText.resize(heightRatio, heightRatio);
        }
        // now let's center the art on the landing zone
        var centerArt = [placedmastBannerIconOnText.left + (placedmastBannerIconOnText.width / 2), placedmastBannerIconOnText.top + (placedmastBannerIconOnText.height / 2)];
        var centerLz = [landingZone.left + (landingZone.width / 2), landingZone.top + (landingZone.height / 2)];
        placedmastBannerIconOnText.translate(centerLz[0] - centerArt[0], centerLz[1] - centerArt[1]);
        // need another centered proportioning to fix it exactly in correct position
        var W = mastBannerIconOnText.width, H = mastBannerIconOnText.height, MW = maxSize.W, MH = maxSize.H, factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
        mastBannerIconOnText.resize(factor, factor);
    }
    placeIconLockup1Correctly(mastBannerIconOnText, { W: 460, H: 460 });
    // delete the landing zone
    landingZoneSquare.remove();
    // mastBannerIconOnText.width = 460;
    // mastBannerIconOnText.height = 460;
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
    // let rectRef = sourceDoc.pathItems.rectangle(-850, -800, 400, 300);
    // let setTextBoxBgColor = new RGBColor();
    // setTextBoxBgColor.red = 141;
    // setTextBoxBgColor.green = 141;
    // setTextBoxBgColor.blue = 141;
    // rectRef.filled = true;
    // rectRef.fillColor = setTextBoxBgColor;
    // svg wtw logo for new purple lockup
    var imagePlacedItem = myMainArtworkLayer.placedItems.add();
    var svgFile = File("".concat(sourceDoc.path, "/../images/wtw_logo.ai"));
    imagePlacedItem.file = svgFile;
    imagePlacedItem.top = -1188;
    imagePlacedItem.left = 62;
    /*@ts-ignore*/
    // svgFile.embed();
    //request a name for the icon, and place that as text on the lockup artboard
    var appName = prompt("What name do you want to put in second the lockup?");
    var textRef = sourceDoc.textFrames.add();
    //use the areaText method to create the text frame
    var pathRef = sourceDoc.pathItems.rectangle(-850, -800, 480, 400);
    /*@ts-ignore*/
    textRef = sourceDoc.textFrames.areaText(pathRef);
    textRef.contents = appName;
    textRef.textRange.characterAttributes.size = 62;
    textRef.textRange.paragraphAttributes.hyphenation = false;
    // textRef.textRange.characterAttributes.horizontalScale = 2299;
    textRef.textRange.characterAttributes.fillColor = colors[whiteIndex][0];
    CSTasks.setFont(textRef, desiredFont);
    //create an outline of the text
    var textGroup = textRef.createOutline();
    //horizontally align the left edge of the text to be 96px to the right of the edge
    var rightEdge = 64;
    var hOffset = CSTasks.getOffset(textGroup.position, [rightEdge, 0]);
    textGroup.translate(-hOffset[0], 0);
    /*@ts-ignore*/
    var mast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
    //resize the artboard to be only a little wider than the text
    var leftMargin = mast.position[0] - sourceDoc.artboards[3].artboardRect[0];
    var newWidth = textGroup.position[0] +
        textGroup.width -
        sourceDoc.artboards[3].artboardRect[0] +
        leftMargin;
    var resizedRect = CSTasks.newRect(sourceDoc.artboards[3].artboardRect[0], -sourceDoc.artboards[3].artboardRect[1], 1024, 512);
    sourceDoc.artboards[3].artboardRect = resizedRect;
    /******************
    Set up purple artboard 4 in main file
    ******************/
    //if there are 3 artboards, create the new fourth lockup artboard
    // creat last artboard in file
    var FourthMainArtboardFirstRect = sourceDoc.artboards[1].artboardRect;
    sourceDoc.artboards.add(
    // this fires but then gets replaced further down
    CSTasks.newRect(FourthMainArtboardFirstRect[1], FourthMainArtboardFirstRect[2] + 772, 800, 400));
    //select the contents on artboard 1
    var selFourthBanner = CSTasks.selectContentsOnArtboard(sourceDoc, 1);
    // make sure all colors are RGB, equivalent of Edit > Colors > Convert to RGB
    app.executeMenuCommand('Colors9');
    if (selFourthBanner.length == 0) {
        //if nothing is in the artboard
        alert("Please try again with artwork on the main second 256x256 artboard.");
        return;
    }
    /********************************
    Add elements to new fourth artboard with lockup
    *********************************/
    //place icon on lockup
    /*@ts-ignore*/
    var fourthBannerMast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
    var fourthBannerMastPos = [
        sourceDoc.artboards[4].artboardRect[0] + iconOffset[0],
        sourceDoc.artboards[4].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(fourthBannerMast, fourthBannerMastPos);
    /********************************
    Custom function to create a landing square to place the icon correctly
    Some icons have width or height less than 256 so it needed special centering geometrically
    you can see the landing zone square by changing fill to true and uncommenting color
    *********************************/
    // create a landing zone square to place icon inside
    //moved it outside the function itself so we can delete it after so it doesn't get exported
    var getArtLayer2 = sourceDoc.layers.getByName('Art');
    var landingZoneSquare2 = getArtLayer2.pathItems.rectangle(-1472, 444, 360, 360);
    function placeIconLockup1Correctly2(fourthBannerMast, maxSize) {
        // let setLandingZoneSquareColor = new RGBColor();
        // setLandingZoneSquareColor.red = 121;
        // setLandingZoneSquareColor.green = 128;
        // setLandingZoneSquareColor.blue = 131;
        // landingZoneSquare2.fillColor = setLandingZoneSquareColor;
        landingZoneSquare2.name = "LandingZone2";
        landingZoneSquare2.filled = false;
        /*@ts-ignore*/
        landingZoneSquare2.move(getArtLayer2, ElementPlacement.PLACEATEND);
        // start moving expressive icon into our new square landing zone
        var placedfourthBannerMast = fourthBannerMast;
        var landingZone = sourceDoc.pathItems.getByName("LandingZone2");
        var preferredWidth = (360);
        var preferredHeight = (360);
        // do the width
        var widthRatio = (preferredWidth / placedfourthBannerMast.width) * 100;
        if (placedfourthBannerMast.width != preferredWidth) {
            placedfourthBannerMast.resize(widthRatio, widthRatio);
        }
        // now do the height
        var heightRatio = (preferredHeight / placedfourthBannerMast.height) * 100;
        if (placedfourthBannerMast.height != preferredHeight) {
            placedfourthBannerMast.resize(heightRatio, heightRatio);
        }
        // now let's center the art on the landing zone
        var centerArt = [placedfourthBannerMast.left + (placedfourthBannerMast.width / 2), placedfourthBannerMast.top + (placedfourthBannerMast.height / 2)];
        var centerLz = [landingZone.left + (landingZone.width / 2), landingZone.top + (landingZone.height / 2)];
        placedfourthBannerMast.translate(centerLz[0] - centerArt[0], centerLz[1] - centerArt[1]);
        // need another centered proportioning to fix it exactly in correct position
        var W = fourthBannerMast.width, H = fourthBannerMast.height, MW = maxSize.W, MH = maxSize.H, factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
        fourthBannerMast.resize(factor, factor);
    }
    placeIconLockup1Correctly2(fourthBannerMast, { W: 360, H: 360 });
    // delete the landing zone
    landingZoneSquare2.remove();
    // fourthBannerMast.width = 360;
    // fourthBannerMast.height = 360;
    // new purple bg
    // Add new layer above Guidelines and fill white
    var secondMainArtworkLayer = sourceDoc.layers.getByName('Art');
    var secondMainPurpleBgLayer = sourceDoc.layers.add();
    secondMainPurpleBgLayer.name = "Main_Purple_BG_layer_two";
    var getsecondMainPurpleBgLayer = sourceDoc.layers.getByName('Main_Purple_BG_layer_two');
    var seondMainRect = getsecondMainPurpleBgLayer.pathItems.rectangle(-1428, 0, 800, 400);
    var secondMainVioletBgColor = new RGBColor();
    secondMainVioletBgColor.red = 72;
    secondMainVioletBgColor.green = 8;
    secondMainVioletBgColor.blue = 111;
    seondMainRect.filled = true;
    seondMainRect.fillColor = secondMainVioletBgColor;
    /*@ts-ignore*/
    getsecondMainPurpleBgLayer.move(secondMainArtworkLayer, ElementPlacement.PLACEATEND);
    /*@ts-ignore*/
    // svgFile.embed();
    var secondResizedRect = CSTasks.newRect(sourceDoc.artboards[4].artboardRect[0], -sourceDoc.artboards[4].artboardRect[1], 800, 400);
    sourceDoc.artboards[4].artboardRect = secondResizedRect;
    /*********************************************************************
    Export an SVG on a white background
    **********************************************************************/
    var rgbDocOnFFF = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 1, DocumentColorSpace.RGB);
    rgbDocOnFFF.swatches.removeAll();
    var rgbGroupOnFFF = iconGroup.duplicate(rgbDocOnFFF.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATBEGINNING);
    var rgbLocOnFFF = [
        rgbDocOnFFF.artboards[0].artboardRect[0] + iconOffset[0],
        rgbDocOnFFF.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(rgbGroupOnFFF, rgbLocOnFFF);
    CSTasks.ungroupOnce(rgbGroupOnFFF);
    // Add a white square bg here to save svg on white
    var getArtLayerOnFFF = rgbDocOnFFF.layers.getByName('Layer 1');
    var myMainPurpleBgLayerOnFFF = rgbDocOnFFF.layers.add();
    myMainPurpleBgLayerOnFFF.name = "Main_Purple_BG_layer";
    var GetMyMainPurpleBgLayerOnFFF = rgbDocOnFFF.layers.getByName('Main_Purple_BG_layer');
    var landingZoneSquareOnFFF = GetMyMainPurpleBgLayerOnFFF.pathItems.rectangle(0, 400, 256, 256);
    var setLandingZoneSquareOnFFFColor = new RGBColor();
    setLandingZoneSquareOnFFFColor.red = 255;
    setLandingZoneSquareOnFFFColor.green = 255;
    setLandingZoneSquareOnFFFColor.blue = 255;
    landingZoneSquareOnFFF.fillColor = setLandingZoneSquareOnFFFColor;
    landingZoneSquareOnFFF.name = "bgSquareForSvgOnFF";
    landingZoneSquareOnFFF.filled = true;
    landingZoneSquareOnFFF.stroked = false;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayerOnFFF.move(getArtLayerOnFFF, ElementPlacement.PLACEATEND);
    //save a master SVG 
    // let svgMasterCoreStartWidthOnFFF =
    //    rgbDocOnFFF.artboards[0].artboardRect[2] - rgbDocOnFFF.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${onWhiteName}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDocOnFFF, destFile, svgMasterCoreStartWidthOnFFF, exportSizes[2]);
    // }
    //close and clean up
    rgbDocOnFFF.close(SaveOptions.DONOTSAVECHANGES);
    rgbDocOnFFF = null;
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
    //save a 1024 JPG
    var jpegStartWidth1024 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_10 = 0; i_10 < exportSizes.length; i_10++) {
        var filename_4 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, ".jpg");
        var destFile_4 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(jpgName)) + filename_4);
        CSTasks.scaleAndExportJPEG(rgbDoc, destFile_4, jpegStartWidth1024, exportSizes[0]);
    }
    //save a master PNG
    var masterStartWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_11 = 0; i_11 < exportSizes.length; i_11++) {
        var filename_5 = "/".concat(iconFilename, "_").concat(expressiveName, ".png");
        var destFile_5 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName)) + filename_5);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_5, masterStartWidth, exportSizes[0]);
    }
    //save a master SVG
    var svgMasterCoreStartWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_12 = 0; i_12 < exportSizes.length; i_12++) {
        var filename_6 = "/".concat(iconFilename, "_").concat(expressiveName, ".svg");
        var destFile_6 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName)) + filename_6);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_6, svgMasterCoreStartWidth, exportSizes[0]);
    }
    // png 1024
    var startWidth1024 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_13 = 0; i_13 < exportSizes.length; i_13++) {
        var filename_7 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(exportSizes[0], ".png");
        var destFile_7 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_7);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_7, startWidth1024, exportSizes[0]);
    }
    // png 512
    var startWidth512 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_14 = 0; i_14 < exportSizes.length; i_14++) {
        var filename_8 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(exportSizes[1], ".png");
        var destFile_8 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_8);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_8, startWidth512, exportSizes[1]);
    }
    // non transparent 1024 png export
    var startWidthonFFF1024 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_15 = 0; i_15 < exportSizes.length; i_15++) {
        var filename_9 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(onWhiteName, "_").concat(exportSizes[0], ".png");
        var destFile_9 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_9);
        CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile_9, startWidthonFFF1024, exportSizes[0]);
    }
    // non transparent 512 png export
    var startWidthonFFF512 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_16 = 0; i_16 < exportSizes.length; i_16++) {
        var filename_10 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(onWhiteName, "_").concat(exportSizes[1], ".png");
        var destFile_10 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_10);
        CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile_10, startWidthonFFF512, exportSizes[1]);
    }
    //save all sizes of SVG into the export folder
    // we dont need to loop svg, they are scaleable
    // let svgCoreStartWidth =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[1];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${exportSizes[i]}.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${expressiveName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgCoreStartWidth, exportSizes[i]);
    // }
    //save EPS into the export folder
    var filename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, ".eps");
    var destFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(epsName)) + filename);
    var rgbSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    rgbSaveOpts.cmykPostScript = false;
    rgbDoc.saveAs(destFile, rgbSaveOpts);
    //index the RGB colors for conversion to CMYK. An inelegant location.
    var colorIndex = CSTasks.indexRGBColors(rgbDoc.pathItems, colors);
    //convert violet to white and save as EPS
    CSTasks.convertColorRGB(rgbDoc.pathItems, colors[violetIndex][0], colors[whiteIndex][0]);
    // let inverseFilename = `/${iconFilename}_${expressiveName}_${inverseName}_${rgbName}.eps`;
    // let inverseFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${expressiveName}/${epsName}`) + inverseFilename);
    // rgbDoc.saveAs(inverseFile, rgbSaveOpts);
    //save inverse 1024 file in all the PNG sizes
    var startWidthonFFFInversed1024 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_17 = 0; i_17 < exportSizes.length; i_17++) {
        var filename_11 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inverseName, "_").concat(rgbName, "_").concat(exportSizes[0], ".png");
        var destFile_11 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_11);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_11, startWidthonFFFInversed1024, exportSizes[0]);
    }
    //save inverse 512 file in all the PNG sizes
    var startWidthonFFFInversed512 = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_18 = 0; i_18 < exportSizes.length; i_18++) {
        var filename_12 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inverseName, "_").concat(rgbName, "_").concat(exportSizes[1], ".png");
        var destFile_12 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_12);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_12, startWidthonFFFInversed512, exportSizes[1]);
    }
    //save inverse file in all the SVG sizes
    var startWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_19 = 0; i_19 < exportSizes.length; i_19++) {
        var filename_13 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inverseName, "_").concat(rgbName, "_").concat(exportSizes[2], ".svg");
        var destFile_13 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)) + filename_13);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_13, startWidth, exportSizes[2]);
    }
    //convert to inactive color (WTW Icon grey at 100% opacity) and save as EPS
    CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 100);
    // let inactiveFilename = `/${iconFilename}_${expressiveName}_${inactiveName}_${rgbName}.eps`;
    // let inactiveFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${expressiveName}/${epsName}`) + inactiveFilename);
    // rgbDoc.saveAs(inactiveFile, rgbSaveOpts);
    // save inactive png 1024
    for (var i_20 = 0; i_20 < exportSizes.length; i_20++) {
        var filename_14 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inactiveName, "_").concat(rgbName, "_").concat(exportSizes[0], ".png");
        var destFile_14 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_14);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_14, startWidth, exportSizes[0]);
    }
    // save inactive png 512
    for (var i_21 = 0; i_21 < exportSizes.length; i_21++) {
        var filename_15 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inactiveName, "_").concat(rgbName, "_").concat(exportSizes[1], ".png");
        var destFile_15 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_15);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_15, startWidth, exportSizes[1]);
    }
    for (var i_22 = 0; i_22 < exportSizes.length; i_22++) {
        var filename_16 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(inactiveName, "_").concat(rgbName, "_").concat(exportSizes[2], ".svg");
        var destFile_16 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)) + filename_16);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_16, startWidth, exportSizes[2]);
    }
    //close and clean up
    rgbDoc.close(SaveOptions.DONOTSAVECHANGES);
    rgbDoc = null;
    /*********************************************************************
    RGB cropped export (JPG, PNGs at 16 and 24 sizes), squares, cropped to artwork
    **********************************************************************/
    var rgbDocCroppedVersion = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 1, DocumentColorSpace.RGB);
    rgbDocCroppedVersion.swatches.removeAll();
    var rgbGroupCropped = iconGroup.duplicate(rgbDocCroppedVersion.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    var rgbLocCropped = [
        rgbDocCroppedVersion.artboards[0].artboardRect[0] + iconOffset[0],
        rgbDocCroppedVersion.artboards[0].artboardRect[1] + iconOffset[1],
    ];
    CSTasks.translateObjectTo(rgbGroupCropped, rgbLocCropped);
    // remove padding here befor exporting
    function resizeCroppedExpressiveIcon(rgbGroupCropped, maxSize) {
        var W = rgbGroupCropped.width, H = rgbGroupCropped.height, MW = maxSize.W, MH = maxSize.H, factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
        rgbGroupCropped.resize(factor, factor);
    }
    resizeCroppedExpressiveIcon(rgbGroupCropped, { W: 256, H: 256 });
    CSTasks.ungroupOnce(rgbGroupCropped);
    // Save a cropped SVG 
    var svgMasterCoreStartWidthCroppedSvg = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    var filenameCroppedSvg = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(exportSizes[8], "_").concat(croppedName, ".svg");
    var destFileCroppedSvg = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)) + filenameCroppedSvg);
    CSTasks.scaleAndExportSVG(rgbDocCroppedVersion, destFileCroppedSvg, svgMasterCoreStartWidthCroppedSvg, exportSizes[8]);
    //convert color to white
    CSTasks.convertColorRGB(rgbDocCroppedVersion.pathItems, colors[violetIndex][0], colors[whiteIndex][0]);
    // no point making an inversed jpg, it is white on white, so futile
    // Save a cropped SVG
    var svgMasterCoreStartWidthCroppedSvgInversed = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    var filenameCroppedSvgInversed = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(exportSizes[8], "_").concat(inverseName, "_").concat(croppedName, ".svg");
    var destFileCroppedSvgInversed = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)) + filenameCroppedSvgInversed);
    CSTasks.scaleAndExportSVG(rgbDocCroppedVersion, destFileCroppedSvgInversed, svgMasterCoreStartWidthCroppedSvgInversed, exportSizes[8]);
    //convert to inactive color
    CSTasks.convertAll(rgbDocCroppedVersion.pathItems, colors[grayIndex][0], 100);
    // Save a cropped SVG
    var svgMasterCoreStartWidthCroppedSvgInactive = rgbDocCroppedVersion.artboards[0].artboardRect[2] - rgbDocCroppedVersion.artboards[0].artboardRect[0];
    var filenameCroppedSvgInactive = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(exportSizes[8], "_").concat(inactiveName, "_").concat(croppedName, ".svg");
    var destFileCroppedSvgInactive = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)) + filenameCroppedSvgInactive);
    CSTasks.scaleAndExportSVG(rgbDocCroppedVersion, destFileCroppedSvgInactive, svgMasterCoreStartWidthCroppedSvgInactive, exportSizes[8]);
    //close and clean up
    rgbDocCroppedVersion.close(SaveOptions.DONOTSAVECHANGES);
    rgbDocCroppedVersion = null;
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
    //save a 1024 JPG
    var jpegStartWidth1024CMYK = cmykDoc.artboards[0].artboardRect[2] - cmykDoc.artboards[0].artboardRect[0];
    for (var i_23 = 0; i_23 < exportSizes.length; i_23++) {
        var filename_17 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(cmykName, ".jpg");
        var destFile_17 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(jpgName)) + filename_17);
        CSTasks.scaleAndExportJPEG(cmykDoc, destFile_17, jpegStartWidth1024CMYK, exportSizes[0]);
    }
    //save EPS into the export folder
    var cmykFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(cmykName, ".eps");
    var cmykDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(epsName)) + cmykFilename);
    var cmykSaveOpts = new EPSSaveOptions();
    cmykDoc.saveAs(cmykDestFile, cmykSaveOpts);
    //convert violet to white and save as EPS
    CSTasks.convertColorCMYK(cmykDoc.pathItems, colors[violetIndex][1], colors[whiteIndex][1]);
    // let cmykInverseFilename = `/${iconFilename}_${expressiveName}_${inverseName}_${cmykName}.eps`;
    // let cmykInverseFile = new File(Folder(`${sourceDoc.path}/${sourceDocName}/${expressiveName}/${epsName}`) + cmykInverseFilename);
    // cmykDoc.saveAs(cmykInverseFile, rgbSaveOpts);
    //close and clean up
    cmykDoc.close(SaveOptions.DONOTSAVECHANGES);
    cmykDoc = null;
    /********************
    Purple Lockup with text export
    ********************/
    //open a new doc and copy and position the icon and the lockup text
    // duplication did not work as expected here. I have used a less elegant solution whereby I recreated the purple banner instead of copying it.
    var mastDoc = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 3, DocumentColorSpace.RGB);
    mastDoc.swatches.removeAll();
    var mastGroup = iconGroup.duplicate(mastDoc.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    // new icon width in rebrand
    // mastGroup.width = 460;
    // mastGroup.height = 460;
    // new icon position
    var mastLoc = [
        mastDoc.artboards[0].artboardRect[0],
        mastDoc.artboards[0].artboardRect[1],
    ];
    CSTasks.translateObjectTo(mastGroup, mastLoc);
    /********************************
    Custom function to create a landing square to place the icon correctly
    Some icons have width or height less than 256 so it needed special centering geometrically
    you can see the landing zone square by changing fill to true and uncommenting color
    *********************************/
    // create a landing zone square to place icon inside
    //moved it outside the function itself so we can delete it after so it doesn't get exported
    var getArtLayer3 = mastDoc.layers.getByName('Layer 1');
    var landingZoneSquare3 = getArtLayer3.pathItems.rectangle(-843, 570, 460, 460);
    function placeIconLockup1Correctly3(mastGroup, maxSize) {
        // let setLandingZoneSquareColor = new RGBColor();
        // setLandingZoneSquareColor.red = 121;
        // setLandingZoneSquareColor.green = 128;
        // setLandingZoneSquareColor.blue = 131;
        // landingZoneSquare3.fillColor = setLandingZoneSquareColor;
        landingZoneSquare3.name = "LandingZone3";
        landingZoneSquare3.filled = false;
        /*@ts-ignore*/
        landingZoneSquare3.move(getArtLayer3, ElementPlacement.PLACEATEND);
        // start moving expressive icon into our new square landing zone
        var placedmastGroup = mastGroup;
        var landingZone = mastDoc.pathItems.getByName("LandingZone3");
        var preferredWidth = (460);
        var preferredHeight = (460);
        // do the width
        var widthRatio = (preferredWidth / placedmastGroup.width) * 100;
        if (placedmastGroup.width != preferredWidth) {
            placedmastGroup.resize(widthRatio, widthRatio);
        }
        // now do the height
        var heightRatio = (preferredHeight / placedmastGroup.height) * 100;
        if (placedmastGroup.height != preferredHeight) {
            placedmastGroup.resize(heightRatio, heightRatio);
        }
        // now let's center the art on the landing zone
        var centerArt = [placedmastGroup.left + (placedmastGroup.width / 2), placedmastGroup.top + (placedmastGroup.height / 2)];
        var centerLz = [landingZone.left + (landingZone.width / 2), landingZone.top + (landingZone.height / 2)];
        placedmastGroup.translate(centerLz[0] - centerArt[0], centerLz[1] - centerArt[1]);
        // need another centered proportioning to fix it exactly in correct position
        var W = mastGroup.width, H = mastGroup.height, MW = maxSize.W, MH = maxSize.H, factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
        mastGroup.resize(factor, factor);
    }
    placeIconLockup1Correctly3(mastGroup, { W: 460, H: 460 });
    // delete the landing zone
    landingZoneSquare3.remove();
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
    var mainRectMastDoc = GetMyMainPurpleBgLayerMastDoc.pathItems.rectangle(-784, 0, 1024, 512);
    var setMainVioletBgColorMastDoc = new RGBColor();
    setMainVioletBgColorMastDoc.red = 72;
    setMainVioletBgColorMastDoc.green = 8;
    setMainVioletBgColorMastDoc.blue = 111;
    mainRectMastDoc.filled = true;
    mainRectMastDoc.fillColor = setMainVioletBgColorMastDoc;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayerMastDoc.move(myMainArtworkLayerMastDoc, ElementPlacement.PLACEATEND);
    // svg wtw logo for new purple lockup
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
    var mainClipRectMastDoc = GetMyCroppingLayerMastDoc.pathItems.rectangle(-784, 0, 1024, 512);
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
    //save a banner JPG
    var jpegStartWidth = mastDoc.artboards[0].artboardRect[2] - mastDoc.artboards[0].artboardRect[0];
    for (var i_24 = 0; i_24 < exportSizes.length; i_24++) {
        var filename_18 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(tenByFive, ".jpg");
        var destFile_18 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(jpgName)) + filename_18);
        CSTasks.scaleAndExportJPEG(mastDoc, destFile_18, jpegStartWidth, exportSizes[0]);
    }
    //save a banner PNG
    for (var i_25 = 0; i_25 < exportSizes.length; i_25++) {
        var filename_19 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(lockupName, "_").concat(rgbName, "_").concat(lockup1, ".png");
        var destFile_19 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_19);
        CSTasks.scaleAndExportPNG(mastDoc, destFile_19, 512, 1024);
    }
    //save a banner SVG 
    for (var i_26 = 0; i_26 < exportSizes.length; i_26++) {
        var filename_20 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(lockupName, "_").concat(rgbName, "_").concat(lockup1, ".svg");
        var destFile_20 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)) + filename_20);
        CSTasks.scaleAndExportSVG(mastDoc, destFile_20, 512, 1024);
    }
    //save RGB EPS into the export folder
    var mastFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(tenByFive, ".eps");
    var mastDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(epsName)) + mastFilename);
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
      Purple Lockup with no text export
      ********************/
    //open a new doc and copy and position the icon
    // duplication did not work as expected here. I have used a less elegant solution whereby I recreated the purple banner instead of copying it.
    var mastDocNoText = CSTasks.duplicateArtboardInNewDoc(sourceDoc, 4, DocumentColorSpace.RGB);
    mastDocNoText.swatches.removeAll();
    var mastGroupNoText = iconGroup.duplicate(mastDocNoText.layers[0], 
    /*@ts-ignore*/
    ElementPlacement.PLACEATEND);
    // new icon width in rebrand
    // mastGroupNoText.width = 360;
    // mastGroupNoText.height = 360;
    // new icon position
    var mastLocNoText = [
        mastDocNoText.artboards[0].artboardRect[0],
        mastDocNoText.artboards[0].artboardRect[1],
    ];
    CSTasks.translateObjectTo(mastGroupNoText, mastLocNoText);
    /********************************
      Custom function to create a landing square to place the icon correctly
      Some icons have width or height less than 256 so it needed special centering geometrically
      you can see the landing zone square by changing fill to true and uncommenting color
      *********************************/
    // create a landing zone square to place icon inside
    //moved it outside the function itself so we can delete it after so it doesn't get exported
    var getArtLayerIn4thArtboard = mastDocNoText.layers.getByName('Layer 1');
    var landingZoneSquareInFourthArtboard = getArtLayerIn4thArtboard.pathItems.rectangle(-1471, 443, 360, 360);
    function placeIconLockupCorrectlyIn4thDoc(mastGroupNoText, maxSize) {
        // let setLandingZoneSquareColor = new RGBColor();
        // setLandingZoneSquareColor.red = 121;
        // setLandingZoneSquareColor.green = 128;
        // setLandingZoneSquareColor.blue = 131;
        // landingZoneSquareInFourthArtboard.fillColor = setLandingZoneSquareColor;
        landingZoneSquareInFourthArtboard.name = "LandingZone4";
        landingZoneSquareInFourthArtboard.filled = false;
        /*@ts-ignore*/
        landingZoneSquareInFourthArtboard.move(getArtLayerIn4thArtboard, ElementPlacement.PLACEATEND);
        // start moving expressive icon into our new square landing zone
        var placedmastGroup = mastGroupNoText;
        var landingZone = mastDocNoText.pathItems.getByName("LandingZone4");
        var preferredWidth = (360);
        var preferredHeight = (360);
        // do the width
        var widthRatio = (preferredWidth / placedmastGroup.width) * 100;
        if (placedmastGroup.width != preferredWidth) {
            placedmastGroup.resize(widthRatio, widthRatio);
        }
        // now do the height
        var heightRatio = (preferredHeight / placedmastGroup.height) * 100;
        if (placedmastGroup.height != preferredHeight) {
            placedmastGroup.resize(heightRatio, heightRatio);
        }
        // now let's center the art on the landing zone
        var centerArt = [placedmastGroup.left + (placedmastGroup.width / 2), placedmastGroup.top + (placedmastGroup.height / 2)];
        var centerLz = [landingZone.left + (landingZone.width / 2), landingZone.top + (landingZone.height / 2)];
        placedmastGroup.translate(centerLz[0] - centerArt[0], centerLz[1] - centerArt[1]);
        // need another centered proportioning to fix it exactly in correct position
        var W = mastGroupNoText.width, H = mastGroupNoText.height, MW = maxSize.W, MH = maxSize.H, factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
        mastGroupNoText.resize(factor, factor);
    }
    placeIconLockupCorrectlyIn4thDoc(mastGroupNoText, { W: 360, H: 360 });
    // delete the landing zone
    landingZoneSquareInFourthArtboard.remove();
    CSTasks.ungroupOnce(mastGroupNoText);
    // add new style purple banner 4 elements
    var myMainArtworkLayerMastDocNoText = mastDocNoText.layers.getByName('Layer 1');
    var myMainPurpleBgLayerMastDocNoText = mastDocNoText.layers.add();
    myMainPurpleBgLayerMastDocNoText.name = "Main_Purple_BG_layer";
    var GetMyMainPurpleBgLayerMastDocNoText = mastDocNoText.layers.getByName('Main_Purple_BG_layer');
    var mainRectMastDocNoText = GetMyMainPurpleBgLayerMastDocNoText.pathItems.rectangle(-1428, 0, 800, 400);
    var setMainVioletBgColorMastDocNoText = new RGBColor();
    setMainVioletBgColorMastDocNoText.red = 72;
    setMainVioletBgColorMastDocNoText.green = 8;
    setMainVioletBgColorMastDocNoText.blue = 111;
    mainRectMastDocNoText.filled = true;
    mainRectMastDocNoText.fillColor = setMainVioletBgColorMastDocNoText;
    /*@ts-ignore*/
    GetMyMainPurpleBgLayerMastDocNoText.move(myMainArtworkLayerMastDocNoText, ElementPlacement.PLACEATEND);
    // we need to make artboard clipping mask here for the artboard to crop expressive icons correctly.
    var myCroppingLayerMastDocNoText = mastDocNoText.layers.add();
    myCroppingLayerMastDocNoText.name = "crop";
    var GetMyCroppingLayerMastDocNoText = mastDocNoText.layers.getByName('crop');
    mastDocNoText.activeLayer = GetMyCroppingLayerMastDocNoText;
    mastDocNoText.activeLayer.hasSelectedArtwork = true;
    // insert clipping rect here
    var mainClipRectMastDocNoText = GetMyCroppingLayerMastDocNoText.pathItems.rectangle(-1428, 0, 800, 400);
    var setClipBgColorMastDocNoText = new RGBColor();
    setClipBgColorMastDocNoText.red = 0;
    setClipBgColorMastDocNoText.green = 255;
    setClipBgColorMastDocNoText.blue = 255;
    mainClipRectMastDocNoText.filled = true;
    mainClipRectMastDocNoText.fillColor = setClipBgColorMastDocNoText;
    // select all for clipping here
    mastDocNoText.selectObjectsOnActiveArtboard();
    // clip!
    app.executeMenuCommand('makeMask');
    //save a banner JPG
    var jpegStartWidth2 = mastDocNoText.artboards[0].artboardRect[2] - mastDocNoText.artboards[0].artboardRect[0];
    for (var i_27 = 0; i_27 < exportSizes.length; i_27++) {
        var filename_21 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(eightByFour, ".jpg");
        var destFile_21 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(jpgName)) + filename_21);
        CSTasks.scaleAndExportJPEG(mastDocNoText, destFile_21, jpegStartWidth2, 800);
    }
    //save a banner PNG
    for (var i_28 = 0; i_28 < exportSizes.length; i_28++) {
        var filename_22 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(lockupName, "_").concat(rgbName, "_").concat(lockup2, ".png");
        var destFile_22 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(pngName)) + filename_22);
        CSTasks.scaleAndExportPNG(mastDocNoText, destFile_22, 400, 800);
    }
    //save a banner SVG
    for (var i_29 = 0; i_29 < exportSizes.length; i_29++) {
        var filename_23 = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(lockupName, "_").concat(rgbName, "_").concat(lockup2, ".svg");
        var destFile_23 = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(svgName)) + filename_23);
        CSTasks.scaleAndExportSVG(mastDocNoText, destFile_23, 400, 800);
    }
    //save RGB EPS into the export folder 
    var mastNoTextFilename = "/".concat(iconFilename, "_").concat(expressiveName, "_").concat(rgbName, "_").concat(eightByFour, ".eps");
    var mastNoTextDestFile = new File(Folder("".concat(sourceDoc.path, "/").concat(sourceDocName, "/").concat(expressiveName, "/").concat(epsName)) + mastNoTextFilename);
    var mastNoTextSaveOpts = new EPSSaveOptions();
    /*@ts-ignore*/
    mastNoTextSaveOpts.cmykPostScript = false;
    /*@ts-ignore*/
    mastNoTextSaveOpts.embedLinkedFiles = true;
    mastDocNoText.saveAs(mastNoTextDestFile, mastNoTextSaveOpts);
    //close and clean up
    mastDocNoText.close(SaveOptions.DONOTSAVECHANGES);
    mastDocNoText = null;
    /************
    Final cleanup
    ************/
    CSTasks.ungroupOnce(iconGroup);
    CSTasks.ungroupOnce(mast);
}
mainExpressive();
// this opens the folder where the assets are saved
var scriptsFolder = Folder(sourceDoc.path + "/");
scriptsFolder.execute();
