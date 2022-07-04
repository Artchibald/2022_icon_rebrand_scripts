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
    [255, 255, 255], // white
];
var CMYKColorElements = [
    [29, 70, 0, 30],
    [0, 0, 0, 25],
    [0, 100, 14, 21],
    [79, 47, 0, 6],
    [74, 0, 9, 14],
    [0, 0, 0, 0], // white
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
// Colors
var rgbName = "RGB";
var cmykName = "CMYK";
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
    //returns an aray [x,y] for the offset between the two points
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
 ***************/
function main() {
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
        // Expressive folder(not in use yet)
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
    // Add new layer above Guidelines and fill white
    // let myWhiteBgLayer = sourceDoc.layers.add();
    // myWhiteBgLayer.name = "White_BG_layer";
    // let getArtworkLayer = sourceDoc.layers.getByName('Artwork');
    // sourceDoc.activeLayer = sourceDoc.layers.getByName("Artwork"); // activates third layer from top
    // sourceDoc.activeLayer.hasSelectedArtwork = true; // selects all in active layer
    // var rect = getArtworkLayer.pathItems.rectangle(
    //    rgbDoc.artboards[0].artboardRect[0],
    //    rgbDoc.artboards[0].artboardRect[1],
    //    256,
    //    256);
    // var setColor = new RGBColor();
    // setColor.red = 155;
    // setColor.green = 155;
    // setColor.blue = 155;
    // rect.filled = true;
    // rect.fillColor = setColor;
    // sourceDoc.selection = null;
    //var artworkLayer = sourceDoc.layers.getByName('Artwork').layers.getByName('<Path>');
    //rect.zOrder(ZOrderMethod.SENDTOBACK);
    // let getWhiteBgLayer = sourceDoc.layers.getByName('Artwork').layers.getByName('<Path>');
    // getWhiteBgLayer.zOrder(ZOrderMethod.SENDTOBACK);
    //artworkLayer.locked;
    //save all sizes of PNG into the export folder
    // let startWidthonFFF =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}_onFFF.png`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
    //    CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile, startWidthonFFF, exportSizes[i]);
    // }
    //sourceDoc.layers.getByName('White_BG_layer').remove();
    //save all sizes of SVG into the export folder
    // let svgCoreStartWidthonFFF =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}_onFFF.svg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
    //    CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgCoreStartWidthonFFF, exportSizes[i]);
    // }
    // //save all sizes of JPEG into the export folder
    // let jpegStartWidthonFFF =
    //    rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    // for (let i = 0; i < exportSizes.length; i++) {
    //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}_onFFF.jpg`;
    //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${jpgName}`) + filename);
    //    CSTasks.scaleAndExportJPEG(rgbDoc, destFile, jpegStartWidthonFFF, exportSizes[i]);
    // }
    // then repeat export loops for SVG PNG JPG Here
    //create a new document with the artboard and contents from artboard 0
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
    var startWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_9 = 0; i_9 < exportSizes.length; i_9++) {
        var filename_3 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[i_9], ".png");
        var destFile_3 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(pngName)) + filename_3);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_3, startWidth, exportSizes[i_9]);
    }
    // non transparent png
    var startWidthonFFF = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_10 = 0; i_10 < exportSizes.length; i_10++) {
        var filename_4 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[i_10], "_onFFF.png");
        var destFile_4 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(pngName)) + filename_4);
        CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile_4, startWidthonFFF, exportSizes[i_10]);
    }
    //save all sizes of SVG into the export folder
    var svgCoreStartWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_11 = 0; i_11 < exportSizes.length; i_11++) {
        var filename_5 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[i_11], ".svg");
        var destFile_5 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(svgName)) + filename_5);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_5, svgCoreStartWidth, exportSizes[i_11]);
    }
    //save all sizes of JPEG into the export folder
    var jpegStartWidth = rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
    for (var i_12 = 0; i_12 < exportSizes.length; i_12++) {
        var filename_6 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(rgbName, "_").concat(exportSizes[i_12], ".jpg");
        var destFile_6 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(jpgName)) + filename_6);
        CSTasks.scaleAndExportJPEG(rgbDoc, destFile_6, jpegStartWidth, exportSizes[i_12]);
    }
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
    for (var i_13 = 0; i_13 < exportSizes.length; i_13++) {
        var filename_7 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inverseName, "_").concat(rgbName, "_").concat(exportSizes[i_13], ".png");
        var destFile_7 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(pngName)) + filename_7);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_7, startWidth, exportSizes[i_13]);
    }
    //save inverse file in all the SVG sizes
    for (var i_14 = 0; i_14 < exportSizes.length; i_14++) {
        var filename_8 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inverseName, "_").concat(rgbName, "_").concat(exportSizes[i_14], ".svg");
        var destFile_8 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(svgName)) + filename_8);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_8, startWidth, exportSizes[i_14]);
    }
    //convert to inactive color (WTW Icon grey at 100% opacity) and save as EPS
    CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 100);
    var inactiveFilename = "/".concat(iconFilename, "_").concat(inactiveName, "_").concat(rgbName, ".eps");
    var inactiveFile = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(epsName)) + inactiveFilename);
    rgbDoc.saveAs(inactiveFile, rgbSaveOpts);
    for (var i_15 = 0; i_15 < exportSizes.length; i_15++) {
        var filename_9 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inactiveName, "_").concat(rgbName, "_").concat(exportSizes[i_15], ".png");
        var destFile_9 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(pngName)) + filename_9);
        CSTasks.scaleAndExportPNG(rgbDoc, destFile_9, startWidth, exportSizes[i_15]);
    }
    for (var i_16 = 0; i_16 < exportSizes.length; i_16++) {
        var filename_10 = "/".concat(iconFilename, "_").concat(coreName, "_").concat(inactiveName, "_").concat(rgbName, "_").concat(exportSizes[i_16], ".svg");
        var destFile_10 = new File(Folder("".concat(sourceDoc.path, "/").concat(coreName, "/").concat(svgName)) + filename_10);
        CSTasks.scaleAndExportSVG(rgbDoc, destFile_10, startWidth, exportSizes[i_16]);
    }
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
    Masthead export (EPS)
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
    var mastFilename = "/".concat(iconFilename, "_Masthead_").concat(rgbName, ".eps");
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
    /************
 Add white BG and save again
 ************/
}
main();
