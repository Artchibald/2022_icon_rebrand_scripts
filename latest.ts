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
// alert("FULL README: https://github.com/Artchibald/2022_icon_rebrand_scripts \n\nVideo set up tutorial available here: XXXTBCXXX \n\n If you run the script again, you should probably delete the previous assets created. \n\nArtboard size must be exactly 256px x 256px. \n\nGuides must be on a layer called exactly 'Guidelines'. \n\n Icons must be on a layer called exactly 'Art'. \n\nMake sure all icons are unlocked to avoid bugs. \n\nWhen the contact sheet is ready to be generated, you can edit the rows and columns. If you increase the columns, you should increase the rows, to avoid overlapping. There is also a separate version of just the contact-sheet script. You must have these 2 folders in the root, for it to work though: /sorted-by-type/SVG/. \n\nExported assets will be saved where the .ai file is saved. \n\nAny issues: archie@archibaldbutler.com.");
/********************************* 
VARIABLES YOU MIGHT NEED TO CHANGE
**********************************/
let sourceDoc = app.activeDocument;
let RGBColorElements = [
   [127, 53, 178], //ultraviolet purple
   [191, 191, 191], //Gray matter light grey
   [201, 0, 172], // Fireworks magenta
   [50, 127, 239], //Stratosphere blue
   [58, 220, 201], // Inifinity turquoise
   [255, 255, 255], // white
   [128, 128, 128], // Dark grey (unused)
];
// New CMYK values dont math rgb exatcly in new branding 2022 so we stopped the exact comparison part of the script.
// Intent is different colors in print for optimum pop of colors
let CMYKColorElements = [
   [65, 91, 0, 0], //ultraviolet purple
   [0, 0, 0, 25], //Gray matter light grey
   [16, 96, 0, 0], // Fireworks magenta
   [78, 47, 0, 0], //Stratosphere blue
   [53, 0, 34, 0], // Inifinity turquoise  
   [0, 0, 0, 0], // white
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

let desiredFont = "Graphik-Regular";
let exportSizes = [1024, 512, 256, 128, 64, 48, 32, 24, 16]; //sizes to export
let violetIndex = 0; //these are for converting to inverse and inactive versions
let grayIndex = 1;
let whiteIndex = 5;
//loop default 
let i;
// folder creations
let coreName = "Core";
let expressiveName = "Expressive";
let inverseName = "Inverse";
let inactiveName = "Inactive";
// Masthead
let mastheadName = "Masthead";
let masthead1 = "Masthead1"
let masthead2 = "Masthead2"
// Colors
let rgbName = "RGB";
let cmykName = "CMYK";
let onWhiteName = "onFFF";
//Folder creations
let pngName = "png";
let jpgName = "jpg";
let svgName = "svg";
let epsName = "eps";
let iconFilename = sourceDoc.name.split(".")[0];
let rebuild = true;
// let gutter = 32;
// hide guides
let guideLayer = sourceDoc.layers["Guidelines"];
let name = sourceDoc.name.split(".")[0];
let destFolder = Folder(sourceDoc.path + "/" + name);

/**********************************
Module for image manipulation tasks 
***********************************/

// create a white bg layer, send to bottom of layer stack 
// let numberOfLayersToBeAdded = 1;

// let artboardRef = sourceDoc.artboards[0];
// let top = artboardRef.artboardRect[1];
// let left = artboardRef.artboardRect[0];
// let width = artboardRef.artboardRect[2] - artboardRef.artboardRect[0];
// let height = artboardRef.artboardRect[1] - artboardRef.artboardRect[3];
// let rect = sourceDoc.pathItems.rectangle(top, left, width, height);
// rect.fillColor = rect.strokeColor = new NoColor();

// then repeat export loops for SVG PNG JPG Here


interface Task {
   getArtboardCorner(artboard: any);
   getOffset(itemPos: any, referencePos: any);
   translateObjectTo(object: any, destination: any);
   clearArtboard(doc: any, index: any);
   selectContentsOnArtboard(doc: any, i: any);
   createGroup(doc: any, collection: any);
   ungroupOnce(group: any);
   newDocument(sourceDoc: any, colorSpace: any);
   duplicateArtboardInNewDoc(sourceDoc: any,
      artboardIndex: number,
      colorspace: any);
   scaleAndExportPNG(doc: any, destFile: any, startWidth: any, desiredWidth: any);
   scaleAndExportNonTransparentPNG(doc: any, destFile: any, startWidth: any, desiredWidth: any);
   scaleAndExportSVG(doc: any, destFile: any, startWidth: any, desiredWidth: any);
   scaleAndExportJPEG(doc: any, destFile: any, startWidth: any, desiredWidth: any);
   newRect(x: any, y: any, width: any, height: any);
   setFont(textRef: any, desiredFont: any);
   createTextFrame(doc: any, message: any, pos: any, size: any);
   initializeColors(RGBArray: any, CMYKArray: any);
   matchRGB(color: any, matchArray: any);
   matchColorsRGB(color1: any, color2: any);
   convertColorCMYK(pathItems: any, startColor: any, endColor: any)
   matchRGB(color: any, matchArray: any);
   matchColorsRGB(color1: any, color2: any);
   convertColorCMYK(pathItems: any, startColor: any, endColor: any)
   matchColorsCMYK(color1: any, color2: any): any;
   convertColorRGB(pathItems: any, startColor: any, endColor: any);
   convertAll(pathItems: any, endColor: any, opcty: any);
   indexRGBColors(pathItems: any, matchArray: any);
   convertToCMYK(doc: any, pathItems: any, colorArray: any, colorIndex: any);
   unique(a: any): any;

}


let CSTasks = (function () {
   let tasks: Task = {} as Task;

   /********************
      POSITION AND MOVEMENT
      ********************/

   //takes an artboard
   //returns its left top corner as an array [x,y]
   tasks.getArtboardCorner = function (artboard) {
      let corner = [artboard.artboardRect[0], artboard.artboardRect[1]];
      return corner;
   };

   //takes an array [x,y] for an item's position and an array [x,y] for the position of a reference point
   //returns an array [x,y] for the offset between the two points
   tasks.getOffset = function (itemPos, referencePos) {
      let offset = [itemPos[0] - referencePos[0], itemPos[1] - referencePos[1]];
      return offset;
   };

   //takes an object (e.g. group) and a destination array [x,y]
   //moves the group to the specified destination
   tasks.translateObjectTo = function (object, destination) {
      let offset = tasks.getOffset(object.position, destination);
      object.translate(-offset[0], -offset[1]);
   };

   //takes a document and index of an artboard
   //deletes everything on that artboard
   tasks.clearArtboard = function (doc, index) {
      //clears an artboard at the given index
      doc.selection = null;
      doc.artboards.setActiveArtboardIndex(index);
      doc.selectObjectsOnActiveArtboard();
      let sel = doc.selection; // get selection

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
      let newGroup = doc.groupItems.add();
      for (i = 0; i < collection.length; i++) {
         collection[i].moveToBeginning(newGroup);
      }
      return newGroup;
   };

   //takes a group
   //ungroups that group at the top layer (no recursion for nested groups)
   tasks.ungroupOnce = function (group) {
      for (i = group.pageItems.length - 1; i >= 0; i--) {
         group.pageItems[i].move(
            group.pageItems[i].layer,
            /*@ts-ignore*/
            ElementPlacement.PLACEATEND
         );
      }
   };

   /****************************
      CREATING AND SAVING DOCUMENTS
      *****************************/

   //take a source document and a colorspace (e.g. DocumentColorSpace.RGB)
   //opens and returns a new document with the source document's units and the specified colorspace
   tasks.newDocument = function (sourceDoc, colorSpace) {
      let preset = new DocumentPreset();
      /*@ts-ignore*/
      preset.colorMode = colorSpace;
      /*@ts-ignore*/
      preset.units = sourceDoc.rulerUnits;
      /*@ts-ignore*/
      let newDoc = app.documents.addDocument(colorSpace, preset);
      newDoc.pageOrigin = sourceDoc.pageOrigin;
      newDoc.rulerOrigin = sourceDoc.rulerOrigin;

      return newDoc;
   };

   //take a source document, artboard index, and a colorspace (e.g. DocumentColorSpace.RGB)
   //opens and returns a new document with the source document's units and specified artboard, the specified colorspace
   tasks.duplicateArtboardInNewDoc = function (
      sourceDoc,
      artboardIndex,
      colorspace
   ) {
      let rectToCopy = sourceDoc.artboards[artboardIndex].artboardRect;
      let newDoc = tasks.newDocument(sourceDoc, colorspace);
      newDoc.artboards.add(rectToCopy);
      newDoc.artboards.remove(0);
      return newDoc;
   };

   //takes a document, destination file, starting width and desired width
   //scales the document proportionally to the desired width and exports as a PNG
   tasks.scaleAndExportPNG = function (doc, destFile, startWidth, desiredWidth) {
      let scaling = (100.0 * desiredWidth) / startWidth;
      let options = new ExportOptionsPNG24();
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
      let scaling = (100.0 * desiredWidth) / startWidth;
      let options = new ExportOptionsPNG24();
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
      let scaling = (100.0 * desiredWidth) / startWidth;
      let options = new ExportOptionsSVG();
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
      let scaling = (100.0 * desiredWidth) / startWidth;
      let options = new ExportOptionsJPEG();
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
      let rect = [];
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
      let foundFont = false;
      /*@ts-ignore*/
      for (let i = 0; i < textFonts.length; i++) {
         /*@ts-ignore*/
         if (textFonts[i].name == desiredFont) {
            /*@ts-ignore*/
            textRef.textRange.characterAttributes.textFont = textFonts[i];
            foundFont = true;
            break;
         }
      }
      if (!foundFont)
         alert(
            "Didn't find the font. Please check if the font is installed or check the script to make sure the font name is right."
         );
   };

   //takes a document, message string, position array and font size
   //creates a text frame with the message
   tasks.createTextFrame = function (doc, message, pos, size) {
      let textRef = doc.textFrames.add();
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
      let colors = new Array(RGBArray.length);

      for (let i = 0; i < RGBArray.length; i++) {
         let rgb = new RGBColor();
         rgb.red = RGBArray[i][0];
         rgb.green = RGBArray[i][1];
         rgb.blue = RGBArray[i][2];

         let cmyk = new CMYKColor();
         cmyk.cyan = CMYKArray[i][0];
         cmyk.magenta = CMYKArray[i][1];
         cmyk.yellow = CMYKArray[i][2];
         cmyk.black = CMYKArray[i][3];

         colors[i] = [rgb, cmyk];
      }
      return colors;
   };

   //take a single RGBColor and an array of corresponding RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
   //returns the index in the array if it finds a match, otherwise returns -1
   tasks.matchRGB = function (color, matchArray) {
      //compares a single color RGB color against RGB colors in [[RGB],[CMYK]] array
      for (let i = 0; i < matchArray.length; i++) {
         if (
            Math.abs(color.red - matchArray[i][0].red) < 1 &&
            Math.abs(color.green - matchArray[i][0].green) < 1 &&
            Math.abs(color.blue - matchArray[i][0].blue) < 1
         ) {
            //can't do equality because it adds very small decimals
            return i;
         }
      }
      return -1;
   };

   //take a single RGBColor and an array of corresponding RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
   //returns the index in the array if it finds a match, otherwise returns -1
   tasks.matchColorsRGB = function (color1, color2) {
      //compares two colors to see if they match
      if (
         Math.abs(color1.red - color2.red) < 1 &&
         Math.abs(color1.green - color2.green) < 1 &&
         Math.abs(color1.blue - color2.blue) < 1
      ) {
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
      if (
         Math.abs(color1.cyan - color2.cyan) < 1 &&
         Math.abs(color1.magenta - color2.magenta) < 1 &&
         Math.abs(color1.yellow - color2.yellow) < 1 &&
         Math.abs(color1.black - color2.black) < 1
      ) {
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
      let colorIndex = new Array(pathItems.length);
      for (i = 0; i < pathItems.length; i++) {
         let itemColor = pathItems[i].fillColor;
         colorIndex[i] = tasks.matchRGB(itemColor, matchArray);
      }
      return colorIndex;
   };




   //takes a doc, collection of pathItems, an array of specified colors and an array of colorIndices
   //converts the fill colors to the indexed CMYK colors and adds a text box with the unmatched colors
   //Note that this only makes sense if you've previously indexed the same path items and haven't shifted their positions in the pathItems array
   tasks.convertToCMYK = function (doc, pathItems, colorArray, colorIndex) {
      let unmatchedColors = [];
      for (i = 0; i < pathItems.length; i++) {
         if (colorIndex[i] >= 0 && colorIndex[i] < colorArray.length)
            pathItems[i].fillColor = colorArray[colorIndex[i]][1];
         else {
            let unmatchedColor =
               "(" +
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
         alert(
            "One or more colors don't match the brand palette and weren't converted."
         );
         unmatchedColors = tasks.unique(unmatchedColors);
         let unmatchedString = "Unconverted colors:";
         for (let i = 0; i < unmatchedColors.length; i++) {
            unmatchedString = unmatchedString + "\n" + unmatchedColors[i];
         }
         let errorMsgPos = [Infinity, Infinity]; //gets the bottom left of all the artboards
         for (let i = 0; i < doc.artboards.length; i++) {
            let rect = doc.artboards[i].artboardRect;
            if (rect[0] < errorMsgPos[0]) errorMsgPos[0] = rect[0];
            if (rect[3] < errorMsgPos[1]) errorMsgPos[1] = rect[3];
         }
         errorMsgPos[1] = errorMsgPos[1] - 20;

         tasks.createTextFrame(doc, unmatchedString, errorMsgPos, 18);
      }
   };

   //takes an array
   //returns a sorted array with only unique elements
   tasks.unique = function (a) {
      let sorted;
      let uniq;
      if (a.length > 0) {
         sorted = a.sort();
         uniq = [sorted[0]];
         for (let i = 1; i < sorted.length; i++) {
            if (sorted[i] != sorted[i - 1]) uniq.push(sorted[i]);
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
   } catch (e) {
      alert(
         "Issue with layer hiding the Guidelines layer (do not change name from exactly Guidelines).",
         e.message
      );
   }
   /*****************************
   create export folder if needed
   ******************************/
   try {
      // Core folder
      new Folder(`${sourceDoc.path}/${coreName}`).create();
      new Folder(`${sourceDoc.path}/${coreName}/${epsName}`).create();
      new Folder(`${sourceDoc.path}/${coreName}/${jpgName}`).create();
      new Folder(`${sourceDoc.path}/${coreName}/${pngName}`).create();
      new Folder(`${sourceDoc.path}/${coreName}/${svgName}`).create();
      // Expressive folder
      new Folder(`${sourceDoc.path}/${expressiveName}`).create();
      new Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`).create();
      new Folder(`${sourceDoc.path}/${expressiveName}/${jpgName}`).create();
      new Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`).create();
      new Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`).create();
   } catch (e) {
      alert(
         "Issues with creating setup folders.",
         e.message
      );
   }
   /******************
   set up artboards
   ******************/


   //if there are two artboards at 256x256, create the new third masthead artboard
   if (
      sourceDoc.artboards.length == 2 &&
      sourceDoc.artboards[0].artboardRect[2] -
      sourceDoc.artboards[0].artboardRect[0] ==
      256 &&
      sourceDoc.artboards[0].artboardRect[1] -
      sourceDoc.artboards[0].artboardRect[3] ==
      256
   ) {
      // alert("More than 2 artboards detected!");
      let firstRect = sourceDoc.artboards[0].artboardRect;
      sourceDoc.artboards.add(
         // CSTasks.newRect(firstRect[1] * 2.5 + gutter, firstRect[2], 2400, 256)
         // CSTasks.newRect(firstRect[1] * 0.5, firstRect[2], 2400, 256)

         CSTasks.newRect(firstRect[1], firstRect[2] + 128, 2400, 256)
      );
   }

   //if the masthead artboard is present, check if rebuilding or just exporting
   else if (
      sourceDoc.artboards.length == 3 &&
      sourceDoc.artboards[1].artboardRect[1] -
      sourceDoc.artboards[1].artboardRect[3] ==
      256
   ) {
      rebuild = confirm(
         "It looks like your artwork already exists. This script will rebuild the masthead and export various EPS and PNG versions. Do you want to proceed?"
      );
      if (rebuild) CSTasks.clearArtboard(sourceDoc, 1);
      else return;
   }

   //otherwise abort
   else {
      alert("Please try again with 2 artboards that are 256x256px.");
      return;
   }
   //select the contents on artboard 0
   let sel = CSTasks.selectContentsOnArtboard(sourceDoc, 0);

   // make sure all colors are RGB, equivalent of Edit > Colors > Convert to RGB
   app.executeMenuCommand('Colors9');

   if (sel.length == 0) {
      //if nothing is in the artboard
      alert("Please try again with artwork on the main 256x256 artboard.");
      return;
   }

   let colors = CSTasks.initializeColors(RGBColorElements, CMYKColorElements); //initialize the colors from the brand palette
   let iconGroup = CSTasks.createGroup(sourceDoc, sel); //group the selection (easier to work with)
   let iconOffset = CSTasks.getOffset(
      iconGroup.position,
      CSTasks.getArtboardCorner(sourceDoc.artboards[0])
   );

   /********************************
   Create new artboard with masthead
   *********************************/

   //place icon on masthead
   /*@ts-ignore*/
   let mast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
   let mastPos = [
      sourceDoc.artboards[2].artboardRect[0] + iconOffset[0],
      sourceDoc.artboards[2].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(mast, mastPos);

   //request a name for the icon, and place that as text on the masthead artboard
   let appName = prompt("What name do you want to put in the masthead?");

   let textRef = sourceDoc.textFrames.add();
   textRef.contents = appName;
   textRef.textRange.characterAttributes.size = 178;
   CSTasks.setFont(textRef, desiredFont);

   //vertically align the baseline to be 64 px above the botom of the artboard
   let bottomEdge =
      sourceDoc.artboards[2].artboardRect[3] +
      0.25 * sourceDoc.artboards[0].artboardRect[2] -
      sourceDoc.artboards[0].artboardRect[0]; //64px (0.25*256px) above the bottom edge of the artboard
   let vOffset = CSTasks.getOffset(textRef.anchor, [0, bottomEdge]);
   textRef.translate(0, -vOffset[1]);

   //create an outline of the text
   let textGroup = textRef.createOutline();

   //horizontally align the left edge of the text to be 96px to the right of the edge
   let rightEdge =
      mast.position[0] +
      mast.width +
      0.375 * sourceDoc.artboards[0].artboardRect[2] -
      sourceDoc.artboards[0].artboardRect[0]; //96px (0.375*256px) right of the icon
   let hOffset = CSTasks.getOffset(textGroup.position, [rightEdge, 0]);
   textGroup.translate(-hOffset[0], 0);

   //resize the artboard to be only a little wider than the text
   let leftMargin = mast.position[0] - sourceDoc.artboards[2].artboardRect[0];
   let newWidth =
      textGroup.position[0] +
      textGroup.width -
      sourceDoc.artboards[2].artboardRect[0] +
      leftMargin;
   let resizedRect = CSTasks.newRect(
      sourceDoc.artboards[2].artboardRect[0],
      -sourceDoc.artboards[2].artboardRect[1],
      newWidth,
      256
   );
   sourceDoc.artboards[2].artboardRect = resizedRect;

   //get the text offset for exporting
   let mastTextOffset = CSTasks.getOffset(
      textGroup.position,
      CSTasks.getArtboardCorner(sourceDoc.artboards[2])
   );



   /*********************************************************************
   RGB export (EPS, PNGs at multiple sizes, inactive EPS and inverse EPS)
   **********************************************************************/
   let rgbDoc = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      0,
      DocumentColorSpace.RGB
   );

   rgbDoc.swatches.removeAll();

   let rgbGroup = iconGroup.duplicate(
      rgbDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   let rgbLoc = [
      rgbDoc.artboards[0].artboardRect[0] + iconOffset[0],
      rgbDoc.artboards[0].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(rgbGroup, rgbLoc);

   CSTasks.ungroupOnce(rgbGroup);

   //save a master PNG
   let masterStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, masterStartWidth, exportSizes[2]);
   }

   //save a master SVG 
   let svgMasterCoreStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgMasterCoreStartWidth, exportSizes[2]);
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
   let filename = `/${iconFilename}_${coreName}_${rgbName}.eps`;
   let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + filename);
   let rgbSaveOpts = new EPSSaveOptions();
   /*@ts-ignore*/
   rgbSaveOpts.cmykPostScript = false;
   rgbDoc.saveAs(destFile, rgbSaveOpts);

   //index the RGB colors for conversion to CMYK. An inelegant location.
   let colorIndex = CSTasks.indexRGBColors(rgbDoc.pathItems, colors);
   //convert violet to white and save as EPS
   CSTasks.convertColorRGB(
      rgbDoc.pathItems,
      colors[violetIndex][0],
      colors[whiteIndex][0]
   );

   let inverseFilename = `/${iconFilename}_${inverseName}_${rgbName}.eps`;
   let inverseFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + inverseFilename);
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

   let inactiveFilename = `/${iconFilename}_${inactiveName}_${rgbName}.eps`;
   let inactiveFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + inactiveFilename);
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
   let cmykDoc = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      0,
      DocumentColorSpace.CMYK
   );
   cmykDoc.swatches.removeAll();

   //need to reverse the order of copying the group to get the right color ordering
   let cmykGroup = iconGroup.duplicate(
      cmykDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATBEGINNING
   );
   let cmykLoc = [
      cmykDoc.artboards[0].artboardRect[0] + iconOffset[0],
      cmykDoc.artboards[0].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(cmykGroup, cmykLoc);
   CSTasks.ungroupOnce(cmykGroup);

   CSTasks.convertToCMYK(cmykDoc, cmykDoc.pathItems, colors, colorIndex);

   //save EPS into the export folder
   let cmykFilename = `/${iconFilename}_${coreName}_${cmykName}.eps`;
   let cmykDestFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + cmykFilename);
   let cmykSaveOpts = new EPSSaveOptions();
   cmykDoc.saveAs(cmykDestFile, cmykSaveOpts);

   //convert violet to white and save as EPS
   CSTasks.convertColorCMYK(
      cmykDoc.pathItems,
      colors[violetIndex][1],
      colors[whiteIndex][1]
   );

   let cmykInverseFilename = `/${iconFilename}_${inverseName}_${cmykName}.eps`;
   let cmykInverseFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + cmykInverseFilename);
   cmykDoc.saveAs(cmykInverseFile, rgbSaveOpts);

   //close and clean up
   cmykDoc.close(SaveOptions.DONOTSAVECHANGES);
   cmykDoc = null;


   /********************
   Masthead export core CMYK (EPS)
   ********************/
   //open a new doc and copy and position the icon and the masthead text
   let mastCMYKDoc = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      1,
      DocumentColorSpace.CMYK
   );
   mastCMYKDoc.swatches.removeAll();

   let mastCMYKGroup = iconGroup.duplicate(
      mastCMYKDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   let mastCMYKLoc = [
      mastCMYKDoc.artboards[0].artboardRect[0] + iconOffset[0],
      mastCMYKDoc.artboards[0].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(mastCMYKGroup, mastCMYKLoc);
   CSTasks.ungroupOnce(mastCMYKGroup);

   let mastCMYKText = textGroup.duplicate(
      mastCMYKDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   let mastCMYKTextLoc = [
      mastCMYKDoc.artboards[0].artboardRect[0] + mastTextOffset[0],
      mastCMYKDoc.artboards[0].artboardRect[1] + mastTextOffset[1],
   ];
   CSTasks.translateObjectTo(mastCMYKText, mastCMYKTextLoc);

   //save CMYK EPS into the export folder
   let mastCMYKFilename = `/${iconFilename}_${mastheadName}_${cmykName}.eps`;
   let mastCMYKDestFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + mastCMYKFilename);
   let mastCMYKSaveOpts = new EPSSaveOptions();
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
   let mastDoc = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      1,
      DocumentColorSpace.RGB
   );
   mastDoc.swatches.removeAll();

   let mastGroup = iconGroup.duplicate(
      mastDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   let mastLoc = [
      mastDoc.artboards[0].artboardRect[0] + iconOffset[0],
      mastDoc.artboards[0].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(mastGroup, mastLoc);
   CSTasks.ungroupOnce(mastGroup);

   let mastText = textGroup.duplicate(
      mastDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   let mastTextLoc = [
      mastDoc.artboards[0].artboardRect[0] + mastTextOffset[0],
      mastDoc.artboards[0].artboardRect[1] + mastTextOffset[1],
   ];
   CSTasks.translateObjectTo(mastText, mastTextLoc);

   //save RGB EPS into the export folder
   let mastFilename = `/${iconFilename}_${mastheadName}_${rgbName}.eps`;
   let mastDestFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + mastFilename);
   let mastSaveOpts = new EPSSaveOptions();
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
   if (
      sourceDoc.artboards.length == 3 &&
      sourceDoc.artboards[1].artboardRect[2] -
      sourceDoc.artboards[1].artboardRect[0] ==
      256 &&
      sourceDoc.artboards[1].artboardRect[1] -
      sourceDoc.artboards[1].artboardRect[3] ==
      256
   ) {
      // IF there are already  3 artboards. Add a 4th one.
      let firstRect = sourceDoc.artboards[1].artboardRect;
      sourceDoc.artboards.add(
         // this fires but then gets replaced further down
         CSTasks.newRect(firstRect[1], firstRect[2] + 128, 1024, 512)
      );
   }

   //if the masthead artboard is present, check if rebuilding or just exporting
   else if (
      sourceDoc.artboards.length == 3 &&
      sourceDoc.artboards[1].artboardRect[1] -
      sourceDoc.artboards[1].artboardRect[3] ==
      512
   ) {
      rebuild = confirm(
         "It looks like your artwork already exists. This script will rebuild the masthead and export various EPS and PNG versions. Do you want to proceed?"
      );
      if (rebuild) CSTasks.clearArtboard(sourceDoc, 3);
      else return;
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
   let sel = CSTasks.selectContentsOnArtboard(sourceDoc, 1);


   // make sure all colors are RGB, equivalent of Edit > Colors > Convert to RGB
   app.executeMenuCommand('Colors9');

   if (sel.length == 0) {
      //if nothing is in the artboard
      alert("Please try again with artwork on the main second 256x256 artboard.");
      return;
   }

   let colors = CSTasks.initializeColors(RGBColorElements, CMYKColorElements); //initialize the colors from the brand palette
   let iconGroup = CSTasks.createGroup(sourceDoc, sel); //group the selection (easier to work with)
   let iconOffset = CSTasks.getOffset(
      iconGroup.position,
      CSTasks.getArtboardCorner(sourceDoc.artboards[1])
   );

   /********************************
   Create new expressive artboard with masthead and text
   *********************************/



   /*@ts-ignore*/
   let mastBannerIconOnText = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);

   let mastBannerIconOnTextPos = [
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
   let getArtLayer = sourceDoc.layers.getByName('Art');
   let landingZoneSquare = getArtLayer.pathItems.rectangle(
      -844,
      571,
      460,
      460);


   function placeIconMasthead1Correctly(mastBannerIconOnText, maxSize) {
      let setLandingZoneSquareColor = new RGBColor();
      setLandingZoneSquareColor.red = 12;
      setLandingZoneSquareColor.green = 28;
      setLandingZoneSquareColor.blue = 151;
      landingZoneSquare.filled = true;
      landingZoneSquare.fillColor = setLandingZoneSquareColor;
      landingZoneSquare.name = "LandingZone"
      /*@ts-ignore*/
      landingZoneSquare.move(getArtLayer, ElementPlacement.PLACEATEND);

      // start moving expressive icon into our new square landing zone
      let placedmastBannerIconOnText = mastBannerIconOnText;
      let landingZone = sourceDoc.pathItems.getByName("LandingZone");
      let preferredWidth = (230);
      let preferredHeight = (460);
      // do the width
      let widthRatio = (preferredWidth / placedmastBannerIconOnText.width) * 100;
      if (placedmastBannerIconOnText.width != preferredWidth) {
         placedmastBannerIconOnText.resize(widthRatio, widthRatio);
      }
      // now do the height
      let heightRatio = (preferredHeight / placedmastBannerIconOnText.height) * 100;
      if (placedmastBannerIconOnText.height != preferredHeight) {
         placedmastBannerIconOnText.resize(heightRatio, heightRatio);
      }
      // now let's center the art on the landing zone
      let centerArt = [placedmastBannerIconOnText.left + (placedmastBannerIconOnText.width / 2), placedmastBannerIconOnText.top + (placedmastBannerIconOnText.height / 2)];
      let centerLz = [landingZone.left + (landingZone.width / 2), landingZone.top + (landingZone.height / 2)];
      placedmastBannerIconOnText.translate(centerLz[0] - centerArt[0], centerLz[1] - centerArt[1]);

      // need another centered proportioning to fix it exactly in correct position
      let W = mastBannerIconOnText.width,
         H = mastBannerIconOnText.height,
         MW = maxSize.W,
         MH = maxSize.H,
         factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
      mastBannerIconOnText.resize(factor, factor);
   }
   placeIconMasthead1Correctly(mastBannerIconOnText, { W: 460, H: 460 });

   // delete the landing zone
   // let docSelected = sourceDoc.layers.getByName('Art').groupItems.getByName("LandingZone");
   //landingZoneSquare.remove();

   // mastBannerIconOnText.width = 460;
   // mastBannerIconOnText.height = 460;



   // new purple bg
   // Add new layer above Guidelines and fill white
   let myMainArtworkLayer = sourceDoc.layers.getByName('Art');
   let myMainPurpleBgLayer = sourceDoc.layers.add();
   myMainPurpleBgLayer.name = "Main_Purple_BG_layer";
   let GetMyMainPurpleBgLayer = sourceDoc.layers.getByName('Main_Purple_BG_layer');
   // mastDoc.activeLayer = GetMyMainPurpleBgLayer;
   // mastDoc.activeLayer.hasSelectedArtwork = true;
   let mainRect = GetMyMainPurpleBgLayer.pathItems.rectangle(
      -784,
      0,
      1024,
      512);
   let setMainVioletBgColor = new RGBColor();
   setMainVioletBgColor.red = 72;
   setMainVioletBgColor.green = 8;
   setMainVioletBgColor.blue = 111;
   mainRect.filled = true;
   mainRect.fillColor = setMainVioletBgColor;
   /*@ts-ignore*/
   GetMyMainPurpleBgLayer.move(myMainArtworkLayer, ElementPlacement.PLACEATEND);

   let rectRef = sourceDoc.pathItems.rectangle(-850, -800, 400, 300);
   let setTextBoxBgColor = new RGBColor();
   setTextBoxBgColor.red = 141;
   setTextBoxBgColor.green = 141;
   setTextBoxBgColor.blue = 141;
   rectRef.filled = true;
   rectRef.fillColor = setTextBoxBgColor;

   // svg wtw logo for new purple masthead

   let imagePlacedItem = myMainArtworkLayer.placedItems.add();
   let svgFile = File(`${sourceDoc.path}/../images/wtw_logo.ai`);
   imagePlacedItem.file = svgFile;
   imagePlacedItem.top = -1188;
   imagePlacedItem.left = 62;
   /*@ts-ignore*/
   // svgFile.embed();



   //request a name for the icon, and place that as text on the masthead artboard
   let appName = prompt("What name do you want to put in second the masthead?");

   let textRef = sourceDoc.textFrames.add();

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
   let textGroup = textRef.createOutline();

   //horizontally align the left edge of the text to be 96px to the right of the edge
   let rightEdge = 64;
   let hOffset = CSTasks.getOffset(textGroup.position, [rightEdge, 0]);
   textGroup.translate(-hOffset[0], 0);


   /*@ts-ignore*/
   let mast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);

   //resize the artboard to be only a little wider than the text
   let leftMargin = mast.position[0] - sourceDoc.artboards[3].artboardRect[0];
   let newWidth =
      textGroup.position[0] +
      textGroup.width -
      sourceDoc.artboards[3].artboardRect[0] +
      leftMargin;
   let resizedRect = CSTasks.newRect(
      sourceDoc.artboards[3].artboardRect[0],
      -sourceDoc.artboards[3].artboardRect[1],
      1024,
      512
   );
   sourceDoc.artboards[3].artboardRect = resizedRect;























































   /******************
   Set up purple artboard 4 in main file
   ******************/


   //if there are 3 artboards, create the new fourth masthead artboard

   // creat last artboard in file
   let FourthMainArtboardFirstRect = sourceDoc.artboards[1].artboardRect;
   sourceDoc.artboards.add(
      // this fires but then gets replaced further down
      CSTasks.newRect(FourthMainArtboardFirstRect[1], FourthMainArtboardFirstRect[2] + 772, 800, 400)
   );

   /* try {
      sourceDoc.artboards.setActiveArtboardIndex(3);//change which artboard you want to crop
      sourceDoc.artboards[3].artboardRect = new Rect(0, 0, 1024, 512);
   } catch (error) {
      alert(error);
   } */
   //select the contents on artboard 1
   let selFourthBanner = CSTasks.selectContentsOnArtboard(sourceDoc, 1);


   // make sure all colors are RGB, equivalent of Edit > Colors > Convert to RGB
   app.executeMenuCommand('Colors9');

   if (selFourthBanner.length == 0) {
      //if nothing is in the artboard
      alert("Please try again with artwork on the main second 256x256 artboard.");
      return;
   }


   /********************************
   Add elements to new fourth artboard with masthead
   *********************************/

   //place icon on masthead
   /*@ts-ignore*/
   let fourthBannerMast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
   let fourthBannerMastPos = [
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
   let getArtLayer2 = sourceDoc.layers.getByName('Art');
   let landingZoneSquare2 = getArtLayer2.pathItems.rectangle(
      -1472,
      444,
      360,
      360);


   function placeIconMasthead1Correctly2(fourthBannerMast, maxSize) {
      let setLandingZoneSquareColor = new RGBColor();
      setLandingZoneSquareColor.red = 121;
      setLandingZoneSquareColor.green = 128;
      setLandingZoneSquareColor.blue = 131;
      landingZoneSquare2.filled = true;
      landingZoneSquare2.fillColor = setLandingZoneSquareColor;
      landingZoneSquare2.name = "LandingZone2"
      /*@ts-ignore*/
      landingZoneSquare2.move(getArtLayer2, ElementPlacement.PLACEATEND);

      // start moving expressive icon into our new square landing zone
      let placedfourthBannerMast = fourthBannerMast;
      let landingZone = sourceDoc.pathItems.getByName("LandingZone2");
      let preferredWidth = (360);
      let preferredHeight = (360);
      // do the width
      let widthRatio = (preferredWidth / placedfourthBannerMast.width) * 100;
      if (placedfourthBannerMast.width != preferredWidth) {
         placedfourthBannerMast.resize(widthRatio, widthRatio);
      }
      // now do the height
      let heightRatio = (preferredHeight / placedfourthBannerMast.height) * 100;
      if (placedfourthBannerMast.height != preferredHeight) {
         placedfourthBannerMast.resize(heightRatio, heightRatio);
      }
      // now let's center the art on the landing zone
      let centerArt = [placedfourthBannerMast.left + (placedfourthBannerMast.width / 2), placedfourthBannerMast.top + (placedfourthBannerMast.height / 2)];
      let centerLz = [landingZone.left + (landingZone.width / 2), landingZone.top + (landingZone.height / 2)];
      placedfourthBannerMast.translate(centerLz[0] - centerArt[0], centerLz[1] - centerArt[1]);

      // need another centered proportioning to fix it exactly in correct position
      let W = fourthBannerMast.width,
         H = fourthBannerMast.height,
         MW = maxSize.W,
         MH = maxSize.H,
         factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
      fourthBannerMast.resize(factor, factor);
   }
   placeIconMasthead1Correctly2(fourthBannerMast, { W: 360, H: 360 });

   // delete the landing zone
   // let docSelected = sourceDoc.layers.getByName('Art').groupItems.getByName("LandingZone");
   //landingZoneSquare.remove();
























   // fourthBannerMast.width = 360;
   // fourthBannerMast.height = 360;

   // new purple bg
   // Add new layer above Guidelines and fill white
   let secondMainArtworkLayer = sourceDoc.layers.getByName('Art');
   let secondMainPurpleBgLayer = sourceDoc.layers.add();
   secondMainPurpleBgLayer.name = "Main_Purple_BG_layer_two";
   let getsecondMainPurpleBgLayer = sourceDoc.layers.getByName('Main_Purple_BG_layer_two');
   // mastDoc.activeLayer = getsecondMainPurpleBgLayer;
   // mastDoc.activeLayer.hasSelectedArtwork = true;
   let seondMainRect = getsecondMainPurpleBgLayer.pathItems.rectangle(
      -1428,
      0,
      800,
      400);
   let secondMainVioletBgColor = new RGBColor();
   secondMainVioletBgColor.red = 72;
   secondMainVioletBgColor.green = 8;
   secondMainVioletBgColor.blue = 111;
   seondMainRect.filled = true;
   seondMainRect.fillColor = secondMainVioletBgColor;
   /*@ts-ignore*/
   getsecondMainPurpleBgLayer.move(secondMainArtworkLayer, ElementPlacement.PLACEATEND);

   /*@ts-ignore*/
   // svgFile.embed();
   let secondResizedRect = CSTasks.newRect(
      sourceDoc.artboards[4].artboardRect[0],
      -sourceDoc.artboards[4].artboardRect[1],
      800,
      400
   );
   sourceDoc.artboards[4].artboardRect = secondResizedRect;








































   /*********************************************************************
   RGB export (EPS, PNGs at multiple sizes, inactive EPS and inverse EPS)
   **********************************************************************/
   let rgbDoc = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      1,
      DocumentColorSpace.RGB
   );

   rgbDoc.swatches.removeAll();

   let rgbGroup = iconGroup.duplicate(
      rgbDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   let rgbLoc = [
      rgbDoc.artboards[0].artboardRect[0] + iconOffset[0],
      rgbDoc.artboards[0].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(rgbGroup, rgbLoc);

   CSTasks.ungroupOnce(rgbGroup);

   //save a master PNG
   let masterStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, masterStartWidth, exportSizes[2]);
   }

   //save a master SVG
   let svgMasterCoreStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgMasterCoreStartWidth, exportSizes[2]);
   }

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
   let filename = `/${iconFilename}_${expressiveName}_${rgbName}.eps`;
   let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + filename);
   let rgbSaveOpts = new EPSSaveOptions();
   /*@ts-ignore*/
   rgbSaveOpts.cmykPostScript = false;
   rgbDoc.saveAs(destFile, rgbSaveOpts);

   //index the RGB colors for conversion to CMYK. An inelegant location.
   let colorIndex = CSTasks.indexRGBColors(rgbDoc.pathItems, colors);
   //convert violet to white and save as EPS
   CSTasks.convertColorRGB(
      rgbDoc.pathItems,
      colors[violetIndex][0],
      colors[whiteIndex][0]
   );

   let inverseFilename = `/${iconFilename}_${expressiveName}_${inverseName}_${rgbName}.eps`;
   let inverseFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + inverseFilename);
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

   let inactiveFilename = `/${iconFilename}_${expressiveName}_${inactiveName}_${rgbName}.eps`;
   let inactiveFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + inactiveFilename);
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
   let cmykDoc = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      0,
      DocumentColorSpace.CMYK
   );
   cmykDoc.swatches.removeAll();

   //need to reverse the order of copying the group to get the right color ordering
   let cmykGroup = iconGroup.duplicate(
      cmykDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATBEGINNING
   );
   let cmykLoc = [
      cmykDoc.artboards[0].artboardRect[0] + iconOffset[0],
      cmykDoc.artboards[0].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(cmykGroup, cmykLoc);
   CSTasks.ungroupOnce(cmykGroup);

   CSTasks.convertToCMYK(cmykDoc, cmykDoc.pathItems, colors, colorIndex);

   //save EPS into the export folder
   let cmykFilename = `/${iconFilename}_${expressiveName}_${cmykName}.eps`;
   let cmykDestFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + cmykFilename);
   let cmykSaveOpts = new EPSSaveOptions();
   cmykDoc.saveAs(cmykDestFile, cmykSaveOpts);

   //convert violet to white and save as EPS
   CSTasks.convertColorCMYK(
      cmykDoc.pathItems,
      colors[violetIndex][1],
      colors[whiteIndex][1]
   );

   let cmykInverseFilename = `/${iconFilename}_${expressiveName}_${inverseName}_${cmykName}.eps`;
   let cmykInverseFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + cmykInverseFilename);
   cmykDoc.saveAs(cmykInverseFile, rgbSaveOpts);

   //close and clean up
   cmykDoc.close(SaveOptions.DONOTSAVECHANGES);
   cmykDoc = null;

   /********************
   Purple Masthead with text export 
   ********************/
   //open a new doc and copy and position the icon and the masthead text
   // duplication did not work as expected here. I have used a less elegant solution whereby I recreated the purple banner instead of copying it.
   let mastDoc = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      3,
      DocumentColorSpace.RGB
   );
   mastDoc.swatches.removeAll();

   let mastGroup = iconGroup.duplicate(
      mastDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   // new icon width in rebrand
   // mastGroup.width = 460;
   // mastGroup.height = 460;

   // new icon position
   let mastLoc = [
      mastDoc.artboards[0].artboardRect[0] + 576,
      mastDoc.artboards[0].artboardRect[1] - 62,
   ];
   CSTasks.translateObjectTo(mastGroup, mastLoc);



















   /********************************
   Custom function to create a landing square to place the icon correctly
   Some icons have width or height less than 256 so it needed special centering geometrically
   you can see the landing zone square by changing fill to true and uncommenting color
   *********************************/

   // create a landing zone square to place icon inside
   //moved it outside the function itself so we can delete it after so it doesn't get exported
   // let getArtLayer2 = sourceDoc.layers.getByName('Art');
   // let landingZoneSquare2 = getArtLayer2.pathItems.rectangle(
   //    -1472,
   //    444,
   //    360,
   //    360);


   // function placeIconMasthead1Correctly2(fourthBannerMast, maxSize) {
   //    let setLandingZoneSquareColor = new RGBColor();
   //    setLandingZoneSquareColor.red = 121;
   //    setLandingZoneSquareColor.green = 128;
   //    setLandingZoneSquareColor.blue = 131;
   //    landingZoneSquare2.filled = true;
   //    landingZoneSquare2.fillColor = setLandingZoneSquareColor;
   //    landingZoneSquare2.name = "LandingZone2"
   //    /*@ts-ignore*/
   //    landingZoneSquare2.move(getArtLayer2, ElementPlacement.PLACEATEND);

   //    // start moving expressive icon into our new square landing zone
   //    let placedfourthBannerMast = fourthBannerMast;
   //    let landingZone = sourceDoc.pathItems.getByName("LandingZone2");
   //    let preferredWidth = (360);
   //    let preferredHeight = (360);
   //    // do the width
   //    let widthRatio = (preferredWidth / placedfourthBannerMast.width) * 100;
   //    if (placedfourthBannerMast.width != preferredWidth) {
   //       placedfourthBannerMast.resize(widthRatio, widthRatio);
   //    }
   //    // now do the height
   //    let heightRatio = (preferredHeight / placedfourthBannerMast.height) * 100;
   //    if (placedfourthBannerMast.height != preferredHeight) {
   //       placedfourthBannerMast.resize(heightRatio, heightRatio);
   //    }
   //    // now let's center the art on the landing zone
   //    let centerArt = [placedfourthBannerMast.left + (placedfourthBannerMast.width / 2), placedfourthBannerMast.top + (placedfourthBannerMast.height / 2)];
   //    let centerLz = [landingZone.left + (landingZone.width / 2), landingZone.top + (landingZone.height / 2)];
   //    placedfourthBannerMast.translate(centerLz[0] - centerArt[0], centerLz[1] - centerArt[1]);

   //    // need another centered proportioning to fix it exactly in correct position
   //    let W = fourthBannerMast.width,
   //       H = fourthBannerMast.height,
   //       MW = maxSize.W,
   //       MH = maxSize.H,
   //       factor = W / H > MW / MH ? MW / W * 100 : MH / H * 100;
   //    fourthBannerMast.resize(factor, factor);
   // }
   // placeIconMasthead1Correctly2(fourthBannerMast, { W: 360, H: 360 });

   // delete the landing zone
   // let docSelected = sourceDoc.layers.getByName('Art').groupItems.getByName("LandingZone");
   //landingZoneSquare.remove();















   CSTasks.ungroupOnce(mastGroup);

   let mastText = textGroup.duplicate(
      mastDoc.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   // text position
   let mastTextLoc = [
      mastDoc.artboards[0].artboardRect[0] + 62,
      mastDoc.artboards[0].artboardRect[1] - 62,
   ];
   CSTasks.translateObjectTo(mastText, mastTextLoc);


   // add new style purple banner elements
   let myMainArtworkLayerMastDoc = mastDoc.layers.getByName('Layer 1');
   let myMainPurpleBgLayerMastDoc = mastDoc.layers.add();
   myMainPurpleBgLayerMastDoc.name = "Main_Purple_BG_layer";
   let GetMyMainPurpleBgLayerMastDoc = mastDoc.layers.getByName('Main_Purple_BG_layer');
   // mastDoc.activeLayer = GetMyMainPurpleBgLayerMastDoc;
   // mastDoc.activeLayer.hasSelectedArtwork = true;
   let mainRectMastDoc = GetMyMainPurpleBgLayerMastDoc.pathItems.rectangle(
      -781,
      0,
      1024,
      512);
   let setMainVioletBgColorMastDoc = new RGBColor();
   setMainVioletBgColorMastDoc.red = 72;
   setMainVioletBgColorMastDoc.green = 8;
   setMainVioletBgColorMastDoc.blue = 111;
   mainRectMastDoc.filled = true;
   mainRectMastDoc.fillColor = setMainVioletBgColorMastDoc;
   /*@ts-ignore*/
   GetMyMainPurpleBgLayerMastDoc.move(myMainArtworkLayerMastDoc, ElementPlacement.PLACEATEND);

   // svg wtw logo for new purple masthead
   let imagePlacedItemMastDoc = myMainArtworkLayerMastDoc.placedItems.add();
   let svgFileMastDoc = File(`${sourceDoc.path}/../images/wtw_logo.ai`);
   imagePlacedItemMastDoc.file = svgFileMastDoc;
   imagePlacedItemMastDoc.top = -1181;
   imagePlacedItemMastDoc.left = 62;

   // we need to make artboard clipping mask here for the artboard to crop expressive icons correctly.
   let myCroppingLayerMastDoc = mastDoc.layers.add();
   myCroppingLayerMastDoc.name = "crop";
   let GetMyCroppingLayerMastDoc = mastDoc.layers.getByName('crop');
   mastDoc.activeLayer = GetMyCroppingLayerMastDoc;
   mastDoc.activeLayer.hasSelectedArtwork = true;
   // insert clipping rect here
   let mainClipRectMastDoc = GetMyCroppingLayerMastDoc.pathItems.rectangle(
      -781,
      0,
      1024,
      512);
   let setClipBgColorMastDoc = new RGBColor();
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
   // for (let i = 0; i < exportSizes.length; i++) {
   //    let filename = `/${iconFilename}_${expressiveName}_${mastheadName}_${rgbName}_${masthead1}.png`;
   //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
   //    CSTasks.scaleAndExportPNG(mastDoc, destFile, 400, 800);
   // }
   //save a banner SVG 
   // for (let i = 0; i < exportSizes.length; i++) {
   //    let filename = `/${iconFilename}_${expressiveName}_${mastheadName}_${rgbName}_${masthead1}.svg`;
   //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
   //    CSTasks.scaleAndExportSVG(mastDoc, destFile, 400, 800);
   // }
   //save RGB EPS into the export folder
   let mastFilename = `/${iconFilename}_${expressiveName}_${mastheadName}_${rgbName}_${masthead1}.eps`;
   let mastDestFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + mastFilename);
   let mastSaveOpts = new EPSSaveOptions();
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
   let mastDocNoText = CSTasks.duplicateArtboardInNewDoc(
      sourceDoc,
      4,
      DocumentColorSpace.RGB
   );
   mastDocNoText.swatches.removeAll();

   let mastGroupNoText = iconGroup.duplicate(
      mastDocNoText.layers[0],
      /*@ts-ignore*/
      ElementPlacement.PLACEATEND
   );
   // new icon width in rebrand
   mastGroupNoText.width = 360;
   mastGroupNoText.height = 360;

   // new icon position
   let mastLocNoText = [
      mastDocNoText.artboards[0].artboardRect[0] + 444,
      mastDocNoText.artboards[0].artboardRect[1] - 44,
   ];
   CSTasks.translateObjectTo(mastGroupNoText, mastLocNoText);
   CSTasks.ungroupOnce(mastGroupNoText);

   // let mastText2 = textGroup.duplicate(
   //    mastDocNoText.layers[0],
   //    /*@ts-ignore*/
   //    ElementPlacement.PLACEATEND
   // );
   // // text position
   // let mastTextLoc2 = [
   //    mastDocNoText.artboards[0].artboardRect[0] + 62,
   //    mastDocNoText.artboards[0].artboardRect[1] - 62,
   // ];
   // CSTasks.translateObjectTo(mastText2, mastTextLoc2);


   // add new style purple banner elements
   let myMainArtworkLayerMastDocNoText = mastDocNoText.layers.getByName('Layer 1');
   let myMainPurpleBgLayerMastDocNoText = mastDocNoText.layers.add();
   myMainPurpleBgLayerMastDocNoText.name = "Main_Purple_BG_layer";
   let GetMyMainPurpleBgLayerMastDocNoText = mastDocNoText.layers.getByName('Main_Purple_BG_layer');
   // mastDoc.activeLayer = GetMyMainPurpleBgLayerMastDocNoText;
   // mastDoc.activeLayer.hasSelectedArtwork = true;
   let mainRectMastDocNoText = GetMyMainPurpleBgLayerMastDocNoText.pathItems.rectangle(
      -1428,
      0,
      800,
      400);
   let setMainVioletBgColorMastDocNoText = new RGBColor();
   setMainVioletBgColorMastDocNoText.red = 72;
   setMainVioletBgColorMastDocNoText.green = 8;
   setMainVioletBgColorMastDocNoText.blue = 111;
   mainRectMastDocNoText.filled = true;
   mainRectMastDocNoText.fillColor = setMainVioletBgColorMastDocNoText;
   /*@ts-ignore*/
   GetMyMainPurpleBgLayerMastDocNoText.move(myMainArtworkLayerMastDocNoText, ElementPlacement.PLACEATEND);

   // svg wtw logo for new purple masthead
   // let imagePlacedItemMastDocNoText = myMainArtworkLayerMastDocNoText.placedItems.add();
   // let svgFileMastDocNoText = File(`${sourceDoc.path}/../images/wtw_logo.ai`);
   // imagePlacedItemMastDocNoText.file = svgFileMastDocNoText;
   // imagePlacedItemMastDocNoText.top = -1181;
   // imagePlacedItemMastDocNoText.left = 62;
   // UNCOMMENT THIS!
   // we need to make artboard clipping mask here for the artboard to crop expressive icons correctly.
   let myCroppingLayerMastDocNoText = mastDocNoText.layers.add();
   myCroppingLayerMastDocNoText.name = "crop";
   let GetMyCroppingLayerMastDocNoText = mastDocNoText.layers.getByName('crop');
   mastDocNoText.activeLayer = GetMyCroppingLayerMastDocNoText;
   mastDocNoText.activeLayer.hasSelectedArtwork = true;
   // insert clipping rect here
   let mainClipRectMastDocNoText = GetMyCroppingLayerMastDocNoText.pathItems.rectangle(
      -1428,
      0,
      800,
      400);
   let setClipBgColorMastDocNoText = new RGBColor();
   setClipBgColorMastDocNoText.red = 0;
   setClipBgColorMastDocNoText.green = 255;
   setClipBgColorMastDocNoText.blue = 255;
   mainClipRectMastDocNoText.filled = true;
   mainClipRectMastDocNoText.fillColor = setClipBgColorMastDocNoText;
   // select all for clipping here
   mastDocNoText.selectObjectsOnActiveArtboard();
   // clip!
   app.executeMenuCommand('makeMask');


   //save a banner PNG
   // for (let i = 0; i < exportSizes.length; i++) {
   //    let filename = `/${iconFilename}_${expressiveName}_${mastheadName}_${rgbName}_${masthead2}.png`;
   //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
   //    CSTasks.scaleAndExportPNG(mastDocNoText, destFile, 400, 800);
   // }
   //save a banner SVG
   // for (let i = 0; i < exportSizes.length; i++) {
   //    let filename = `/${iconFilename}_${expressiveName}_${mastheadName}_${rgbName}_${masthead2}.svg`;
   //    let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
   //    CSTasks.scaleAndExportSVG(mastDocNoText, destFile, 400, 800);
   // }
   //save RGB EPS into the export folder 
   let mastNoTextFilename = `/${iconFilename}_${expressiveName}_${mastheadName}_${rgbName}_${masthead2}.eps`;
   let mastNoTextDestFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + mastNoTextFilename);
   let mastNoTextSaveOpts = new EPSSaveOptions();
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
   // CSTasks.ungroupOnce(iconGroup);
   // CSTasks.ungroupOnce(mast);
}

mainExpressive();


