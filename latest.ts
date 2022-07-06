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
let sourceDoc = app.activeDocument;
let RGBColorElements = [
   [127, 53, 178], //ultraviolet purple
   [191, 191, 191], //Gray matter light grey
   [201, 0, 172], // Fireworks magenta
   [50, 127, 239], //Stratosphere blue
   [58, 220, 201], // Inifinity turquoise
   [255, 255, 255], // white
];
let CMYKColorElements = [
   [29, 70, 0, 30], //ultraviolet purple
   [0, 0, 0, 25], //Gray matter light grey
   [0, 100, 14, 21], // Fireworks magenta
   [79, 47, 0, 6], //Stratosphere blue
   [74, 0, 9, 14], // Inifinity turquoise  
   [0, 0, 0, 0], // white
];

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

// var artboardRef = sourceDoc.artboards[0];
// var top = artboardRef.artboardRect[1];
// var left = artboardRef.artboardRect[0];
// var width = artboardRef.artboardRect[2] - artboardRef.artboardRect[0];
// var height = artboardRef.artboardRect[1] - artboardRef.artboardRect[3];
// var rect = sourceDoc.pathItems.rectangle(top, left, width, height);
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
      artboardIndex: any,
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
   //returns an aray [x,y] for the offset between the two points
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
         // NOTE: Don't perform the Artboard Creation Work if there are unmatched colors.
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
 * MAIN SCRIPT START
 ***************/

function main() {
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
      // Expressive folder(not in use yet)
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
      let filename = `/${iconFilename}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, masterStartWidth, exportSizes[2]);
   }

   //save a master SVG 
   let svgMasterCoreStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];

   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgMasterCoreStartWidth, exportSizes[2]);
   }

   //save all sizes of PNG into the export folder
   let startWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }
   // non transparent png exports
   let startWidthonFFF =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${rgbName}_${onWhiteName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
      CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile, startWidthonFFF, exportSizes[i]);
   }

   //save all sizes of SVG into the export folder
   let svgCoreStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];

   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgCoreStartWidth, exportSizes[i]);
   }

   //save all sizes of JPEG into the export folder
   let jpegStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}.jpg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${jpgName}`) + filename);
      CSTasks.scaleAndExportJPEG(rgbDoc, destFile, jpegStartWidth, exportSizes[i]);
   }

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
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${inverseName}_${rgbName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

   //save inverse file in all the SVG sizes
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${inverseName}_${rgbName}_${exportSizes[i]}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

   //convert to inactive color (WTW Icon grey at 100% opacity) and save as EPS
   CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 100);

   let inactiveFilename = `/${iconFilename}_${inactiveName}_${rgbName}.eps`;
   let inactiveFile = new File(Folder(`${sourceDoc.path}/${coreName}/${epsName}`) + inactiveFilename);
   rgbDoc.saveAs(inactiveFile, rgbSaveOpts);

   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${inactiveName}_${rgbName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${pngName}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${coreName}_${inactiveName}_${rgbName}_${exportSizes[i]}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

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
   Masthead export (EPS)
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


   /************
PARKING FOR NOW Add white BG and save again, needs to go inside a new layer with the artwork group inside with the rect
************/

   // Add new layer above Guidelines and fill white
   // let myArtworkLayer = sourceDoc.layers.getByName('Artwork');
   // let myWhiteBgLayer = sourceDoc.layers.add();
   // myWhiteBgLayer.name = "White_BG_layer";
   // let GetMyWhiteBgLayer = sourceDoc.layers.getByName('White_BG_layer');
   // // sourceDoc.activeLayer = GetMyWhiteBgLayer;
   // // sourceDoc.activeLayer.hasSelectedArtwork = true;
   // let rect = GetMyWhiteBgLayer.pathItems.rectangle(
   //    0,
   //    0,
   //    256,
   //    256);
   // let setColor = new RGBColor();
   // setColor.red = 191;
   // setColor.green = 191;
   // setColor.blue = 191;
   // rect.filled = true;
   // rect.fillColor = setColor;
   // sourceDoc.selection = null;


   // /*@ts-ignore*/
   // GetMyWhiteBgLayer.move(myArtworkLayer, ElementPlacement.PLACEATEND);

   // sourceDoc.activeLayer = myArtworkLayer;
   // sourceDoc.activeLayer.hasSelectedArtwork = true;

   // let rgbDoc2 = CSTasks.duplicateArtboardInNewDoc(
   //    sourceDoc,
   //    0,
   //    DocumentColorSpace.RGB
   // );
   // rgbDoc2.swatches.removeAll();

   // let rgbGroup2 = iconGroup.duplicate(
   //    rgbDoc2.layers[0],
   //    /*@ts-ignore*/
   //    ElementPlacement.PLACEATEND
   // );
   // let rgbLoc2 = [
   //    rgbDoc2.artboards[0].artboardRect[0] + iconOffset[0],
   //    rgbDoc2.artboards[0].artboardRect[1] + iconOffset[1],
   // ];
   // CSTasks.translateObjectTo(rgbGroup2, rgbLoc2);

   // CSTasks.ungroupOnce(rgbGroup2);

   // //save all sizes of SVG into the export folder
   // let svgCoreStartWidthonFFF =
   //    rgbDoc2.artboards[0].artboardRect[2] - rgbDoc2.artboards[0].artboardRect[0];

   // for (let i = 0; i < exportSizes.length; i++) {

   //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}_onFFF.svg`;
   //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
   //    CSTasks.scaleAndExportSVG(rgbDoc2, destFile, svgCoreStartWidthonFFF, exportSizes[i]);
   // }
   // //close and clean up
   // rgbDoc2.close(SaveOptions.DONOTSAVECHANGES);
   // rgbDoc2 = null;
   // // artworkLayer.locked;



}

main();


// expressive


function mainExpressive() {

   /******************
   set up artboards
   ******************/


   //if there are two artboards at 256x256, create the new third masthead artboard
   if (
      sourceDoc.artboards.length == 3 &&
      sourceDoc.artboards[1].artboardRect[2] -
      sourceDoc.artboards[1].artboardRect[0] ==
      256 &&
      sourceDoc.artboards[1].artboardRect[1] -
      sourceDoc.artboards[1].artboardRect[3] ==
      256
   ) {
      // alert("More than 2 artboards detected!");
      let firstRect = sourceDoc.artboards[1].artboardRect;
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
      alert("Please try again with 2 artboards that are 256x256px.");
      return;
   }
   //select the contents on artboard 1
   let sel = CSTasks.selectContentsOnArtboard(sourceDoc, 1);

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
      CSTasks.getArtboardCorner(sourceDoc.artboards[1])
   );

   /********************************
   Create new artboard with masthead
   *********************************/

   //place icon on masthead
   /*@ts-ignore*/
   let mast = iconGroup.duplicate(iconGroup.layer, ElementPlacement.PLACEATEND);
   let mastPos = [
      sourceDoc.artboards[3].artboardRect[0] + iconOffset[0],
      sourceDoc.artboards[3].artboardRect[1] + iconOffset[1],
   ];
   CSTasks.translateObjectTo(mast, mastPos);

   //request a name for the icon, and place that as text on the masthead artboard
   let appName = prompt("What name do you want to put in second the masthead?");



   // sourceDoc.activeLayer = myArtworkLayer;
   // sourceDoc.activeLayer.hasSelectedArtwork = true;

   // let rgbDoc2 = CSTasks.duplicateArtboardInNewDoc(
   //    sourceDoc,
   //    0,
   //    DocumentColorSpace.RGB
   // );
   // rgbDoc2.swatches.removeAll();

   // let rgbGroup2 = iconGroup.duplicate(
   //    rgbDoc2.layers[0],
   //    /*@ts-ignore*/
   //    ElementPlacement.PLACEATEND
   // );
   // let rgbLoc2 = [
   //    rgbDoc2.artboards[0].artboardRect[0] + iconOffset[0],
   //    rgbDoc2.artboards[0].artboardRect[1] + iconOffset[1],
   // ];
   // CSTasks.translateObjectTo(rgbGroup2, rgbLoc2);

   // CSTasks.ungroupOnce(rgbGroup2);

   // //save all sizes of SVG into the export folder
   // let svgCoreStartWidthonFFF =
   //    rgbDoc2.artboards[0].artboardRect[2] - rgbDoc2.artboards[0].artboardRect[0];

   // for (let i = 0; i < exportSizes.length; i++) {

   //    let filename = `/${iconFilename}_${coreName}_${rgbName}_${exportSizes[i]}_onFFF.svg`;
   //    let destFile = new File(Folder(`${sourceDoc.path}/${coreName}/${svgName}`) + filename);
   //    CSTasks.scaleAndExportSVG(rgbDoc2, destFile, svgCoreStartWidthonFFF, exportSizes[i]);
   // }
   // //close and clean up
   // rgbDoc2.close(SaveOptions.DONOTSAVECHANGES);
   // rgbDoc2 = null;
   // artworkLayer.locked;

   let textRef = sourceDoc.textFrames.add();
   textRef.contents = appName;
   textRef.textRange.characterAttributes.size = 178;
   CSTasks.setFont(textRef, desiredFont);

   //vertically align the baseline to be 64 px above the botom of the artboard
   let bottomEdge =
      sourceDoc.artboards[3].artboardRect[3] +
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
   let leftMargin = mast.position[0] - sourceDoc.artboards[3].artboardRect[0];
   let newWidth =
      textGroup.position[0] +
      textGroup.width -
      sourceDoc.artboards[3].artboardRect[0] +
      leftMargin;
   let resizedRect = CSTasks.newRect(
      sourceDoc.artboards[3].artboardRect[0],
      -sourceDoc.artboards[3].artboardRect[1],
      newWidth,
      256
   );
   sourceDoc.artboards[3].artboardRect = resizedRect;

   //get the text offset for exporting
   let mastTextOffset = CSTasks.getOffset(
      textGroup.position,
      CSTasks.getArtboardCorner(sourceDoc.artboards[3])
   );

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
   let startWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }
   // non transparent png exports
   let startWidthonFFF =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${onWhiteName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
      CSTasks.scaleAndExportNonTransparentPNG(rgbDoc, destFile, startWidthonFFF, exportSizes[i]);
   }

   //save all sizes of SVG into the export folder
   let svgCoreStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[1];

   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${exportSizes[i]}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, svgCoreStartWidth, exportSizes[i]);
   }

   //save all sizes of JPEG into the export folder
   let jpegStartWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[1];
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${rgbName}_${exportSizes[i]}.jpg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${jpgName}`) + filename);
      CSTasks.scaleAndExportJPEG(rgbDoc, destFile, jpegStartWidth, exportSizes[i]);
   }

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
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${inverseName}_${rgbName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

   //save inverse file in all the SVG sizes
   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${inverseName}_${rgbName}_${exportSizes[i]}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

   //convert to inactive color (WTW Icon grey at 100% opacity) and save as EPS
   CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 100);

   let inactiveFilename = `/${iconFilename}_${expressiveName}_${inactiveName}_${rgbName}.eps`;
   let inactiveFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + inactiveFilename);
   rgbDoc.saveAs(inactiveFile, rgbSaveOpts);

   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${inactiveName}_${rgbName}_${exportSizes[i]}.png`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${pngName}`) + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

   for (let i = 0; i < exportSizes.length; i++) {
      let filename = `/${iconFilename}_${expressiveName}_${inactiveName}_${rgbName}_${exportSizes[i]}.svg`;
      let destFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${svgName}`) + filename);
      CSTasks.scaleAndExportSVG(rgbDoc, destFile, startWidth, exportSizes[i]);
   }

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
   Masthead export (EPS)
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



   // new purple bg
   // Add new layer above Guidelines and fill white
   let myArtworkLayer = mastDoc.layers.getByName('Layer 1');
   let myWhiteBgLayer = mastDoc.layers.add();
   myWhiteBgLayer.name = "White_BG_layer";
   let GetMyWhiteBgLayer = mastDoc.layers.getByName('White_BG_layer');
   // mastDoc.activeLayer = GetMyWhiteBgLayer;
   // mastDoc.activeLayer.hasSelectedArtwork = true;
   let rect = GetMyWhiteBgLayer.pathItems.rectangle(
      -784,
      0,
      1024,
      512);
   let setColor = new RGBColor();
   setColor.red = 72;
   setColor.green = 8;
   setColor.blue = 111;
   rect.filled = true;
   rect.fillColor = setColor;
   // mastDoc.selection = null;
   /*@ts-ignore*/
   GetMyWhiteBgLayer.move(myArtworkLayer, ElementPlacement.PLACEATEND);
   // mastDoc.selection = null;



   //save RGB EPS into the export folder
   let mastFilename = `/${iconFilename}_${expressiveName}_${mastheadName}_${rgbName}.eps`;
   let mastDestFile = new File(Folder(`${sourceDoc.path}/${expressiveName}/${epsName}`) + mastFilename);
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


   // create purple bg on last artboard 

}

mainExpressive();


