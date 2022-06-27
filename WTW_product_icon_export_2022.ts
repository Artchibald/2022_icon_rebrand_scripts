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
   [50, 127, 239], //Stratosphere blue
   [58, 220, 201], // Inifity turquoise
   [201, 0, 172], // Fireworks magenta
   [191, 191, 191], //Gray matter light grey
   [128, 128, 128], // Gray matter Dark dark grey
];
let CMYKColorElements = [
   [65, 91, 0, 0], //ultraviolet purple
   [78, 47, 0, 0], //Stratosphere blue
   [53, 0, 34, 0], // Inifity turquoise 
   [16, 96, 0, 0], // Fireworks magenta
   [0, 0, 0, 25], //Gray matter light grey
   [0, 0, 0, 50], // Gray matter Dark dark grey
];

// let desiredFont = "NHaasGroteskTXStd-55Rg";
let desiredFont = "Monaco";
let PNGSizes = [1024, 256, 128, 64, 48, 32]; //sizes to export
let violetIndex = 0; //these are for converting to inverse and inactive versions
let grayIndex = 1;
let whiteIndex = 5;
//loop default 
let i;

/********************************** 
Module for image manipulation tasks 
***********************************/
// let artboardsInfo = app.activeDocument.artboards[1].name;
// alert(artboardsInfo);

if (sourceDoc.artboards[0].name = "Core") {

}

let CSTasks = (function () {
   let tasks: any = {};

   /********************
      POSITION AND MOVEMENT
      ********************/

   //takes an artboard
   //returns its left top corner as an array [x,y]

   tasks.getArtboardCorner = function (artboard) {
      let corner = [sourceDoc.artboards[0].artboardRect[0], artboard.artboardRect[1]];
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
      let k;
      for (k = 0; k < collection.length; k++) {
         collection[k].moveToBeginning(newGroup);
      }
      return newGroup;
   };

   //takes a group
   //ungroups that group at the top layer (no recursion for nested groups)
   tasks.ungroupOnce = function (group) {
      let i;
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
      // broken, not needed if font is found
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
      let i;
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
 ***************/

function main() {


   /*****************************
     create export folder if needed
     ******************************/
   let name = sourceDoc.name.split(".")[0];
   let destFolder = Folder(sourceDoc.path + "/" + name);
   if (!destFolder.exists) destFolder.create();

   /******************
     set up artboards
     ******************/
   let rebuild = true;
   let gutter = 32;

   //if there is one artboard at 256x256, create the new artboard
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
         CSTasks.newRect(firstRect[2] * 2.2 + gutter, firstRect[1], 2400, 256)
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

   //vertically align the baseline to be 64 px above the bottom of the artboard
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

   //create a new document with the artboard and contents from artboard 0
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

   //save all sizes of PNG into the export folder
   let startWidth =
      rgbDoc.artboards[0].artboardRect[2] - rgbDoc.artboards[0].artboardRect[0];
   for (let i = 0; i < PNGSizes.length; i++) {
      let filename = "/" + name + "_Core_RGB_" + PNGSizes[i] + ".png";
      let destFile = new File(destFolder + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, PNGSizes[i]);
   }

   //save EPS into the export folder
   let filename = "/" + name + "_Core_RGB.eps";
   let destFile = new File(destFolder + filename);
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

   let inverseFilename = "/" + name + "_Inverse_RGB.eps";
   let inverseFile = new File(destFolder + inverseFilename);
   rgbDoc.saveAs(inverseFile, rgbSaveOpts);

   //save inverse file in all the PNG sizes
   for (let i = 0; i < PNGSizes.length; i++) {
      let filename = "/" + name + "_Inverse_RGB_" + PNGSizes[i] + ".png";
      let destFile = new File(destFolder + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, PNGSizes[i]);
   }

   //convert to inactive color (WTW Icon grey at 50% opacity) and save as EPS
   CSTasks.convertAll(rgbDoc.pathItems, colors[grayIndex][0], 50);

   let inactiveFilename = "/" + name + "_Inactive_RGB.eps";
   let inactiveFile = new File(destFolder + inactiveFilename);
   rgbDoc.saveAs(inactiveFile, rgbSaveOpts);

   for (let i = 0; i < PNGSizes.length; i++) {
      let filename = "/" + name + "_Inactive_RGB_" + PNGSizes[i] + ".png";
      let destFile = new File(destFolder + filename);
      CSTasks.scaleAndExportPNG(rgbDoc, destFile, startWidth, PNGSizes[i]);
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
   let cmykFilename = "/" + name + "_Core_CMYK.eps";
   let cmykDestFile = new File(destFolder + cmykFilename);
   let cmykSaveOpts = new EPSSaveOptions();
   cmykDoc.saveAs(cmykDestFile, cmykSaveOpts);

   //convert violet to white and save as EPS
   CSTasks.convertColorCMYK(
      cmykDoc.pathItems,
      colors[violetIndex][1],
      colors[whiteIndex][1]
   );

   let cmykInverseFilename = "/" + name + "_Inverse_CMYK.eps";
   let cmykInverseFile = new File(destFolder + cmykInverseFilename);
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
   let mastFilename = "/" + name + "_Masthead_RGB.eps";
   let mastDestFile = new File(destFolder + mastFilename);
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

main();
