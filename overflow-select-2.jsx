/*global app: false, alert: false, MeasurementUnits: false, PageSideOptions:false, LocationOptions: false */
// Put overflow from document as new doc name
// Check to see if there is overflow before opening new document
// Alert when there is no overflow
// Before moving on to the next text frame, check the new text frame for overflow and add an additional page if overflowing
var a, b, c;
a = app.activeDocument;

function calcOverset(txt) {
    var textLength, textId, allStory, overflowLength, pNumber;
    textLength = txt.characters.item(-1).index;
    textId = txt.parentStory.characters.length;
    allStory = txt.parentStory;
    overflowLength = textId - textLength;
    pNumber = txt.parentPage.name;
    return {
        tObject: allStory,
        oLength: textLength,
        page: pNumber
    };
}

function getOversetText(doc) {
    var overFlows = [],
        charLength = [],
        counter = 0,
        i = 0,
        textFramesO = [];
    textFramesO = doc.textFrames;
    for (i = 0; i < textFramesO.length; i += 1) {
        if (textFramesO[i].overflows === true && textFramesO[i].parent.constructor.name === "Spread" && textFramesO[i].parentPage !== null) {
            overFlows[counter] = calcOverset(textFramesO[i]);
            counter += 1;
        }
    }
    return overFlows;
}

function myGetBounds(myDocument, myPage) {
    var myPageWidth, myPageHeight, myX2, myX1, myY1, myY2;
    myPageWidth = myDocument.documentPreferences.pageWidth;
    myPageHeight = myDocument.documentPreferences.pageHeight;
    if (myPage.side === PageSideOptions.leftHand) {
        myX2 = myPage.marginPreferences.left;
        myX1 = myPage.marginPreferences.right;
    } else {
        myX1 = myPage.marginPreferences.left;
        myX2 = myPage.marginPreferences.right;
    }
    myY1 = myPage.marginPreferences.top;
    myX2 = myPageWidth - myX2;
    myY2 = myPageHeight - myPage.marginPreferences.bottom;
    return [myY1, myX1, myY2, myX2];
}

function printOverflow(overflows) {
    var newDoc = app.documents.add(),
        i = 0,
        pageItem,
        textFrameA,
        startChar,
        endChar,
        indent,
        mHeight,
        tBounds,
        overflowText,
        textFrameB,
        pageSelect;
    for (i = 0; i < overflows.length; i += 1) {
        // add pages
        if (newDoc.pages.length < i + 1) {
            newDoc.pages.add();
        }

        pageItem = newDoc.pages.item(i);
        newDoc.documentPreferences.facingPages = false;
        newDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
        newDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
        textFrameA = newDoc.pages.item(i).textFrames.add({
            geometricBounds: myGetBounds(newDoc, pageItem)
        });
        textFrameA.textFramePreferences.textColumnCount = 2;
        //Move text and remove beginning of story
        overflows[i].tObject.duplicate(LocationOptions.atBeginning, textFrameA.insertionPoints.item(0));
        startChar = 0;
        endChar = overflows[i].oLength;
        indent = false;
        if (textFrameA.parentStory.characters.item(endChar).contents === "\r") {
            indent = true;
        } else {
            indent = false;
        }
        textFrameA.parentStory.characters.itemByRange(startChar, endChar).remove();
        if (indent === false) {
            textFrameA.characters.item(0).firstLineIndent = 0;
        }
        mHeight = 60;
        tBounds = textFrameA.geometricBounds;
        overflowText = "Overflow from page: " + overflows[i].page;
        textFrameB = newDoc.pages.item(i).textFrames.add({
            geometricBounds: [tBounds[0], tBounds[1], mHeight, tBounds[3]],
            contents: overflowText
        });
        textFrameA.geometricBounds = [tBounds[0] + 25, tBounds[1], tBounds[2], tBounds[3]];
        pageSelect = textFrameB.characters.itemByRange(textFrameB.characters.item(0), textFrameB.characters.item(-1));
        pageSelect.pointSize = 15;
        pageSelect.fontStyle = "Bold";
    }
    newDoc.print();
}

function inspectObject(targ) {
    var key;
    for (key in targ) {
        if (targ.hasOwnProperty(key)) {
            alert(key + " -> " + targ[key]);
        }
    }
}

if (app.selection.length === 1) {
    //Evaluate the selection based on its type.
    if (app.selection[0].constructor.name === "TextFrame") {
        if (app.selection[0].overflows === true && app.selection[0].parent.constructor.name === "Spread" && app.selection[0].parentPage !== null) {
            b = [calcOverset(app.selection[0])];
            printOverflow(b);
        } else {
            if (app.selection[0].parentStory.overflows === true) {
                c = app.selection[0].parentStory.textContainers;
                b = [calcOverset(c[c.length - 1])];
                printOverflow(b);
            } else {
                alert("Your selection does not appear to contain overflow");
            }
        }
    } else {
        alert("Select a text frame to print overflow");
    }
} else {
    b = getOversetText(a);
    printOverflow(b);
}