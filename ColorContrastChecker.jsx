// ColorContrastChecker.jsx
// An InDesign Javascript

var scriptName = "WCAG Contrast Checker";

var selection = app.selection;

if (selection.length === 1) {
    var selection1 = selection[0];
    // Check if single selection is a textframe
    if (selection[0] == '[object TextFrame]' || selection[0] == '[object Cell]') {

        // Run the script, consolidate entire script into one Undo command
        if (parseFloat(app.version) < 6)
            checkColor();
        else
            app.doScript(checkColor, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Check Color Contrast");

        // Revert color to original values by undoing the script
        app.activeDocument.undo();

    }
    // If not, prompt to select a textframe or 2 shapes
    else {
        var myDialog = app.dialogs.add({
            name: scriptName
        });
        //Add a dialog column.
        with (myDialog.dialogColumns.add()) {
            staticTexts.add({
                staticLabel: "Select a textframe or two shapes to compare colors."
            });
        }
        //Show the dialog box.
        var myResult = myDialog.show();
        myDialog.destroy();
    }

}

if (selection.length === 2) {
    var selection1 = selection[0];
    var selection2 = selection[1];

    // Find object type of selection and returns true if it is an image
    var imageCheck1 = selection1.allGraphics == "[object Image]" ? true : false;
    var imageCheck2 = selection2.allGraphics == "[object Image]" ? true : false;

    if (imageCheck1 === false && imageCheck2 === false) {
        // Run the script, consolidate entire script into one Undo command
        if (parseFloat(app.version) < 6)
            checkColor();
        else
            app.doScript(checkColor, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Check Color Contrast");

        // Revert color to original values by undoing the script
        app.activeDocument.undo();
    }
    // Show error dialog if an image is selected
    else {
        var myDialog = app.dialogs.add({
            name: scriptName
        });
        //Add a dialog column.
        with (myDialog.dialogColumns.add()) {
            staticTexts.add({
                staticLabel: "Do not select an image"
            });
        }
        //Show the dialog box.
        var myResult = myDialog.show();
        myDialog.destroy();
    }
}
// Show error dialog if there are not exactly 2 items selected
if (selection.length > 2) {
    var myDialog = app.dialogs.add({
        name: scriptName
    });
    //Add a dialog column.
    with (myDialog.dialogColumns.add()) {
        staticTexts.add({
            staticLabel: "Select at most 2 items to check contrast"
        });
    }
    //Show the dialog box.
    var myResult = myDialog.show();
    myDialog.destroy();
}


// Function to compare the two colors
function checkColor() {
    var compareMethod;

    // Comparing two shapes
    if (selection1 && selection2) {
        // Set the compare method
        compareMethod = 'Fill Color v. Fill Color';
        // Create duplicates of the selections.
        var dupSelection1 = selection1.duplicate();
        var dupSelection2 = selection2.duplicate();
        var color1, color2;

        // Extract RGB value from selections. Manual values added for uneditable colors.
        if (dupSelection1 == '[object TextFrame]' && dupSelection1.words.length > 0) {
            var c = dupSelection1.texts[0].fillColor;
            color1 = assignColor(c);
        } else {
            var c = dupSelection1.fillColor;
            color1 = assignColor(c);
        }

        if (dupSelection2 == '[object TextFrame]') {
            var c = dupSelection2.texts[0].fillColor;
            color2 = assignColor(c);
        } else {
            var c = dupSelection2.fillColor;
            color2 = assignColor(c);
        }
    }

    // Comparing text fill vs. frame/cell fill
    if (selection1 && !selection2) {
        var color1, color2;
        if (selection1 == '[object TextFrame]') {
            // Set the compare method
            compareMethod = 'Text v. Textframe Fill';
            var dupSelection1 = selection1.duplicate();

            var c1 = dupSelection1.texts[0].fillColor;
            var c2 = dupSelection1.fillColor;

            color1 = assignColor(c1);
            color2 = assignColor(c2);
        }

        if (selection1 == '[object Cell]') {
            compareMethod = 'Text v. Cell Fill'
            var c1 = selection1.texts[0].fillColor
            var c2 = selection1.fillColor;

            color1 = checkForDefaultSwatch(c1);
            color2 = checkForDefaultSwatch(c2);
            if (color1 === false) {
                color1 = getRGBFromCell(c1);
            }
            if (color2 === false) {
                color2 = getRGBFromCell(c2);
            }
        }
    }


    var color1luminance = luminance(color1);
    var color2luminance = luminance(color2);

    // Calculate the contrast ratio
    var ratio = color1luminance > color2luminance ?
        ((color2luminance + 0.05) / (color1luminance + 0.05)) :
        ((color1luminance + 0.05) / (color2luminance + 0.05));

    var AA_Large = ratio < 1 / 3 ? 'PASS' : 'FAIL';
    var AA_Normal = ratio < 1 / 4.5 ? 'PASS' : 'FAIL';
    var AAA_Large = ratio < 1 / 4.5 ? 'PASS' : 'FAIL';
    var AAA_Normal = ratio < 1 / 7 ? 'PASS' : 'FAIL';

    myDialog(compareMethod, color1, color2, AA_Large, AA_Normal, AAA_Large, AAA_Normal);

    if (dupSelection1) { dupSelection1.remove() }
    if (dupSelection2) { dupSelection2.remove(); }
}


// Calculate luminance for each color
function luminance(rgb) {
    var a = [];
    for (var i = 0; i < 3; i++) {
        rgb[i] /= 255;
        a.push(rgb[i] <= 0.03928 ? rgb[i] / 12.92 : Math.pow((rgb[i] + 0.055) / 1.055, 2.4));
    }
    return a[0] * 0.2126 + a[1] * 0.7152 + a[2] * 0.0722;
}

function assignColor(c) {
    var color = checkForDefaultSwatch(c);
    if (color === false) {
        c.space = ColorSpace.RGB;
        color = c.colorValue;
    }
    return color;
}

function checkForDefaultSwatch(c) {
    var color;
    if (c.name == 'Registration' || c.name == 'Black') {
        color = [0, 0, 0];
        return color;
    } else if (c.name == 'None') {
        color = [255, 255, 255];
        return color;
    }
    return false;
}

function getRGBFromCell(c) {
    var rectangle = app.activeDocument.rectangles.add({ fillColor: c });
    rectangle.fillColor.space = ColorSpace.RGB;
    var color = rectangle.fillColor.colorValue;
    rectangle.remove();
    return color;
}

// ScriptUI Dialog Window
function myDialog(compareMethod, color1, color2, AA_Large, AA_Normal, AAA_Large, AAA_Normal) {
    // DIALOG
    // ======
    var dialog = new Window("dialog");
    dialog.text = "WCAG Contrast Checker";
    dialog.orientation = "column";
    dialog.alignChildren = ["center", "top"];
    dialog.spacing = 10;
    dialog.margins = 16;

    // GROUP1
    // ======
    var group1 = dialog.add("group", undefined, { name: "group1" });
    group1.orientation = "row";
    group1.alignChildren = ["left", "center"];
    group1.spacing = 10;
    group1.margins = 0;

    // PANEL1
    // ======
    var panel1 = group1.add("panel", undefined, undefined, { name: "panel1" });
    panel1.text = "Compare Method";
    panel1.orientation = "column";
    panel1.alignChildren = ["left", "top"];
    panel1.spacing = 10;
    panel1.margins = 10;

    var statictext1 = panel1.add("statictext", undefined, undefined, { name: "statictext1" });
    statictext1.text = compareMethod;

    // GROUP2
    // ======
    var group2 = dialog.add("group", undefined, { name: "group2" });
    group2.orientation = "row";
    group2.alignChildren = ["left", "center"];
    group2.spacing = 10;
    group2.margins = 0;

    // PANEL2
    // ======
    var panel2 = group2.add("panel", undefined, undefined, { name: "panel2" });
    panel2.text = "Color One";
    panel2.orientation = "column";
    panel2.alignChildren = ["left", "top"];
    panel2.spacing = 10;
    panel2.margins = 10;
    panel2.graphics.backgroundColor = panel2.graphics.newBrush(panel2.graphics.BrushType.SOLID_COLOR, color1);

    // PANEL3
    // ======
    var panel3 = group2.add("panel", undefined, undefined, { name: "panel3" });
    panel3.text = "Color Two";
    panel3.orientation = "column";
    panel3.alignChildren = ["left", "top"];
    panel3.spacing = 10;
    panel3.margins = 10;
    panel3.graphics.backgroundColor = panel3.graphics.newBrush(panel3.graphics.BrushType.SOLID_COLOR, color2);

    // GROUP3
    // ======
    var group3 = dialog.add("group", undefined, { name: "group3" });
    group3.orientation = "row";
    group3.alignChildren = ["left", "center"];
    group3.spacing = 10;
    group3.margins = 0;

    // PANEL4
    // ======
    var panel4 = group3.add("panel", undefined, undefined, { name: "panel4" });
    panel4.text = "Contrast";
    panel4.orientation = "column";
    panel4.alignChildren = ["left", "top"];
    panel4.spacing = 10;
    panel4.margins = 10;
    panel4.alignment = ["left", "top"];

    var aaLargeText = panel4.add("statictext", undefined, undefined, { name: "aaLargeText" });
    aaLargeText.text = "AA Large: " + AA_Large;
    if (AA_Large === 'PASS') {
        aaLargeText.graphics.foregroundColor = aaLargeText.graphics.newPen(aaLargeText.graphics.PenType.SOLID_COLOR, [0, 1, 0], 1);
    }
    if (AA_Large === 'FAIL') {
        aaLargeText.graphics.foregroundColor = aaLargeText.graphics.newPen(aaLargeText.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
    }

    var aaNormalText = panel4.add("statictext", undefined, undefined, { name: "aaNormalText" });
    aaNormalText.text = "AA Normal: " + AA_Normal;
    if (AA_Normal === 'PASS') {
        aaNormalText.graphics.foregroundColor = aaNormalText.graphics.newPen(aaNormalText.graphics.PenType.SOLID_COLOR, [0, 1, 0], 1);
    }
    if (AA_Normal === 'FAIL') {
        aaNormalText.graphics.foregroundColor = aaNormalText.graphics.newPen(aaNormalText.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
    }

    var aaaLargeText = panel4.add("statictext", undefined, undefined, { name: "aaaLargeText" });
    aaaLargeText.text = "AAA Large: " + AAA_Large;
    if (AAA_Large === 'PASS') {
        aaaLargeText.graphics.foregroundColor = aaaLargeText.graphics.newPen(aaaLargeText.graphics.PenType.SOLID_COLOR, [0, 1, 0], 1);
    }
    if (AAA_Large === 'FAIL') {
        aaaLargeText.graphics.foregroundColor = aaaLargeText.graphics.newPen(aaaLargeText.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
    }

    var aaaNormalText = panel4.add("statictext", undefined, undefined, { name: "aaaNormalText" });
    aaaNormalText.text = "AAA Normal: " + AAA_Normal;
    if (AAA_Normal === 'PASS') {
        aaaNormalText.graphics.foregroundColor = aaaNormalText.graphics.newPen(aaaNormalText.graphics.PenType.SOLID_COLOR, [0, 1, 0], 1);
    }
    if (AAA_Normal === 'FAIL') {
        aaaNormalText.graphics.foregroundColor = aaaNormalText.graphics.newPen(aaaNormalText.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
    }

    dialog.show();
}