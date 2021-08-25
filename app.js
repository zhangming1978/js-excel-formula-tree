window.onload = function () {
    initFunction();
};

function initFunction() {
    var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
    var spreadForShow = new GC.Spread.Sheets.Workbook(
        document.getElementById("show")
    );
    initShowSpread(spreadForShow);
    buildNodeTreeAndPaint(spread, spreadForShow);
}

function initShowSpread(spreadForShow) {
    var sheetForShow = spreadForShow.getActiveSheet();
    spreadForShow.suspendPaint();
    var spreadOptions = spreadForShow.options,
        sheetOptions = sheetForShow.options;
    spreadOptions.allowContextMenu = false;
    spreadOptions.scrollbarMaxAlign = true;
    spreadOptions.tabStripVisible = false;
    spreadOptions.allowUserResize = false;
    spreadOptions.allowUserDragDrop = false;
    spreadOptions.allowUserDragFill = false;
    spreadOptions.allowUserZoom = false;
    spreadOptions.grayAreaBackColor = "#ccddff";
    sheetOptions.colHeaderVisible = false;
    sheetOptions.rowHeaderVisible = false;
    sheetOptions.selectionBackColor = "transparent";
    sheetOptions.selectionBorderColor = "transparent";
    sheetOptions.gridline = {
        showVerticalGridline: false,
        showHorizontalGridline: false,
    };
    sheetForShow
        .getCell(1, 0)
        .foreColor("white")
        .text("Formula Precedents Tree")
        .font("bold italic 24pt Calibri")
        .vAlign(GC.Spread.Sheets.VerticalAlign.center)
        .textIndent(2);
    sheetForShow.getRange(0, 0, 100, 100).backColor("#ccddff");
    sheetOptions.isProtected = true;
    spreadForShow.resumePaint();
}

function buildNodeTreeAndPaint(spread, spreadForShow) {
    var sd = data;
    if (sd.length > 0) {
        spread.fromJSON(sd[0]);
        var sheet = spread.getActiveSheet();
        var sheetForShow = spreadForShow.getActiveSheet();
        sheet.bind(GC.Spread.Sheets.Events.SelectionChanging, function (e, info) {
            sheetForShow.shapes.clear();
            var row = info.newSelections[0].row;
            var col = info.newSelections[0].col;
            var nodeTree = creatNodeTree(row, col, sheet);
            paintDataTree(sheetForShow, nodeTree);
        });
    }
}

function creatNodeTree(row, col, sheet) {
    var _comment = sheet.getCell(row, col).comment();
    var node = {
        value: sheet.getValue(row, col),
        position: sheet.name() +
            "!" +
            GC.Spread.Sheets.CalcEngine.rangeToFormula(
                sheet.getRange(row, col, 1, 1)
            ),
        description: _comment && _comment.text(),
    };
    var childNodeArray = addChildNode(row, col, sheet);
    if (childNodeArray.length > 0) {
        node.childNodes = childNodeArray;
    }
    return node;
}

function addChildNode(row, col, sheet) {
    var childNodeArray = [];
    var childNodes = sheet.getPrecedents(row, col);
    if (childNodes.length >= 1) {
        childNodes.forEach(function (node) {
            var row = node.row,
                col = node.col,
                rowCount = node.rowCount,
                colCount = node.colCount,
                _sheet = sheet.parent.getSheetFromName(node.sheetName);
            if (rowCount > 1 || colCount > 1) {
                for (var r = row; r < row + rowCount; r++) {
                    for (var c = col; c < col + colCount; c++) {
                        childNodeArray.push(creatNodeTree(r, c, _sheet));
                    }
                }
            } else {
                childNodeArray.push(creatNodeTree(row, col, _sheet));
            }
        });
    }
    return childNodeArray;
}

function getRectShape(sheetForShow, name, x, y, width, height) {
    var rectShape = sheetForShow.shapes.add(
        name,
        GC.Spread.Sheets.Shapes.AutoShapeType.rectangle,
        x,
        y,
        width,
        height
    );
    var oldStyle = rectShape.style();
    oldStyle.textEffect.color = "white";
    oldStyle.fill.color = "#0065ff";
    oldStyle.textEffect.font = "bold 15px Calibri";
    oldStyle.textFrame.vAlign = GC.Spread.Sheets.VerticalAlign.top;
    oldStyle.textFrame.hAlign = GC.Spread.Sheets.HorizontalAlign.left;
    oldStyle.line.beginArrowheadWidth = 2;
    oldStyle.line.endArrowheadWidth = 2;
    rectShape.style(oldStyle);
    return rectShape;
}

function getConnectorShape(sheetForShow) {
    var connectorShape = sheetForShow.shapes.addConnector(
        "",
        GC.Spread.Sheets.Shapes.ConnectorType.elbow
    );
    var LineStyle = connectorShape.style();
    var line = LineStyle.line;
    line.beginArrowheadWidth = GC.Spread.Sheets.Shapes.ArrowheadWidth.wide;
    line.endArrowheadWidth = GC.Spread.Sheets.Shapes.ArrowheadWidth.wide;
    line.color = "#FF6600";
    connectorShape.style(LineStyle);
    return connectorShape;
}

function paintDataTree(
    sheetForShow,
    nodeTree,
    index,
    childLength,
    fatherShape
) {
    var rectWidth = 260,
        rectHeight = nodeTree.description ? 65 : 45;

    var spacingWidth = 300;
    var convertArray = [-0.75, 0.75, -2.25, 2.25, -2.25, 2.25, -4, 4, -5, 5];
    var spacingHeightMapping = [145, 135, 125, 50, 50];
    var name = Math.random().toString();
    var rectShape;
    if (fatherShape) {
        var x = fatherShape.x(),
            y = fatherShape.y();
        rectShape = getRectShape(
            sheetForShow,
            name,
            x + spacingWidth,
            y + convertArray[index] * spacingHeightMapping[childLength],
            rectWidth,
            rectHeight
        );
        var connectorShape = getConnectorShape(sheetForShow);
        connectorShape.startConnector({
            name: fatherShape.name(),
            index: 3,
        });
        connectorShape.endConnector({
            name: rectShape.name(),
            index: 1,
        });
    } else {
        rectShape = getRectShape(
            sheetForShow,
            name,
            200,
            250,
            rectWidth,
            rectHeight
        );
    }
    var _description =
        "值: " +
        nodeTree.value +
        "\n单元格: " +
        nodeTree.position +
        (nodeTree.description !== null ?
            "\n备注: " + nodeTree.description :
            "");
    rectShape.text(_description);
    var childNodes = nodeTree.childNodes;
    if (childNodes) {
        childNodes.forEach(function (node, index) {
            //if (node.description) {
            paintDataTree(sheetForShow, node, index, childNodes.length, rectShape);
            //}
        });
    }
}