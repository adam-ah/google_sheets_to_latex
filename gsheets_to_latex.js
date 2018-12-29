function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createAddonMenu();
  menu.addItem("Selection to LaTeX table", "latexify").addToUi();
}

function latexify() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var values = range.getDisplayValues();
  var columnStyle = Array(numCols);

  for (var k = 0; k < numCols; k++) {
    columnStyle[k] = "l";
  }

  var latex = "\\begin{table}\n";
  latex += "\\centering\n";
  latex += "\\begin{tabular}";
  latex += "{|" + columnStyle.join("|") + "|}\n";
  latex += "\\hline\n";

  var skips = getSkips(range);
  var nohlinefor = 0;

  for (var row = 0; row < numRows; row++) {
    var rowValues = values[row];
    var hline = "\\hline";
    nohlinefor--;

    for (var col = 0; col < numCols; col++) {
      var cell = rowValues[col];
      if (!cell) {
        continue;
      }
      cell = cell.replace(/[\u2018\u2019]/g, "'");
      cell = cell.replace("\\", "\\textbackslash");
      cell = cell.replace("~", "\\textasciitilde");
      cell = cell.replace(/([&%$#_{}])/g, "\\$1");

      splitCell = String(cell).split(/\n/);
      if (splitCell.length == 1) {
        cell = splitCell;
      } else {
        lines = splitCell.join("\\\\");
        cell = "\\begin{tabular}[c]{@{}l@{}}" + lines + "\\end{tabular}";
      }

      var skipkey = row + "-" + col;
      var rowskip = skips.rowskips[skipkey];

      if (rowskip) {
        cell = "\\multirow{" + rowskip + "}{*}{" + cell + "}";
        nohlinefor = rowskip;
      }

      if (nohlinefor > 1 && col < 2) {
        hline = "\\cline{2-" + numCols + "}";
      }

      var colskip = skips.colskips[skipkey];
      if (colskip) {
        for (var i = col + 1; i < col + colskip; i++) {
          rowValues[i] = null;
        }
        cell = "\\multicolumn{" + colskip + "}{c|}{" + cell + "}";
      }

      rowValues[col] = cell;
    }

    rowValues = rowValues.filter(function(v) {
      return v != null;
    });
    latex += rowValues.join(" & ");
    latex += " \\\\\n" + hline + "\n";
  }

  latex += "\\end{tabular}\n";
  latex += "\\label{table:table}\n";
  latex += "\\caption{\\small{}} \n";
  latex += "\\end{table}\n";

  var ui = SpreadsheetApp.getUi();

  ui.alert(
    "LaTeX table (don't forget to add \\usepackage{multirow})",
    latex,
    ui.ButtonSet.OK
  );
}

function getSkips(range) {
  var mergedRanges = range.getMergedRanges();

  var rowskips = {};
  var colskips = {};

  for (k = 0; k < mergedRanges.length; k++) {
    numRows = mergedRanges[k].getNumRows();
    numCols = mergedRanges[k].getNumColumns();

    row = mergedRanges[k].getRow() - range.getRow();
    col = mergedRanges[k].getColumn() - range.getColumn();

    var skipkey = row + "-" + col;
    if (numRows > 1) {
      rowskips[skipkey] = numRows;
    }
    if (numCols > 1) {
      colskips[skipkey] = numCols;
    }
  }
  return { rowskips: rowskips, colskips: colskips };
}
