//button functions
//button event listeners
let addColumnBtn = document.getElementById("addColumnBtn");
let addRowBtn = document.getElementById("addRowBtn");
let delColumnBtn = document.getElementById("delColumnBtn");
let delRowBtn = document.getElementById("delRowBtn");
let exportBtn = document.getElementById("exportBtn");
let importBtn = document.getElementById("importBtn");

addColumnBtn.addEventListener("click", function () {
  addCol(this);
});
addRowBtn.addEventListener("click", function () {
  addRow(this);
});
delColumnBtn.addEventListener("click", function () {
  delCol(this);
});
delRowBtn.addEventListener("click", function () {
  delRow(this);
});
exportBtn.addEventListener("click", function () {
  exportCSV(this);
});
importBtn.addEventListener("click", function () {
  importFromCSV(this);
});


const addRow = () => {
  console.log("Add Row Function");
  if (selectedRow > 0) {
    let spreadSheetData = this.fetchspreadSheetData();
    const colCount = spreadSheetData[0].length;
    let rowPos = parseInt(selectedRow) + 1;
    const newRow = new Array(colCount).fill("");
    updateTable(selectedRow, "addRow");
    spreadSheetData.splice(rowPos, 0, newRow);
    initialRowCount++;
    this.displaySpreadSheet("update");
    selectedRow = -1;
  }
};

const addCol = () => {
  console.log("Add col func");
  console.log(selectedCol);
  let colPos = parseInt(selectedCol) + 1;
  if (colPos < 27 && colPos > 0 && initialColCount < 26) {
    let spreadSheetData = this.fetchspreadSheetData();
    updateTable(selectedCol, "addCol");
    for (let i = 0; i <= initialRowCount; i++) {
      spreadSheetData[i].splice(colPos, 0, "");
    }
    initialColCount++;
    this.displaySpreadSheet("update");
    selectedCol = -1;
  }
};

const delRow = () => {
  console.log("delRow Function");
  if (selectedRow > 0 && initialRowCount > 1) {
    let spreadSheetData = this.fetchspreadSheetData();
    updateTable(selectedRow, "delRow");
    spreadSheetData.splice(selectedRow, 1);
    initialRowCount--;
    this.displaySpreadSheet("update");
    selectedRow = -1;
  }
};

const delCol = () => {
  console.log("delCol Function");
  let colPos = parseInt(selectedCol);
  if (colPos < 27 && colPos > 0 && initialColCount > 1) {
    let spreadSheetData = this.fetchspreadSheetData();
    updateTable(selectedCol, "delCol");
    for (let i = 0; i <= initialRowCount; i++) {
      spreadSheetData[i].splice(colPos, 1);
    }
    initialColCount--;
    this.displaySpreadSheet("update");
    selectedCol = -1;
  }
};

disableButtons = (value) => {
  addColumnBtn.style.visibility = "hidden";
  addRowBtn.style.visibility = "hidden";
  delColumnBtn.style.visibility = "hidden";
  delRowBtn.style.visibility = "hidden";
};

// It will be called after every row/column  update/delete and will update formulas accordingly
updateTable = (selectedRowOrColumnIndex, operation) => {
  let spreadSheetData = this.fetchspreadSheetData();
  for (let i = 1; i < spreadSheetData.length; i++) {
    spreadSheetData[i].forEach((cell) => {
      if ((cell.formula || "").startsWith("=")) {
        let splittedAcutalValue = cell.formula.split("");
        if (operation == "addCol") {
          splittedAcutalValue.forEach(function (item, index) {
            if (
              selectedRowOrColumnIndex <= colHeaderArr.indexOf(item) &&
              colHeaderArr.indexOf(item) >= 0 &&
              colHeaderArr.indexOf(item) < spreadSheetData[0].length - 1
            ) {
              let desiredIndex = colHeaderArr.indexOf(item) + 1;
              this[index] = colHeaderArr[desiredIndex] || "";
            }
          }, splittedAcutalValue);
        }
        if (operation == "addRow") {
          splittedAcutalValue.forEach(function (item, index) {
            if (
              parseInt(item) > selectedRowOrColumnIndex &&
              parseInt(item) > 0
            ) {
              this[index] = parseInt(item) + 1;
            }
          }, splittedAcutalValue);
        }
        if (operation == "delRow") {
          splittedAcutalValue.forEach(function (item, index) {
            if (
              parseInt(item) > selectedRowOrColumnIndex &&
              parseInt(item) > 0
            ) {
              this[index] = parseInt(item) - 1;
            }
          }, splittedAcutalValue);
        }
        if (operation == "delCol") {
          splittedAcutalValue.forEach(function (item, index) {
            if (
              colHeaderArr.indexOf(item) > selectedRowOrColumnIndex - 1 &&
              colHeaderArr.indexOf(item) >= 0
            ) {
              let desiredIndex = colHeaderArr.indexOf(item) - 1;
              this[index] = colHeaderArr[desiredIndex] || "";
            } else if (
              colHeaderArr.indexOf(item) ===
              selectedRowOrColumnIndex - 1
            ) {
              cell.displayValue = "!ERR";
            }
          }, splittedAcutalValue);
        }
        if (cell.displayValue !== "!ERR") {
          cell.formula = splittedAcutalValue.join("");
        }
      }
    });
  }
};


//import sheet from CSV func
const importFromCSV = () => {
  console.log("importFromCSV Function");
  var fileUpload = document.getElementById("fileUpload");
  var regex = /^([a-zA-Z0-9\s_\\.\-:()])+(.csv)$/;
  if (regex.test(fileUpload.value.toLowerCase())) {
    if (typeof FileReader != "undefined") {
      var reader = new FileReader();
      reader.onload = function (e) {
        var rows = e.target.result.split("\n");
        spreadSheetData = [];
        for (let i = 0; i < rows.length; i++) {
          let eachRow = [];
          let row = rows[i].split(",");
          for (let j = 0; j < row.length; j++) {
            let currCell;
            if (row[j].startsWith("=")) {
              currCell = new DataElement("", row[j], []);
            } else {
              currCell = new DataElement(row[j], "", []);
            }
            eachRow.push(currCell);
          }
          spreadSheetData.push(eachRow);
        }
        let cols = spreadSheetData[0].length;
        spreadSheetData.splice(0, 0, Array(cols).fill(""));
        for (let i = 0; i < spreadSheetData.length; i++) {
          spreadSheetData[i].splice(0, 0, "");
        }
        for (let i = 1; i < spreadSheetData.length; i++) {
          for (let j = 1; j < spreadSheetData[0].length; j++) {
            if (spreadSheetData[i][j].formula.startsWith("=")) {
              let rowIndex = i;
              let colIndex = j - 1;
              let id = String(colHeaderArr[colIndex] + rowIndex);
              let item = document.createElement("td");
              item.setAttribute("id", id);
              item.innerHTML = spreadSheetData[i][j].formula;
              evaluateEachCell(item);
            }
          }
        }
        //displaySpreadSheet("initialise");
        getTableData();
      };
      reader.readAsText(fileUpload.files[0]);
    }
  } else {
    alert("Wrong CSV file");
  }
};

//export to CSV file func
const exportCSV = () => {
  console.log("exportCSV Function");
  var csv = [];
  const spreadsheetData = this.fetchspreadSheetData();
  for (var i = 1; i < spreadsheetData.length; i++) {
    var row = [];
    var col = spreadsheetData[i];
    for (var j = 1; j < col.length; j++) {
      if (col[j].formula != "") {
        row.push(col[j].formula);
      } else {
        row.push(col[j].displayValue);
      }
    }
    csv.push(row.join(","));
  }
  downloadCsv(csv.join("\n"), filename);
};

downloadCsv = (csv, filename) => {
  var csvFile;
  var downloadLink;
  csvFile = new Blob([csv], { type: "text/csv" });
  downloadLink = document.createElement("a");
  downloadLink.download = filename;
  downloadLink.href = window.URL.createObjectURL(csvFile);
  downloadLink.style.display = "none";
  document.body.appendChild(downloadLink);
  downloadLink.click();
};


//Evaluation functions
evaluateEachCell = (item) => {
  let cellExpression = item.innerHTML;
  let result = "";
  let indices = getIndices(item);
  if (cellExpression.startsWith("=")) {
    console.log("starts with a =");
    if (cellExpression.indexOf("SUM") > 0) {
      console.log("Group SUM");
      result = evaluateGroupSum(item);
    } else {
      console.log("Math equation");
      result = evaluateMathEqn(item);
    }
  } else if (cellExpression != "") {
    console.log("value");
    updateValue(item);
    return;
  } else {
    console.log("empty cell");
    let cell =
      spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1];
    cell.formula = item.innerHTML;
    cell.displayValue = item.innerHTML;
    removeObs(cell);
    console.log("Cell evaluation successful");
    return;
  }
  if (result != "") {
    let cell =
      spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1];
    cell.formula = item.innerHTML;
    if (result == "!ERR") {
      cell.observables = [];
    }
    if (result == "Infinity") {
      cell.displayValue = "DIV/0";
    } else {
      cell.displayValue = result;
    }
    console.log("Cell evaluation successful");
  } else {
    console.log("Result is empty");
  }
};

//SUM(:) evaluation
evaluateGroupSum = (item) => {
  console.log("evaluateGroupSum");
  let sum = 0;
  let valueArr = [];
  let cellExpression = item.innerHTML;
  let c = cellExpression
    .replace("=SUM", "")
    .replace("(", "")
    .replace(")", "")
    .split(":");
  console.log(c + " " + c.length);
  //Format mismatch as c should be an array with start and end cell location
  if (c.length < 2) {
    console.log("Invalid Group Sum Formula");
    return error;
  }
  let patternInd = /(\d{1,2}:?)/i;
  let patternAlphabet = /(\w{1}:?)/i;
  let startCell = c[0];
  let endCell = c[1];
  let startCol = patternAlphabet.exec(startCell)[0];
  let endCol = patternAlphabet.exec(endCell)[0];
  let startRow = patternInd.exec(startCell)[0];
  let endRow = patternInd.exec(endCell)[0];
  //for item location
  let indices = getIndices(item);
  console.log(
    "startCol = " +
      startCol +
      " endCol = " +
      endCol +
      " startRow = " +
      startRow +
      " endRow = " +
      endRow
  );
  //check boundary condition and if headers are excluded
  if (
    !(
      colHeaderArr.includes(startCol) &&
      colHeaderArr.includes(endCol) &&
      startRow >= 1 &&
      endRow <= spreadSheetData.length
    )
  ) {
    console.log("boundary condition error");
    return error;
  }
  if (selfReferenced(c, item)) {
    return error;
  }
  //remove old observables
  let cell = spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1];
  removeObs(cell);
  //check if col sum / row sum/ wrong selection
  if (startCol == endCol && startRow != endRow) {
    console.log("col group sum");
    for (let i = startRow; i <= endRow; i++) {
      valueArr.push(
        checkCellValueFormat(
          String(document.getElementById(`${startCol}${i}`).innerHTML)
        )
      );
      cell.observables.push(this.getCellObs(String(startCol + i)));
    }
  } else if (startCol != endCol && startRow == endRow) {
    console.log("row group sum");
    for (
      let j = colHeaderArr.indexOf(startCol) + 1;
      j <= colHeaderArr.indexOf(endCol) + 1;
      j++
    ) {
      valueArr.push(
        checkCellValueFormat(
          String(
            document.getElementById(`${colHeaderArr[j - 1]}${startRow}`)
              .innerHTML
          )
        )
      );
      cell.observables.push(
        this.getCellObs(String(colHeaderArr[j - 1] + startRow))
      );
    }
  } else {
    console.log(" Wrong selection");
    return error;
  }
  sum = valueArr.reduce(function (a, b) {
    if (b == "!ERR") {
      return error;
    }
    return parseInt(a) + parseInt(b);
  }, 0);
  console.log("sum is " + sum);
  return sum;
};

//Algebraic Eval
evaluateMathEqn = (item) => {
  console.log("evaluateMathEqn");
  let result = 0;
  let valueArr = [];
  let cellExpression = item.innerHTML;
  let eqn = cellExpression.split(/([=*/%+-])/g).splice(2);
  console.log(eqn + " " + eqn.length);
  //Format mismatch as a should be an array with atleast start and end cell location
  if (eqn.length < 3) {
    console.log("Invalid Math Eqn");
    return error;
  }
  //check expression starts with a special char
  if (/^[*+-//]/.test(cellExpression) || /$[*+-//]/.test(cellExpression)) {
    console.log("Starts/Ends with special Char");
    return error;
  }
  //get equation
  let EqnWithValue = cellExpression.replace(/([=*/%+-])/g, " ").split(" ");
  EqnWithValue.shift();
  if (eqn.length < 2) {
    console.log("Contains less than two operand");
    return error;
  }
  //get cell indices involved in equation for future resolution
  let cellIndices = [];
  for (let i = 0; i < EqnWithValue.length; i++) {
    if (EqnWithValue[i].length > 1) {
      cellIndices.push(EqnWithValue[i]);
    } else {
      console.log("Encountered a value in the eqn");
    }
  }
  //check if cell is self referrenced
  if (selfReferenced(cellIndices, item)) {
    console.log("Self referrence error");
    return error;
  }
  for (let i = 0; i < cellIndices.length; i++) {
    let patternInd = /(\d{1,2}:?)/i;
    let patternAlphabet = /(\w{1}:?)/i;
    let rowIndex = patternInd.exec(cellIndices[i]);
    let colIndex = patternAlphabet.exec(cellIndices[i]);
    console.log(rowIndex[0] + " " + colIndex[0]);
    //check boundary condition and if headers are excluded
    if (
      !(
        colHeaderArr.includes(colIndex[0]) &&
        rowIndex[0] >= 1 &&
        rowIndex[0] <= spreadSheetData.length
      )
    ) {
      console.log("Boundary condition error");
      return error;
    }
  }

  //for item location
  let indices = getIndices(item);
  let cell = spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1];
  removeObs(cell);
  for (let i = 0; i < eqn.length; i++) {
    if (cellIndices.includes(eqn[i])) {
      console.log("eqn[i] is " + eqn[i]);
      //resole the indices to values
      let patternInd = /(\d{1,2}:?)/i;
      let patternAlphabet = /(\w{1}:?)/i;
      let rowIndex = patternInd.exec(eqn[i])[0];
      let colIndex = patternAlphabet.exec(eqn[i])[0];
      console.log(rowIndex[0] + " " + colIndex[0]);
      valueArr.push(
        checkCellValueFormat(
          String(document.getElementById(`${colIndex}${rowIndex}`).innerHTML)
        )
      );
      cell.observables.push(this.getCellObs(String(`${colIndex}${rowIndex}`)));
    } else {
      console.log("operator/Operand");
      valueArr.push(String(eqn[i]));
    }
  }
  console.log(valueArr);
  result = getEqnOutput(valueArr);
  return result;
};

//helper functions
//use BODMAS rule to calc result of the eqn
getEqnOutput = (valueArr) => {
  //Handle Multiply
  for (i = 0; i <= valueArr.length; i++) {
    cItem = valueArr[i];
    if (cItem == "*") {
      tLeft = parseFloat(valueArr[i - 1]);
      tRight = parseFloat(valueArr[i + 1]);

      nVal = tLeft * tRight;
      valueArr[i - 1] = nVal;
      valueArr.splice(i, 2);
      i = valueArr.length;
    }
  }
  console.log(valueArr);
  //Handle Divide
  for (i = 0; i <= valueArr.length; i++) {
    cItem = valueArr[i];
    if (cItem == "/") {
      tLeft = parseFloat(valueArr[i - 1]);
      tRight = parseFloat(valueArr[i + 1]);

      nVal = tLeft / tRight;
      valueArr[i - 1] = nVal;
      valueArr.splice(i, 2);
      i = valueArr.length;
    }
  }
  console.log(valueArr);
  //Handle Plus and Minus
  var tResult = parseFloat(valueArr[0]);
  for (i = 1; i < valueArr.length; i++) {
    if (valueArr[i] == "+") {
      tResult = tResult + parseFloat(valueArr[i + 1]);
    } else if (valueArr[i] == "-") {
      tResult = tResult - parseFloat(valueArr[i + 1]);
    }
    i++;
  }
  console.log(tResult);
  return tResult;
};

//check self reference func
selfReferenced = (c, item) => {
  if (c.includes(item.id)) {
    return true;
  } else {
    return false;
  }
};

//check format of the value before passing to arr func
checkCellValueFormat = (value) => {
  if (value == "!ERR") {
    return "!ERR";
  } else if (
    isNaN(value) ||
    value == "" ||
    value.match(/[\s{?}]/gi) ||
    value.match(/[a-zA-Z]/gi)
  ) {
    return 0;
  } else {
    return value;
  }
};

//return row and column index of an item using its id
getIndices = (item) => {
  let patternInd = /(\d{1,2}:?)/i;
  let patternAlphabet = /(\w{1}:?)/i;
  let rowIndex = patternInd.exec(item.id)[0];
  let colIndex = patternAlphabet.exec(item.id)[0];
  let indices = [];
  indices.push(rowIndex, colIndex);
  console.log(item.id);
  console.log(rowIndex);
  console.log(colIndex);
  return indices;
};

//func to update the formula bar
function updateFormulaBar(formula, displayValue) {
  var ps = document.getElementsByClassName("messages");
  let value = "";
  if (formula && displayValue) {
    value = "Formula " + formula + " Value=" + displayValue;
  } else if (formula) {
    value = formula;
  } else {
    value = displayValue;
  }
  for (var i = 0, len = ps.length; i < len; i++) {
    ps[i].innerHTML = value;
  }
}

//clear the formula bar when cell is deselected
function clearMessages() {
  var ps = document.getElementsByClassName("messages");
  for (var i = 0, len = ps.length; i < len; i++) {
    ps[i].innerHTML = "";
  }
}

//called when a formula is removed/updated
function removeObs(cell) {
  if (cell.observables.length > 0) {
    console.log(cell.observables);
    console.log(
      "Observables exist but formula is removed, unsubscribe and delete"
    );
    cell.observables.forEach((observable) => {
      try {
        observable.unsubscribe();
      } catch (err) {
        console.log(" not subscribed");
      }
    });
  }
  cell.observables = [];
}

//Initialize base parameters
//Contain entire sheet details in an array
let spreadSheetData = [];
let error = "!ERR";
const { takeWhile } = rxjs.operators;
//Genarate alphabet array for column headings
let colHeaderArr = [];
for (let i = 0; i < 26; i++) {
  colHeaderArr.push(String.fromCharCode(65 + i));
}
let filename = "webdesAssignment7.csv";
// Initial row and column count
let initialRowCount = 10;
let initialColCount = 10;

//Index of selected row, column to get the index of the selected cell
let selectedRow = -1;
let selectedCol = -1;

class DataElement {
  constructor(displayValue, formula, observables) {
    this.displayValue = displayValue;
    this.formula = formula;
    this.observables = observables;
  }
}

//Table functions
// Initialize and get sheet Data from spreadSheetData Array
fetchspreadSheetData = () => {
  if (
    spreadSheetData === undefined ||
    spreadSheetData === null ||
    spreadSheetData.length === 0
  ) {
    for (let i = 0; i <= initialRowCount; i++) {
      const eachRow = [];
      for (let j = 0; j <= initialColCount; j++) {
        eachRow.push("");
      }
      spreadSheetData.push(eachRow);
    }
  }
  return spreadSheetData;
};

displaySpreadSheet = (op) => {
  console.log("displaySpreadSheet Function");
  let tableBody = document.getElementById("spreadsheetTable");
  let thead = "";
  let table = "";
  if (op === "update") {
    const spreadSheetData = this.fetchspreadSheetData();
    let tableHeaders = document.getElementById("tableheader");
    table = tableBody.cloneNode(true);
    tableBody.parentNode.replaceChild(table, tableBody);
    thead = tableHeaders.cloneNode(true);
    tableHeaders.parentNode.replaceChild(thead, tableHeaders);
    table.innerHTML = "";
    thead.innerHTML = "";
    initialRowCount = spreadSheetData.length - 1;
    initialColCount = spreadSheetData[0].length - 1;
  } else {
    table = document.getElementById("spreadsheetTable");
    thead = document.createElement("thead");
    thead.setAttribute("id", "tableheader");
  }
  //initialize spreadsheetdata
  fetchspreadSheetData();
  //generate col header (ABC...)
  thead.innerHTML = generateColHeader(thead, initialColCount);
  table.appendChild(thead);
  //generate row header (123...)
  generateRows(table, initialRowCount, initialColCount);
  //initialize cells with cell obj display value and check for obs
  getTableData();
  //disable buttons when using the table
  table.addEventListener("focusin", function (e) {
    if (e.target && e.target.nodeName === "TD") {
      unselect();
      disableButtons(true);
      let item = e.target;
      let indices = getIndices(item);
      updateFormulaBar(
        spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1]
          .formula,
        spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1]
          .displayValue
      );
      if (
        spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1]
          .formula != ""
      ) {
        item.innerHTML =
          spreadSheetData[indices[0]][
            colHeaderArr.indexOf(indices[1]) + 1
          ].formula;
      } else {
        item.innerHTML =
          spreadSheetData[indices[0]][
            colHeaderArr.indexOf(indices[1]) + 1
          ].displayValue;
      }
    }
  });
  // attach focusout event listener to whole table body container
  table.addEventListener("focusout", function (e) {
    if (e.target && e.target.nodeName === "TD") {
      unselect();
      let item = e.target;
      clearMessages();
      //when mouse is clicked outside, evaluate the cells and populate the table again
      evaluateEachCell(item);
      getTableData(item);
    }
  });
  // Attach click event listener to table body
  table.addEventListener("click", function (e) {
    if (e.target && e.target.className === "rowHeader") {
      unselect();
      console.log("rowHeader found");
      addRowBtn.style.visibility = "visible";
      delRowBtn.style.visibility = "visible";
      e.target.classList.add("highlight");
      e.target.parentNode.classList.add("highlight");
      selectedRow = e.target.parentNode.id.split("-")[1];
      selectRow(e.target.parentNode);
    }
  });
  // Attach click event listener to table headers
  thead.addEventListener("click", function (e) {
    addColumnBtn.style.visibility = "visible";
    delColumnBtn.style.visibility = "visible";
    unselect();
    if (e.target && e.target.className === "colHeader") {
      e.target.classList.add("highlight");
      selectedCol = e.target.id.split("-")[1];
      //console.log(selectedCol)
      selectCol(selectedCol);
    }
  });
  disableButtons(true);
};

//get column data
generateColHeader = (thead, initialColCount) => {
  console.log("generateColHeader");
  let tr = document.createElement("tr");
  for (let i = 0; i <= initialColCount; i++) {
    let th = document.createElement("th");
    th.contentEditable = false;
    th.setAttribute("class", "colHeader");
    if (i == 0) {
      th.innerHTML = " ";
    } else {
      th.setAttribute("id", `col-${i}`);
      th.innerHTML = colHeaderArr[i - 1];
    }
    tr.appendChild(th);
  }
  thead.appendChild(tr);
  return thead.innerHTML;
};

//get row data
generateRows = (table, initialRowCount, initialColCount) => {
  console.log("generateRows");
  for (let i = 1; i <= initialRowCount; i++) {
    let tr = document.createElement("tr");
    tr.setAttribute("id", `r-${i}`);
    let td = document.createElement("td");
    td.setAttribute("class", "rowHeader");
    td.innerHTML = i;
    td.contentEditable = false;
    tr.appendChild(td);
    for (let j = 1; j <= initialColCount; j++) {
      let cell = document.createElement("td");
      cell.innerHTML = "";
      cell.contentEditable = true;
      cell.setAttribute("id", `${colHeaderArr[j - 1]}${i}`);
      cell.id = `${colHeaderArr[j - 1]}${i}`;
      tr.appendChild(cell);
      // console.log("cell id is"+ `${colHeaderArr[j-1]}${i}`)
      if (!spreadSheetData[i][j]) {
        // console.log("Creating a Cell obj at "+i+","+j);
        spreadSheetData[i][j] = new DataElement("", "", []);
      }
    }
    table.appendChild(tr);
  }
};

//if only value was updated (not empty and not a formula)
updateValue = (item) => {
  let indices = getIndices(item);
  let cell = spreadSheetData[indices[0]][colHeaderArr.indexOf(indices[1]) + 1];
  cell.formula = "";
  cell.displayValue = item.innerHTML;
  removeObs(cell);
  console.log("value updated in spreadSheetArray");
};

// get each cell data with is actual and display values and create observables
getTableData = (item) => {
  const spreadSheetData = this.fetchspreadSheetData();
  if (spreadSheetData === undefined || spreadSheetData === null) return;
  for (let i = 1; i <= spreadSheetData.length - 1; i++) {
    for (let j = 1; j <= spreadSheetData[i].length - 1; j++) {
      // console.log(`${colHeaderArr[j-1]}${i}`)
      let cell = document.getElementById(`${colHeaderArr[j - 1]}${i}`);
      // console.log(cell)
      cell.innerHTML = spreadSheetData[i][j].displayValue;
      if (
        spreadSheetData[i][j].formula &&
        spreadSheetData[i][j].formula.startsWith("=")
      ) {
        console.log("check if formula exists");
        if (spreadSheetData[i][j].observables.length > 0) {
          console.log("check if observables exists ");
          spreadSheetData[i][j].observables.forEach((observable) => {
            observable.subscribe(() => {
              console.log("in obs in pop table func");
              let rowIndex = i;
              let colIndex = j - 1;
              let id = String(colHeaderArr[colIndex] + rowIndex);
              let item = document.createElement("td");
              item.setAttribute("id", id);
              item.innerHTML = spreadSheetData[i][j].formula;
              evaluateEachCell(item);
              item.remove();
            });
          });
        } else {
          console.log("No obs for formula");
          spreadSheetData[i][j].displayValue = "!ERR";
          cell.innerHTML = "!ERR";
          spreadSheetData[i][j].observables = [];
        }
      }
    }
  }
};

// get observable for each cell
getCellObs = (cellId) => {
  console.log(" getCellObs  func");
  return rxjs.fromEvent(document.getElementById(cellId), "focusout");
};

//Reset selection
unselect = () => {
  selectedRow = -1;
  selectedCol = -1;
  document.querySelectorAll(".highlight").forEach((node) => {
    node.classList.remove("highlight");
  });
};

//set class as highlighted
selectCol = (colId) => {
  let spreadSheetData = this.fetchspreadSheetData();
  for (let i = 1; i < spreadSheetData.length; i++) {
    let doc = document.getElementById(`${colHeaderArr[colId - 1]}${i}`);
    doc.className = "highlight";
  }
};

selectRow = (parent) => {
  console.log(parent);
  console.log(children);
  var children = parent.childNodes;
  let arr = [...children];
  for (let i = 0; i < arr.length; i++) {
    console.log(arr[i]);
    arr[i].classList.toggle("highlight");
  }
};

//start
displaySpreadSheet("initialise");
