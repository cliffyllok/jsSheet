////////////////////////Constant definition section /////////////////////
//Ops regex for function matching
const RegexSUM = /^SUM\([A-Z]+[0-9]+([\:]*[A-Z]+[0-9]+)*\)/g;
const RegexAVG = /^AVG\([A-Z]+[0-9]+([\:]*[A-Z]+[0-9]+)*\)/g;

const WIDTH = 100;
const HEIGHT = 100;

////////////////////////Main cell object class starts here /////////////////////
class SheetCell {
  constructor(el, columnNo, rowNo) {
    this.textboxItem = el;
    
    if (el){      
    
      this.textboxItem.addEventListener("change", this.onTextboxValueChange, false);
      this.textboxItem.addEventListener("focus", this.onTextboxFocus, false);
      this.textboxItem.addEventListener("focusout", this.onTextboxFocusout, false);
    }
    this.formula = '';
    this.dependencyList=[];
    this.columnNo=columnNo;
    this.rowNo=rowNo;
  }
  
  changeTextbox(el){
    this.textboxItem = el;
    
    if (el){
      this.textboxItem.addEventListener("change", this.onTextboxValueChange, false);
      this.textboxItem.addEventListener("focus", this.onTextboxFocus, false);
      this.textboxItem.addEventListener("focusout", this.onTextboxFocusout, false);
    }
  }
  get displayValue(){
    //Get only property for calculating display value on the fly
    return this.calculateValue(this.formula);
  }
  refreshValue(){
    this.textboxItem.value = this.displayValue;
  }
  
  onTextboxFocusout = e=>{
    let cellValue=  this.textboxItem.value;
        this.textboxItem.value = this.displayValue;
  }
  
  onTextboxValueChange = e=>{
    let cellValue=  this.textboxItem.value;
    this.formula = cellValue;
    this.textboxItem.value = this.displayValue;
  }
  onTextboxFocus = e=>{
    let cellValue=  this.textboxItem.value;
    this.textboxItem.value = this.formula;
  }

  calculateValue(textinput){
    //Clear the dependencyList before recalculate the formula/function
    this.clearDependency();
    if (!textinput && textinput.length == 0)
      return '';
    if (textinput.startsWith('=')){
      if (textinput.length > 1)
        return this.executeFunction(textinput.substring(1)) ;
    }else{
      return textinput;
    }
  }
  
  executeFunction(textinput){
    try{
      if (textinput && textinput.length>0){
        let trimmedinput = textinput.trim().toUpperCase();
        return this.processOps(trimmedinput);
      }
    }
    catch(ex){
      //could add alert message, otherwise just set the display value to ''
    }
    return '';    
  }
  
  processOps(trimmedinput){      
      const regexSimpleOp = /^[A-Z]*[0-9]+[\+\-\*\/][A-Z]*[0-9]+/g;
			const found = trimmedinput.match(regexSimpleOp);
			if (found){
        try{
        	return this.simpleOps(trimmedinput);
        }catch (ex){
        	return '';
        }
      }
      return this.aggregateOps(trimmedinput);
  }
  aggregateOps(trimmedinput){
    let result = this.processAggregateOps(trimmedinput,RegexSUM, this.sumAll);
    if (result){
      return result;
    }
    result = this.processAggregateOps(trimmedinput,RegexAVG, this.avgAll);
    if (result){
      return result;
    }
    return trimmedinput;
  }
  
  processAggregateOps(trimmedinput,regexType, opfunction){
    //let regexSUM = /^SUM\([A-Z]+[0-9]+([\:]*[A-Z]+[0-9]+)*\)/g;
		let found = trimmedinput.match(regexType);
    
    if (found){      
      
      let rangeop = trimmedinput.substring(trimmedinput.indexOf('(')+1, trimmedinput.indexOf(')'));
      let rangeArray = this.parseRange(rangeop);
      this.dependencyList = rangeArray;
      rangeArray.forEach(obj => { obj.textboxItem.addEventListener("change", this.onTextboxFocusout, false)});
      return opfunction(rangeArray);
    }
    return null;
  }
  //////////////Actual function calculation logic////////////////////
  sumAll(cellrangearray){
    let result = 0;
    cellrangearray.forEach((obj) => {
  		result += parseInt(obj.displayValue.trim());
		});
    return result;
  }
  avgAll(cellrangearray){
    let result = 0;
    cellrangearray.forEach((obj) => {
  		result += parseInt(obj.displayValue.trim());
		});
    return result/cellrangearray.length;
  }
  /////////////Extract all objects in the range////////////////////
  parseRange(rangeop){
    let result =[];
    if(rangeop){
    	let trimrangeop = rangeop.trim();
    	let splitop = trimrangeop.split(':');
      if (splitop.length == 0){
        let cellobj = CellArray[splitop];
        if (cellobj){
        	result.push(CellArray[splitop]); 
        }
      }else if (splitop.length == 2){
        if(CellArray[splitop[0]] && CellArray[splitop[1]]){
          this.convertCellRangeToCellArray(CellArray[splitop[0]], CellArray[splitop[1]]).forEach((e)=>{result.push(e)});
        }
      }
    }
    return result;   
  }
  convertCellRangeToCellArray(range1, range2){
    let result = [];
    let maxColumn = range1.columnNo > range2.columnNo?range1.columnNo : range2.columnNo;
    let minColumn = range1.columnNo > range2.columnNo?range2.columnNo : range1.columnNo;
    let maxRow = range1.rowNo > range2.rowNo?range1.rowNo : range2.rowNo;
    let minRow = range1.rowNo > range2.rowNo?range2.rowNo : range1.rowNo;
    //1 based index
    for(let i =minColumn; i <= maxColumn; i++){
      for(let j=minRow; j <= maxRow; j++){
        result.push(CellArray[intToBase26String(i)+j]);
      }
    }
    return result;
  }
  ///////////////////////////Section for Simple Formula Operation//////////////////////
  simpleOps(trimmedinput){
    try {
    
      let op = this.convertOps(trimmedinput);
      switch(op[2]){
        case '+':
          return op[0] + op[1];  
				case '-':
          return op[0] - op[1];  
        case '*':
          return op[0] * op[1];  
				case '/':
          return op[0] / op[1];  
      }
      throw "invalid formula";
    }
    catch(ex){
      throw ex;
    }
  }
  convertOps(trimmedinput){
      let operandArray = trimmedinput.split('+');
      
      if (operandArray.length == 2){
        let convertedOpArray = this.parseOperands(operandArray);
        convertedOpArray[2] = '+';
        return convertedOpArray;
      }
      operandArray = trimmedinput.split('-');
      
      if (operandArray.length == 2){
        let convertedOpArray = this.parseOperands(operandArray);
        convertedOpArray[2] = '+';
        return convertedOpArray;
      }      
      operandArray = trimmedinput.split('*');
      
      if (operandArray.length == 2){
        let convertedOpArray = this.parseOperands(operandArray);
        convertedOpArray[2] = '*';
        return convertedOpArray;
      }
      operandArray = trimmedinput.split('/');
      if (operandArray.length == 2){
        let convertedOpArray = this.parseOperands(operandArray);
        convertedOpArray[2] = '/';
        return convertedOpArray;
      }
      throw "Invalid formula";
  }
  parseOperands(operandArray){
    let resultArray = [];
    try{
      if (operandArray.length == 2){
        let firstOperand = parseInt(operandArray[0]);
        if (isNaN(firstOperand)){
          firstOperand = CellArray[operandArray[0]];
          if(firstOperand){
            this.dependencyList.push(firstOperand);
            firstOperand.textboxItem.addEventListener("change", this.onTextboxFocusout, false)
            firstOperand = parseInt(firstOperand.displayValue);
            if (isNaN(firstOperand )){
              throw "invalid formula";
            }
          }else{
            throw "invalid formula";
          }
        }
        resultArray[0]=firstOperand;

        let secondOperand = parseInt(operandArray[1]);
        if (isNaN(secondOperand)){
          secondOperand = CellArray[operandArray[1]];
          if(secondOperand){
            this.dependencyList.push(secondOperand);
            secondOperand.textboxItem.addEventListener("change", this.onTextboxFocusout, false);
            secondOperand = parseInt(secondOperand.displayValue);
            if (isNaN(secondOperand)){
              throw "invalid formula";
            }
          }else{
            throw "invalid formula";
          }
        }
        resultArray[1]=secondOperand;
        return resultArray;
      }
    }catch(ex){
      throw ex;
    }
    return resultArray;
  }
  clearDependency(){
    this.dependencyList.forEach ((e)=>{
      e.textboxItem.removeEventListener("change", this.onTextboxFocusout);
    })
    this.dependencyList.splice(0, this.dependencyList.length);
	}
}
//////////////////End of SheetCell Class///////////////////////

///////////////////Page load starts here///////////////////////
const CellArray = {};

generateTable(WIDTH, HEIGHT);

let formObj = document.querySelector('#worksheetform');

let refreshButton = document.createElement("button");
refreshButton.setAttribute("id", "refreshButton");
refreshButton.setAttribute("name", "refreshButton");
refreshButton.setAttribute("class", "refreshButton");
refreshButton.addEventListener("click", RefreshGrid);
refreshButton.innerHTML = "Refresh";
formObj.appendChild(refreshButton);
///////////////////Page load ends here///////////////////////

///////////////////////////Refresh Grid////////////////////
function RefreshGrid(event){
  for (const property in CellArray) {
    CellArray[property].textboxItem = null;
	}
  document.querySelector('#worksheet').innerHTML='';
  generateTable(WIDTH, HEIGHT);  
  event.preventDefault();
}
/////////////////////////Table generation code
//////////////Top level table
function generateTable(width, height) {
  
  let tableObj = document.querySelector('#worksheet');
  generateHeaders(tableObj, width);
  generateBody(tableObj, width,height);
}

/////////////Header row
function generateHeaders(tableObj, width) {
  let header = tableObj.createTHead();
  let row = header.insertRow(0);    
  let headerrow='';
  let firstCell = row.insertCell(0);
  for (let i = 0; i < width; i++) {
    let header = generateHeader(i);
    let cell = row.insertCell(i+1);
    cell.innerHTML = header;
  }
}
////////////Get header string
function generateHeader(columnNo) {
  return intToBase26String(columnNo + 1);
}
//Turn 1 based index to Base26String
function intToBase26String(columnNo) {
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let result = '';
  // no letters for 0 or less
  if (columnNo < 1) {
    return result;
  }
  let quotient = columnNo;
  let remainder=0;
  // until we have a 0 quotient
  while (quotient !== 0) {
    //since start with 1, need to adjust to 0 based index
    let decremented = quotient - 1;
    quotient = Math.floor(decremented / 26);
    remainder = decremented % 26;
    result = letters[remainder] + result;
  }
  return result;
}

function generateBody(tableObj, width,height) {
  let body = tableObj.createTBody();   
  
  for (let j =0; j< height; j++){
    let row = body.insertRow(j);
    let firstCell = row.insertCell(0);
    let numberlabel = document.createElement("label");
    numberlabel.setAttribute("name", "verticallabel" + (j+1));
    numberlabel.setAttribute("id", "verticalLabel" + (j+1));
    numberlabel.innerHTML = (j+1);
    firstCell.appendChild(numberlabel);
    for (let i = 0; i < width; i++) {
      let cell = row.insertCell(i+1);
      let input = document.createElement("input");
      input.setAttribute("type", "text");
      input.setAttribute("id", "cell" + intToBase26String(i+1)+(j+1));
      input.setAttribute("name", "cell");
      input.value = '';
      if (CellArray[intToBase26String(i+1) + (j+1)]) {
        CellArray[intToBase26String(i+1) + (j+1)].changeTextbox(input);
        CellArray[intToBase26String(i+1) + (j+1)].refreshValue();
      }else{
        CellArray[intToBase26String(i+1) + (j+1)] = new SheetCell(input, i+1, j+1);
      }
      
    
      cell.appendChild(input);
    }  
  }
}

