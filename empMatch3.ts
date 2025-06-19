interface Benefit {
  account: string;
  description: string;
  amount: number;
}

class Employee {
  public id: number;
  public ukgId: number;
  public amounts: Benefit[];

  constructor(employeeID: number, public firstName: string, public lastName: string) {
    this.id = employeeID;
    this.amounts = [];
  }

  compareNames(name: string): boolean {
    if (cleanString(name).includes(cleanString(this.lastName))) {
      if (cleanString(name).includes(cleanString(this.firstName))) {
        return true;
      }
    }
    return false;
  }
}

class Report {
  public header: Array<string | number | boolean>;
  public data: Array<string | number | boolean>[];
  constructor(protected reportData: Array<string | number | boolean>[]) {
    this.header = reportData.shift();
    this.data = reportData;
  }

  getColumn(text: string) {
    const cleanHeader = this.header.map(h => cleanString(h.toString()));
    return cleanHeader.indexOf(text);
  }
}

class RosterReport extends Report {
  public empIDCol: number;
  public firstNameCol: number;
  public lastNameCol: number;
  constructor(public reportData: Array<string | number | boolean>[]) {
    super(reportData);
    this.empIDCol = this.getColumn("emp");
    this.firstNameCol = this.getColumn("firstname");
    this.lastNameCol = this.getColumn("lastname");
  }

  getAllEmployees() {
    return this.data.map(row => new Employee(Number(row[this.empIDCol]), String(row[this.firstNameCol]), String(row[this.lastNameCol])));
  }
}

class BeneReport extends Report {
  public empNameCol: number;
  public ukgEmpIDCol: number;
  public debitAmountCol: number;
  public creditAmountCol: number;
  public acctIDCol: number;
  public acctNameCol: number;
  public payDate: number;
  constructor(public reportData: Array<string | number | boolean>[]) {
    super(reportData);
    this.empNameCol = this.getColumn("employeename");
    this.ukgEmpIDCol = this.getColumn("employeeno");
    this.debitAmountCol = this.getColumn("debitamount");
    this.creditAmountCol = this.getColumn("creditamount");
    this.acctIDCol = this.getColumn("glaccountno");
    this.acctNameCol = this.getColumn("codedescriptionall");
    this.payDate = Number(this.data[0][this.getColumn("paydate")]);
  }

  getAllEmployees() {
    return this.data.map(row => {
      return { 
        name: String(row[this.empNameCol]),
        ukgID: Number(row[this.ukgEmpIDCol]),
        gl: String(row[this.acctIDCol]),
        desc: String(row[this.acctNameCol]),
        amount: Number(row[this.debitAmountCol]) === 0 ? -Number(row[this.creditAmountCol]) : Number(row[this.debitAmountCol])
      }
    });
  }
}

class NewSheet {
  public sheet: ExcelScript.Worksheet;
  constructor(workbook: ExcelScript.Workbook, sheetName: string, private data: Array<string | number>[]) {
    this.sheet = workbook.addWorksheet(sheetName);
    this.data = data;
  }

  build() {
    this.sheet.getRangeByIndexes(0, 0, this.data.length, this.data[0].length).setValues(this.data);
  }

  showData() {
    console.log(this.data);
  }
}

function cleanString(text: string): string {
  const alphaonly = text.replace(/\W/ig, "");
  const finalText = alphaonly.replace(/\s/g, "");
  return finalText.toLowerCase();
}

function dateToRef(date: number): string {
  const newDate = new Date(Math.round((date - 25569) * 86400 * 1000)).toJSON();
  const splitDate = String(newDate).split('T')[0];
  const [year, month, day] = splitDate.split('-');

  if (Number(day) - 15 > 5) {
    return `EOM${month}${year.slice(2)}`;
  } else {
    return `PR${month}${day}${year.slice(2)}`;
  }
}

function getSheets(workbook: ExcelScript.Workbook) {
  let sheets = workbook.getWorksheets();
  sheets.forEach(sheet => {
    if (sheet.getTabColor() != "") {
      const d = sheet.getUsedRange().getValues();
      const h = d.shift();

      switch (h.length) {
        case 8: sheet.setName("ROSTER");
          break;
        default: sheet.setName("REPORT");
          break;
      }
    }
  });
  return workbook.getWorksheets();
}

function main(workbook: ExcelScript.Workbook) {
  const sheets = getSheets(workbook);
  const rosterReport = workbook.getWorksheet("ROSTER").getUsedRange().getValues();
  const beneReport = workbook.getWorksheet("REPORT").getUsedRange().getValues();
  const roster = new RosterReport(rosterReport);
  const benes = new BeneReport(beneReport);
  const all_roster_employees = roster.getAllEmployees();
  const all_bene_employees = benes.getAllEmployees();

  all_bene_employees.forEach(row => {
    all_roster_employees.forEach(emp => {
      if(emp.compareNames(row.name)) {
        emp.ukgId = row.ukgID;
        emp.amounts.push({ account: row.gl, description: row.desc, amount: row.amount });
      }
    });
  });


  const matchedEmployees = all_roster_employees.filter(emp => emp.amounts.length > 0);
  const uploadSheetHeader = ["Reference #", "G/L Account", "Amount", "Control #", "Description"];
  const uploadSheetRows: Array<string|number>[] = [];
  matchedEmployees.forEach(emp => {
    emp.amounts.forEach(acct => {
      uploadSheetRows.push([dateToRef(benes.payDate), acct.account, acct.amount, emp.id, acct.description]);
    });
  });
  const uploadSheetData = [uploadSheetHeader, ...uploadSheetRows];
  const uploadSheet = new NewSheet(workbook, "UPLOAD", uploadSheetData);
  uploadSheet.build();

  const noMatch = all_roster_employees.filter(emp => emp.amounts.length === 0);
  const noMatchSheetHeader = ["Name", "Reynolds #"];
  const noMatchRows = noMatch.map(emp => {
    const name = [emp.firstName, emp.lastName].join(' ');
    return [name, emp.id];
  });
  const noMatchSheetData = [noMatchSheetHeader, ...noMatchRows];
  const noMatchSheet = new NewSheet(workbook, "No Match", noMatchSheetData);
  noMatchSheet.build();
}
