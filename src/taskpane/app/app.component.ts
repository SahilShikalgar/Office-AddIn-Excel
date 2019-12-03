import { Component } from "@angular/core";
import { async } from "q";
const template = require("./app.component.html");
/* global console, Excel, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";
  sorted = false;

  async run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }

  async createTable() {
    try {
      await Excel.run(async context => {
        const currentWorkSheet = context.workbook.worksheets.getActiveWorksheet();
        const expenseTable = currentWorkSheet.tables.add("A1:D1", true /*hasHeaders*/);
        expenseTable.name = "ExpensesTable";

        const now = new Date();

        expenseTable.getHeaderRowRange().values = [[
          "Date", "Merchant", "Category", "Amount"
        ]];

        const rowsData = [];

        for (let i = 0; i < 7; i++) {
          rowsData.push(new Array(
            now.toUTCString(), `The Phone Company ${i+1}`, `Communication Channel ${i+1}`, (23 * i+1).toString()
          ));
        }
  
        expenseTable.rows.add(null /*add at the end */, rowsData);

        expenseTable.columns.getItemAt(3).getRange().numberFormat = [["$#,##0.00"]];
        expenseTable.getRange().format.autofitColumns();
        expenseTable.getRange().format.autofitRows();

        return context.sync();
      });
    } catch (error) {
      this.clearTable();
      console.error(error);
    }
  }

  async clearTable() {
    await Excel.run(async context => {
      const currentWorkSheet = context.workbook.worksheets.getActiveWorksheet();
      const expenseTable = currentWorkSheet.tables.getItem("ExpensesTable");
      expenseTable.delete();

      return context.sync();
    });
    this.createTable();
  }

  async filterTable() {
    try {
      await Excel.run(async context => {
        const currentWorkSheet = context.workbook.worksheets.getActiveWorksheet();
        const expenseTable = currentWorkSheet.tables.getItem("ExpensesTable");
        const AmountFilter = expenseTable.columns.getItem("Category").filter;
        AmountFilter.applyValuesFilter(["Communication Channel 3", "Communication Channel 4", "Communication Channel 7"]);

        return context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }

  async clearFilter() {
    try {
      await Excel.run(async context => {
        const currentWorkSheet = context.workbook.worksheets.getActiveWorksheet();
        const expenseTable = currentWorkSheet.tables.getItem("ExpensesTable");
        // const AmountFilter = expenseTable.columns.getItem("Category").filter;
        // AmountFilter.clear();
        expenseTable.clearFilters();

        return context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }

  async sortTable() {
    try {  
      await Excel.run(async context => {
        const currentWorkSheet = context.workbook.worksheets.getActiveWorksheet();
        const expenseTable = currentWorkSheet.tables.getItem("ExpensesTable");

        const sortingType = !this.sorted;
  
        const sortingFields = [
          {
            key: 1,
            ascending: sortingType
          }
        ];

        this.sorted = sortingType;

        expenseTable.sort.apply(sortingFields);
        return context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }

  async createChart() {
    console.log("createChart");
    try {
      await Excel.run(async context => {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
        const dataRange = expensesTable.getDataBodyRange();

        const chart = currentWorksheet.charts.add('ColumnClustered', dataRange);
        
        chart.setPosition("A11", "F30");
        chart.title.text = "Expenses";
        chart.legend.position = "Right"
        chart.legend.format.fill.setSolidColor("white");
        chart.dataLabels.format.font.size = 12;
        chart.dataLabels.format.font.color = "black";
        chart.series.getItemAt(0).name = 'Value in €';

        return context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }

  async freezeHeader() {
    try {
      await Excel.run(async context => {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        currentWorksheet.freezePanes.freezeRows(1);

        return context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }

  async cancelFreezeHeader() {
    try {
      await Excel.run(async context => {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        currentWorksheet.freezePanes.unfreeze();

        return context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }
}
