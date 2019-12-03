import { Component } from "@angular/core";
const template = require("./app.component.html");
/* global console, Excel, require */
import * as moment from 'moment-msdate';
import { ContextReplacementPlugin } from "webpack";

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";
  isFilterApplied

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
}
