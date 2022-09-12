import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "ARS Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};
// eslint-disable-next-line office-addins/no-office-initialize
// Office.initialize = async () => {
//   // eslint-disable-next-line no-undef
//   try {
//     //await this.getUsers();
//     await Excel.run(async (context) => {
//       // Getting current worksheet
//       //debugger;
//       const worksheet = context.workbook.worksheets.getActiveWorksheet();
//       // Changing sheet name
//       worksheet.name = "Users";
//       const tableExists = worksheet.tables.getItemOrNullObject("UserTable").load("items");
//       await context.sync();
//       console.log("Table Exists : ", tableExists.isNullObject);

//       if (!tableExists.isNullObject) {
//         tableExists.delete();
//       }
//       const userTable = worksheet.tables.add("A1:B1", true);
//       userTable.name = "UserTable";
//       // userTable.onChanged = onTableChanged();
//       userTable.getHeaderRowRange().values = [["Name", "Email"]];
//       userTable.rows.add(null, this.state.users, true);
//       worksheet.getUsedRange().format.autofitColumns();
//       worksheet.getUsedRange().format.autofitRows();
//       await context.sync();
//     });
//   } catch (error) {
//     console.error(error);
//   }
// };
/* Render application after Office initializes */
Office.onReady(() => {
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  isOfficeInitialized = true;
  render(App);
});

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
