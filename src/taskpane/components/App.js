import * as React from "react";
import PropTypes from "prop-types";
import Progress from "./Progress";
import { users } from "../common/data";
// import axios from "../common/axios";
/* global console, Excel, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      users: users,
    };
  }

  async componentDidMount() {
    try {
      //await this.getUsers();
      await Excel.run(async (context) => {
        // Getting current worksheet
        //debugger;
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        // Changing sheet name
        worksheet.name = "Users";
        const tableExists = worksheet.tables.getItemOrNullObject("UserTable").load("items");
        await context.sync();
        console.log("Table Exists : ", tableExists.isNullObject);

        if (!tableExists.isNullObject) {
          tableExists.delete();
        }
        const userTable = worksheet.tables.add("A1:B1", true);
        userTable.name = "UserTable";
        // userTable.onChanged = onTableChanged();
        userTable.getHeaderRowRange().values = [["Name", "Email"]];
        userTable.rows.add(null, this.state.users, true);
        worksheet.getUsedRange().format.autofitColumns();
        worksheet.getUsedRange().format.autofitRows();
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <h1>Welcome to My-Add-In</h1>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
