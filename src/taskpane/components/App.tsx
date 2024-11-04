import { Button } from "@fluentui/react-components";
import * as React from "react";

import { faker } from "@faker-js/faker";
const ROW_MAX = 1_000;

class TableDataRow {

    private constructor(
        readonly prefix: string,
        readonly firstName: string,
        readonly familyName: string,
        readonly dob: string,
        readonly jobTitle: string,
    ) {}

    public static build() {
        return new TableDataRow(
            faker.person.prefix("male"),
            faker.person.firstName(),
            faker.person.lastName(),
            faker.date.birthdate().toISOString(),
            faker.person.jobTitle()
        );
    }
}

const columns = ["prefix", "firstName", "familyName", "dob", "jobTitle"];

export function App() {

    const [tableID, setTableID] = React.useState<string>();

    async function doCreateTable() {
        const tableID = await Excel.run(async (context) => {
            
            const table = context.workbook.tables.add("Sheet1!A1:E1", true);
            table.getHeaderRowRange().values = [columns];

            for(let i=0; i < ROW_MAX; i++) {
                const rowData = TableDataRow.build();
                const rowValues = columns.map(fieldName => (rowData as any)[fieldName]);
                table.rows.add(-1, [rowValues]);
            }
            await context.sync();           
            return table.id; 
        });
        setTableID(tableID);
    }

    return (
        <>
        <h1>Table Listener reproduction.</h1>

        <h2>Instructions</h2>
        <p>Click the <Button 
            disabled={!!tableID} 
            appearance="primary" 
            onClick={doCreateTable}>Create</Button> button to create a table holding random data.</p>
        </>
    );


}