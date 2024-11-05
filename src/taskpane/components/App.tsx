import { Button } from "@fluentui/react-components";
import * as React from "react";

import { faker } from "@faker-js/faker";
import { TableDataRow } from "./TableDataRow";
const ROW_MAX = 1_000;

export function App() {

    const [tableID, setTableID] = React.useState<string>();

    async function doCreateTable() {
        const tableID = await Excel.run(async (context) => {

            // Experiment - add the onDeleted EV handler
            context.workbook.tables.onDeleted.add(async (ev) => {
                if (ev.tableId === table.id) {
                    console.log(`Deleted table ${table.id}`);
                }
            });
            
            const leftColumn = 'A';
            const rightColumn = String.fromCharCode(leftColumn.charCodeAt(0) + TableDataRow.columns.length -1);
            const table = context.workbook.tables.add(`Sheet1!${leftColumn}1:${rightColumn}1`, true);
            table.getHeaderRowRange().values = [TableDataRow.columns];

            const rows = Array.from({ length: ROW_MAX }, () => {
                const rowData = TableDataRow.build();
                return TableDataRow.columns.map(fieldName => (rowData as any)[fieldName]);
            });
            table.rows.add(-1, rows);
            await context.sync();

            const eventListener = async (ev: Excel.TableChangedEventArgs) => {
                console.log(`ev.changeType=${ev.changeType}`);
                console.dir(ev);
            };
            table.onChanged.add(eventListener);
            
            return table.id; 
        });
        setTableID(tableID);
    }

    async function changeColumnData() {
        Excel.run(async (context) => {
            const table = context.workbook.tables.getItem(tableID!);
            const titleColumn = table.columns.getItem("jobTitle");
            titleColumn.load("values");
            await context.sync();

            const newValues = titleColumn.values.map((row, index) => {
                if (index === 0) return row; // Skip the header row
                return [faker.person.jobTitle()];
            });
        
            // Set the new values to the column
            titleColumn.values = newValues;
            await context.sync();
        })
    }

    return (
        <>
        <h1>Table Listener reproduction.</h1>

        <h2>Instructions</h2>

        <ol>
        <li>Delete any existing tables on the active worksheet.</li>
        <li>Click the <Button 
            disabled={!!tableID} 
            appearance="primary" 
            size="small"
            onClick={doCreateTable}>Create</Button> button to create a table holding random data.</li>
        <li>Open the JS console.</li>
        <li>Add or remove columns and rows and observe the events.</li>
        <li>Click this <Button appearance="primary" size="small" disabled={!tableID} onClick={changeColumnData}>button</Button> to programmatically change the <code>jobTitle</code> column , and observe the events.</li>
        </ol>

        </>
    );


}