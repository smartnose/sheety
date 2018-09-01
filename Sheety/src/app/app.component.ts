import { Component } from '@angular/core';
import * as OfficeHelpers from '@microsoft/office-js-helpers';

const template = require('./app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';
    app = this;
    headerJson = "";

    async run() {
        try {
            await Excel.run(async context => {
                /**
                 * Insert your Excel code here
                 */
                const range = context.workbook.getSelectedRange();
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const headerRange = sheet.getRange('A1:Z1')
                headerRange.load('values')
                await context.sync()

                // Get the header values and trim away the empty cells
                const headers = headerRange
                .values
                .filter((v) => v.length > 0)
                .map((v) => v.filter(w => w !== '' && w !== false))
                const headersJson = JSON.stringify(headers)
                console.log(`headers: ${headersJson}`)
                this.headerJson = headersJson

                // Read the range address
                range.load('address');

                // Update the fill color
                range.format.fill.color = 'yellow';

                await context.sync();
                console.log(`The range address was ${range.address}.`);
            });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        };
    }
}