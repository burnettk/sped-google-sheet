function onEdit(e) {
    const sheetName = 'Log';
    const dataSheetName = 'Accommodations';
    const studentColumn = 1;
    const accommodationColumn = 2;
    const dateColumn = 5;
    const studentRange = 'A:A';
    const accommodationRange = 'B:B';

    const sheet = e.range.getSheet();
    if (sheet.getName() !== sheetName) return;

    const editedCell = e.range;
    const editedColumn = editedCell.getColumn();

    if (editedColumn === studentColumn) {
        const student = editedCell.getValue();
        const dataSheet = e.source.getSheetByName(dataSheetName);

        const accommodations = dataSheet.getRange(studentRange).getValues()
            .reduce((acc, row, index) => {
                if (row[0] === student) {
                    acc.push(dataSheet.getRange(accommodationRange).getCell(index + 1, 1).getValue());
                }
                return acc;
            }, []);

        const accommodationCell = sheet.getRange(editedCell.getRow(), accommodationColumn);
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(accommodations).build();
        accommodationCell.clearContent().setDataValidation(rule);

        const dateCell = sheet.getRange(editedCell.getRow(), dateColumn);
        if (!dateCell.getValue()) {
            dateCell.setValue(new Date());
        }
    }
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('SPED Actions')
        .addItem('Fill Accommodations for Student', 'fillAccommodations')
        .addItem('Shortened accommodations', 'fillShortestString')
        .addToUi();
}

function fillShortestString() {
    const sheetName = 'Accommodations'; // Ensure this is the correct sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    
    for (let row = 2; row <= lastRow; row++) { // Start from row 2 to skip the header row
        const colBValue = sheet.getRange(row, 2).getValue().toString().trim(); // Column B

        if (colBValue) {
            continue; // Skip rows where Column B already has a value
        }

        const colCValue = sheet.getRange(row, 3).getValue().toString().trim(); // Column C
        const colDValue = sheet.getRange(row, 4).getValue().toString().trim(); // Column D

        let shortestString = '';

        if (colCValue && colDValue) {
            shortestString = colCValue.length <= colDValue.length ? colCValue : colDValue;
        } else if (colCValue) {
            shortestString = colCValue;
        } else if (colDValue) {
            shortestString = colDValue;
        }

        if (shortestString) {
            sheet.getRange(row, 2).setValue(shortestString); // Fill column 2 with the shortest string
        }
    }
}

function fillAccommodations() {
    const sheetName = 'Log';  // Capitalized Log sheet
    const dataSheetName = 'Accommodations';
    const studentColumn = 1;
    const accommodationColumn = 2;
    const dateColumn = 5;
    const studentRange = 'A:A';
    const accommodationRange = 'B:B';

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
    
    const lastRow = sheet.getLastRow();
    let targetRow = null;

    for (let row = 1; row <= lastRow; row++) {
        const student = sheet.getRange(row, studentColumn).getValue();
        const accommodation = sheet.getRange(row, accommodationColumn).getValue();
        const nextRowStudent = sheet.getRange(row + 1, studentColumn).getValue();
        
        if (student && !accommodation && !nextRowStudent) {
            targetRow = row;
            break;
        }
    }

    if (targetRow === null) {
        SpreadsheetApp.getUi().alert('No suitable row found to fill accommodations.');
        return;
    }

    const student = sheet.getRange(targetRow, studentColumn).getValue();
    const accommodations = dataSheet.getRange(studentRange).getValues()
        .reduce((acc, row, index) => {
            if (row[0] === student) {
                acc.push(dataSheet.getRange(accommodationRange).getCell(index + 1, 1).getValue());
            }
            return acc;
        }, []);
    
    if (accommodations.length === 0) {
        SpreadsheetApp.getUi().alert('No accommodations found for the selected student.');
        return;
    }
    
    accommodations.forEach((accommodation, index) => {
        const currentRow = targetRow + index;
        sheet.getRange(currentRow, studentColumn).setValue(student);
        sheet.getRange(currentRow, accommodationColumn).setValue(accommodation);
        const dateCell = sheet.getRange(currentRow, dateColumn);
        if (!dateCell.getValue()) {
            dateCell.setValue(new Date());
        }
    });

    // Clear the student cell in the row below the inserted data to keep a blank row
    sheet.getRange(targetRow + accommodations.length, studentColumn).clearContent();
}
