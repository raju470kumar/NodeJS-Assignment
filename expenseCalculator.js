// Requiring the module
const reader = require('xlsx')

// Reading our expense sheet
const file = reader.readFile('./NodeJS Assessment L1.xlsx')


let A = [0];
let B = [0];
let C = [0];
let D = [0];
let E = [0];

let expenseAmount = 0;
let payerPerson;

const sheets = file.SheetNames
let arrayTwod = new Array(5);
let originalSheet;

for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
    originalSheet = temp;
    for (const res of temp) {
        if (res.Date == undefined) break;
        for (const element in res) {
            if (element == 'Person') payerPerson = res[element];

            if (element == 'Expense') expenseAmount = res[element];

            if (element == 'Split between') {
                let contributors = res[element].split(',');
                let length = contributors.length;
                for (contri of contributors) {
                    let index;
                    if (contri == 'A') index = 0;
                    else if (contri == 'B') index = 1;
                    else if (contri == 'C') index = 2;
                    else if (contri == 'D') index = 3;
                    else if (contri == 'E') index = 4;

                    if (payerPerson == 'A') {
                        if (A[index] != undefined) A[index] = A[index] + Number(expenseAmount / length);
                        else A[index] = Number(expenseAmount / length);
                    }
                    else if (payerPerson == 'B') {
                        if (B[index] != undefined) B[index] = B[index] + Number(expenseAmount / length);
                        else B[index] = Number(expenseAmount / length);
                    }
                    else if (payerPerson == 'C') {
                        if (C[index] != undefined) C[index] = C[index] + Number(expenseAmount / length);
                        else C[index] = Number(expenseAmount / length);
                    }
                    else if (payerPerson == 'D') {
                        if (D[index] != undefined) D[index] = D[index] + Number(expenseAmount / length);
                        else D[index] = Number(expenseAmount / length);
                    }
                    else if (payerPerson == 'E') {
                        if (E[index] != undefined) E[index] = E[index] + Number(expenseAmount / length);
                        else E[index] = Number(expenseAmount / length);
                    }
                }
            }
        }
    }
}


// Loop to create 2D array using 1D array
for (let i = 0; i < arrayTwod.length; i++) {
    arrayTwod[i] = [];
}

//Make array into 2d array to calculate individual owes
for (let i = 0; i < 5; i++) {
    for (let j = 0; j < 5; j++) {
        if (i == 0) arrayTwod[i][j] = A[j];
        if (i == 1) arrayTwod[i][j] = B[j];
        if (i == 2) arrayTwod[i][j] = C[j];
        if (i == 3) arrayTwod[i][j] = D[j];
        if (i == 4) arrayTwod[i][j] = E[j];

    }
}

//main logic to get indivial owes
for (let i = 0; i < 5; i++) {
    for (let j = 0; j < 5; j++) {
        arrayTwod[i][j] = arrayTwod[j][i] - arrayTwod[i][j];
    }
}

for (const res of originalSheet) {
    let firstPersonIndex = String(res['Expected Output']).charCodeAt(0) - 65;
    let secondPersonIndex = String(res['__EMPTY_2']).charCodeAt(0) - 65;

    if (firstPersonIndex > 5) continue;
    if (arrayTwod[firstPersonIndex][secondPersonIndex] < 0) res.__EMPTY_1 = "will receive from";
    res.__EMPTY_3 = Math.abs(arrayTwod[firstPersonIndex][secondPersonIndex]);
}

const ws = reader.utils.json_to_sheet(originalSheet);
reader.utils.book_append_sheet(file, ws, `Sheet2`);

reader.writeFile(file, `./NodeJS Assessment L1-op.xlsx`);