const fs = require('fs');
const xlsx = require('xlsx');    //library to read excel sheet

function employeeData(filePath){

const workbook = xlsx.readFile(filePath);  
const sheetname = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetname];

const jsonData = xlsx.utils.sheet_to_json(sheet);  //converting data to json format



function excelDateToJSDate(excelDate) {
    return new Date((excelDate - (25567 + 1)) * 86400 * 1000);
}

function sameDay(startDate,endDate){
    return startDate.toDateString() === endDate.toDateString();

}

    function employeeWork(){
        console.log('Started...')

        // console.log('jsonData:', jsonData);
        console.log('Number of entries:', jsonData.length);

        // Sort entries based on 'Time' column
    jsonData.sort((entry1, entry2) => {
        const time1 = excelDateToJSDate(entry1['Time']);
        const time2 = excelDateToJSDate(entry2['Time']);
        return time1 - time2;
    });

    


        for(const entry of jsonData)
        {
            //  console.log('Entry:', entry);
            const position = entry['Position ID'];
            const name = entry['Employee Name'];
            const timeIn = excelDateToJSDate(entry['Time']);
            const timeOut = excelDateToJSDate(entry['Time Out']);
            const timeCard = entry['Timecard Hours (as Time)']
            const timecardHoursNumeric = parseInt(timeCard.split(':')[0], 10);
            console.log();
            const workTime = (timeOut - timeIn) / (1000 * 60 * 60);

            const payCycleStart = new Date(entry['Pay Cycle Start']);
            const payCycleEnd = new Date(entry['Pay Cycle End']);
            
            let consecutiveDays = 0;

            let currentDate = payCycleStart;

        while (currentDate <= payCycleEnd) {

        if (sameDay(currentDate, timeIn) || sameDay(currentDate, timeOut)) {
            consecutiveDays++;
         }

        currentDate.setDate(currentDate.getDate() + 1);  //+1 to add the endDate
    }


            console.log('Consecutive Days Worked:', consecutiveDays);

            if(consecutiveDays >= 7)
            {
                console.log(`${position} ${name} - Worked 7 consecutive days.`)
            }

            if(workTime < 10 && workTime > 1)
            {
                console.log(`${position} ${name} - Worked less than 10 hours but greater than an hour.`)
            }

            if(timecardHoursNumeric > 14)
            {
                console.log(`${position} ${name} -  14 HOURS IN A SINGLE SHIFT.`)
            }
        }
    }

    employeeWork();
}

const filePath = 'data.xlsx';
employeeData(filePath);