const fs = require("fs");
const xlsx = require("xlsx");

// Function to convert Excel serial number to JavaScript Date object
function excelSerialNumberToDate(serialNumber) {
  const baseDate = new Date("1899-12-30");
  const millisecondsPerDay = 24 * 60 * 60 * 1000; 
  const milliseconds = Math.round(serialNumber * millisecondsPerDay); 
  return new Date(baseDate.getTime() + milliseconds);
}

function loadExcelData(file) {
  try {
    const workbook = xlsx.readFile(file);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    return data;
  } catch (error) {
    console.error(`Error loading data: ${error.message}`);
    return null;
  }
}

function convertToMinutes(timeString) {
  const [hours, minutes] = timeString.split(":").map(Number);
  return hours * 60 + minutes;
}

//parse time values and convert them to JavaScript Date objects
function parseDate(timeValue) {
  const excelSerialNumber = parseFloat(timeValue);
  return excelSerialNumberToDate(excelSerialNumber);
}

function analyzeData(data) {

  let consecutiveDaysCount = 0;
  let lastEmployee = null;
  let lastDay = null;
  const consecutiveDaysSet = new Set();
  const shiftsLessThan10HoursSet = new Set();
  const moreThan14HoursSet = new Set();
  let lastTimeOut = null;

  for (const row of data) {
    const employee = row["Employee Name"];
    const timeOut = parseDate(row["Time Out"]);
    const time = parseDate(row["Time"]);
    // Handle invalid date (e.g., when the 'Time Out' value is not a valid Excel serial number)
    if (isNaN(timeOut.getTime()) || isNaN(time.getTime())) {
      console.error(`Invalid date for employee ${employee}`);
      continue;
    }

    const day = time.toISOString().split("T")[0];

    // Check for consecutive days
    // console.log(`Employee: ${employee}, Day: ${day}`, "-------------------", "Last Employee: " + lastEmployee, "Last Day: " + lastDay);
    if (lastEmployee === employee && lastDay !== day) {
      const lastDate = new Date(lastDay);
      const currentDate = new Date(day);

      // Check if the current date is exactly one day after the last date
      if (lastDate.getTime() + 24 * 60 * 60 * 1000 === currentDate.getTime()) {
        consecutiveDaysCount++;
        // console.log(consecutiveDaysCount, "consecutive days");
      } else {
        consecutiveDaysCount = 1; // Reset count for a non-consecutive day
        // console.log("reset");
        // console.log(consecutiveDaysCount, "consecutive days");
      }
    }
    if (lastEmployee !== employee) {
    //   console.log("employee changed");
      consecutiveDaysCount = 1; // Reset count for a new employee
    //   console.log(consecutiveDaysCount, "consecutive days");
    }

    // a) Store employees who have worked for 7 consecutive days
    if (consecutiveDaysCount >= 7) {
      consecutiveDaysSet.add(employee);
    }

    // b) Store employees with less than 10 hours between shifts but greater than 1 hour
    if (employee === lastEmployee) {
      const timeDiff = lastTimeOut ? timeOut - lastTimeOut : 0;
      if (timeDiff < 10 * 60 * 60 * 1000 && timeDiff > 1 * 60 * 60 * 1000) {
        shiftsLessThan10HoursSet.add(employee);
      }
    }
    // c) Store employees who have worked for more than 14 hours in a single shift
    const timecardHours = row["Timecard Hours (as Time)"];
    const totalMinutes =
      typeof timecardHours === "string"
        ? convertToMinutes(timecardHours)
        : null;
    if (totalMinutes && totalMinutes > 14 * 60) {
      moreThan14HoursSet.add(employee);
    }

    // Update lastEmployee and lastTimeOut for the next iteration
    lastEmployee = employee;
    lastTimeOut = timeOut;
    lastDay = day;
  }

  // Check the last row for consecutive days
  if (consecutiveDaysCount === 7) {
    consecutiveDaysSet.add(lastEmployee);
  }

  return {
    consecutiveDays: Array.from(consecutiveDaysSet),
    shiftsLessThan10Hours: Array.from(shiftsLessThan10HoursSet),
    moreThan14Hours: Array.from(moreThan14HoursSet),
  };
}

function writeToOutputFile(output) {
  const outputPath = "output.txt";

  // Write to output file
  const formattedOutput = Object.entries(output)
    .map(
      ([category, employees]) =>
        `Category: ${category}\n${employees.join("\n")}`
    )
    .join("\n\n");

  fs.writeFileSync(outputPath, formattedOutput);
  console.log(`Results written to ${outputPath}`);
}

// Replace 'path/to/your/file.xlsx' with the actual path to your Excel file
const filePath = "./Assignment_Timecard.xlsx";
const excelData = loadExcelData(filePath);

if (excelData) {
  const analysisOutput = analyzeData(excelData);
  writeToOutputFile(analysisOutput);
  console.log("---------------------------------------------------");
  for(let key in analysisOutput) {
    console.log("Category: " + key);
    for(let i = 0; i < analysisOutput[key].length; i++) {
      console.log(analysisOutput[key][i]);
    }
    console.log("---------------------------------------------------");
  }
}
