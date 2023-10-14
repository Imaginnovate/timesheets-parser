const _ = require("lodash");
const csvToJson = require("csv-file-to-json");
const moment = require("moment");
const excel = require('write-excel-file/node');
const logger = require('node-color-log');
const commander = require('commander');


main();

function main()
{
    commander
        .version('1.0.0', '-v, --version')
        .usage('[OPTIONS]...')
        .option('-i, --input <value>', 'Input CSV file name')
        .option('-o, --output <value>', 'Optional output file name')
        .parse(process.argv);

    const options = commander.opts();

    const inputFile = options.input;
    if (!inputFile) {
        logger.color('red').log('Please enter input file using -i option');
        return false;
    }

    logger.color('green').reverse().log('This script will parse same month time sheets effectively...');
    logger.log('');
    const dataInJSON = csvToJson({filePath: inputFile});

    const sortedData = _.sortBy(dataInJSON, ['Emp Code', 'Date']);

    if (sortedData.length < 1) {
        logger.color('red').log('No Data Found');
        return false;
    }

    const tempDate = moment(sortedData[0].Date, 'MMM DD, YYYY');
    const summary = [];
    const sheetData = []
    const sheetNames = ['Summary']
    const uniqueNames = _.uniqWith(sortedData.map(data => _.pick(data, ['Emp Code', 'Emp Name'])), _.isEqual);
    // Create Summary Object, Sheet Names Object and Empty Sheet Data.
    _.forEach(uniqueNames, name => {
        summary.push({
            code: _.get(name, 'Emp Code'),
            name: _.get(name, 'Emp Name'),
            totalHours: 0,
            daysAdded: 0,
            status: 'Existing Member'
        })
        sheetNames.push(_.get(name, 'Emp Name'));
        sheetData.push([]);
    })

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const startDate = tempDate.clone().startOf('month');
    const endDate = tempDate.clone().endOf('month');
    const month = tempDate.month();
    const year = tempDate.year();
    const holidays = _.filter(csvToJson({filePath: 'holidays.csv'}), h => {
        return _.startsWith(h.Date, `${year}-${_.padStart(_.toString(month+1), 2, "0")}`);
    })
    logger.color('blue').log(`Month identified as ${months[month]} ${year}`);
    _.forEach(holidays, h => {
        logger.color('blue').log(`Holiday found on ${h.Date} - ${h.Name}`);
    })
    logger.log('');
    let tempEmpName = '';
    _.forEach(sortedData, row => {
        const code = _.get(row, 'Emp Code');
        const index = _.findIndex(summary, ['code', code]);
        const empName = _.get(row, 'Emp Name');
        if (empName != tempEmpName) {
            tempEmpName = empName;
            logger.color('green').log(`Processing time sheets for ${empName} ...`);
        }
        const rowDate = moment(_.get(row, 'Date'), 'MMM DD, YYYY');
        if (rowDate < startDate || rowDate > endDate) {
            logger.color('red').log(`Out of range date found for ${empName}. Date: ${rowDate.format('MMM DD, YYYY')}`);
            return;
        }
        const rowDay = rowDate.date();
        let daysAdded = summary[index].daysAdded;
        const daysDifference = rowDay - daysAdded;
        // Add empty hours for non-working days
        if (daysDifference != 1) {
            for (i = daysAdded + 1; i < rowDay; i++) {
                sheetData[index].push(getEmptyRow(_.get(row, 'Emp Name'), i, months[month], year, holidays, month))
            }
            summary[index].daysAdded = i - 1;
        }
        const hours = _.toNumber(_.get(row, 'Task Hours'), 0);
        sheetData[index].push({
            name: empName,
            date: _.get(row, 'Date'),
            title: _.get(row, 'Title'),
            details: _.get(row, 'Description'),
            hours: hours,
        })
        summary[index].daysAdded = rowDay;
        summary[index].totalHours += hours;
    })

    // Add empty hours for non-working days at last
    logger.log('');
    logger.color('green').log('Preparing summary sheet...');
    let totalTeamHours = 0;
    const daysInMonth = tempDate.daysInMonth();
    _.forEach(summary, (row, index) => {
        const daysAdded = summary[index].daysAdded;
        const daysDifference = daysInMonth - daysAdded;
        for (i = daysAdded + 1; i <= daysInMonth; i++) {
            sheetData[index].push(getEmptyRow(_.get(row, 'name'), i, months[month], year, holidays, month))
        }
        sheetData[index].push({});
        sheetData[index].push({
            name: 'Total',
            hours: summary[index].totalHours,
        })
        totalTeamHours += summary[index].totalHours;
    })
    summary.push({
        name: 'Total',
        totalHours: totalTeamHours,
    })


    const outputFile = options.output || `${months[month]}_${year}_Time_Sheets.xlsx`;
    excel([summary].concat(sheetData), {
        schema: getSchemaArray(summary),
        sheets: sheetNames,
        stickyRowsCount: 1,
        filePath: outputFile,
    })

    logger.log('');
    logger.color('green').reverse().log(`Excel file created: ${outputFile}`);
}

function getEmptyRow(empName, day, month, year, holidays, monthNo) {
    const dateString = `${year}-${_.padStart(monthNo + 1, 2, '0')}-${_.padStart(day, 2, '0')}`;
    console.log(dateString);
    const holidayIndex = _.findIndex(holidays, ['Date', dateString]);
    let title = '';
    let details = '';
    if (holidayIndex != -1) {
        title = 'Holiday';
        details = holidays[holidayIndex].Name;
    } else {
        const dayOfWeek = moment(dateString, 'YYYY-MM-DD').isoWeekday();
        if (dayOfWeek > 5) {
            title = 'Weekend';
        } else {
            title = 'Leave';
        }
    }
    return {
        name: empName,
        date: `${month} ${_.padStart(_.toString(day), 2, '0')}, ${year}`,
        title,
        details,
        hours: '',
    }
}

function getSchemaArray(summary) {
    const schema = [
        {
            column: 'Name',
            type: String,
            value: data => _.get(data, 'name'),
            borderColor: '#000000',
            width: 20,
        },
        {
            column: 'Date',
            type: String,
            value: data => _.get(data, 'date'),
            borderColor: '#000000',
            width: 12,
        },
        {
            column: 'Title',
            type: String,
            value: data => _.get(data, 'title'),
            borderColor: '#000000',
            width: 18,
        },
        {
            column: 'Details',
            type: String,
            value: data => _.get(data, 'details'),
            borderColor: '#000000',
            width: 80,
        },
        {
            column: 'Hours',
            type: Number,
            value: data => _.get(data, 'hours'),
            borderColor: '#000000',
            width: 10,
        },
    ];

    const schemaArray = _.map(summary, s => {
        return schema;
    });

    const summarySchema = [
        {
            column: 'Name',
            type: String,
            value: data => _.get(data, 'name'),
            borderColor: '#000000',
            width: 20,
        },
        {
            column: 'Hours',
            type: Number,
            value: data => _.get(data, 'totalHours'),
            borderColor: '#000000',
            width: 10,
        },
        {
            column: 'New/ Existing member',
            type: String,
            value: data => _.get(data, 'status'),
            borderColor: '#000000',
            width: 20,
        },
    ]

    return [summarySchema].concat(schemaArray);
}
