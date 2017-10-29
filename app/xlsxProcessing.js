app.service('XlsxProcService', ['$q', function ($q) {
    var service = {};
    var isIE = false;
    var dbColumnsMetaData = [
        {
            normalizedColumnName: 'COUNTRY',
            actualColumnName: 'Country',
            displayColumnName: '',
            type: 'number'	
        },
        {
            normalizedColumnName: 'CURRENCY',
            actualColumnName: 'Currency ',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'COUNTRYTOTALREVENUETURNOVER',
            actualColumnName: 'Country Total Revenue / Turnover',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'UNITSINDUSTRYSECTOR',
            actualColumnName: 'Unit\'s Industry Sector',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'UNITTOTALNUMBEROFEMPLOYEES',
            actualColumnName: 'Unit Total Number of Employees',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEE JOBTITLE',
            actualColumnName: 'Employee Job Title',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEGRADEBAND',
            actualColumnName: 'Employee Grade / Band',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEIDDONOTUSEEMPLOYEESNAMEORSOCIALSECURITYNUMBERDUETOINTERNATIONALPRIVACYLAW',
            actualColumnName: 'Employee ID                          (do not use employee\'s name or social security number due to international privacy law)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'MANAGEREMPLOYEEIDDONOTUSEMANAGERSNAMEORSOCIALSECURITYNUMBERDUETOINTERNATIONALPRIVACYLAW',
            actualColumnName: 'Manager Employee ID                          (do not use manager\'s name or social security number due to international privacy law)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'REPORTINGLEVELFROMCEO',
            actualColumnName: 'Reporting Level from CEO',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'PERFORMANCERANKING',
            actualColumnName: 'Performance Ranking',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GENDER',
            actualColumnName: 'Gender',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'CURRENTLEVELJOBSTARTDATE',
            actualColumnName: 'Current Level/Job Start Date',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'COMPANYHIREDATEDATESTARTEDWITHCOMPANYNOTSPECIFICTOROLE',
            actualColumnName: 'Company Hire Date(Date started with company,not specific to role)'
        },
        {
            normalizedColumnName: 'FULLTIMEPARTTIMESTATUS',
            actualColumnName: 'Full-time/Part-time Status',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'FTEPERCENTAGE',
            actualColumnName: 'FTE percentage',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEWORKLOCATIONZIPPOSTALCODE',
            actualColumnName: 'Employee WorkLocation / Zip/Postal Code',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'JOBFAMILYSUBFAMILYCODENOTREQUIREDIFREFJOBCODEISPROVIDEDINCOLUMNM',
            actualColumnName: 'Job Family / Subfamily Code (not required if Ref. Job Code is provided in column M)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'REFERENCELEVELNOTREQUIREDIFREFJOBCODEISPROVIDEDINCOLUMNM',
            actualColumnName: 'Reference Level  (not required if Ref. Job Code is provided in Column M)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'REFERENCEJOBCODE',
            actualColumnName: 'Reference Job Code',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'HAYGROUPTOTALPOINTS',
            actualColumnName: 'Hay Group Total Points',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'BASICPAYMENTS',
            actualColumnName: 'Basic Payments',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'FIXEDPAYMENTS',
            actualColumnName: 'Fixed Payments',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMINIMUM',
            actualColumnName: 'Salary Structure Minimum',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMIDPOINT',
            actualColumnName: 'Salary Structure Midpoint',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMAXIMUM',
            actualColumnName: 'Salary Structure Maximum',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'TOTALANNUALSHIFTPREMIUMSPAID',
            actualColumnName: 'Total Annual Shift Premiums Paid',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'SHORTTERMVARIABLEPAYMENTELIGIBILITYYN',
            actualColumnName: 'Short-term Variable Payment Eligibility (Y/N)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'ACTUALANNUALSHORTTERMVARIABLEPAYMENT',
            actualColumnName: 'Actual Annual Short-term Variable Payment',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'TARGETSHORTTERMVARIABLEPAYMENTOFBASESALARY',
            actualColumnName: 'Target Short-term Variable Payment (% of Base Salary)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'CARELIGIBILITYYN',
            actualColumnName: 'CarEligibility (Y/N)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'LONGTERMINCENTIVEELIGIBILITYYN',
            actualColumnName: 'Long-term Incentive Eligibility (Y/N)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVESOSARRSPSPU',
            actualColumnName: 'Type of Long-term Incentive (SO, SAR, RS, PS, PU)'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED',
            actualColumnName: 'Number of Shares / Options Granted ',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE',
            actualColumnName: 'Grant Date',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY',
            actualColumnName: 'Grant    Price Currency',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE',
            actualColumnName: 'Grant Price',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVESOSARRSPSPU',
            actualColumnName: 'Type of Long-term Incentive (SO, SAR, RS, PS, PU)'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED',
            actualColumnName: 'Number of Shares / Options Granted ',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE',
            actualColumnName: 'Grant Date',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY',
            actualColumnName: 'Grant    Price Currency',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE',
            actualColumnName: 'Grant Price',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVESOSARRSPSPU',
            actualColumnName: 'Type of Long-term Incentive (SO, SAR, RS, PS, PU)'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED',
            actualColumnName: 'Number of Shares / Options Granted ',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE',
            actualColumnName: 'Grant Date',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY',
            actualColumnName: 'Grant    Price Currency',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE',
            actualColumnName: 'Grant Price',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'ALLOWANCESELIGIBILITYYN',
            actualColumnName: 'AllowancesEligibility (Y/N)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'CARALLOWANCE',
            actualColumnName: 'Car Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'TRANSPORTATIONCOMMUTINGALLOWANCE',
            actualColumnName: 'Transportation / Commuting Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'REPRESENTATIONALLOWANCE',
            actualColumnName: 'Representation Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'HOUSINGALLOWANCE',
            actualColumnName: 'Housing Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'MEALALLOWANCE',
            actualColumnName: 'Meal Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEEDUCATIONALLOWANCE',
            actualColumnName: 'Employee Education Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'DEPENDENTEDUCATIONALLOWANCE',
            actualColumnName: 'Dependent Education Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'TELECOMMUNICATIONALLOWANCE',
            actualColumnName: 'Telecommuni-cation Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'CLUBMEMBERSHIPALLOWANCE',
            actualColumnName: 'Club Membership Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'ALLOTHERALLOWANCE',
            actualColumnName: 'All Other Allowance',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'NAMEOFCCNLAPPLIEDITALYONLY',
            actualColumnName: 'Name of CCNL Applied(Italy Only)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEECATEGORYITALYONLY',
            actualColumnName: 'Employee Category (Italy Only)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'RETIREMENTPREMIUM',
            actualColumnName: 'Retirement Premium ',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'MANDATORYPROFITSHARINGPARTICIPATIONELIGIBILITYFRANCEPERUONLYYN',
            actualColumnName: 'Mandatory Profit-Sharing Participation Eligibility (France / Peru Only)(Y/N)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'MANDATORYPROFITSHARINGPARTICIPATIONPAYMENTFRANCEPERUONLY',
            actualColumnName: 'Mandatory Profit-Sharing Participation Payment(France / Peru Only)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'VOLUNTARYPROFITSHARINGINTRESSEMENTELIGIBILITYFRANCEONLYYN',
            actualColumnName: 'Voluntary Profit-Sharing Intéressement Eligibility (France Only)(Y/N)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'VOLUNTARYPROFITSHARINGINTRESSEMENTPAYMENTFRANCEONLY',
            actualColumnName: 'Voluntary Profit-Sharing Intéressement Payment(France Only)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'EXPATYN',
            actualColumnName: 'Expat (Y/N)',
            displayColumnName: '',
            type: 'number'
        },
        {
            normalizedColumnName: 'NATIONALITY',
            actualColumnName: 'Nationality',
            displayColumnName: '',
            type: 'number'
        }
    ];

    // Thie following constant defines the number of rows from the top of the range to search for the header row
    var MAX_ROWS_TO_SEARCH_FOR_HEADER = 10;
    // The following constant defines the number of cols in the file that should be matched to consider it to be the header row
    var PCT_COLS_TO_MATCH_FOR_HEADER = 10;
    // THe following constant defines the maximum characters from the left of the column name in the Xlsx file that must match the predefined column name to identify the column 
    var MAX_COL_LEN_TO_MATCH_FOR_HEADER = 20;

    service.processXlsx = function(file) {
        var def = $q.defer();

        readXlsx(file).then(function(workbook) {
            console.log('workbook', workbook);
            // RM: All processing of the workbook happens here...
            var eeDataSheetName = 'Employee Data Requirements';
            var worksheet = workbook.Sheets[eeDataSheetName];
            if (worksheet) {
                var dataRange = XLSX.utils.decode_range(worksheet['!ref']);
                console.log('dataRange', dataRange);
                var headerRow = getIdentifiedHeaderRow(worksheet, dataRange);
                console.log('headerRow', headerRow);

                if (headerRow) {
                    def.resolve({ headerRow: headerRow });
                } else {
                    def.reject({ error: 'Unable to find header Row in Xlsx File' });
                }
            } else {
                def.reject({ error: 'Unable to find the worksheet \'Employee Data Requirements\' in the Xlsx File' });
            }
/*
            var row = 7;
            var countryColAddr = 'B';
            var currencyColAddr = 'C';
            var revenueColAddr = 'D';
            var industryColAddr = 'E';
            var numEesColAddr = 'F';
            var eeJobTitleColAddr = 'G';
            var eeGradeColAddr = 'H';
            var eeIdColAddr = 'I';

            var country    = getCellValueByRowCol(worksheet, countryColAddr, row);
            var currency   = getCellValueByRowCol(worksheet, currencyColAddr, row);
            var revenue    = getCellValueByRowCol(worksheet, revenueColAddr, row);
            var industry   = getCellValueByRowCol(worksheet, industryColAddr, row);
            var numEes     = getCellValueByRowCol(worksheet, numEesColAddr, row);
            var eeJobTitle = getCellValueByRowCol(worksheet, eeJobTitleColAddr, row);
            var eeGrade    = getCellValueByRowCol(worksheet, eeGradeColAddr, row);
            var eeId       = getCellValueByRowCol(worksheet, eeIdColAddr, row);

            console.log('country', country);
            console.log('currency', currency);
            console.log('revenue', revenue);
            console.log('industry', industry);
            console.log('numEes', numEes);
            console.log('eeJobTitle', eeJobTitle);
            console.log('eeGrade', eeGrade);
            console.log('eeId', eeId);
*/
        });

        return def.promise;
    }

    function getIdentifiedHeaderRow(worksheet, range) {
        var rowRangeBegin = range.s.r;
        var rowRangeEnd   = range.e.r < MAX_ROWS_TO_SEARCH_FOR_HEADER ? range.e.r: MAX_ROWS_TO_SEARCH_FOR_HEADER;
        var colRangeBegin = range.s.c;
        var colRangeEnd   = range.e.c;

        for (var row = rowRangeBegin; row < rowRangeEnd; row++) {
            var numColsIdentified = 0;
            for (var col = colRangeBegin; col < colRangeEnd; col++) {
                var cell = getCellByRowCol(worksheet, row, col);
                var value = cell ? getCellValue(cell): null;
                var type = cell ? getCellType(cell): null;
                if (value && type && type === 's' && isColumnIdentified(value)) {
                    numColsIdentified++;
                    console.log('getPctColsIdentified(numColsIdentified)', getPctColsIdentified(numColsIdentified));
                    console.log('numColsIdentified', numColsIdentified);
                    console.log('dbColumnsMetaData.length', dbColumnsMetaData.length);
                    if (getPctColsIdentified(numColsIdentified) >= PCT_COLS_TO_MATCH_FOR_HEADER) {
                        return row;
                    }
                }
            }
        }

        return null;
    }

    function getPctColsIdentified(numColsIdentified) {
        return numColsIdentified / dbColumnsMetaData.length * 100;
    }

    function isColumnIdentified(potentialColumnName) {
        var normalizedColumnName = getNormalizedColumnName(potentialColumnName);
        return dbColumnsMetaData.find(function(columnMetaData) {
            return (columnMetaData.normalizedColumnName.substring(0, MAX_COL_LEN_TO_MATCH_FOR_HEADER) === normalizedColumnName.substring(0, MAX_COL_LEN_TO_MATCH_FOR_HEADER)) ? true: false;
        });
    }

    function getNormalizedColumnName(columnName) {
        return columnName.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
    }

    function readXlsx(file) {
        // RM: You could replace the promise returned by $q with native promises if using this in node
        var def = $q.defer();
        var reader = new FileReader();
        reader.onload = function (e) {
            var bstr = isIE ? reader.content : reader.result;
            var workbook = XLSX.read(bstr, {type:'binary'});
            def.resolve(workbook);
        };
        reader.readAsBinaryString(file);
        return def.promise;
    }

    function getCellAddr(row, col) {
        return XLSX.utils.encode_cell({r: row, c: col});
    }

    function getCellValue(cell) {
        return cell.v;
    }

    function getCellType(cell) {
        return cell.t;
    }

    function getCellByAddr(worksheet, addr) {
        return (worksheet[addr] ? worksheet[addr] : undefined);
    }

    function getCellByRowCol(worksheet, row, col) {
        return getCellByAddr(worksheet, getCellAddr(row, col));
    }

    function isIEEvaluation() {
        var retVal = false;
        if (navigator.appName == 'Microsoft Internet Explorer' ||  !!(navigator.userAgent.match(/Trident/) || navigator.userAgent.match(/rv:11/)) || (typeof $.browser !== "undefined" && $.browser.msie == 1)) {
            retVal = true;
        }
        return retVal;
    }

    function init() {
        isIE = isIEEvaluation();
    }

    init();

    return service;
}]);
