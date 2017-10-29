app.service('XlsxProcService', ['$q', function ($q) {
    var service = {};
    var isIE = false;

    // Thie following constant defines the number of rows from the top of the range to search for the header row
    var MAX_ROWS_TO_SEARCH_FOR_HEADER = 10;
    // The following constant defines the number of cols in the file that should be matched to consider it to be the header row
    var PCT_COLS_TO_MATCH_FOR_HEADER = 10;
    // THe following constant defines the maximum characters from the left of the column name in the Xlsx file that must match the predefined column name to identify the column 
    var MAX_COL_LEN_TO_MATCH_FOR_HEADER = 20;
    // The default worksheet name that is assumed to contain the data to be parsed.
    var DEFAULT_WORKSHEET_NAME = 'Employee Data Requirements';

    var dbColumnsMetaData = [
        {
            normalizedColumnName: 'COUNTRY',
            actualColumnName: 'Country',
            displayColumnName: 'Country',
            datAttributeName: 'country',
            type: 'number'	
        },
        {
            normalizedColumnName: 'CURRENCY',
            actualColumnName: 'Currency ',
            displayColumnName: 'Currency',
            datAttributeName: 'currency',
            type: 'number'
        },
        {
            normalizedColumnName: 'COUNTRYTOTALREVENUETURNOVER',
            actualColumnName: 'Country Total Revenue / Turnover',
            displayColumnName: 'Country Total Revenue / Turnover',
            datAttributeName: 'countryTotalRevenueTurnOver',
            type: 'number'
        },
        {
            normalizedColumnName: 'UNITSINDUSTRYSECTOR',
            actualColumnName: 'Unit\'s Industry Sector',
            displayColumnName: 'Unit\'s Inustry Sector',
            datAttributeName: 'unitsIndustrySector',
            type: 'number'
        },
        {
            normalizedColumnName: 'UNITTOTALNUMBEROFEMPLOYEES',
            actualColumnName: 'Unit Total Number of Employees',
            displayColumnName: 'Unit Total Number of Employees',
            datAttributeName: 'unitTotalNumEmployees',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEE JOBTITLE',
            actualColumnName: 'Employee Job Title',
            displayColumnName: 'Employee Job Title',
            datAttributeName: 'employeeJobTitle',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEGRADEBAND',
            actualColumnName: 'Employee Grade / Band',
            displayColumnName: 'Employee Grade / Band',
            datAttributeName: 'employeeGradeBand',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEIDDONOTUSEEMPLOYEESNAMEORSOCIALSECURITYNUMBERDUETOINTERNATIONALPRIVACYLAW',
            actualColumnName: 'Employee ID                          (do not use employee\'s name or social security number due to international privacy law)',
            displayColumnName: 'Employee ID',
            datAttributeName: 'employeeId',
            type: 'number'
        },
        {
            normalizedColumnName: 'MANAGEREMPLOYEEIDDONOTUSEMANAGERSNAMEORSOCIALSECURITYNUMBERDUETOINTERNATIONALPRIVACYLAW',
            actualColumnName: 'Manager Employee ID                          (do not use manager\'s name or social security number due to international privacy law)',
            displayColumnName: 'Manager Employee ID',
            datAttributeName: 'managerEmployeeId',
            type: 'number'
        },
        {
            normalizedColumnName: 'REPORTINGLEVELFROMCEO',
            actualColumnName: 'Reporting Level from CEO',
            displayColumnName: 'Reporting Level from CEO',
            datAttributeName: 'reportingLevelFromCeo',
            type: 'number'
        },
        {
            normalizedColumnName: 'PERFORMANCERANKING',
            actualColumnName: 'Performance Ranking',
            displayColumnName: 'Performance Ranking',
            datAttributeName: 'performanceRanking',
            type: 'number'
        },
        {
            normalizedColumnName: 'GENDER',
            actualColumnName: 'Gender',
            displayColumnName: 'Gender',
            datAttributeName: 'gender',
            type: 'number'
        },
        {
            normalizedColumnName: 'CURRENTLEVELJOBSTARTDATE',
            actualColumnName: 'Current Level/Job Start Date',
            displayColumnName: 'Current Level / Job Start Date',
            datAttributeName: 'currentlevelJobStartDate',
            type: 'number'
        },
        {
            normalizedColumnName: 'COMPANYHIREDATEDATESTARTEDWITHCOMPANYNOTSPECIFICTOROLE',
            actualColumnName: 'Company Hire Date(Date started with company,not specific to role)',
            displayColumnName: 'Company Hire Date',
            datAttributeName: 'companyHireDate',
            type: 'number'
        },
        {
            normalizedColumnName: 'FULLTIMEPARTTIMESTATUS',
            actualColumnName: 'Full-time/Part-time Status',
            displayColumnName: 'Full-time / Part-time Status',
            datAttributeName: 'fullTimePartTimeStatus',
            type: 'number'
        },
        {
            normalizedColumnName: 'FTEPERCENTAGE',
            actualColumnName: 'FTE percentage',
            displayColumnName: 'FTE percentage',
            datAttributeName: 'ftePercentage',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEWORKLOCATIONZIPPOSTALCODE',
            actualColumnName: 'Employee WorkLocation / Zip/Postal Code',
            displayColumnName: 'Employee WorkLocation',
            datAttributeName: 'employeeWorkLocation',
            type: 'number'
        },
        {
            normalizedColumnName: 'JOBFAMILYSUBFAMILYCODENOTREQUIREDIFREFJOBCODEISPROVIDEDINCOLUMNM',
            actualColumnName: 'Job Family / Subfamily Code (not required if Ref. Job Code is provided in column M)',
            displayColumnName: 'Job Family / Subfamily Code',
            datAttributeName: 'jobFamily',
            type: 'number'
        },
        {
            normalizedColumnName: 'REFERENCELEVELNOTREQUIREDIFREFJOBCODEISPROVIDEDINCOLUMNM',
            actualColumnName: 'Reference Level  (not required if Ref. Job Code is provided in Column M)',
            displayColumnName: 'Reference Level',
            datAttributeName: 'referenceLevel',
            type: 'number'
        },
        {
            normalizedColumnName: 'REFERENCEJOBCODE',
            actualColumnName: 'Reference Job Code',
            displayColumnName: 'Reference Job Code',
            datAttributeName: 'referenceJobCode',
            type: 'number'
        },
        {
            normalizedColumnName: 'HAYGROUPTOTALPOINTS',
            actualColumnName: 'Hay Group Total Points',
            displayColumnName: 'Hay Group Total Points',
            datAttributeName: 'haygroupTotalPoints',
            type: 'number'
        },
        {
            normalizedColumnName: 'BASICPAYMENTS',
            actualColumnName: 'Basic Payments',
            displayColumnName: 'Basic Payments',
            datAttributeName: 'basicPayments',
            type: 'number'
        },
        {
            normalizedColumnName: 'FIXEDPAYMENTS',
            actualColumnName: 'Fixed Payments',
            displayColumnName: 'Fixed Payments',
            datAttributeName: 'fixedPayments',
            type: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMINIMUM',
            actualColumnName: 'Salary Structure Minimum',
            displayColumnName: 'Salary Structure Minimum',
            datAttributeName: 'salaryStructureMinimum',
            type: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMIDPOINT',
            actualColumnName: 'Salary Structure Midpoint',
            displayColumnName: 'Salary Structure Midpoint',
            datAttributeName: 'salaryStructureMidpoint',
            type: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMAXIMUM',
            actualColumnName: 'Salary Structure Maximum',
            displayColumnName: 'Salary Structure Maximum',
            datAttributeName: 'salaryStructureMaximum',
            type: 'number'
        },
        {
            normalizedColumnName: 'TOTALANNUALSHIFTPREMIUMSPAID',
            actualColumnName: 'Total Annual Shift Premiums Paid',
            displayColumnName: '',
            datAttributeName: 'totalAnnualShiftPremiumsPaid',
            type: 'number'
        },
        {
            normalizedColumnName: 'SHORTTERMVARIABLEPAYMENTELIGIBILITYYN',
            actualColumnName: 'Short-term Variable Payment Eligibility (Y/N)',
            displayColumnName: 'Short-term Variable Payment Eligibility',
            datAttributeName: 'shortTermVariablePaymentEligibility',
            type: 'number'
        },
        {
            normalizedColumnName: 'ACTUALANNUALSHORTTERMVARIABLEPAYMENT',
            actualColumnName: 'Actual Annual Short-term Variable Payment',
            displayColumnName: 'Actual Annual Short-term Variable Payment',
            datAttributeName: 'actualAnnualShortTermVariablePayment',
            type: 'number'
        },
        {
            normalizedColumnName: 'TARGETSHORTTERMVARIABLEPAYMENTOFBASESALARY',
            actualColumnName: 'Target Short-term Variable Payment (% of Base Salary)',
            displayColumnName: 'Target Short-term Variable Payment',
            datAttributeName: 'targetShortTermVariablePayment',
            type: 'number'
        },
        {
            normalizedColumnName: 'CARELIGIBILITYYN',
            actualColumnName: 'CarEligibility (Y/N)',
            displayColumnName: 'Car Eligibility',
            datAttributeName: 'carEligibility',
            type: 'number'
        },
        {
            normalizedColumnName: 'LONGTERMINCENTIVEELIGIBILITYYN',
            actualColumnName: 'Long-term Incentive Eligibility (Y/N)',
            displayColumnName: 'Long-term Incentive Eligibility',
            datAttributeName: 'longTermIncentiveEligibility',
            type: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVESOSARRSPSPU',
            actualColumnName: 'Type of Long-term Incentive (SO, SAR, RS, PS, PU)',
            displayColumnName: 'Type of Long-term Incentive',
            datAttributeName: 'typeOfLongTermIncentives1',
            type: 'number'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED',
            actualColumnName: 'Number of Shares / Options Granted',
            displayColumnName: 'Number of Shares / Options Granted - 1',
            datAttributeName: 'numberOfSharesOptionsGranted1',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE',
            actualColumnName: 'Grant Date',
            displayColumnName: 'Grant Date - 1',
            datAttributeName: 'grantedDate1',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY',
            actualColumnName: 'Grant    Price Currency',
            displayColumnName: 'Grant Price Currency - 1',
            datAttributeName: 'grantPriceCurrency1',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE',
            actualColumnName: 'Grant Price',
            displayColumnName: 'Grant Price - 1',
            datAttributeName: 'grantPrice1',
            type: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVESOSARRSPSPU',
            actualColumnName: 'Type of Long-term Incentive (SO, SAR, RS, PS, PU)',
            displayColumnName: 'Type of Long-term Incentive - 2',
            datAttributeName: 'typeOfLongTermIncentives2',
            type: 'number'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED',
            actualColumnName: 'Number of Shares / Options Granted',
            displayColumnName: 'Number of Shares / Options Granted - 2',
            datAttributeName: 'numberOfSharesOptionsGranted2',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE',
            actualColumnName: 'Grant Date',
            displayColumnName: 'Grant Date - 2',
            datAttributeName: 'grantDate2',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY',
            actualColumnName: 'Grant    Price Currency',
            displayColumnName: 'Grant Price Currency - 2',
            datAttributeName: 'grantPriceCurrency2',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE',
            actualColumnName: 'Grant Price',
            displayColumnName: 'Grant Price - 2',
            datAttributeName: 'grantPrice2',
            type: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVESOSARRSPSPU',
            actualColumnName: 'Type of Long-term Incentive (SO, SAR, RS, PS, PU)',
            displayColumnName: 'Type of Long-term Incentive - 3',
            datAttributeName: 'typeOfLongTermIncentives3',
            type: 'number'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED',
            actualColumnName: 'Number of Shares / Options Granted',
            displayColumnName: 'Number of Shares / Options Granted - 3',
            datAttributeName: 'numberOfSharesOptionsGranted3',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE',
            actualColumnName: 'Grant Date',
            displayColumnName: 'Grant Date - 3',
            datAttributeName: 'grantedDate3',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY',
            actualColumnName: 'Grant    Price Currency',
            displayColumnName: 'Grant Price Currency - 3',
            datAttributeName: 'grantPriceCurrency3',
            type: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE',
            actualColumnName: 'Grant Price',
            displayColumnName: 'Grant Price - 3',
            datAttributeName: 'grantPrice3',
            type: 'number'
        },
        {
            normalizedColumnName: 'ALLOWANCESELIGIBILITYYN',
            actualColumnName: 'AllowancesEligibility (Y/N)',
            displayColumnName: 'Allowances Eligibility',
            datAttributeName: 'allowancesEligibility',
            type: 'number'
        },
        {
            normalizedColumnName: 'CARALLOWANCE',
            actualColumnName: 'Car Allowance',
            displayColumnName: 'Car Allowance',
            datAttributeName: 'carAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'TRANSPORTATIONCOMMUTINGALLOWANCE',
            actualColumnName: 'Transportation / Commuting Allowance',
            displayColumnName: 'Transportation / Commuting Allowance',
            datAttributeName: 'transportationCommutingAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'REPRESENTATIONALLOWANCE',
            actualColumnName: 'Representation Allowance',
            displayColumnName: 'Representation Allowance',
            datAttributeName: 'representationAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'HOUSINGALLOWANCE',
            actualColumnName: 'Housing Allowance',
            displayColumnName: 'Housing Allowance',
            datAttributeName: 'housingAlloance',
            type: 'number'
        },
        {
            normalizedColumnName: 'MEALALLOWANCE',
            actualColumnName: 'Meal Allowance',
            displayColumnName: 'Meal Allowance',
            datAttributeName: 'mealAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEEDUCATIONALLOWANCE',
            actualColumnName: 'Employee Education Allowance',
            displayColumnName: 'Employee Education Allowance',
            datAttributeName: 'employeeEducationAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'DEPENDENTEDUCATIONALLOWANCE',
            actualColumnName: 'Dependent Education Allowance',
            displayColumnName: 'Dependent Education Allowance',
            datAttributeName: 'dependentEducationAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'TELECOMMUNICATIONALLOWANCE',
            actualColumnName: 'Telecommuni-cation Allowance',
            displayColumnName: 'Telecommunication Allowance',
            datAttributeName: 'telecommunicationAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'CLUBMEMBERSHIPALLOWANCE',
            actualColumnName: 'Club Membership Allowance',
            displayColumnName: 'Club Membership Allowance',
            datAttributeName: 'clubMembershipAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'ALLOTHERALLOWANCE',
            actualColumnName: 'All Other Allowance',
            displayColumnName: 'All Other Allowance',
            datAttributeName: 'allOtherAllowance',
            type: 'number'
        },
        {
            normalizedColumnName: 'NAMEOFCCNLAPPLIEDITALYONLY',
            actualColumnName: 'Name of CCNL Applied(Italy Only)',
            displayColumnName: 'Name of CCNL Applied (Italy Only)',
            datAttributeName: 'nameOfCCNLAppliedItalyOnly',
            type: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEECATEGORYITALYONLY',
            actualColumnName: 'Employee Category (Italy Only)',
            displayColumnName: 'Employee Category (Italy Only)',
            datAttributeName: 'employeeCategoryItalyOnly',
            type: 'number'
        },
        {
            normalizedColumnName: 'RETIREMENTPREMIUM',
            actualColumnName: 'Retirement Premium',
            displayColumnName: 'Retirement Premium',
            datAttributeName: 'retirementPremium',
            type: 'number'
        },
        {
            normalizedColumnName: 'MANDATORYPROFITSHARINGPARTICIPATIONELIGIBILITYFRANCEPERUONLYYN',
            actualColumnName: 'Mandatory Profit-Sharing Participation Eligibility (France / Peru Only)(Y/N)',
            displayColumnName: 'Mandatory Profit-Sharing Participation Eligibility (France / Peru Only)',
            datAttributeName: 'mandatoryProfitSharingParticipationEligibilityFrancePeruOnly',
            type: 'number'
        },
        {
            normalizedColumnName: 'MANDATORYPROFITSHARINGPARTICIPATIONPAYMENTFRANCEPERUONLY',
            actualColumnName: 'Mandatory Profit-Sharing Participation Payment(France / Peru Only)',
            displayColumnName: 'Mandatory Profit-Sharing Participation Payment(France / Peru Only)',
            datAttributeName: 'mandatoryProfitSharingParticipationPaymentFrancePeruOnly',
            type: 'number'
        },
        {
            normalizedColumnName: 'VOLUNTARYPROFITSHARINGINTRESSEMENTELIGIBILITYFRANCEONLYYN',
            actualColumnName: 'Voluntary Profit-Sharing Intéressement Eligibility (France Only)(Y/N)',
            displayColumnName: 'Voluntary Profit-Sharing Intéressement Eligibility (France Only)',
            datAttributeName: 'voluntaryProfitSharingIntressementEligibilityFranceOnly',
            type: 'number'
        },
        {
            normalizedColumnName: 'VOLUNTARYPROFITSHARINGINTRESSEMENTPAYMENTFRANCEONLY',
            actualColumnName: 'Voluntary Profit-Sharing Intéressement Payment(France Only)',
            displayColumnName: 'Voluntary Profit-Sharing Intéressement Payment (France Only)',
            datAttributeName: 'voluntaryProfitSharingIntressementPaymentFranceOnly',
            type: 'number'
        },
        {
            normalizedColumnName: 'EXPATYN',
            actualColumnName: 'Expat (Y/N)',
            displayColumnName: 'Expat',
            datAttributeName: 'expat',
            type: 'number'
        },
        {
            normalizedColumnName: 'NATIONALITY',
            actualColumnName: 'Nationality',
            displayColumnName: 'Nationality',
            datAttributeName: 'nationality',
            type: 'number'
        }
    ];

    service.processXlsx = function(file) {
        var def = $q.defer();

        readXlsx(file).then(function(workbook) {
            // RM: All processing of the workbook happens here...
            var worksheet = workbook.Sheets[DEFAULT_WORKSHEET_NAME];
            if (!worksheet) {
                def.reject({ error: 'Unable to find the worksheet \'' + DEFAULT_WORKSHEET_NAME + '\' in the Xlsx File' });
                return false;
            } 
            
            if (!worksheet['!ref']) {
                def.reject({ error: 'Unable to find the header Row or content in the Xlsx File' });
                return false;
            }

            var dataRange = XLSX.utils.decode_range(worksheet['!ref']);
            var headerRow = getIdentifiedHeaderRow(worksheet, dataRange);
            if (!headerRow) {
                def.reject({ error: 'Unable to find header Row in Xlsx File' });
                return false;
            }

            var header = getHeader(worksheet, datRange, headerRow);

//            var worksheetdata = getWorksheetData(worksheet, dataRange, headerRow);

            console.log('dataRange', dataRange);
            console.log('headerRow', headerRow);
            def.resolve({ headerRow: headerRow });
            return true;
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
        }, function(e) {
            def.reject({ error: 'Unable to interpret file as a Microsoft Excel (.xlsx) Workbook' });
        });

        return def.promise;
    }

    function getWorksheetData(worksheet, dataRange, headerRow) {
        var rowRangeBegin = headerRow + 1;
        var rowRangeEnd   = range.e.r;
        var colRangeBegin = range.s.c;
        var colRangeEnd   = range.e.c;
        
        for (var row = rowRangeBegin; row < rowRangeEnd; row++) {
            for (var col = colRangeBegin; col < colRangeEnd; col++) {
            }
        }        
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
            var workbook = null;
            try {
                var bstr = isIE ? reader.content : reader.result;
                workbook = XLSX.read(bstr, {type:'binary'});
                if (workbook) {
                    def.resolve(workbook);
                } else {
                    throw 'Unable to create workbook object';
                }
            } catch(e) {
                def.reject(e);
            }
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
