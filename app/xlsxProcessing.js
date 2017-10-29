app.service('XlsxProcService', ['$q', function ($q) {
    var service = {};
    var isIE = false;

    // Thie following constant defines the number of rows from the top of the range to search for the header row
    var MAX_ROWS_TO_SEARCH_FOR_HEADER = 10;
    // The following constant defines the number of cols in the file that should be matched to consider it to be the header row
    var PCT_COLS_TO_MATCH_FOR_HEADER = 10;
    // THe following constant defines the maximum characters from the left of the column name in the Xlsx file that must match the predefined column name to identify the column 
    var MAX_COL_LEN_TO_MATCH_FOR_HEADER = 40;
    // The default worksheet name that is assumed to contain the data to be parsed.
    var DEFAULT_WORKSHEET_NAME = 'Employee Data Requirements';

    var DB_COLUMNS_METADATA = [
        {
            normalizedColumnName: 'COUNTRY',
            xlsxDisplayColumnName: 'Country',
            displayColumnName: 'Country',
            dataAttributeName: 'country',
            expectedType: 'number'	
        },
        {
            normalizedColumnName: 'CURRENCY',
            xlsxDisplayColumnName: 'Currency ',
            displayColumnName: 'Currency',
            dataAttributeName: 'currency',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'COUNTRYTOTALREVENUETURNOVER',
            xlsxDisplayColumnName: 'Country Total Revenue / Turnover',
            displayColumnName: 'Country Total Revenue / Turnover',
            dataAttributeName: 'countryTotalRevenueTurnOver',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'UNITSINDUSTRYSECTOR',
            xlsxDisplayColumnName: 'Unit\'s Industry Sector',
            displayColumnName: 'Unit\'s Inustry Sector',
            dataAttributeName: 'unitsIndustrySector',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'UNITTOTALNUMBEROFEMPLOYEES',
            xlsxDisplayColumnName: 'Unit Total Number of Employees',
            displayColumnName: 'Unit Total Number of Employees',
            dataAttributeName: 'unitTotalNumEmployees',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEJOBTITLE',
            xlsxDisplayColumnName: 'Employee Job Title',
            displayColumnName: 'Employee Job Title',
            dataAttributeName: 'employeeJobTitle',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEGRADEBAND',
            xlsxDisplayColumnName: 'Employee Grade / Band',
            displayColumnName: 'Employee Grade / Band',
            dataAttributeName: 'employeeGradeBand',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEIDDONOTUSEEMPLOYEESNAMEORSOCIALSECURITYNUMBERDUETOINTERNATIONALPRIVACYLAW',
            xlsxDisplayColumnName: 'Employee ID                          (do not use employee\'s name or social security number due to international privacy law)',
            displayColumnName: 'Employee ID',
            dataAttributeName: 'employeeId',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'MANAGEREMPLOYEEIDDONOTUSEMANAGERSNAMEORSOCIALSECURITYNUMBERDUETOINTERNATIONALPRIVACYLAW',
            xlsxDisplayColumnName: 'Manager Employee ID                          (do not use manager\'s name or social security number due to international privacy law)',
            displayColumnName: 'Manager Employee ID',
            dataAttributeName: 'managerEmployeeId',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'REPORTINGLEVELFROMCEO',
            xlsxDisplayColumnName: 'Reporting Level from CEO',
            displayColumnName: 'Reporting Level from CEO',
            dataAttributeName: 'reportingLevelFromCeo',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'PERFORMANCERANKING',
            xlsxDisplayColumnName: 'Performance Ranking',
            displayColumnName: 'Performance Ranking',
            dataAttributeName: 'performanceRanking',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GENDER',
            xlsxDisplayColumnName: 'Gender',
            displayColumnName: 'Gender',
            dataAttributeName: 'gender',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'CURRENTLEVELJOBSTARTDATE',
            xlsxDisplayColumnName: 'Current Level/Job Start Date',
            displayColumnName: 'Current Level / Job Start Date',
            dataAttributeName: 'currentlevelJobStartDate',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'COMPANYHIREDATEDATESTARTEDWITHCOMPANYNOTSPECIFICTOROLE',
            xlsxDisplayColumnName: 'Company Hire Date(Date started with company,not specific to role)',
            displayColumnName: 'Company Hire Date',
            dataAttributeName: 'companyHireDate',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'FULLTIMEPARTTIMESTATUS',
            xlsxDisplayColumnName: 'Full-time/Part-time Status',
            displayColumnName: 'Full-time / Part-time Status',
            dataAttributeName: 'fullTimePartTimeStatus',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'FTEPERCENTAGE',
            xlsxDisplayColumnName: 'FTE percentage',
            displayColumnName: 'FTE percentage',
            dataAttributeName: 'ftePercentage',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEWORKLOCATIONZIPPOSTALCODE',
            xlsxDisplayColumnName: 'Employee WorkLocation / Zip/Postal Code',
            displayColumnName: 'Employee WorkLocation',
            dataAttributeName: 'employeeWorkLocation',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'JOBFAMILYSUBFAMILYCODENOTREQUIREDIFREFJOBCODEISPROVIDEDINCOLUMNM',
            xlsxDisplayColumnName: 'Job Family / Subfamily Code (not required if Ref. Job Code is provided in column M)',
            displayColumnName: 'Job Family / Subfamily Code',
            dataAttributeName: 'jobFamily',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'REFERENCELEVELNOTREQUIREDIFREFJOBCODEISPROVIDEDINCOLUMNM',
            xlsxDisplayColumnName: 'Reference Level  (not required if Ref. Job Code is provided in Column M)',
            displayColumnName: 'Reference Level',
            dataAttributeName: 'referenceLevel',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'REFERENCEJOBCODE',
            xlsxDisplayColumnName: 'Reference Job Code',
            displayColumnName: 'Reference Job Code',
            dataAttributeName: 'referenceJobCode',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'HAYGROUPTOTALPOINTS',
            xlsxDisplayColumnName: 'Hay Group Total Points',
            displayColumnName: 'Hay Group Total Points',
            dataAttributeName: 'haygroupTotalPoints',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'BASICPAYMENTS',
            xlsxDisplayColumnName: 'Basic Payments',
            displayColumnName: 'Basic Payments',
            dataAttributeName: 'basicPayments',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'FIXEDPAYMENTS',
            xlsxDisplayColumnName: 'Fixed Payments',
            displayColumnName: 'Fixed Payments',
            dataAttributeName: 'fixedPayments',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMINIMUM',
            xlsxDisplayColumnName: 'Salary Structure Minimum',
            displayColumnName: 'Salary Structure Minimum',
            dataAttributeName: 'salaryStructureMinimum',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMIDPOINT',
            xlsxDisplayColumnName: 'Salary Structure Midpoint',
            displayColumnName: 'Salary Structure Midpoint',
            dataAttributeName: 'salaryStructureMidpoint',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'SALARYSTRUCTUREMAXIMUM',
            xlsxDisplayColumnName: 'Salary Structure Maximum',
            displayColumnName: 'Salary Structure Maximum',
            dataAttributeName: 'salaryStructureMaximum',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'TOTALANNUALSHIFTPREMIUMSPAID',
            xlsxDisplayColumnName: 'Total Annual Shift Premiums Paid',
            displayColumnName: '',
            dataAttributeName: 'totalAnnualShiftPremiumsPaid',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'SHORTTERMVARIABLEPAYMENTELIGIBILITYYN',
            xlsxDisplayColumnName: 'Short-term Variable Payment Eligibility (Y/N)',
            displayColumnName: 'Short-term Variable Payment Eligibility',
            dataAttributeName: 'shortTermVariablePaymentEligibility',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'ACTUALANNUALSHORTTERMVARIABLEPAYMENT',
            xlsxDisplayColumnName: 'Actual Annual Short-term Variable Payment',
            displayColumnName: 'Actual Annual Short-term Variable Payment',
            dataAttributeName: 'actualAnnualShortTermVariablePayment',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'TARGETSHORTTERMVARIABLEPAYMENTOFBASESALARY',
            xlsxDisplayColumnName: 'Target Short-term Variable Payment (% of Base Salary)',
            displayColumnName: 'Target Short-term Variable Payment',
            dataAttributeName: 'targetShortTermVariablePayment',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'CARELIGIBILITYYN',
            xlsxDisplayColumnName: 'CarEligibility (Y/N)',
            displayColumnName: 'Car Eligibility',
            dataAttributeName: 'carEligibility',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'LONGTERMINCENTIVEELIGIBILITYYN',
            xlsxDisplayColumnName: 'Long-term Incentive Eligibility (Y/N)',
            displayColumnName: 'Long-term Incentive Eligibility',
            dataAttributeName: 'longTermIncentiveEligibility',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVE1SOSARRSPSPU',
            xlsxDisplayColumnName: 'Type of Long-term Incentive - 1 (SO, SAR, RS, PS, PU)',
            displayColumnName: 'Type of Long-term Incentive - 1',
            dataAttributeName: 'TypeOfLongTermIncentives1',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED1',
            xlsxDisplayColumnName: 'Number of Shares / Options Granted - 1',
            displayColumnName: 'Number of Shares / Options Granted - 1',
            dataAttributeName: 'numberOfSharesOptionsGranted1',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE1',
            xlsxDisplayColumnName: 'Grant Date - 1',
            displayColumnName: 'Grant Date - 1',
            dataAttributeName: 'grantedDate1',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY1',
            xlsxDisplayColumnName: 'Grant    Price Currency - 1',
            displayColumnName: 'Grant Price Currency - 1',
            dataAttributeName: 'grantPriceCurrency1',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE1',
            xlsxDisplayColumnName: 'Grant Price - 1',
            displayColumnName: 'Grant Price - 1',
            dataAttributeName: 'grantPrice1',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVE2SOSARRSPSPU',
            xlsxDisplayColumnName: 'Type of Long-term Incentive - 2 (SO, SAR, RS, PS, PU)',
            displayColumnName: 'Type of Long-term Incentive - 2',
            dataAttributeName: 'TypeOfLongTermIncentives2',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED2',
            xlsxDisplayColumnName: 'Number of Shares / Options Granted - 2',
            displayColumnName: 'Number of Shares / Options Granted - 2',
            dataAttributeName: 'numberOfSharesOptionsGranted2',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE2',
            xlsxDisplayColumnName: 'Grant Date - 2',
            displayColumnName: 'Grant Date - 2',
            dataAttributeName: 'grantDate2',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY2',
            xlsxDisplayColumnName: 'Grant    Price Currency - 2',
            displayColumnName: 'Grant Price Currency - 2',
            dataAttributeName: 'grantPriceCurrency2',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE2',
            xlsxDisplayColumnName: 'Grant Price - 2',
            displayColumnName: 'Grant Price - 2',
            dataAttributeName: 'grantPrice2',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'TYPEOFLONGTERMINCENTIVE3SOSARRSPSPU',
            xlsxDisplayColumnName: 'Type of Long-term Incentive - 3 (SO, SAR, RS, PS, PU)',
            displayColumnName: 'Type of Long-term Incentive - 3',
            dataAttributeName: 'TypeOfLongTermIncentives3',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'NUMBEROFSHARESOPTIONSGRANTED3',
            xlsxDisplayColumnName: 'Number of Shares / Options Granted - 3',
            displayColumnName: 'Number of Shares / Options Granted - 3',
            dataAttributeName: 'numberOfSharesOptionsGranted3',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTDATE3',
            xlsxDisplayColumnName: 'Grant Date - 3',
            displayColumnName: 'Grant Date - 3',
            dataAttributeName: 'grantedDate3',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICECURRENCY3',
            xlsxDisplayColumnName: 'Grant    Price Currency - 3',
            displayColumnName: 'Grant Price Currency - 3',
            dataAttributeName: 'grantPriceCurrency3',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'GRANTPRICE3',
            xlsxDisplayColumnName: 'Grant Price - 3',
            displayColumnName: 'Grant Price - 3',
            dataAttributeName: 'grantPrice3',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'ALLOWANCESELIGIBILITYYN',
            xlsxDisplayColumnName: 'AllowancesEligibility (Y/N)',
            displayColumnName: 'Allowances Eligibility',
            dataAttributeName: 'allowancesEligibility',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'CARALLOWANCE',
            xlsxDisplayColumnName: 'Car Allowance',
            displayColumnName: 'Car Allowance',
            dataAttributeName: 'carAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'TRANSPORTATIONCOMMUTINGALLOWANCE',
            xlsxDisplayColumnName: 'Transportation / Commuting Allowance',
            displayColumnName: 'Transportation / Commuting Allowance',
            dataAttributeName: 'transportationCommutingAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'REPRESENTATIONALLOWANCE',
            xlsxDisplayColumnName: 'Representation Allowance',
            displayColumnName: 'Representation Allowance',
            dataAttributeName: 'representationAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'HOUSINGALLOWANCE',
            xlsxDisplayColumnName: 'Housing Allowance',
            displayColumnName: 'Housing Allowance',
            dataAttributeName: 'housingAlloance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'MEALALLOWANCE',
            xlsxDisplayColumnName: 'Meal Allowance',
            displayColumnName: 'Meal Allowance',
            dataAttributeName: 'mealAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEEEDUCATIONALLOWANCE',
            xlsxDisplayColumnName: 'Employee Education Allowance',
            displayColumnName: 'Employee Education Allowance',
            dataAttributeName: 'employeeEducationAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'DEPENDENTEDUCATIONALLOWANCE',
            xlsxDisplayColumnName: 'Dependent Education Allowance',
            displayColumnName: 'Dependent Education Allowance',
            dataAttributeName: 'dependentEducationAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'TELECOMMUNICATIONALLOWANCE',
            xlsxDisplayColumnName: 'Telecommuni-cation Allowance',
            displayColumnName: 'Telecommunication Allowance',
            dataAttributeName: 'telecommunicationAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'CLUBMEMBERSHIPALLOWANCE',
            xlsxDisplayColumnName: 'Club Membership Allowance',
            displayColumnName: 'Club Membership Allowance',
            dataAttributeName: 'clubMembershipAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'ALLOTHERALLOWANCE',
            xlsxDisplayColumnName: 'All Other Allowance',
            displayColumnName: 'All Other Allowance',
            dataAttributeName: 'allOtherAllowance',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'NAMEOFCCNLAPPLIEDITALYONLY',
            xlsxDisplayColumnName: 'Name of CCNL Applied(Italy Only)',
            displayColumnName: 'Name of CCNL Applied (Italy Only)',
            dataAttributeName: 'nameOfCCNLAppliedItalyOnly',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'EMPLOYEECATEGORYITALYONLY',
            xlsxDisplayColumnName: 'Employee Category (Italy Only)',
            displayColumnName: 'Employee Category (Italy Only)',
            dataAttributeName: 'employeeCategoryItalyOnly',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'RETIREMENTPREMIUM',
            xlsxDisplayColumnName: 'Retirement Premium',
            displayColumnName: 'Retirement Premium',
            dataAttributeName: 'retirementPremium',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'MANDATORYPROFITSHARINGPARTICIPATIONELIGIBILITYFRANCEPERUONLYYN',
            xlsxDisplayColumnName: 'Mandatory Profit-Sharing Participation Eligibility (France / Peru Only)(Y/N)',
            displayColumnName: 'Mandatory Profit-Sharing Participation Eligibility (France / Peru Only)',
            dataAttributeName: 'mandatoryProfitSharingParticipationEligibilityFrancePeruOnly',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'MANDATORYPROFITSHARINGPARTICIPATIONPAYMENTFRANCEPERUONLY',
            xlsxDisplayColumnName: 'Mandatory Profit-Sharing Participation Payment(France / Peru Only)',
            displayColumnName: 'Mandatory Profit-Sharing Participation Payment(France / Peru Only)',
            dataAttributeName: 'mandatoryProfitSharingParticipationPaymentFrancePeruOnly',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'VOLUNTARYPROFITSHARINGINTRESSEMENTELIGIBILITYFRANCEONLYYN',
            xlsxDisplayColumnName: 'Voluntary Profit-Sharing Intéressement Eligibility (France Only)(Y/N)',
            displayColumnName: 'Voluntary Profit-Sharing Intéressement Eligibility (France Only)',
            dataAttributeName: 'voluntaryProfitSharingIntressementEligibilityFranceOnly',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'VOLUNTARYPROFITSHARINGINTRESSEMENTPAYMENTFRANCEONLY',
            xlsxDisplayColumnName: 'Voluntary Profit-Sharing Intéressement Payment(France Only)',
            displayColumnName: 'Voluntary Profit-Sharing Intéressement Payment (France Only)',
            dataAttributeName: 'voluntaryProfitSharingIntressementPaymentFranceOnly',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'EXPATYN',
            xlsxDisplayColumnName: 'Expat (Y/N)',
            displayColumnName: 'Expat',
            dataAttributeName: 'expat',
            expectedType: 'number'
        },
        {
            normalizedColumnName: 'NATIONALITY',
            xlsxDisplayColumnName: 'Nationality',
            displayColumnName: 'Nationality',
            dataAttributeName: 'nationality',
            expectedType: 'number'
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

            var header = getHeader(worksheet, dataRange, headerRow);
//            console.log('header', header);

            var unidentifiedHeaderColumns = getUnidentifiedHeaderColumns(header);
            if (unidentifiedHeaderColumns.length > 0) {
                console.log('unidentifiedHeaderColumns', unidentifiedHeaderColumns);
                def.reject({ unidentifiedHeaderColumns: unidentifiedHeaderColumns });
                return false;
            }
//            var worksheetdata = getWorksheetData(worksheet, dataRange, headerRow);

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

    function getUnidentifiedHeaderColumns(header) {
        var retVal = [];
        header.forEach(function(column) {
            if (!column.isIdentified) {
                retVal.push(column);
            }
        });
        return retVal;
    }

    function getHeader(worksheet, range, headerRow) {
        var header = $.extend(true, [], DB_COLUMNS_METADATA);
        var colRangeBegin = range.s.c;
        var colRangeEnd   = range.e.c;

        for (var col = colRangeBegin; col <= colRangeEnd; col++) {
            var cell = getCellByRowCol(worksheet, headerRow, col);
            var value = cell ? getCellValue(cell): null;
            if (!value) {
                continue;
            }

            var column = getIdentifiedColumn(value, header);
            if (!column) {
                continue;
            }

            column.isIdentified = true;
            column.xlsxCell = cell;
        }

        return header;
    }

    function getWorksheetData(worksheet, dataRange, headerRow) {
        var rowRangeBegin = headerRow + 1;
        var rowRangeEnd   = range.e.r;
        var colRangeBegin = range.s.c;
        var colRangeEnd   = range.e.c;
        
        for (var row = rowRangeBegin; row <= rowRangeEnd; row++) {
            for (var col = colRangeBegin; col <= colRangeEnd; col++) {
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
                if (value && type && type === 's' && getIdentifiedColumn(value, DB_COLUMNS_METADATA)) {
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
        return numColsIdentified / DB_COLUMNS_METADATA.length * 100;
    }

    function getIdentifiedColumn(potentialColumnName, source) {
        var normalizedColumnName = getNormalizedColumnName(potentialColumnName);
//        console.log('normalizedColumnName', normalizedColumnName);

        return source.find(function(columnMetaData) {
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
