(function () {
    'use strict';

    var factoryId = 'excelFactory';

    angular
        .module('app')
        .factory(factoryId, factory);

    factory.$inject = [
        '$q', '$log', '$timeout'
    ];

    function factory($q, $log, $timeout) {
        var loggerSource = '[' + factoryId + ']';

        var fc = {
            parseExcel: parseExcel,
            recalcTable: recalcTable
        };

        init();

        return fc;

        ////////////

        function init() {
            $log.debug(loggerSource, 'Фабрика загружена');
        }

        function parseExcel(workbook, assignmentKidConst, executionDocPercentConst, executionDocKidConst) {
            var deferred = $q.defer();
            
            var startTable = 0;
            var isCorrectFile = false;
            
            for (var i = 1; i < 50; i++) {
                if (workbook.Sheets.Sheet1['A' + i] !== undefined && workbook.Sheets.Sheet1['A' + i].v.trim() == '№') {
                    startTable = i + 1;
                    isCorrectFile = true;
                    break;
                }
            }

            if (!isCorrectFile) {
                deferred.reject();
                return deferred.promise;
            }

            try {
                var range = XLSX.utils.decode_range(workbook.Sheets.Sheet1['!ref']);
                var users = [];

                var doneCount = 0;
                var overdueCount = 0;
                var docsCount = 0;
                var assignmentDoneCount = 0;
                var assignmentOverdueCount = 0;
                var assignmentDocsCount = 0;
                var userNumber = startTable;
                var userName = null;

                for (var i = startTable; i <= range.e.r; i++) {
                    if (i == startTable) {
                        userName = workbook.Sheets.Sheet1['B' + i].v;
                    }

                    if (workbook.Sheets.Sheet1['A' + (i + 1)] == undefined && i < range.e.r) {
                        if (workbook.Sheets.Sheet1['C' + i].v.trim() == 'Поручение') {
                            assignmentDoneCount = parseInt(workbook.Sheets.Sheet1['D' + i].v);
                            assignmentOverdueCount = parseInt(workbook.Sheets.Sheet1['E' + i].v);
                        } else {
                            doneCount += parseInt(workbook.Sheets.Sheet1['D' + i].v);
                            overdueCount += parseInt(workbook.Sheets.Sheet1['E' + i].v);
                        }
                    } else {
                        doneCount += parseInt(workbook.Sheets.Sheet1['D' + i].v);
                        overdueCount += parseInt(workbook.Sheets.Sheet1['E' + i].v);

                        docsCount = overdueCount + doneCount;
                        var kid = getKid(assignmentDoneCount, assignmentOverdueCount, docsCount, doneCount, assignmentKidConst, executionDocPercentConst, executionDocKidConst);

                        users.push({
                            Name: userName,
                            DoneCount: doneCount,
                            OverdueCount: overdueCount,
                            DocsCount: docsCount,
                            AssignmentDoneCount: assignmentDoneCount,
                            AssignmentOverdueCount: assignmentOverdueCount,
                            DoneCountPercent: (kid.DoneCountPercent * 100).toFixed(2) + '%',
                            AssignmentKid: kid.AssignmentKid.toFixed(2) + '%',
                            ExecutionDocPercent: (kid.ExecutionDocPercent * 100).toFixed(2) + '%',
                            ExecutionDocKid: kid.ExecutionDocKid.toFixed(2) + '%',
                            TotalKid: kid.TotalKid.toFixed(2) + '%'
                        });

                        if (workbook.Sheets.Sheet1['B' + (i + 1)]) {
                            userName = workbook.Sheets.Sheet1['B' + (i + 1)].v;
                        }

                        doneCount = 0;
                        overdueCount = 0;
                    }
                }

                deferred.resolve(users);
            } catch (error) {
                deferred.reject();
            } finally {
                return deferred.promise;
            }
        }

        function getKid(assignmentDoneCount, assignmentOverdueCount, docsCount, doneCount, assignmentKidConst, executionDocPercentConst, executionDocKidConst) {
            executionDocPercentConst = executionDocPercentConst * 0.01;

            var doneCountPercent = ((assignmentDoneCount + assignmentOverdueCount) == 0) ? 1 : (assignmentDoneCount / (assignmentDoneCount + assignmentOverdueCount));
            var executionDocPercent = (docsCount == 0 ? 1 : (doneCount / docsCount));
            var assignmentKid = doneCountPercent * assignmentKidConst;
            var executionDocKid = ((executionDocPercentConst + executionDocPercent) < 1 ? (executionDocPercentConst + executionDocPercent) * executionDocKidConst : executionDocKidConst);
            var totalKid = assignmentKid + executionDocKid;

            return {
                DoneCountPercent: doneCountPercent,
                ExecutionDocPercent: executionDocPercent,
                AssignmentKid: assignmentKid,
                ExecutionDocKid: executionDocKid,
                TotalKid: totalKid
            };
        }

        function recalcTable(users, assignmentKid, executionDocPercent, executionDocKid) {
            return users.map(function (item) {
                var kid = getKid(
                    item.AssignmentDoneCount,
                    item.AssignmentOverdueCount,
                    item.DocsCount,
                    item.DoneCount,
                    assignmentKid,
                    executionDocPercent,
                    executionDocKid
                );

                item.DoneCountPercent = (kid.DoneCountPercent * 100).toFixed(2) + '%',
                item.AssignmentKid = kid.AssignmentKid.toFixed(2) + '%',
                item.ExecutionDocPercent = (kid.ExecutionDocPercent * 100).toFixed(2) + '%',
                item.ExecutionDocKid = kid.ExecutionDocKid.toFixed(2) + '%',
                item.TotalKid = kid.TotalKid.toFixed(2) + '%'

                return item;
            });
        }
    }
})();