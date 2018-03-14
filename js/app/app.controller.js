(function() {
    'use strict';

    var controllerId = 'excelController';

    angular
        .module('app')
        .controller(controllerId, excelController);

    excelController.$inject = ['$log', 'excelFactory', '$window', '$scope', '$cookies'];
    function excelController($log, excelFactory, $window, $scope, $cookies) {
        var vm = this;
        var loggerSource = '[' + controllerId + ']';
        var exportFileName = 'export.xlsx';
        var exportSheetName = 'Sheet1';
        var cookieName = 'KID_NUMBERS';

        vm.loadExcel = loadExcel;
        vm.loading = false;
        vm.isError = false;
        vm.excelTable = [];
        vm.exportExcel = exportExcel;
        vm.reloadTable = reloadTable;

        vm.assignmentKid = 45;
        vm.executionDocPercent = 10;
        vm.executionDocKid = 35;

        activate();

        ////////////////

        function activate() {
            $log.debug(loggerSource, 'Контроллер загружен');

            var cookie = $cookies.getObject(cookieName);
            if (cookie != undefined && cookie != null) {
                vm.assignmentKid = cookie.AssignmentKid;
                vm.executionDocPercent = cookie.ExecutionDocPercent;
                vm.executionDocKid = cookie.ExecutionDocKid;
            }
        }

        function loadExcel(workbook) {
            vm.excelTable = [];
            vm.loading = true;
            vm.isError = false;
            $scope.$apply();
            
            if ($window.Worker) {
                var worker = new Worker('/_layouts/15/xlsx-kid/js/app/parse-excel.worker.js');

                worker.onmessage = function(e) {
                    if (e.data == null) {
                        vm.isError = true;
                        vm.loading = false;
                        $scope.$apply();
                    } else {
                        excelFactory.parseExcel(e.data, vm.assignmentKid, vm.executionDocPercent, vm.executionDocKid)
                            .then(function (data) {
                                vm.excelTable = data;
                            })
                            .catch(function () {
                                vm.isError = true;
                            })
                            .finally(function () {
                                vm.loading = false;
                            });
                    }
                }

                worker.postMessage([workbook]);
            }
        }

        function exportExcel() {
            var ws = XLSX.utils.json_to_sheet(vm.excelTable);
            
            for (var prop in ws) {
                if (ws.hasOwnProperty(prop)) {
                    if (ws[prop].v && angular.isString(ws[prop].v) && ws[prop].v.indexOf('%') !== -1) {
                        ws[prop].v = ws[prop].v.replace(/\%/g, '');
                        ws[prop].t = 'n';
                        angular.extend(ws[prop], { z: '0.00\\%' });
                    }
                } 
            }

            var headers = [
                'ФИО',
                'Выполнено в срок (кол-во)',
                'Просрочено (кол-во)',
                'Всего',
                'Поручений выполнено в срок (кол-во)',
                'Поручений просрочено (кол-во)',
                'Процент выполнения поручений',
                'КИД по поручениям',
                'Процент исполнения по процедурам документооборота',
                'КИД по процедурам документооборота',
                'Итого КИД'
            ];

            var wscols = [
                {wpx: 250},
                {hidden: true},
                {hidden: true},
                {hidden: true},
                {hidden: true},
                {hidden: true},
                {wpx: 100},
                {wpx: 100},
                {wpx: 100},
                {wpx: 100},
                {wpx: 100}
            ];

            ws['!cols'] = wscols;
            ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'].forEach(function (item, index) {
                ws[item + '1'].v = headers[index];
            });

            var wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, exportSheetName);

            XLSX.writeFile(wb, exportFileName);
        }

        function reloadTable() {
            var expireDate = new Date();
            expireDate.setDate(expireDate.getDate() + 30);

            $cookies.putObject(cookieName, {
                AssignmentKid: vm.assignmentKid,
                ExecutionDocPercent: vm.executionDocPercent,
                ExecutionDocKid: vm.executionDocKid
            }, {'expires': expireDate});

            vm.excelTable = excelFactory.recalcTable(vm.excelTable, vm.assignmentKid, vm.executionDocPercent, vm.executionDocKid);
        }
    }
})();