(function() {
    'use strict';

    var directiveId = 'importSheetJs';

    angular
        .module('app')
        .directive(directiveId, importSheetJs);

    importSheetJs.$inject = ['$log'];

    function importSheetJs($log) {
        var loggerSource = '[' + directiveId + ']';
        var directive = {
            link: link,
            restrict: 'A',
            scope: {
                onLoadExcel: '&'
            }
        };
        return directive;
        
        function link(scope, element, attrs) {
            $log.debug(loggerSource, 'Директива загружена');
            
            element.on('change', function (changeEvent) {
                var reader = new FileReader();
        
                reader.onload = function (e) {
                    var bstr = e.target.result;
                    scope.onLoadExcel({workbook: bstr});
                };
        
                if (changeEvent.target.files[0]) {
                    reader.readAsArrayBuffer(changeEvent.target.files[0]);
                }
            });
        }
    }
})();