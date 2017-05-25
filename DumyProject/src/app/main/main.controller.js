(function() {
  'use strict';

  angular
    .module('dumyProject')
    .controller('MainController', MainController);

  /** @ngInject */
  function MainController($scope,$window,$timeout) {
    var vm = this;
      vm.dataList = [];
      vm.myInterval = 3000;
      vm.noWrapSlides = false;
      vm.active = 0;
      vm.slides = [
    {
      image: 'http://lorempixel.com/400/200/'
    },
    {
      image: 'http://lorempixel.com/400/200/food'
    },
    {
      image: 'http://lorempixel.com/400/200/sports'
    },
    {
      image: 'http://lorempixel.com/400/200/people'
    }
    ];
    console.log($scope.excel);
    /*vm.downloadFile = function(tableId,worksheetName){
         var uri='data:application/vnd.ms-excel;base64,',
            template='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
            base64=function(s){return $window.btoa(unescape(encodeURIComponent(s)));},
            format=function(s,c){return s.replace(/{(\w+)}/g,function(m,p){return c[p];})};
        return {
            tableToExcel:function(tableId,worksheetName){
                var table=$(tableId),
                     ctx={worksheet:worksheetName,table:table.html()},
                    href=uri+base64(format(template,ctx));
                return href;
            }
        };
    }

    $scope.exportToExcel=function(tableId){ // ex: '#my-table'
            var exportHref=Excel.tableToExcel(tableId,'WireWorkbenchDataExport');
            $timeout(function(){location.href=exportHref;},100); // trigger download
        }*/
    //console.log(excel);
    vm.download = function(tableId,worksheetName){
       var uri='data:application/vnd.ms-excel;base64,',
          template='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
            base64=function(s){return $window.btoa(unescape(encodeURIComponent(s)));},
            format=function(s,c){return s.replace(/{(\w+)}/g,function(m,p){return c[p];})};
                var table=$(tableId),
                    ctx={worksheet:worksheetName,table:table.html()},
                    href=uri+base64(format(template,ctx));
                return href; 
    } 
     vm.exportToExcel = function(tableId){ 
            var exportHref=vm.download(tableId,'WireWorkbenchDataExport');
            $timeout(function(){location.href=exportHref;},100); 
        }
    
    vm.setTable = function(data){
      vm.dataList = data;
      $scope.$apply();
    }
    $scope.read = function (workbook) {
        var sheetName = workbook.SheetNames[0];
        var sheet = workbook.Sheets[sheetName];
        console.log(sheet['!ref']);
        var s = XLSX.utils.sheet_to_json(sheet);
        vm.setTable(s);
       
      }

      $scope.error = function (e) {
        /* DO SOMETHING WHEN ERROR IS THROWN */
        console.log(e);
      }
       } 
})();
