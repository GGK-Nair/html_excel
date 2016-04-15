'use strict';

angular.module('myApp.view1', ['ngRoute'])

    .config(['$routeProvider', function ($routeProvider) {
        $routeProvider.when('/view1', {
            templateUrl: 'view1/view1.html',
            controller: 'View1Ctrl'
        });
    }])

    // .controller('View1Ctrl', [function() {
    //
    // }]);
    // .factory('Excel', function ($window) {
    //     var tablesToExcel = (function () {
    //         var uri = 'data:application/vnd.ms-excel;base64,'
    //             , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>'
    //             , templateend = '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head>'
    //             , body = '<body>'
    //             , tablevar = '<table>{table'
    //             , tablevarend = '}</table>'
    //             , bodyend = '</body></html>'
    //             , worksheet = '<x:ExcelWorksheet><x:Name>'
    //             , worksheetend = '</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>'
    //             , worksheetvar = '{worksheet'
    //             , worksheetvarend = '}'
    //             , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
    //             , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
    //             , wstemplate = ''
    //             , tabletemplate = '';
    //
    //         return function (table, name, filename) {
    //             var tables = table;
    //
    //             for (var i = 0; i < tables.length; ++i) {
    //                 wstemplate += worksheet + worksheetvar + i + worksheetvarend + worksheetend;
    //                 tabletemplate += tablevar + i + tablevarend;
    //             }
    //
    //             var allTemplate = template + wstemplate + templateend;
    //             var allWorksheet = body + tabletemplate + bodyend;
    //             var allOfIt = allTemplate + allWorksheet;
    //
    //             var ctx = {};
    //             for (var j = 0; j < tables.length; ++j) {
    //                 ctx['worksheet' + j] = name[j];
    //             }
    //
    //             for (var k = 0; k < tables.length; ++k) {
    //                 var exceltable;
    //                 if (!tables[k].nodeType) exceltable = document.getElementById(tables[k]);
    //                 ctx['table' + k] = exceltable.innerHTML;
    //             }
    //
    //             //document.getElementById("dlink").href = uri + base64(format(template, ctx));
    //             //document.getElementById("dlink").download = filename;
    //             //document.getElementById("dlink").click();
    //
    //             window.location.href = uri + base64(format(allOfIt, ctx));
    //
    //         }
    //     })();
    // })
    .controller('View1Ctrl', function ($timeout, $scope) {
        $scope.tablesToExcel_Dynamic = function () {
            var cellSpan = 0;
            var uri = 'data:application/vnd.ms-excel;base64,'
                , tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
                + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author></Author><Created>{created}</Created></DocumentProperties>'
                + '<Styles>'

                + '<Style ss:ID="center">'
                + '<Alignment ss:Horizontal="Center" ss:Vertical="Center"/>'
                + '</Style>'
                + '<Style ss:ID="right">'
                + '<Alignment ss:Horizontal="Right" ss:Vertical="Center"/>'
                + '</Style>'
                + '<Style ss:ID="left">'
                + '<Alignment ss:Horizontal="Left" ss:Vertical="Center"/>'
                + '</Style>'
                + '</Styles>'
                + '{worksheets}</Workbook>'
                , tmplWorksheetXML = '<Worksheet ss:Name="Report"><Table><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/><Column ss:AutoFitWidth="0" ss:Width="135"/>{rows}</Table></Worksheet>'
                , tmplCellXML = '<Cell ss:MergeAcross="{colSpan}" ss:MergeDown="{rowSpan}" ss:StyleID="{styleId}" {cellSsIndex}><Data ss:Type="{nameType}">{data}</Data></Cell>'
                , base64 = function (s) {
                return window.btoa(unescape(encodeURIComponent(s)))
            }
                , format = function (s, c) {
                return s.replace(/{(\w+)}/g, function (m, p) {
                    return c[p];
                })
            }
            var ctx = "";
            var workbookXML = "";
            var worksheetsXML = "";
            var rowsXML = "";
            var colSpan = 0;
            var rowSpan = 0;

            var tables = document.getElementsByTagName('table');//get all the tables from DOM

            for (var i = 0; i < tables.length; i++) {
                for (var j = 0; j < tables[i].rows.length; j++) {
                    if (i > 0 && j == 0) {
                        var k = tables[i - 1].rows.length + 3;
                        rowsXML += '<Row ss:Index="' + k + '">';
                    } else {
                        rowsXML += '<Row>'
                    }

                    for (var k = 0; k < tables[i].rows[j].cells.length; k++) {
                        // var dataType = tables[i].rows[j].cells[k].getAttribute("data-type");
                        var dataValue = tables[i].rows[j].cells[k].innerHTML;
                        colSpan = tables[i].rows[j].cells[k].colSpan - 1;
                        var align = tables[i].rows[j].cells[k].align.toLowerCase() || 'left';
                        rowSpan = tables[i].rows[j].cells[k].rowSpan - 1;

                        var cellSsIndex = '';

                        if (j > 0 && angular.isDefined(tables[i].rows[j - 1].cells[k])) {
                            if (tables[i].rows[(j - 1)].cells[k].rowSpan > 1 && k>cellSpan) {

                                cellSpan = k;
                                do {

                                    cellSpan++;
                                } while (angular.isDefined(tables[i].rows[(j - 1)].cells[cellSpan]) && tables[i].rows[(j - 1)].cells[cellSpan].rowSpan > 1);

                                cellSsIndex = 'ss:Index="' + (cellSpan + 1) + '"';
                                //k=cellSpan;
                            }

                        }

                        var dataType = typeof(dataValue);
                        ctx = {
                            nameType: 'String',
                            data: dataValue,
                            colSpan: colSpan,
                            rowSpan: rowSpan,
                            styleId: align,
                            cellSsIndex: cellSsIndex
                        };

                        rowsXML += format(tmplCellXML, ctx);//replacing nameType, Data, mergeAcross in tmplCellXMl
                    }
                    rowsXML += '</Row>'
                }
            }
            ctx = {rows: rowsXML};
            worksheetsXML += format(tmplWorksheetXML, ctx);//replacing rows in tmplWorksheetXML
            rowsXML = "";
            ctx = {created: (new Date()).getTime(), worksheets: worksheetsXML};
            workbookXML = format(tmplWorkbookXML, ctx);//replacing created, worksheets in tmplWorkbookXML
            var link = document.createElement("A");
            link.href = uri + base64(workbookXML);
            link.download = 'Workbook' + "_" + (new Date()).toString().replace(/\s/g, '_') + ".xls";//Name of file

            link.target = '_blank';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
        ////////////////////////////////////////////////////////////////


        $scope.tablesToExcel = function (tables, wsnames, wbname, appname) { // ex: '#my-table'

            var uri = 'data:application/vnd.ms-excel;base64,'
                , tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
                + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Gopi</Author><Created>{created}</Created></DocumentProperties>'
                + '{worksheets}</Workbook>'
                , tmplWorksheetXML = '<Worksheet ss:Name="Report"><Table>{rows}</Table></Worksheet>'
                , tmplCellXML = '<Cell><Data ss:Type="{nameType}">{data}</Data></Cell>'
                , base64 = function (s) {
                return window.btoa(unescape(encodeURIComponent(s)))
            }
                , format = function (s, c) {
                return s.replace(/{(\w+)}/g, function (m, p) {
                    return c[p];
                })
            }
            var ctx = "";
            var workbookXML = "";
            var worksheetsXML = "";
            var rowsXML = "";

            for (var i = 0; i < tables.length; i++) {
                if (!tables[i].nodeType) tables[i] = document.getElementById(tables[i]);
                for (var j = 0; j < tables[i].rows.length; j++) {
                    rowsXML += '<Row>'
                    for (var k = 0; k < tables[i].rows[j].cells.length; k++) {
                        // var dataType = tables[i].rows[j].cells[k].getAttribute("data-type");
                        var dataValue = tables[i].rows[j].cells[k].innerHTML;
                        var dataType = typeof(dataValue);

                        ctx = {
                            nameType: 'String',
                            data: dataValue

                        };
                        rowsXML += format(tmplCellXML, ctx);
                    }
                    rowsXML += '</Row>'
                }

            }
            ctx = {rows: rowsXML, nameWS: wsnames[i] || 'Sheet' + i};
            worksheetsXML += format(tmplWorksheetXML, ctx);
            rowsXML = "";
            ctx = {created: (new Date()).getTime(), worksheets: worksheetsXML};
            workbookXML = format(tmplWorkbookXML, ctx);

            console.log(workbookXML);

            var link = document.createElement("A");
            link.href = uri + base64(workbookXML);
            link.download = wbname || 'Workbook.xls';
            link.target = '_blank';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    });