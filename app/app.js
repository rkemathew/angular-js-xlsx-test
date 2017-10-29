var app = angular.module('app', [ 'angularFileUpload' ]);

app.directive('importSheetJs', function() {
    return {
        scope: { },
        link: function ($scope, $elm, $attrs) {
            $elm.on('change', function (changeEvent) {
                $scope.$emit('file:uploaded', changeEvent.target.files[0])
            });
        }
    };
});

app.controller('MainCtrl', ['$scope', '$q', 'FileUploader', 'XlsxProcService', function ($scope, $q, FileUploader, XlsxProcService) {
    var uploader = $scope.uploader = new FileUploader({});
    uploader.onAfterAddingFile = function(fileItem) {
        XlsxProcService.processXlsx(fileItem._file).then(function(data) {
            $scope.data = data;
        }, function(error) {
            $scope.error = error;
        });
    };

    $scope.$on('file:uploaded', function(event, file) {
        XlsxProcService.processXlsx(file).then(function(data) {
            $scope.data = data;
        }, function(error) {
            $scope.error = error;
        });
    });
}]);
