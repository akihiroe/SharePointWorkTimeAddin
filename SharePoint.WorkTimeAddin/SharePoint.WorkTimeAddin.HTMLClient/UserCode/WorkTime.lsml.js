/// <reference path="~/GeneratedArtifacts/viewModel.js" />

myapp.AddEditCurrentUserWorkTimeItem.created = function (entity) {
    // Write code here.
    entity.WorkDate = new Date();
    entity.StartTime = "09:00";
    entity.EndTime = "17:00";
    // Using a Promise object we can call the CallGetUserName function
    msls.promiseOperation(CallGetUserName).then(function PromiseSuccess(PromiseResult) {
        // Set the result of the CallGetUserName function to the 
        // UserName of the entity
        entity.UserId = PromiseResult;
    });
};


// This function will be wrapped in a Promise object
function CallGetUserName(operation) {
    $.ajax({
        type: 'post',
        data: {},
        url: '../GetUserName.ashx',
        success: operation.code(function AjaxSuccess(AjaxResult) {
            operation.complete(AjaxResult);
        })
    });
}

$(function () {
    var promiseop;
    $('#upload').fileupload({
        singleFileUploads: true,
        add: function (e, data) {
            var promise = msls.promiseOperation(function (op) {
                data.submit()
                promiseop = op;
            });
            msls.showProgress(promise);
        },
        done: function (e, data) {
            promiseop.complete();
            if (data || data.result) {
                msls.showMessageBox(data.result).then(function () {
                    CurrentUserWorkTime.refresh();
                });
            } else {
                CurrentUserWorkTime.refresh();
            }
        }
    });
});