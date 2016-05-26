/// <reference path="~/GeneratedArtifacts/viewModel.js" />

myapp.BrowseCurrentUserWorkTime.created = function (screen) {
    var thisMonth = new Date();
    thisMonth.setDate(1);
    thisMonth.setHours(0, 0, 0, 0);
    screen.year = thisMonth.getFullYear();
    screen.month = thisMonth.getMonth() + 1;
    msls.promiseOperation(CallGetUserName).then(function PromiseSuccess(PromiseResult) {
        // Set the result of the CallGetUserName function to the 
        // UserName of the entity
        screen.CurrentUser = PromiseResult;
    });
};

myapp.BrowseCurrentUserWorkTime.AddCurrentUserWorkTimeItem_Tap_execute = function (screen) {
    // Write code here.
    myapp.showAddEditCurrentUserWorkTimeItem(null, {
        beforeShown: function (addEditWorkTimeScreen) {
            var entity = screen.CurrentUserWorkTime.addNew();
            var today = new Date();
            today.setHours(0, 0, 0, 0);
            entity.WorkDate = today;
            entity.StartTime = "09:00";
            entity.EndTime = "17:00";
            // Using a Promise object we can call the CallGetUserName function
            msls.promiseOperation(CallGetUserName).then(function PromiseSuccess(PromiseResult) {
                // Set the result of the CallGetUserName function to the 
                // UserName of the entity
                entity.UserId = PromiseResult;
            });
            addEditWorkTimeScreen.WorkTime = entity;
        },
        afterClosed: function (addEditScreen, navigationAction) {
            // If the user commits the change,
            // update the selected order on the Main screen
            if (navigationAction === msls.NavigateBackAction.commit) {
                screen.CurrentUserWorkTime.refresh();
            }
        }
    });

};

myapp.BrowseCurrentUserWorkTime.Delete_execute = function (screen) {

    var resp = msls.showMessageBox("データを削除しますか？", { title: 'Wanna delete?!', buttons: msls.MessageBoxButtons.yesNo });
    resp.then(function (val) {
        if (val == msls.MessageBoxResult.yes) {
            screen.CurrentUserWorkTime.deleteSelected();
            myapp.commitChanges().then(function success() {
            }, function fail(e) {
                // If error occurs,
                msls.showMessageBox(e.message, { title: e.title }).then(function () {
                    // Cancel Changes
                    myapp.cancelChanges();
                });
            });
        }
    });
};
myapp.BrowseCurrentUserWorkTime.WorkDate_postRender = function (element, contentItem) {
    // Write code here.
    contentItem.dataBind("value", function (value) {
        if (value) {
            $(element).text(moment(value).format("MM/DD(ddd)"));
        }
    });
};
myapp.BrowseCurrentUserWorkTime.Upload_execute = function (screen) {
    CurrentUserWorkTime = screen.CurrentUserWorkTime;
    $('#upl').click();
};

myapp.BrowseCurrentUserWorkTime.Submit_execute = function (screen) {
    var promise = msls.promiseOperation(function (op) {
        $.ajax({
            type: "POST",
            url: "../SubmitWorkTime.ashx",
            data: "email=" + screen.CurrentUser + "&year=" + screen.year + "&month=" + screen.month,
            success: function (message) {
                if (message) {
                    msls.showMessageBox(message, { title: "通知" });
                }
                op.complete();
            },
            fail: function () {
                msls.showMessageBox("作業時間の送信に失敗しました", { title: "失敗" });
                op.complete();
            }
        });
    });
    msls.showProgress(promise);
};
myapp.BrowseCurrentUserWorkTime.Download_execute = function (screen) {
    var url = "../SheetDownload.ashx?email=" + screen.CurrentUser + "&year=" + screen.year + "&month=" + screen.month;
    window.open(url);
};