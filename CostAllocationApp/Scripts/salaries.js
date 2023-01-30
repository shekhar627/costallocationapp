function onSalaryInactiveClick() {
    let salaryIds = GetCheckedIds("salary_list_tbody");
    var apiurl = '/api/utilities/SalaryCount?salaryIds=' + salaryIds;
    $.ajax({
        url: apiurl,
        type: 'Get',
        dataType: 'json',
        success: function (data) {
            $('.salary_count').empty();
            $.each(data, function (key, item) {
                $('.salary_count').append(`<li class='text-info'>${item}</li>`);
            });
        },
        error: function (data) {
        }
    });
}

$(document).ready(function () {
    GetSalaries();

    $('#salary_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("salary_list_tbody");
        id = id.slice(0, -1);        
        $.ajax({
            url: '/api/Salaries?salaryIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);
                GetSalaries();                
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });

        $('#inactive_salary_modal').modal('toggle');
    });

    $('#salary_inactive_btn').on('click', function (event) {

        let id = GetCheckedIds("salary_list_tbody");
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        }
    });

});
function GetSalaries(){
    $.getJSON('/api/Salaries/')
    .done(function (data) {
        $('#salary_list_tbody').empty();
        $.each(data, function (key, item) {
            $('#salary_list_tbody').append(`<tr><td><input type="checkbox" class="salary_list_chk" data-id='${item.Id}' /></td><td>${item.SalaryLowPointWithComma} ～ ${item.SalaryHighPointWithComma}</td><td>${item.SalaryGrade}</td></tr>`);
        });
    });
}    
function InsertSalaries() {
    var apiurl = "/api/Salaries/";
    let lowUnitPrice = $("#lowUnitPrice").val().trim();
    let highUnitPrice = $("#hightUnitPrice").val().trim();
    let gradePoints = $("#gradePoints").val().trim();

    let isValidRequest = true;
    //if (lowUnitPrice == "") {
    //    $("#lowPrice").show();
    //    isValidRequest = false;
    //}
    //else {
    //    $("#lowPrice").hide();
    //}
    //if (highUnitPrice == "") {
    //    $("#highPrice").show();
    //    isValidRequest = false;
    //} else {
    //    $("#highPrice").hide();
    //}
    if (gradePoints == "") {
        $("#salaryGradePoints").show();
        isValidRequest = false;
    } else {
        $("#salaryGradePoints").hide();
    }

    if (isValidRequest) {
        var data = {
            SalaryLowPoint: lowUnitPrice,
            SalaryHighPoint: highUnitPrice,
            SalaryGrade: gradePoints
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {

                ToastMessageSuccess(data);
                GetSalaries();    
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });
    }
}