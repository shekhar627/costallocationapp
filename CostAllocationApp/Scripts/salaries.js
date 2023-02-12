//function onSalaryInactiveClick() {
//    let salaryIds = GetCheckedIds("salary_list_tbody");
//    var apiurl = '/api/utilities/SalaryCount?salaryIds=' + salaryIds;
//    $.ajax({
//        url: apiurl,
//        type: 'Get',
//        dataType: 'json',
//        success: function (data) {
//            $('.salary_count').empty();
//            $.each(data, function (key, item) {
//                $('.salary_count').append(`<li class='text-info'>${item}</li>`);
//            });
//        },
//        error: function (data) {
//        }
//    });
//}

$(document).ready(function () {
    //GetSalaries();

    //$('#salary_inactive_confirm_btn').on('click', function (event) {
    //    event.preventDefault();
    //    let id = GetCheckedIds("salary_list_tbody");
    //    id = id.slice(0, -1);        
    //    $.ajax({
    //        url: '/api/Salaries?salaryIds=' + id,
    //        type: 'DELETE',
    //        success: function (data) {
    //            ToastMessageSuccess(data);
    //            GetSalaries();                
    //        },
    //        error: function (data) {
    //            ToastMessageFailed(data);
    //        }
    //    });

    //    $('#inactive_salary_modal').modal('toggle');
    //});

    //$('#salary_inactive_btn').on('click', function (event) {

    //    let id = GetCheckedIds("salary_list_tbody");
    //    if (id == "") {
    //        alert("Please check first to delete.");
    //        return false;
    //    }
    //});


    $.getJSON('/api/departments/')
        .done(function (data) {
            $('#salary_department').empty();
            $('#salary_department').append(`<option value=''>Select Department</option>`);
            $.each(data, function (key, item) {
                $('#salary_department').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
            });
        });
    $.getJSON('/api/UnitPriceTypes/')
        .done(function (data) {
            $('#salary_salaryType').empty();
            $('#salary_salaryType').append(`<option value=''>Select Salary Type</option>`);
            $.each(data, function (key, item) {
                $('#salary_salaryType').append(`<option value='${item.Id}'>${item.SalaryTypeName}</option>`);
            });
        });
    $.getJSON('/api/Grades/')
        .done(function (data) {
            $('#salary_gradeId').empty();
            $('#salary_gradeId').append(`<option value=''>Select Grade</option>`);
            $.each(data, function (key, item) {
                $('#salary_gradeId').append(`<option value='${item.Id}'>${item.GradeName}</option>`);
            });
        });

});
//function GetSalaries(){
//    $.getJSON('/api/Salaries/')
//    .done(function (data) {
//        $('#salary_list_tbody').empty();
//        $.each(data, function (key, item) {
//            let tempHighVal = "";
//            let tempLowVal = "";
//            if (item.SalaryLowPointWithComma == '' || item.SalaryLowPointWithComma == null) {
//                tempLowVal = "-"
//            } else {
//                tempLowVal = item.SalaryLowPointWithComma
//            }
//            if (item.SalaryHighPointWithComma == '' || item.SalaryHighPointWithComma == null) {
//                tempHighVal = "-"
//            } else {
//                tempHighVal = item.SalaryHighPointWithComma
//            }
//            $('#salary_list_tbody').append(`<tr><td><input type="checkbox" class="salary_list_chk" data-id='${item.Id}' /></td><td>${tempLowVal} ～ ${tempHighVal}</td><td>${item.SalaryGrade}</td></tr>`);
//        });
//    });
//}    
function CreateGradeSalaryType() {
    var apiurl = "/api/Salaries/";
    let year = $('#salary_year').val();
    let departmentId = $('#salary_department').val();
    let salaryTypeId = $('#salary_salaryType').val();
    let gradeId = $('#salary_gradeId').val();
    let beginning = $('#salary_beginning').val();
    let revision = $('#salary_revision').val();

    if (year == undefined || year == null || year == "") {
        $(".salary_year_err").show();
        return false;
    }
    else {
        $(".salary_year_err").hide();
    }

    if (departmentId == undefined || departmentId == null || departmentId == "") {
        $(".salary_department_name_err").show();
        return false;
    }
    else {
        $(".salary_department_name_err").hide();
    }

    if (salaryTypeId == undefined || salaryTypeId == null || salaryTypeId == "") {
        $(".salary_salaryType_err").show();
        return false;
    }
    else {
        $(".salary_salaryType").hide();
    }

    if (gradeId == undefined || gradeId == null || gradeId == "") {
        $(".salary_gradeId_name_err").show();
        return false;
    }
    else {
        $(".salary_gradeId_name_err").hide();
    }

    if (beginning<0) {
        $(".salary_beginning_name_err").show();
        return false;
    }
    else {
        $(".salary_beginning_name_err").hide();
    }
    if (revision < 0) {
        $(".salary_revision_name_err").show();
        return false;
    }
    else {
        $(".salary_revision_name_err").hide();
    }

    var data = {
        GradeId: gradeId,
        GradeLowPoints: beginning,
        GradeHighPoints: revision,
        DepartmentId: departmentId,
        Year: year,
        SalaryTypeId: salaryTypeId,
    };

    $.ajax({
        url: apiurl,
        type: 'POST',
        dataType: 'json',
        data: data,
        success: function (data) {

            ToastMessageSuccess(data);
            //GetSalaries();    
        },
        error: function (data) {
            //ToastMessageFailed(data);
        }
    });
    
}