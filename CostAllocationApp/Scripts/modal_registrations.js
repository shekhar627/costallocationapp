//sections insert
function InsertSection() {
    var apiurl = "/api/sections/";
    let sectionName = $("#section-name").val().trim();
    if (sectionName == "") {
        $(".section_name_err").show();
        return false;
    } else {
        var data = {
            SectionName: sectionName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function(data) {
                $("#page_load_after_modal_close").val("yes");
                ToastMessageSuccess(data);

                $('#section-name').val('');
                GetSectionList();
            },
            error: function(data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}
//Get section list
function GetSectionList() {
    $.getJSON('/api/sections/')
        .done(function(data) {
            $('#section_list_tbody').empty();
            $.each(data, function(key, item) {
                $('#section_list_tbody').append(`<tr><td><input type="checkbox" class="section_list_chk" data-id='${item.Id}' /></td><td>${item.SectionName}</td></tr>`);
            });
        });
}

//Load Section 
$('#department_modal_href').click(function () {
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_list').empty();
            $('#section_list').append(`<option value=''>Select Section</option>`)
            $.each(data, function (key, item) {
                $('#section_list').append(`<option value='${item.Id}'>&nbsp;&nbsp; ${item.SectionName}</option>`)
            });
        });
});

//insert department
function InsertDepartment() {
    var apiurl = "/api/Departments/";
    let departmentName = $("#department_name").val().trim();
    let sectionId = $("#section_list").val().trim();

    let isValidRequest = true;

    if (departmentName == "") {
        $(".dept_name_err").show();
        isValidRequest = false;
    } else {
        $(".dept_name_err").hide();
    }
    if (sectionId == "" || sectionId < 0) {
        $("#section_list_error").show();
        isValidRequest = false;
    } else {
        $("#section_list_error").hide();
    }

    if (isValidRequest) {
        var data = {
            DepartmentName: departmentName,
            SectionId: sectionId
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function(data) {
                $("#page_load_after_modal_close").val("yes");
                $("#department_name").val('');
                $("#section_list").val('');

                ToastMessageSuccess(data);
                $('#section-name').val('');
                GetDepartments();
            },
            error: function(data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//get department list
function GetDepartments() {
    $.getJSON('/api/departments/')
        .done(function(data) {
            $('#department_list_tbody').empty();
            $.each(data, function(key, item) {
                $('#department_list_tbody').append(`<tr><td><input type="checkbox" class="department_list_chk" data-id='${item.Id}' /></td><td>${item.DepartmentName}</td><td>${item.SectionName}</td></tr>`);
            });
        });
}

//insert department
function InsertInCharge() {
    var apiurl = "/api/incharges/";
    let in_charge_name = $("#in_charge_name").val().trim();
    if (in_charge_name == "") {
        $(".incharge_name_err").show();
        return false;
    } else {
        $(".incharge_name_err").hide();
        var data = {
            InChargeName: in_charge_name
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function(data) {
                $("#page_load_after_modal_close").val("yes");
                $("#in_charge_name").val('');
                ToastMessageSuccess(data)
                $('#section-name').val('');

                GetInchargeList();
            },
            error: function(data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//get department list
function GetInchargeList() {
    $.getJSON('/api/InCharges/')
        .done(function(data) {
            $('#incharge_list_tbody').empty();
            $.each(data, function(key, item) {
                $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.InChargeName}</td></tr>`);
            });
        });
}

//insert role
function InsertRoles() {
    var apiurl = "/api/Roles/";
    let roleName = $("#role_name").val().trim();
    if (roleName == "") {
        $(".role_name_err").show();
        return false;
    } else {
        $(".role_name_err").hide();
        var data = {
            RoleName: roleName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function(data) {
                $("#page_load_after_modal_close").val("yes");
                $("#role_name").val('');
                ToastMessageSuccess(data);
                $('#section-name').val('');
                GetRoleList();
            },
            error: function(data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//Get role list
function GetRoleList() {
    $.getJSON('/api/Roles/')
        .done(function(data) {
            $('#role_list_tbody').empty();
            $.each(data, function(key, item) {
                // Add a list item for the product.
                $('#role_list_tbody').append(`<tr><td><input type="checkbox" class="role_list_chk" data-id='${item.Id}' /></td><td>${item.RoleName}</td></tr>`);
            });
        });
}

function GetExplanationList() {
    $.getJSON('/api/Explanations/')
        .done(function(data) {
            $('#explanations_list_tbody').empty();
            $.each(data, function(key, item) {
                $('#explanations_list_tbody').append(`<tr><td><input type="checkbox" class="explanation_list_chk" data-id='${item.Id}' /></td><td>${item.ExplanationName}</td></tr>`);
            });
        });
}

function InsertExplanations() {
    var apiurl = "/api/Explanations/";
    let explanationName = $("#explanation_name").val().trim();
    if (explanationName == "") {
        $(".explanations_name_err").show();
        return false;
    } else {
        $(".explanations_name_err").hide();
        var data = {
            ExplanationName: explanationName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function(data) {
                $("#page_load_after_modal_close").val("yes");
                $("#explanation_name").val('');
                ToastMessageSuccess(data);
                GetExplanationList();
            },
            error: function(data) {
                ToastMessageFailed(data);
            }
        });
    }
}

function GetCompanyList() {
    $.getJSON('/api/Companies/')
        .done(function(data) {
            $('#company_list_tbody').empty();
            $.each(data, function(key, item) {
                $('#company_list_tbody').append(`<tr><td><input type="checkbox" class="company_list_chk" data-id='${item.Id}' /></td><td>${item.CompanyName}</td></tr>`);
            });
        });
}

function InsertCompanies() {
    var apiurl = "/api/Companies/";
    let companyName = $("#companyName").val().trim();
    if (companyName == "") {
        $(".company_name_err").show();
        return false;
    } else {
        $(".company_name_err").hide();
        var data = {
            CompanyName: companyName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function(data) {
                $("#page_load_after_modal_close").val("yes");
                $("#companyName").val('');
                ToastMessageSuccess(data);
                GetCompanyList();
            },
            error: function(data) {
                ToastMessageFailed(data);
            }
        });
    }
}

function GetSalaries() {
    $.getJSON('/api/Salaries/')
        .done(function(data) {
            $('#salary_list_tbody').empty();
            $.each(data, function(key, item) {
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
    if (lowUnitPrice == "") {
        $("#lowPrice").show();
        isValidRequest = false;
    } else {
        $("#lowPrice").hide();
    }
    if (highUnitPrice == "") {
        $("#highPrice").show();
        isValidRequest = false;
    } else {
        $("#highPrice").hide();
    }
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
            success: function(data) {

                ToastMessageSuccess(data);
                GetSalaries();
            },
            error: function(data) {
                ToastMessageFailed(data);
            }
        });
    }
}