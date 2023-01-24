//sections insert
function InsertSection() {
    var apiurl = "/api/sections/";
    let sectionName = $("#section-name").val().trim();
    if (sectionName == "") {
        $(".section_name_err").show();
        return false;
    } else {
        $(".section_name_err").hide();
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
$('#department_modal_href').click(function() {
    $.getJSON('/api/sections/')
        .done(function(data) {
            $('#section_list').empty();
            $('#section_list').append(`<option value=''>Select Section</option>`)
            $.each(data, function(key, item) {
                $('#section_list').append(`<option value='${item.Id}'>&nbsp;&nbsp; ${item.SectionName}</option>`)
            });
        });
});

//insert department
function InsertDepartment() {
    var apiurl = "/api/Departments/";
    let departmentName = $("#department_name").val().trim();
    //let sectionId = $("#section_list").val().trim();

    let isValidRequest = true;

    if (departmentName == "") {
        $(".dept_name_err").show();
        isValidRequest = false;
    } else {
        $(".dept_name_err").hide();
    }
    //if (sectionId == "" || sectionId < 0) {
    //    $("#section_list_error").show();
    //    isValidRequest = false;
    //} else {
    //    $("#section_list_error").hide();
    //}

    if (isValidRequest) {
        var data = {
            DepartmentName: departmentName
            //,
            //SectionId: sectionId
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function(data) {
                $("#page_load_after_modal_close").val("yes");
                $("#department_name").val('');
                //$("#section_list").val('');

                ToastMessageSuccess(data);
                //$('#section-name').val('');
                //GetDepartments();
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

$('#section_registration').on('hidden.bs.modal', function () {
    var isPageLoad = $("#page_load_after_modal_close").val();
    if (isPageLoad == "yes") {
        //SetRegistrationValueInHiddenField();
        //window.location.reload(true);
        //RetainRegistrationValue()
        FillDropdownOfNameRegistration();
    }
})
$('#department_registration').on('hidden.bs.modal', function () {
    var isPageLoad = $("#page_load_after_modal_close").val();
    if (isPageLoad == "yes") {
        //SetRegistrationValueInHiddenField();
        //window.location.reload(true);
        //RetainRegistrationValue()
        FillDropdownOfNameRegistration();
    }
})
$('#in-charge_registration').on('hidden.bs.modal', function () {
    var isPageLoad = $("#page_load_after_modal_close").val();
    if (isPageLoad == "yes") {
        //SetRegistrationValueInHiddenField();
        //window.location.reload(true);
        //RetainRegistrationValue();
        FillDropdownOfNameRegistration();
    }
})
$('#role_registration_modal').on('hidden.bs.modal', function () {
    var isPageLoad = $("#page_load_after_modal_close").val();
    if (isPageLoad == "yes") {
        //SetRegistrationValueInHiddenField();
        //window.location.reload(true);
        //RetainRegistrationValue();
        FillDropdownOfNameRegistration();
    }
})
$('#explanation_modal').on('hidden.bs.modal', function () {
    var isPageLoad = $("#page_load_after_modal_close").val();
    if (isPageLoad == "yes") {
        //SetRegistrationValueInHiddenField();
        //window.location.reload(true);
        //RetainRegistrationValue()
        FillDropdownOfNameRegistration();
    }
})
$('#company_reg_modal').on('hidden.bs.modal', function () {
    var isPageLoad = $("#page_load_after_modal_close").val();
    if (isPageLoad == "yes") {
        //window.location.reload(true);
        //RetainRegistrationValue()
        FillDropdownOfNameRegistration();
    }
})
function FillDropdownOfNameRegistration() {
    SetRegistrationValueInHiddenField();

    //SectionDropdown fill
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_multi_search').empty();
            var sectionId = $("#hid_sectionId").val();
            console.log(sectionId);
            $.each(data, function (key, item) {
                if (sectionId == item.Id) {
                    $('#section_multi_search').append(`<option selected class='section_checkbox' id="section_checkbox_${item.Id}" value='${item.Id}' >${item.SectionName}</option>`);
                } else {
                    $('#section_multi_search').append(`<option class='section_checkbox' id="section_checkbox_${item.Id}" value='${item.Id}' >${item.SectionName}</option>`)
                }
            });
            $('#section_multi_search').multiselect('rebuild');
        });

    $.getJSON(`/api/Departments/`)
        .done(function (data) {
            $('#dept_multi_search').empty();
            var departmentId = $("#hid_departmentId").val();
            $.each(data, function (key, item) {
                if (departmentId == item.Id) {
                    $('#dept_multi_search').append(`<option selected class='department_checkbox' id="department_checkbox_${item.Id}" value='${item.Id}'>${item.DepartmentName}</option>`)
                } else {
                    $('#dept_multi_search').append(`<option class='department_checkbox' id="department_checkbox_${item.Id}" value='${item.Id}'>${item.DepartmentName}</option>`)
                }
            });
            $('#dept_multi_search').multiselect('rebuild');
        });

    //InCharge dropdown fill
    $.getJSON('/api/incharges/')
        .done(function (data) {
            $('#incharge_multi_search').empty();
            var inchargeId = $("#hid_inchargeId").val();
            $.each(data, function (key, item) {
                if (inchargeId == item.Id) {
                    $('#incharge_multi_search').append(`<option selected class='incharge_checkbox' id="incharge_checkbox_${item.Id}" value='${item.Id}'>${item.InChargeName}</option>`)
                } else {
                    $('#incharge_multi_search').append(`<option class='incharge_checkbox' id="incharge_checkbox_${item.Id}" value='${item.Id}'>${item.InChargeName}</option>`)
                }
            });
            $('#incharge_multi_search').multiselect('rebuild');
        });
    //Role dropdown fill
    $.getJSON('/api/Roles/')
        .done(function (data) {
            $('#role_multi_search').empty();
            var roleId = $("#hid_roleId").val();
            $.each(data, function (key, item) {
                if (roleId == item.Id) {
                    $('#role_multi_search').append(`<option selected class='role_checkbox' id="role_checkbox_${item.Id}" value='${item.Id}'>${item.RoleName}</option>`)
                } else {
                    $('#role_multi_search').append(`<option class='role_checkbox' id="role_checkbox_${item.Id}" value='${item.Id}'>${item.RoleName}</option>`)
                }
            });
            $('#role_multi_search').multiselect('rebuild');
        });

    //Explanation dropdown fill
    $.getJSON('/api/Explanations/')
        .done(function (data) {
            $('#explanation_multi_search').empty();
            var explantionId = $("#hid_explanationId").val();
            $.each(data, function (key, item) {
                if (explantionId == item.Id) {
                    $('#explanation_multi_search').append(`<option selected class='explanation_checkbox' id="explanation_checkbox_${item.Id}" value='${item.Id}'>${item.ExplanationName}</option>`)
                } else {
                    $('#explanation_multi_search').append(`<option class='explanation_checkbox' id="explanation_checkbox_${item.Id}" value='${item.Id}'>${item.ExplanationName}</option>`)
                }
            });
            $('#explanation_multi_search').multiselect('rebuild');
        });
    //Company dropdown fill
    $.getJSON('/api/Companies/')
        .done(function (data) {
            $('#company_multi_search').empty();
            var companyid = $("#hid_companyId").val();
            $.each(data, function (key, item) {
                if (companyid == item.Id) {
                    $('#company_multi_search').append(`<option selected class='comopany_checkbox' id="comopany_checkbox_${item.Id}" value='${item.Id}'>${item.CompanyName}</option>`)
                } else {
                    $('#company_multi_search').append(`<option class='comopany_checkbox' id="comopany_checkbox_${item.Id}" value='${item.Id}'>${item.CompanyName}</option>`)
                }
            });
            $('#company_multi_search').multiselect('rebuild');
        });
}
function SetRegistrationValueInHiddenField() {
    let name = $("#name_search").val();
    let sectionId = $('#section_multi_search').find(":selected").val();
    let departmentId = $('#dept_multi_search').find(":selected").val();
    let inchargeId = $('#incharge_multi_search').find(":selected").val();
    let roleId = $('#role_multi_search').find(":selected").val();
    let explantionId = $('#explanation_multi_search').find(":selected").val();
    let companyid = $('#company_multi_search').find(":selected").val();
    let companyName = $('#company_multi_search').find(":selected").text();

    $("#hid_name").val(name);
    $("#hid_sectionId").val(sectionId);

    $("#hid_departmentId").val(departmentId);
    $("#hid_inchargeId").val(inchargeId);
    $("#hid_roleId").val(roleId);
    $("#hid_explanationId").val(explantionId);
    $("#hid_companyId").val(companyid);
    $("#hid_companyName").val(companyName);
}