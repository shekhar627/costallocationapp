function CreateMultipleInput() {
    var multipleHtmlInputs = "";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='row'> ";

    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='section-code' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='division-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a section code.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";

    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-4'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='divison-name'>Section Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='divison-name' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='division-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a section name.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";

    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code' id='plus-icon'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <button type='button' class='form-control' onclick='CreateMultipleInput();' style='border:none;padding-right:139px;'><i class='fa fa-plus'></i></button>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";

    $('#main-div').append(multipleHtmlInputs);
}

function Department_MultiRow() {
    var multipleHtmlInputs = "";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='row'> ";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='divison-name'>Department Code<span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='divison-name' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a department code  .";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code'>Department Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='section-code' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a department name.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <label for='section-name'>Section Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "        <select id='section-name' class='form-control col-12' style='width:100% !important; border:1px solid #e3e3e3 !important;'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Select Section</option>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Section-1</option>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Section-2</option>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Section-3</option>";
    multipleHtmlInputs = multipleHtmlInputs + "        </select>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "            Please provide a department name.";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code' id='plus-icon'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <button type='button' class='form-control' onclick='Department_MultiRow();' style='border:none;padding-right:139px;'><i class='fa fa-plus'></i></button>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";

    $('#main-div-department').append(multipleHtmlInputs);
}
function CreateMultipleInputInCharge() {
    var multipleHtmlInputs = "";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='row'> ";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='divison-name'>In-Charge Code<span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='divison-name' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a In-Charge code  .";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-4'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code'>In-Charge Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='section-code' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a In-Charge name.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code' id='plus-icon'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <button type='button' class='form-control' onclick='CreateMultipleInput();' style='border:none;padding-right:139px;'><i class='fa fa-plus'></i></button>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";

    $('#main-div-in-charge').append(multipleHtmlInputs);
}



var expanded = false;

function SectionCheck() {
    var checkboxes = document.getElementById("sectionChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function DepartmentCheck() {
    var checkboxes = document.getElementById("departmentChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function InchargeCheck() {
    var checkboxes = document.getElementById("inchargeChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function RoleCheck() {
    var checkboxes = document.getElementById("RoleChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function ExplanationCheck() {
    var checkboxes = document.getElementById("ExplanationChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function CompanyCheck() {
    var checkboxes = document.getElementById("CompanyChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

//---------------------------master setting: get delete id list ------------------------------//
function GetCheckedIds(department_list_id) {
    var id = '';
    var sectionIds = $("#" + department_list_id + " tr input[type='checkbox']:checked").map(function () {
        return $(this).data('id')
    }).get();

    $.each(sectionIds, (index, data) => {
        id += data + ",";
    });
    return id;
}

//---------------modal insert--------------------//
function AddDivision() {
    var apiurl = "/api/sections/";
    let sectionName = $("#section-name").val().trim();
    if (sectionName == "") {
        $(".division-name-error").show();
        return false;

    }
    else {

        var data = {
            SectionName: sectionName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {

                Command: toastr["success"](data, "Success")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }
                $('#section-name').val('');

                $.getJSON('/api/sections/')
                    .done(function (data) {
                        $('#section_list_tbody').empty();
                        $.each(data, function (key, item) {
                            // Add a list item for the product.
                            $('#section_list_tbody').append(`<tr><td><input type="checkbox" class="section_list_chk" data-id='${item.Id}' /></td><td>${item.SectionName}</td></tr>`);
                        });
                    });
            },
            error: function () {
                alert("Error please try again");
            }
        });
    }
}

function AddDepartment() {
    var apiurl = "/api/Departments/";
    let departmentName = $("#department_name").val().trim();
    let sectionId = $("#section_list").val().trim();

    let isValidRequest = true;

    if (departmentName == "") {
        $(".department-name-error").show();
        isValidRequest = false;
    } else {
        $(".department-name-error").hide();
    }
    if (sectionId == "" || sectionId < 0) {
        $("#section_ist_error").show();
        isValidRequest = false;
    } else {
        $("#section_ist_error").hide();
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
            success: function (data) {

                Command: toastr["success"](data, "Success")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }
                $('#section-name').val('');

                $.getJSON('/api/departments/')
                    .done(function (data) {
                        $('#department_list_tbody').empty();
                        $.each(data, function (key, item) {
                            // Add a list item for the product.
                            $('#department_list_tbody').append(`<tr><td><input type="checkbox" class="department_list_chk" data-id='${item.Id}' /></td><td>${item.DepartmentName}</td><td>${item.SectionName}</td></tr>`);
                        });
                    });
            },
            error: function () {
                alert("Error please try again");
            }
        });
    }
}

function AddInCharge() {
    var apiurl = "/api/incharges/";
    let in_charge_name = $("#in_charge_name").val().trim();
    if (in_charge_name == "") {
        $(".department-name-error").show();
        return false;
    }
    else {
        $(".department-name-error").hide();
        var data = {
            InChargeName: in_charge_name
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {

                Command: toastr["success"](data, "Success")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }
                $('#section-name').val('');

                $.getJSON('/api/incharges/')
                    .done(function (data) {
                        $('#incharge_list_tbody').empty();
                        $.each(data, function (key, item) {
                            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.InChargeName}</td></tr>`);
                        });
                    });
            },
            error: function () {
                alert("Error please try again");
            }
        });
    }
}
function AddRoles() {
    var apiurl = "/api/Roles/";
    let roleName = $("#role_name").val().trim();
    if (roleName == "") {
        $(".department-name-error").show();
        return false;
    }
    else {
        $(".department-name-error").hide();
        var data = {
            RoleName: roleName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {

                Command: toastr["success"](data, "Success")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }
                $('#section-name').val('');

                $.getJSON('/api/Roles/')
                    .done(function (data) {
                        $('#incharge_list_tbody').empty();
                        $.each(data, function (key, item) {
                            // Add a list item for the product.
                            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.RoleName}</td></tr>`);
                        });
                    });
            },
            error: function () {
                alert("Error please try again");
            }
        });
    }
}
function AddExplanations() {
    var apiurl = "/api/Explanations/";
    let explanationName = $("#explanation_name").val().trim();
    if (explanationName == "") {
        $(".department-name-error").show();
        return false;
    }
    else {
        $(".department-name-error").hide();
        var data = {
            ExplanationName: explanationName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {

                Command: toastr["success"](data, "Success")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }
                $('#section-name').val('');

                $.getJSON('/api/Explanations/')
                    .done(function (data) {
                        $('#incharge_list_tbody').empty();
                        $.each(data, function (key, item) {
                            // Add a list item for the product.
                            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.ExplanationName}</td></tr>`);
                        });
                    });
            },
            error: function () {
                alert("Error please try again");
            }
        });
    }
}
function AddCompany() {
    var apiurl = "/api/Companies/";
    let companyName = $("#companyName").val().trim();
    if (companyName == "") {
        $(".department-name-error").show();
        return false;
    }
    else {
        $(".department-name-error").hide();
        var data = {
            CompanyName: companyName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {

                Command: toastr["success"](data, "Success")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }
                $('#section-name').val('');

                $.getJSON('/api/companies/')
                    .done(function (data) {
                        $('#incharge_list_tbody').empty();
                        $.each(data, function (key, item) {
                            // Add a list item for the product.
                            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.CompanyName}</td></tr>`);
                        });
                    });
            },
            error: function () {
                alert("Error please try again");
            }
        });
    }
}
function AddSalary() {
    var apiurl = "/api/Salaries/";
    let lowUnitPrice = $("#lowUnitPrice").val().trim();
    let highUnitPrice = $("#hightUnitPrice").val().trim();
    let gradePoints = $("#gradePoints").val().trim();

    let isValidRequest = true;
    if (lowUnitPrice == "") {
        $("#lowPrice").show();
        isValidRequest = false;
    }
    else {
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
            success: function (data) {

                Command: toastr["success"](data, "Success")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }
                $('#section-name').val('');

                $.getJSON('/api/Salaries/')
                    .done(function (data) {
                        $('#incharge_list_tbody').empty();
                        $.each(data, function (key, item) {
                            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.SalaryLowPoint} ～ ${item.SalaryHighPoint}</td><td>${item.SalaryGrade}</td></tr>`);
                        });
                    });
            },
            error: function () {
                alert("Error please try again");
            }
        });
    }
}
