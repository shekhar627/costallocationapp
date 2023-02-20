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





//$('#section_click').click(function (e) {
//    //alert("test");
//    //$(this).siblings('#sectionChks').fadeToggle(100);
//    var checkboxes = document.getElementById("sectionChks");
//    if (!expanded) {
//        checkboxes.style.display = "block";
//        //$(this).find("#sectionChks").slideToggle("fast");
//        expanded = true;
//    } else {
//        checkboxes.style.display = "none";
//        expanded = false;
//    }
//});

function DismissOtherDropdown(requestType) {
    var section_display = "";
    if (requestType.toLowerCase() != "section") {
        section_display = $('#sectionChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("sectionChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "dept") {
        section_display = "";
        section_display = $('#departmentChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("departmentChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "incharge") {
        section_display = "";
        section_display = $('#inchargeChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("inchargeChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "role") {
        section_display = "";
        section_display = $('#RoleChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("RoleChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "explanation") {
        section_display = "";
        section_display = $('#ExplanationChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("ExplanationChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "company") {
        section_display = "";
        section_display = $('#CompanyChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("CompanyChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }
}
//var expanded = false;
function SectionCheck() {
    DismissOtherDropdown("section");

    var checkboxes = document.getElementById("sectionChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        //$(this).find("#sectionChks").slideToggle("fast");
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function DepartmentCheck() {
    DismissOtherDropdown("dept");
    // section_display = $('#sectionChks');
    // if (section_display.css("display") == "block") {
    //     let checkboxes = document.getElementById("sectionChks");
    //     checkboxes.style.display = "none";
    //     expanded = false;
    // }

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
    DismissOtherDropdown("incharge");
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
    DismissOtherDropdown("role");
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
    DismissOtherDropdown("explanation");
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
    DismissOtherDropdown("company");
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
                $("#page_load_after_modal_close").val("yes");

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
            error: function (data) {
                alert(data.responseJSON.Message);
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
                $("#page_load_after_modal_close").val("yes");
                $("#department_name").val('');
                $("#section_list").val('');

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
            error: function (data) {
                alert(data.responseJSON.Message);
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
                $("#page_load_after_modal_close").val("yes");
                $("#in_charge_name").val('');
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
            error: function (data) {
                alert(data.responseJSON.Message);
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
                $("#page_load_after_modal_close").val("yes");
                $("#role_name").val('');
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
            error: function (data) {
                alert(data.responseJSON.Message);
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
                $("#page_load_after_modal_close").val("yes");
                $("#explanation_name").val('');

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
                $("#page_load_after_modal_close").val("yes");
                $("#companyName").val('');
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
            error: function (data) {
                alert(data.responseJSON.Message);
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
                            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.SalaryLowPointWithComma} ～ ${item.SalaryHighPointWithComma}</td><td>${item.SalaryGrade}</td></tr>`);
                        });
                    });
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}
function LoadGradeValue(rowId,sel) {
    CompanySectionOnChange_AddName(rowId,sel);
    // var _rowId = $("#add_name_row_no_hidden").val();
    // var _companyName = $('#company_row_' + _rowId).find(":selected").text();
    // var _sectionName = $('#section_row_' + _rowId).find(":selected").text();
    // var _unitPrice = $("#add_name_unit_price_hidden").val();

    // $.ajax({
    //     url: `/api/utilities/CompareGrade/${_unitPrice}`,
    //     type: 'GET',
    //     dataType: 'json',
    //     //data: {
    //     //    unitPrice: _unitPrice
    //     //},
    //     success: function (data) {
    //         $('#grade_edit_hidden').val(data.Id);
    //         if (_companyName.toLowerCase().indexOf("mw") > 0 || _sectionName.toLowerCase().indexOf("mw")>0) {
    //             $('#grade_row_' + _rowId).val(data.SalaryGrade);
    //         } else {
    //             $('#grade_row_' + _rowId).val('');
    //         }
    //     },
    //     error: function () {
    //         $('#grade_row_' + _rowId).val('');
    //     }
    // });
}

function NameListSort(sort_asc, sort_desc) {
    var nameAsc = $('#' + sort_asc).css('display');
    if (nameAsc == 'inline-block') {
        $('#' + sort_asc).css('display', 'none');
        $('#' + sort_desc).css('display', 'inline-block');
    }
    else {
        $('#' + sort_asc).css('display', 'inline-block');
        $('#' + sort_desc).css('display', 'none');
    }
}
function NameListSort_OnChange() {
    $('#name_asc').css('display', 'inline-block');
    $('#name_desc').css('display', 'none');

    $('#section_asc').css('display', 'inline-block');
    $('#section_desc').css('display', 'none');

    $('#department_asc').css('display', 'inline-block');
    $('#department_desc').css('display', 'none');

    $('#incharge_asc').css('display', 'inline-block');
    $('#incharge_desc').css('display', 'none');

    $('#role_asc').css('display', 'inline-block');
    $('#role_desc').css('display', 'none');

    $('#explanation_asc').css('display', 'inline-block');
    $('#explanation_desc').css('display', 'none');

    $('#company_asc').css('display', 'inline-block');
    $('#company_desc').css('display', 'none');

    $('#grade_asc').css('display', 'inline-block');
    $('#grade_desc').css('display', 'none');

    $('#unit_asc').css('display', 'inline-block');
    $('#unit_desc').css('display', 'none');
}

function NameList_DatatableLoad(data) {
    var tempCompanyName = "";
    var tempSubCode = "";
    var redMark = "";

    $('#name-list').DataTable({
        destroy: true,
        data: data,
        ordering: true,
        orderCellsTop: true,
        pageLength: 100,
        searching: false,
        bLengthChange: false,
        columns: [
            //{
            //    data: 'MarkedAsRed',
            //    render: function (markedAsRed) {
            //        redMark = markedAsRed;
            //        return null;
            //    }
            //},
            {
                data: 'EmployeeNameWithCodeRemarks',
                render: function (employeeNameWithCodeRemarks) {
                    var splittedString = employeeNameWithCodeRemarks.split('$');
                    if (splittedString[2] == 'true') {
                        return `<span style='color:red;' class='namelist_addname' onClick="loadSingleAssignmentDataForExistingEmployee('${splittedString[0]}')" data-toggle="modal" data-target="#modal_add_name">${splittedString[1]}</span>`;
                    }
                    else {
                        return `<span class='namelist_addname' onClick="loadSingleAssignmentDataForExistingEmployee('${splittedString[0]}')" data-toggle="modal" data-target="#modal_add_name">${splittedString[1]}</span>`;
                    }

                }
            },
            {
                data: 'SectionName'
            },
            {
                data: 'DepartmentName'
            },
            {
                data: 'InchargeName'
            },
            {
                data: 'RoleName'
            },
            {
                data: 'ExplanationName'
            },
            {
                data: 'CompanyName',
                render: function (companyName) {
                    tempCompanyName = companyName;
                    return companyName;
                }
            },
            {
                data: 'GradePoint',
                render: function (grade) {
                    if (tempCompanyName.toLowerCase() == "mw") {
                        return grade;
                    } else {
                        return "<span style='display:none;'>" + grade + "</span>";
                    }
                }
            },
            {
                data: 'UnitPrice'
            },
            {
                data: 'Id',
                render: function (Id) {
                    if (tempCompanyName.toLowerCase() == 'mw') {
                        return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-toggle="modal" data-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                    } else {
                        return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-toggle="modal" data-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                    }
                }
            }
        ]
    });
}

//$("#section_registration").on('hide', function () {
//    window.location.reload();
//});

function RetainRegistrationValue() {
    let name = $("#hid_name").val();
    let memo = $("#hid_memo").val(memo);
    let sectionId = $("#hid_sectionId").val(sectionId);
    let departmentId = $("#hid_departmentId").val(departmentId);
    let inchargeId = $("#hid_inchargeId").val(inchargeId);
    let roleId = $("#hid_roleId").val(roleId);
    let explantionId = $("#hid_explanationId").val(explantionId);
    let companyid = $("#hid_companyId").val();
    let companyName = $("#hid_companyName").val();
    let grade = $("#hid_grade").val();
    //let unitPrice = $("#hid_gradeId").val(unitPrice);
    let unitPrice = $("#hid_unitPrice").val();

    $("#identity_new").val(name);
    $("#memo_new").val(memo);
    // $('#section_new').find(":selected").val();
    // $('#department_new').find(":selected").val();
    // $('#incharge_new').find(":selected").val();
    // $('#role_new').find(":selected").val();
    // $('#explanation_new').find(":selected").val();
    // $('#company_new').find(":selected").val();
    if (companyName.toLowerCase() == 'mw') {
        $("#grade_new").val(grade);
    } else {
        $("#grade_new").val('');
    }

    $("#unitprice_new").val(unitPrice);

}


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

function ForecastSearchDropdownInLoad() {
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#sectionChks').empty();
            $('#sectionChks').append(`<label for='chk_sec_all'><input id="chk_sec_all" type="checkbox" checked data-checkwhat="chkSelect" value='sec_all'/>All</label>`)
            $.each(data, function (key, item) {
                $('#sectionChks').append(`<label for='section_checkbox_${item.Id}'><input class='section_checkbox' id="section_checkbox_${item.Id}"  type="checkbox" checked value='${item.Id}'/>${item.SectionName}</label>`)
            });
        });
    $.getJSON('/api/Departments/')
        .done(function (data) {
            $('#departmentChks').empty();
            $('#departmentChks').append(`<label for='chk_dept_all'><input id="chk_dept_all" type="checkbox" checked value='dept_all'/>All</label>`)
            $.each(data, function (key, item) {
                $('#departmentChks').append(`<label for='department_checkbox_${item.Id}'><input class='department_checkbox' id="department_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.DepartmentName}</label>`)
            });
        });
    $.getJSON('/api/InCharges/')
        .done(function (data) {
            $('#inchargeChks').empty();
            $('#inchargeChks').append(`<label for='chk_incharge_all'><input id="chk_incharge_all" type="checkbox" checked value='incharge_all'/>All</label>`);
            $.each(data, function (key, item) {
                $('#inchargeChks').append(`<label for='incharge_checkbox_${item.Id}'><input class='incharge_checkbox' id="incharge_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.InChargeName}</label>`)
            });
        });
    $.getJSON('/api/Roles/')
        .done(function (data) {
            $('#RoleChks').empty();
            $('#RoleChks').append(`<label for='chk_role_all'><input id="chk_role_all" type="checkbox" checked value='role_all'/>All</label>`);
            $.each(data, function (key, item) {
                $('#RoleChks').append(`<label for='role_checkbox_${item.Id}'><input class='role_checkbox' id="role_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.RoleName}</label>`)
            });
        });
    $.getJSON('/api/Explanations/')
        .done(function (data) {
            $('#ExplanationChks').empty();
            $('#ExplanationChks').append(`<label for='chk_explanation_all'><input id="chk_explanation_all" type="checkbox" checked value='explanation_all'/>All</label><br>`);
            $.each(data, function (key, item) {
                $('#ExplanationChks').append(`<label for='explanation_checkbox_${item.Id}'><input class='explanation_checkbox' id="explanation_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.ExplanationName}</label><br>`);

            });
        });
    $.getJSON('/api/Companies/')
        .done(function (data) {
            $('#CompanyChks').empty();
            $('#CompanyChks').append(`<label for='chk_comopany_all'><input id="chk_comopany_all" type="checkbox" checked value='comopany_all'/>All</label>`);
            $.each(data, function (key, item) {
                $('#CompanyChks').append(`<label for='comopany_checkbox_${item.Id}'><input class='comopany_checkbox' id="comopany_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.CompanyName}</label>`)
            });
        });
}

function checkPoint_click(e) {
    if (e.value <= 0) {
        $("#" + e.id).val('');
    }
}

function LoaderShow() {
    $("#addname_table").css("display", "none");
    $("#loading").css("display", "block");
}
function LoaderHide() {
    $("#addname_table").css("display", "block");
    $("#loading").css("display", "none");
}

let previousAssignmentRow = 0;
let unitPrices = [];

function loadSingleAssignmentData(id) {    
    var ajax_companyName = "";
    var ajax_sectionName = "";   

    $.ajax({
        url: '/api/Utilities/AssignmentById/' + id,
        type: 'GET',
        dataType: 'json',
        success: function (assignmentData) {
            $('#edit_model_span').html(assignmentData.EmployeeName);
            $('#name_edit').val(assignmentData.EmployeeName);
            $('#memo_edit').val(assignmentData.Remarks);
            $('#row_id_hidden_edit').val(assignmentData.Id);
            $.getJSON('/api/sections/')
                .done(function (data) { 
                    $('#section_edit').empty();   
                    $('#section_edit').append(`<option value=''>Select Section</option>`);
                    $.each(data, function (key, item) {
                        if (item.Id == assignmentData.SectionId) {
                            $('#section_edit').append(`<option value='${item.Id}' selected>${item.SectionName}</option>`);
                        } else {
                            $('#section_edit').append(`<option value='${item.Id}'>${item.SectionName}</option>`);
                        }

                    });
                });
            ajax_companyName = assignmentData.CompanyName;
            ajax_sectionName = assignmentData.SectionName;
            GetAssignedGradeId(assignmentData.GradeId);
            var isGradeShow = IsGradeShow(ajax_companyName,ajax_sectionName);

            $.getJSON(`/api/departments`)
                .done(function (data) {
                    $('#department_edit').empty();
                    $('#department_edit').append(`<option value=''>Select Department</option>`);
                    $.each(data, function (key, item) {
                        if (item.Id == assignmentData.DepartmentId) {
                            $('#department_edit').append(`<option value='${item.Id}' selected>${item.DepartmentName}</option>`);
                        } else {
                            $('#department_edit').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
                        }

                    });
                });            

            $.getJSON('/api/Explanations/')
                .done(function (data) {
                    $('#explanation_edit').empty();
                    $('#explanation_edit').append(`<option value=''>Select Explanation</option>`);
                    $.each(data, function (key, item) {
                        if (item.Id == assignmentData.ExplanationId) {
                            $('#explanation_edit').append(`<option value='${item.Id}' selected>${item.ExplanationName}</option>`);
                        } else {
                            $('#explanation_edit').append(`<option value='${item.Id}'>${item.ExplanationName}</option>`);
                        }

                    });
                });

            $.getJSON('/api/Companies/')
                .done(function (data) {
                    $('#company_edit').empty();
                    $('#company_edit').append(`<option value=''>Select Company</option>`);
                    $.each(data, function (key, item) {
                        if (item.Id == assignmentData.CompanyId) {
                            $('#company_edit').append(`<option value='${item.Id}' selected>${item.CompanyName}</option>`);
                        } else {
                            $('#company_edit').append(`<option value='${item.Id}'>${item.CompanyName}</option>`);
                        }

                    });
                });                        
                            
            if(isGradeShow){                
                $('#unitprice_edit').attr("readonly",true);
                $('#unitprice_edit').css('opacity', '0.6');
                $('#unitprice_edit').val(assignmentData.UnitPrice);
            }else{                
                $('#unitprice_edit').attr("readonly",false) 
                $('#unitprice_edit').css('opacity', '');
                $('#unitprice_edit').val(assignmentData.UnitPrice);
            }            

            $.getJSON(`/api/Utilities/GetAllSalaries/`)
                .done(function (data) {
                    $('#grade_edit').empty();
                    $('#grade_edit').append(`<option value='-1'>Select Grade</option>`);
                    $.each(data, function (key, item) {
                        $('#grade_edit').append(`<option value='${item.GradeId}'>${item.GradeName}</option>`);
                    });                                        
                    //var tempGradeIdWithSalaryType = $("#hid_gradeIdWithSalaryType").val();
                    var tempGradeIdWithSalaryType = $("#hid_gradeIdWithSalaryType").val();

                    if ((assignmentData.CompanyName != '' && assignmentData.CompanyName != null) || assignmentData.SectionName != '' && assignmentData.SectionName != null) {
                        if (isGradeShow) {
                            $("#grade_edit option[value=" + tempGradeIdWithSalaryType + "]").attr("selected", "selected");
                        } else {
                            $("#grade_edit").attr('style', 'pointer-events: none;opacity:.7;');
                            $("#grade_edit").attr('onclick', 'return false;');
                            $("#grade_edit").attr('onkeydown', 'return false;');
                        }
                    } else {
                        $("#grade_edit").attr('style', 'pointer-events: none;opacity:.7;');
                        $("#grade_edit").attr('onclick', 'return false;');
                        $("#grade_edit").attr('onkeydown', 'return false;');
                    }
                });
            $('#grade_edit_hidden').val(assignmentData.GradeId);
            //$('#unitprice_edit').val(assignmentData.GradeId);
            //if (assignmentData.CompanyName != '' && assignmentData.CompanyName != null) {
            //    if (assignmentData.CompanyName.toLowerCase() == "mw" || assignmentData.SectionName.toLowerCase() == 'mw') {
            //        $('#grade_edit').val(assignmentData.GradePoint);
            //    } else {
            //        $('#grade_edit').val("");
            //    }
            //} else {
            //    $('#grade_edit').val("");
            //}



        },
        error: function () {
            //$('#add_name_table_1 tbody').empty();
        }
    });
}

function loadSingleAssignmentDataForExistingEmployee(employeeName) {
    $('#add_name_table_2 thead .sub_thead').remove();
    $('#add_name_table_2 tbody').empty();
    $('#add_name_add').css('display', 'inline-block');
    $.ajax({
        url: `/api/Utilities/GetEmployeesByName/${employeeName}`,
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function (assignmentData) {
            $('#fixed_hidden_name').val(assignmentData[0].EmployeeName);

            var count = 1;
            var firstRowUnitPrice = '';

            $.each(assignmentData, function (index, item) {
                unitPrices.push(item.UnitPrice);
                previousAssignmentRow = item.SubCode;
                if (count == 1) {
                    firstRowUnitPrice = item.UnitPrice;
                }

                tr += '<tr class="sub_thead">';
                tr += `<td>${item.SubCode}</td>`;
                var colorMark = '';
                if (firstRowUnitPrice.toString() == item.UnitPrice.toString()) {
                    tr += `<td><input type='text' class=" col-12" style='width: 141px;' value='${item.EmployeeName}' readonly/></td>`;
                    tr += `<td><input type='text' class=" col-12" style='width: 47px;' value='${item.SubCode}' readonly /></td>`;
                } else {
                    tr += `<td><input type='text' class=" col-12" style='width: 141px;color:red;' value='${item.EmployeeName}' readonly/></td>`;
                    tr += `<td><input type='text' class=" col-12" style='width: 47px;color:red;' value='${item.SubCode}' readonly /></td>`;
                }


                tr += `<td><input type='text' class=" col-12" style='width: 89px;' value='${item.Remarks}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 87px;' value='${item.SectionName}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 89px;' value='${item.DepartmentName}' readonly /></td>`;                        
                tr += `<td><input type='text' class=" col-12" style='width: 190px;' value='${item.ExplanationName}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width:81px;'   value='${item.CompanyName}' readonly /></td>`;
                if (item.CompanyName.toLowerCase() == "mw") {
                    tr += `<td><input type='text' class=" col-12" value='${item.GradePoint}' readonly style='width: 55px;' /></td>`;
                }
                else {
                    tr += `<td><input type='text' class=" col-12" value='' readonly style='width: 55px;' /></td>`;
                }
                tr += `<td><input type='text' class="col-12" id="add_name_year_${count}" style='width: 72px;' value='${2022}' readonly /></td>`;
                tr += `<td><input type='text' class="col-12" id="add_name_unit_price_${count}" style='width: 72px;' value='${item.UnitPrice}' readonly /></td>`;
                if (count == 1) {
                    $("#unit_price_first_project_hid").val(item.UnitPrice);
                    tr += `<td rowspan='${assignmentData.length}' style='text-align:center;'><a href='javascript:void(0);' onClick="addNew()"> <i class="fa fa-plus" aria-hidden="true"></i></a></td>`;
                }

                tr += '</tr>';
                count++;
                $('#add_name_table_2 thead').append(tr);
                var tr = "";
            });

        },
        error: function () {

        }
    });
}

function addNew() {
    //
    previousAssignmentRow++;
    var text = `
                <tr data-id="${previousAssignmentRow}" id="row_${previousAssignmentRow}">
                    <td>${previousAssignmentRow}</td>
                    <td><input class=" col-12" id="identity_row_${previousAssignmentRow}" style='width: 141px;' value="${$('#fixed_hidden_name').val()}" readonly /></td>
                    <td><input class="" value="${previousAssignmentRow}"  style='width: 47px;' id='sub_code_row_${previousAssignmentRow}' readonly /></td>
                    <td><input class=" col-12" style='width: 89px;' id="memo_row_${previousAssignmentRow}" /></td>`;

    // for add new
    $.ajax({
        url: '/api/Sections/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function (data) {
            text += '<td>';
            text += `
                                    <select id="section_row_${previousAssignmentRow}" style='width: 87px;'class=" col-12 section_row" onchange="LoadGradeValue(${previousAssignmentRow},this);">
                                        <option value=''>Select Section</option>
                                `;
            $.each(data, function (key, item) {
                text += `<option value = '${item.Id}'> ${item.SectionName}</option>`;
            });
            text += '</select></td>';
        },
        error: function () {

        }
    });

    $.ajax({
        url: '/api/departments/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function (data) {
            text += '<td>';
            text += `
                                <select id="department_row_${previousAssignmentRow}" class=" col-12">
                                    <option value=''>Select Department</option>
                            `;
            $.each(data, function (key, item) {
                text += `<option value = '${item.Id}'> ${item.DepartmentName}</option>`;
            });
            text += '</select></td>';
        },
        error: function () {

        }
    });            
    $.ajax({
        url: '/api/Explanations/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function (data) {
            text += '<td>';
            text += `
                                    <select id="explain_row_${previousAssignmentRow}" style='width: 190px;' class=" col-12">
                                        <option value=''>Select Allocation</option>
                                `;
            $.each(data, function (key, item) {
                text += `<option value = '${item.Id}'> ${item.ExplanationName}</option>`;
            });
            text += '</select></td>';
        },
        error: function () {

        }
    });

    $.ajax({
        url: '/api/Companies/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function (data) {
            text += '<td>';
            text += `
                                    <select id="company_row_${previousAssignmentRow}" style='width:81px;' class=" col-12 add_company" onchange="LoadGradeValue(${previousAssignmentRow},this);">
                                        <option value=''>Select Company</option>
                                `;
            $.each(data, function (key, item) {
                text += `<option value = '${item.Id}'> ${item.CompanyName}</option>`;
            });
            text += '</select></td>';
        },
        error: function () {

        }
    });
    $.ajax({
        url: '/api/Utilities/GetAllSalaries/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function (data) {
            text += '<td>';
            text += `<select id="grade_row_${previousAssignmentRow}" style='width: 113px;' class=" col-12" onchange="LoadGradeValue(${previousAssignmentRow},this);"><option value=''>Select Grade</option>`;
            $.each(data, function (key, item) {
                text += `<option value = '${item.GradeId}'> ${item.GradeName}</option>`;
            });
            text += '</select></td>';
        },
        error: function () {

        }
    });
    text += `<td><select id='year_row_${previousAssignmentRow}' class='col-12'><option value='-1'>Select Year</option><option value='2022'>2022</option></select>`;
    //text += `<td><input class=" col-12" id="grade_row_${previousAssignmentRow}" readonly style='width: 55px;'/><input type='hidden' id='grade_row_new_${previousAssignmentRow}' /></td>`;
    //text += `<td><input class=" col-12" id="grade_row_${previousAssignmentRow}" readonly style='width: 55px;'/><input type='hidden' id='grade_row_new_${previousAssignmentRow}' /></td>`;
    text += `<td><input class=" col-12" id="unitprice_row_${previousAssignmentRow}" style='width: 72px;' onChange="unitPriceChange(${previousAssignmentRow},this)"/></td>
            <td style='text-align:center;'><a href='javascript:void(0);' id='remove_row_${previousAssignmentRow}' onClick="removeTr(${previousAssignmentRow})"><i class="fa fa-trash-o" aria-hidden="true"></i></a></td>
            </tr>`;            
    // text += `
    //                  <td><input class=" col-12" id="grade_row_${previousAssignmentRow}" readonly style='width: 55px;'/><input type='hidden' id='grade_row_new_${previousAssignmentRow}' /></td>
    //                  <td><input class=" col-12" id="unitprice_row_${previousAssignmentRow}" style='width: 72px;' onChange="unitPriceChange(${previousAssignmentRow},this)"/></td>
    //                  <td style='text-align:center;'><a href='javascript:void(0);' id='remove_row_${previousAssignmentRow}' onClick="removeTr(${previousAssignmentRow})"><i class="fa fa-trash-o" aria-hidden="true"></i></a></td>
    //             </tr>

    //             `;

    $(`#remove_row_${previousAssignmentRow - 1}`).css('display', 'none');

    $('#add_name_table_2 tbody').append(text);

}

$('.add_company').on('change', function () {
    var _companyName = $(this).val();
    var _unitPrice = $("#unitprice_edit").val();            
});

function removeTr(rowId) {
    var rowElement = $('#row_' + rowId);
    rowElement.remove();
    previousAssignmentRow--;
    $(`#remove_row_${previousAssignmentRow}`).css('display', 'block');
}

function unitPriceChange(rowId, element) {
    let _unitPrice = $(element).val();
    $("#add_name_unit_price_hidden").val(_unitPrice);
    $("#add_name_row_no_hidden").val(rowId);
    let _unitPrideFirstProject = $("#unit_price_first_project_hid").val();
    let _unitPrideFirstProjectWithComma = $("#add_name_unit_price_1").val();

    var _rowId = $("#add_name_row_no_hidden").val();
    var _companyName = $('#company_row_' + _rowId).find(":selected").text();
    var _sectionName = $('#section_row_' + _rowId).find(":selected").text();
    console.log("_sectionName: "+_sectionName);

    if (_unitPrideFirstProject != _unitPrice && _unitPrideFirstProjectWithComma != _unitPrice) {
        $('#identity_row_' + rowId).css('color', 'red');
        $('#sub_code_row_' + rowId).css('color', 'red');
    } else {
        $('#identity_row_' + rowId).css('color', '');
        $('#sub_code_row_' + rowId).css('color', '');
    }

    $.ajax({
        url: `/api/utilities/CompareGrade/${_unitPrice}`,
        type: 'GET',
        dataType: 'json',
        success: function (data) {
            $('#grade_edit_hidden').val(data.Id);
            
            if (_companyName.toLowerCase().indexOf("mw") > 0 || _sectionName.toLowerCase().indexOf("mw")> 0) {
                $('#grade_row_' + rowId).attr('data-id', data.Id);
                $('#grade_row_new_' + rowId).val(data.Id);
                $('#grade_row_' + _rowId).val(data.SalaryGrade);
            } else {
                $('#grade_row_' + rowId).attr('data-id', data.Id);
                $('#grade_row_new_' + rowId).val(data.Id);
                $('#grade_row_' + _rowId).val('');
            }
        },
        error: function () {
            $('#grade_row_' + rowId).attr('data-id', '');
            $('#grade_row_' + rowId).val('');
            $('#grade_row_new_' + rowId).val('');
        }
    });
}


$(document).ready(function () {
    $('#add_name_name').click(function () {
        NameListSort("name_asc", "name_desc");
    });
    $('#add_name_subcode').click(function () {
        NameListSort("subcode_asc", "subcode_desc");
    });
    $('#add_name_memo').click(function () {
        NameListSort("memo_asc", "memo_desc");
    });
    $('#add_name_section').click(function () {
        NameListSort("section_asc", "section_desc");
    });
    $('#add_name_department').click(function () {
        NameListSort("department_asc", "department_desc");
    });
    $('#add_name_incharge').click(function () {
        NameListSort("incharge_asc", "incharge_desc");
    });
    $('#add_name_role').click(function () {
        NameListSort("role_asc", "role_desc");
    });
    $('#add_name_explanation').click(function () {
        NameListSort("explanation_asc", "explanation_desc");
    });
    $('#add_name_company').click(function () {
        NameListSort("company_asc", "company_desc");
    });
    $('#add_name_grade').click(function () {
        NameListSort("grade_asc", "grade_desc");
    });
    $('#add_name_unit').click(function () {
        NameListSort("unit_asc", "unit_desc");
    });            
    // for add new
    $.getJSON('/api/Explanations/')
        .done(function (data) {
            $('#explanation_new').empty();
            $('#explanation_new').append(`<option value=''>Select Explanation</option>`);
            $.each(data, function (key, item) {
                $('#explanation_new').append(`<option value='${item.Id}'>${item.ExplanationName}</option>`);
            });
        });
    // for search
    $.getJSON('/api/Explanations/')
        .done(function (data) {
            $('#explanation_search').empty();
            $('#explanation_search').append(`<option value=''>Select Allocation</option>`);
            $.each(data, function (key, item) {
                $('#explanation_search').append(`<option value='${item.Id}'>${item.ExplanationName}</option>`);
            });
        });
    // for add new
    $.getJSON('/api/Companies/')
        .done(function (data) {
            $('#company_new').empty();
            $('#company_new').append(`<option value=''>Select Company</option>`);
            $.each(data, function (key, item) {
                $('#company_new').append(`<option value='${item.Id}'>${item.CompanyName}</option>`);
            });
        });
    // for search
    $.getJSON('/api/Companies/')
        .done(function (data) {
            $('#company_search').empty();
            $('#company_search').append(`<option value=''>Select Company</option>`);
            $.each(data, function (key, item) {
                $('#company_search').append(`<option value='${item.Id}'>${item.CompanyName}</option>`);
            });
        });
    // for add new
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_new').empty();
            $('#section_new').append(`<option value=''>Select Section</option>`);
            $.each(data, function (key, item) {
                $('#section_new').append(`<option value='${item.Id}'>${item.SectionName}</option>`);
            });
        });
    // for search
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_search').empty();
            $('#section_search').append(`<option value=''>Select Section</option>`);
            $.each(data, function (key, item) {
                $('#section_search').append(`<option value='${item.Id}'>${item.SectionName}</option>`);
            });
        });
    // for search
    $.getJSON('/api/Departments/')
        .done(function (data) {
            $('#department_search').empty();
            $('#department_search').append(`<option value=''>Select Departments</option>`);
            $.each(data, function (key, item) {
                $('#department_search').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
            });
        });

    // for new
    $.getJSON('/api/Departments/')
        .done(function (data) {
            $('#department_new').empty();
            $('#department_new').append(`<option value=''>Select Departments</option>`);
            $.each(data, function (key, item) {
                $('#department_new').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
            });
        });

    // for edit
    $.getJSON('/api/Departments/')
        .done(function (data) {
            $('#department_edit').empty();
            $('#department_edit').append(`<option value=''>Select Departments</option>`);
            $.each(data, function (key, item) {
                $('#department_edit').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
            });
        });

    $('#unitprice_new').on('change', function () {
        var _unitPrice = $(this).val();

        $.ajax({
            url: `/api/utilities/CompareGrade/${_unitPrice}`,
            type: 'GET',
            dataType: 'json',
            //data: {
            //    unitPrice: _unitPrice
            //},
            success: function (data) {
                $('#grade_new_hidden').val(data.Id);
                $('#grade_new').val(data.SalaryGrade);
            },
            error: function () {
                $('#grade_new_hidden').val('');
                $('#grade_new').val('');
            }
        });
    });
    
    $('#unitprice_edit').on('change', function () {
        var _unitPrice = $(this).val();
        var _companyName = $("#company_edit option:selected").text();
        var _sectionName = $("#section_edit option:selected").text();
        console.log("_companyName: "+_companyName);
        console.log("_sectionName: "+_sectionName);

        $.ajax({
            url: `/api/utilities/CompareGrade/${_unitPrice}`,
            type: 'GET',
            dataType: 'json',                   
            success: function (data) {                        
                if (data != null) {
                    $('#grade_edit_hidden').val(data.Id);
                    if (_companyName.toLowerCase() == "mw" || _sectionName.toLowerCase() == "mw") {
                        $('#grade_edit').val(data.SalaryGrade);
                    } else {
                        $('#grade_edit').val('');
                    }
                } else {
                    if (_companyName.toLowerCase() == "mw" || _sectionName.toLowerCase() == "mw") {
                        $('#grade_edit').val();
                    } else {
                        $('#grade_edit').val('');
                    }
                }   
            },
            error: function () {
                $('#grade_edit').val('');
                $('#grade_edit_hidden').val('');
            }
        });
    });

    $('#company_edit').on('change', function () {
        var _companyName = $("#company_edit option:selected").text();
        var _sectionName = $("#section_edit option:selected").text();
        var _unitPrice = $("#unitprice_edit").val();

        $.ajax({
            url: `/api/utilities/CompareGrade/${_unitPrice}`,
            type: 'GET',
            dataType: 'json',
            success: function (data) {
                if (data != null) {
                    $('#grade_edit_hidden').val(data.Id);
                    if (_companyName.toLowerCase() == "mw" || _sectionName.toLowerCase() == "mw") {
                        $('#grade_edit').val(data.SalaryGrade);
                    } else {
                        $('#grade_edit').val('');
                    }
                } else {
                    if (_companyName.toLowerCase() == "mw" || _sectionName.toLowerCase() == "mw") {
                        $('#grade_edit').val();
                    } else {
                        $('#grade_edit').val('');
                    }
                } 

            },
            error: function () {
                $('#grade_edit').val('');
                $('#grade_edit_hidden').val('');
            }
        });
    });
    $('#section_edit').on('change', function () {
        var _companyName = $("#company_edit option:selected").text();
        var _sectionName = $("#section_edit option:selected").text();
        var _unitPrice = $("#unitprice_edit").val();

        $.ajax({
            url: `/api/utilities/CompareGrade/${_unitPrice}`,
            type: 'GET',
            dataType: 'json',
            success: function (data) {
                if (data != null) {
                    $('#grade_edit_hidden').val(data.Id);
                    if (_companyName.toLowerCase() == "mw" || _sectionName.toLowerCase() == "mw") {
                        $('#grade_edit').val(data.SalaryGrade);
                    } else {
                        $('#grade_edit').val('');
                    }
                } else {
                    if (_companyName.toLowerCase() == "mw" || _sectionName.toLowerCase() == "mw") {
                        $('#grade_edit').val();
                    } else {
                        $('#grade_edit').val('');
                    }
                } 
            },
            error: function () {
                $('#grade_edit').val('');
                $('#grade_edit_hidden').val('');
            }
        });
    });
    $('#add_name_search_button').on('click', function () {
        AddNameSearchResult();             
    });

    $('#add_name_add_new').on('click', function () {
        var employeeName = $('#identity_new').val();
        var sectionId = $('#section_new').find(":selected").val();
        var inchargeId = $('#incharge_new').find(":selected").val();
        var departmentId = $('#department_new').find(":selected").val();
        var roleId = $('#role_new').find(":selected").val();
        var companyId = $('#company_new').find(":selected").val();
        var explanationId = $('#explanation_new').find(":selected").val();
        var unitPrice = $('#unitprice_new').val();
        var gradeId = $('#grade_new_hidden').val();
        var remarks = $('#memo_new').val();

        var data = {
            EmployeeName: employeeName,
            SectionId: sectionId,
            DepartmentId: departmentId,
            InchargeId: inchargeId,
            RoleId: roleId,
            ExplanationId: explanationId,
            CompanyId: companyId,
            UnitPrice: unitPrice,
            GradeId: gradeId,
            Remarks: remarks,
            SubCode: 1
        };
        $.ajax({
            url: '/api/Employees',
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
                alert('Success');
                location.reload();
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });

    });

    $('#add_name_edit').on('click', function () {

        var sectionId = $('#section_edit').find(":selected").val();
        var inchargeId = $('#incharge_edit').find(":selected").val();
        var departmentId = $('#department_edit').find(":selected").val();
        var roleId = $('#role_edit').find(":selected").val();
        var companyId = $('#company_edit').find(":selected").val();
        var explanationId = $('#explanation_edit').find(":selected").val();
        var unitPrice = $('#unitprice_edit').val();
        var gradeId = $('#grade_edit_hidden').val();
        var rowId = $('#row_id_hidden_edit').val();
        var remarks = $('#memo_edit').val();

        var dataSingleAssignmentUpdate = {
            Id: rowId,
            SectionId: sectionId,
            DepartmentId: departmentId,
            InchargeId: inchargeId,
            RoleId: roleId,
            ExplanationId: explanationId,
            CompanyId: companyId,
            UnitPrice: unitPrice,
            GradeId: gradeId,
            Remarks: remarks,
        };

        $.ajax({
            url: '/api/Employees',
            type: 'PUT',
            dataType: 'json',
            data: dataSingleAssignmentUpdate,
            success: function (data) {

                //ToastMessageSuccess(data);
                alert('Success');
                $('#modal_edit_name').modal('hide');
                AddNameSearchResult();
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    });
});

function onSave() {
    const rows = document.querySelectorAll("#add_name_table_2 > tbody > tr");
    //$('#add_name_add').css('display','none');
    console.log(rows);
    $.each(rows, function (index, data) {
        var rowId = data.dataset.id;

        var employeeName = $('#identity_row_' + rowId).val();
        var sectionId = $('#section_row_' + rowId).find(":selected").val();
        var sectionId = $('#section_row_' + rowId).find(":selected").text();
        var departmentId = $('#department_row_' + rowId).find(":selected").val();
        var inchargeId = $('#incharge_row_' + rowId).find(":selected").val();
        var roleId = $('#role_row_' + rowId).find(":selected").val();
        var companyId = $('#company_row_' + rowId).find(":selected").val();
        var explanationId = $('#explain_row_' + rowId).find(":selected").val();
        var unitPrice = $('#unitprice_row_' + rowId).val();
        var gradeId = $('#grade_row_new_' + rowId).val();
        var remarks = $('#memo_row_' + rowId).val();
        var subCode = $('#sub_code_row_' + rowId).val();
        
        var isGradeShow = IsGradeShow();


        var data = {
            EmployeeName: employeeName,
            SectionId: sectionId,
            DepartmentId: departmentId,
            InchargeId: inchargeId,
            RoleId: roleId,
            ExplanationId: explanationId,
            CompanyId: companyId,
            UnitPrice: unitPrice,
            GradeId: gradeId,
            Remarks: remarks,
            SubCode: subCode,

        };
        $.ajax({
            url: '/api/Employees',
            type: 'POST',
            async: false,
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
                alert('Success');
                location.reload();
                //$('#modal_add_name_new').modal('toggle');
            },
            error: function () {
                alert("Error please try again");
                //$('#modal_add_name_new').modal('toggle');
            }
        });
    });

}
function AddNameSearchResult(){
    LoaderShow();
    var sectionId = $('#section_search').find(":selected").val();
    var inchargeId = $('#incharge_search').find(":selected").val();
    var departmentId = $('#department_search').find(":selected").val();
    var roleId = $('#role_search').find(":selected").val();
    var companyId = $('#company_search').find(":selected").val();
    var explanationId = $('#explanation_search').find(":selected").val();
    var employeeName = $('#identity_search').val();

    if (departmentId == undefined) {
        departmentId = '';
    }
    var data_info = { employeeName: employeeName, sectionId: sectionId, departmentId: departmentId, inchargeId: '', roleId: '', explanationId: explanationId, companyId: companyId, status: '' };
    
    $.ajax({
        url: `/api/utilities/SearchEmployee`,
        type: 'GET',
        dataType: 'json',
        data: data_info,
        success: function (data) {
            var employeeNameTemp = '';
            var companyNameTemp = '';
            $('#add_name_table_1 tbody').empty();
            let redMark = false;


            $('#add_name_table_1').DataTable({
                destroy: true,
                data: data,
                ordering: true,
                pageLength: 100,
                searching: false,
                bLengthChange: false,
                dom: 'lifrtip',
                columns: [
                    {
                        data: 'MarkedAsRed',
                        render: function (markedAsRed) {
                            redMark = markedAsRed;
                            return null;
                        }
                    },
                    {
                        data: 'EmployeeName',
                        render: function (employeeName) {
                            employeeNameTemp = employeeName;
                            matchColor = "";
                            if (redMark) {
                                return `<span style='color:red'>${employeeName}</span>`;
                            }
                            else {
                                return `<span>${employeeName}</span>`;
                            }
                        },
                        orderable: true,
                    },
                    {
                        data: 'AddNameSubCode',
                        render: function (subCode) {
                            if (redMark) {
                                return `<span style='color:red;'>${subCode}</span></td>`;
                            }
                            else {
                                return `<span >${subCode}</span>`;
                            }
                        },
                        searching: false,
                    },
                    {
                        data: 'Remarks',
                    },
                    {
                        data: 'SectionName',
                    },
                    {
                        data: 'DepartmentName',

                    },
                    {
                        data: 'ExplanationName',
                    },
                    {
                        data: 'CompanyName',
                        render: function (companyName) {
                            companyNameTemp = companyName;
                            return companyName;
                        },
                    },
                    {
                        data: 'GradePoint',
                        render: function (gradePoint) {
                            if (companyNameTemp.toLowerCase() == 'mw') {
                                return gradePoint;
                            }
                            else {
                                return "<span style='display:none;'>" + gradePoint + "</span>";
                            }
                        },
                        orderable: true,
                    },
                    {
                        data: 'UnitPrice',
                        render: function (unitPrice) {
                            return `<span style='text-align:right;display:block;padding-right:2px;'>${unitPrice}</span>`;
                        },
                        orderable: true,
                    },
                    {
                        data: 'Id',
                        render: function (id) {
                            return `<td class="add_name_btn_group"><a href="javascript:void(0);" onClick="loadSingleAssignmentDataForExistingEmployee('${employeeNameTemp}')" class="link_add" data-toggle="modal" data-target="#modal_add_name">add name</a> <a href="javascript:void(0);" onClick="loadSingleAssignmentData(${id})" class="link_edit" data-toggle="modal" data-target="#modal_edit_name">edit</a></td>`;
                        },
                        searching: false,
                        orderable: false,
                    }
                ]
            });

            LoaderHide();
        },
        error: function () {
            $('#add_name_table_1 tbody').empty();
        }
    });
}
function CompanySectionOnChange_AddName(rowId,sel){
    $("#add_name_row_no_hidden").val(rowId);
    var _rowId = $("#add_name_row_no_hidden").val();
    console.log("_rowId: "+_rowId);

    
    var _companyName = $('#company_row_' + _rowId).find(":selected").text();
    var _sectionName = $('#section_row_' + _rowId).find(":selected").text();
    var _unitPrice = $("#add_name_unit_price_hidden").val();
    var _gradePoint = $("#grade_row_" + _rowId).val();
    var _selectDepartment = $("#department_row_"+ _rowId).val();
    var _selectYear = $("#year_row_"+ _rowId).val();
    if(_selectYear == '' || _selectYear <0){
        _selectYear = 2022;
    }
    var isGradeShow = IsGradeShow(_companyName,_sectionName);
    var selectGradeId = "";
    var selectGradeLowPoints = "";

    var tempHidGradeVal = $('#grade_addname_hidden').val();

    $.ajax({
        url: `/api/utilities/GetUnitPrice`,
        type: 'GET',
        dataType: 'json',        
        data:'gradeId='+_gradePoint+"&departmentId="+_selectDepartment+"&year="+_selectYear,
        success: function (data) {
            if (isGradeShow) {                
                if(data.Id != '' && data.Id >0){
                    selectGradeId = data.Id;
                    selectGradeLowPoints = data.GradeLowPoints;                    
                }else{
                    selectGradeId = -1;
                    selectGradeLowPoints = '';
                }                
                $('#unitprice_row_'+ _rowId).val(selectGradeLowPoints);                 

                $('#unitprice_row_'+ _rowId).attr("readonly",true);
                $('#unitprice_row_'+ _rowId).css('opacity', '0.6');

                $("#grade_row_"+ _rowId).attr('style', '');
                $("#grade_row_"+ _rowId).attr('onclick', '');
                $("#grade_row_"+ _rowId).attr('onkeydown', '');

                // if(tempHidGradeVal >0){
                //     $("#grade_edit option[value=" + tempHidGradeVal + "]").attr("selected", false);
                // }

                $("#grade_row_"+_rowId+" option[value=" + selectGradeId + "]").attr("selected", "selected");
            } else {                                                    
                $("#grade_row_"+_rowId+" option").attr("selected", false);
                $("#grade_row_"+_rowId+" option[value='-1']").attr("selected", "selected");

                $('#unitprice_row_'+ _rowId).attr("readonly",false);
                $('#unitprice_row_'+ _rowId).css('opacity', '');
                
                $("#grade_row_"+ _rowId).attr('style', 'pointer-events: none;opacity:.7;');
                $("#grade_row_"+ _rowId).attr('onclick', 'return false;');
                $("#grade_row_"+ _rowId).attr('onkeydown', 'return false;');
            }
            $('#grade_edit_hidden').val(data.Id);

            
            if (_companyName.toLowerCase().indexOf("mw") > 0 | _sectionName.toLowerCase().indexOf("mw") > 0) {
                $('#grade_row_' + _rowId).attr('data-id', data.Id);
                $('#grade_row_' + _rowId).val(data.SalaryGrade);
                $('#grade_row_new_' + _rowId).val(data.Id);
            } else {
                $('#grade_row_' + _rowId).attr('data-id', '');
                $('#grade_row_' + _rowId).val('');
                $('#grade_row_new_' + _rowId).val('');
            }
        },
        error: function () {
            $('#grade_row_' + _rowId).attr('data-id', '');
            $('#grade_row_' + _rowId).val('');
            $('#grade_row_new_' + _rowId).val('');
        }
    });            
}
function IsGradeShow(ajax_companyName,ajax_sectionName){
    var isGradeShow = false;
    if(ajax_companyName.toLowerCase() == 'mw' || ajax_sectionName.toLowerCase() == 'mw'){
        isGradeShow = true;
    }
    return isGradeShow;
}
function GetAssignedGradeId(salaryTypeId){
    var returnGradeId = "";
    $.ajax({
        url: `/api/utilities/GetGradeId/${salaryTypeId}`,
        type: 'GET',
        dataType: 'text',
        success: function (data) {
            if (data != null) {
                returnGradeId = data;
                $("#hid_gradeIdWithSalaryType").val(returnGradeId);
            } 
        },
        error: function () {
        }
    });    
}
