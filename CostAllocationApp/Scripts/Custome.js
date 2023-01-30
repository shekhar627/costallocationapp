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

function DismissOtherDropdown(requestType){
    var section_display = "";
    if(requestType.toLowerCase() != "section"){
        section_display = $('#sectionChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("sectionChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }
    
    if(requestType.toLowerCase() != "dept"){
        section_display = "";
        section_display = $('#departmentChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("departmentChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }
    
    if(requestType.toLowerCase() != "incharge"){
        section_display = "";
        section_display = $('#inchargeChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("inchargeChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }
    
    if(requestType.toLowerCase() != "role"){
        section_display = "";
        section_display = $('#RoleChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("RoleChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }
    
    if(requestType.toLowerCase() != "explanation"){
        section_display = "";
        section_display = $('#ExplanationChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("ExplanationChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }
    
    if(requestType.toLowerCase() != "company"){
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
    DismissOtherDropdown ("section");

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
    DismissOtherDropdown ("dept");
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
    DismissOtherDropdown ("incharge");
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
    DismissOtherDropdown ("role");
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
    DismissOtherDropdown ("explanation");
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
    DismissOtherDropdown ("company");
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
function LoadGradeValue(sel) {
    var _rowId = $("#add_name_row_no_hidden").val();
    var _companyName = $('#company_row_' + _rowId).find(":selected").text();
    console.log("_companyName: " + _companyName);
    var _unitPrice = $("#add_name_unit_price_hidden").val();

    console.log("_unitPrice: " + _unitPrice);

    $.ajax({
        url: `/api/utilities/CompareGrade/${_unitPrice}`,
        type: 'GET',
        dataType: 'json',
        //data: {
        //    unitPrice: _unitPrice
        //},
        success: function (data) {
            $('#grade_edit_hidden').val(data.Id);
            console.log("_companyName.toLowerCase(): " + _companyName.toLowerCase());
            console.log("_companyName.toLowerCase(): " + typeof (_companyName));
            console.log("_rowId: " + _rowId);
            //if (_companyName.toLowerCase == 'mti') {
            if (_companyName.toLowerCase().indexOf("mti") > 0) {
                console.log("test- equal");
                //$('#grade_new_hidden').val(data.Id);
                $('#grade_row_' + _rowId).val(data.SalaryGrade);
            } else {
                console.log("test- not equal");
                $('#grade_row_' + _rowId).val('');
            }
        },
        error: function () {
            $('#grade_row_' + _rowId).val('');
        }
    });
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
                    if (tempCompanyName.toLowerCase() == "mti") {
                        return grade;
                    } else {
                        return "<span style='display:none;'>" + grade+"</span>";
                    }
                }
            },
            {
                data: 'UnitPrice'
            },
            {
                data: 'Id',
                render: function (Id) {
                    if (tempCompanyName.toLowerCase() == 'mti') {
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
    if (companyName.toLowerCase() == 'mti') {
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
    if (e.value <=0) {
        $("#" + e.id).val('');
    } 
}
